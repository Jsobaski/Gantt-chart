"use strict";

import "./../style/visual.less";
import * as d3 from "d3";
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService, formattingSettings as fs } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import DataView = powerbi.DataView;
import DataViewTable = powerbi.DataViewTable;
import IViewport = powerbi.IViewport;

import { VisualFormattingSettingsModel } from "./settings";

interface SeriesDef {
    name: string;
    queryName: string;
    color: string;
}

interface TooltipField {
    name: string;
    value: string;
}

interface BarValue {
    start: Date | null;
    finish: Date | null;
    // Fields from "Series-Specific Tooltip Fields" name-matched to this bar only.
    extra: TooltipField[];
}

interface MilestoneValue {
    date: Date | null;
    // Fields from "Series-Specific Tooltip Fields" name-matched to this milestone only.
    extra: TooltipField[];
}

interface GanttRow {
    location: string;
    projectName: string;
    // Aligned by index with this.milestoneDefs
    milestones: MilestoneValue[];
    // Aligned by index with this.barDefs
    bars: BarValue[];
    // Shown on every bar/milestone tooltip ("Additional Tooltip Fields")
    extraFields: TooltipField[];
    // Index into the raw dataView.table.rows this row came from — some raw rows are
    // skipped (blank location/project, no bar/milestone data), so this can't be inferred
    // from GanttRow array position. Needed to rebuild a table-row-scoped SelectionId
    // for right-click drill-through.
    rawRowIndex: number;
}

interface LocationGroup {
    name: string;
    rows: GanttRow[];
}

type DisplayRow =
    | { type: "location"; name: string }
    | { type: "project"; data: GanttRow };

export class Visual implements IVisual {
    private events: IVisualEventService;
    private host: IVisualHost;
    private container: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private allRows: GanttRow[] = [];
    // Dynamic series definitions (name + auto-assigned color), rebuilt each update()
    // from whichever columns are currently bound to the milestoneDate / barStart+barFinish roles.
    private milestoneDefs: SeriesDef[] = [];
    private barDefs: SeriesDef[] = [];
    private userDateFrom: Date | null = null;
    private userDateTo: Date | null = null;
    // Right-click drill-through support. dataTable is the raw table backing the current
    // dataView (needed to rebuild row-scoped SelectionIds); the two query-name arrays are
    // aligned with barDefs/milestoneDefs and hold the name-matched "Series Drill-through
    // Field" column for that series, if one was bound.
    private selectionManager: ISelectionManager;
    private dataTable: DataViewTable | null = null;
    private barDrillQueryNames: (string | undefined)[] = [];
    private msDrillQueryNames: (string | undefined)[] = [];
    // Live viewer toggle (not a persisted report setting) so anyone looking at the
    // report can flip between fiscal-year and calendar-year framing on the fly.
    private useFiscalYear = false;

    private clearEl(el: HTMLElement): void {
        while (el.firstChild) el.removeChild(el.firstChild);
    }

    private mkHR(): HTMLElement {
        const hr = document.createElement("div");
        hr.style.cssText = "border-top:1px solid #2a5080;margin:5px 0;height:0;";
        return hr;
    }

    // Layout constants
    private readonly LABEL_WIDTH = 295;
    private readonly MIN_ROW_HEIGHT = 34;
    private readonly FY_LANE_HEIGHT = 17;
    private readonly QUARTER_LANE_HEIGHT = 17;
    private readonly MONTH_LANE_HEIGHT = 24;
    private readonly CONTROLS_HEIGHT = 44;
    private readonly BAR_GAP = 2;
    private readonly MILESTONE_RADIUS = 7;
    private readonly LANE_GAP = 5;
    private readonly ROW_V_PADDING = 16;
    // Fiscal year starts Oct 1 (e.g. Oct 2025 - Sep 2026 is FY2026).
    private readonly FISCAL_YEAR_START_MONTH = 9;

    // Bar appearance (defaults overridden by format panel)
    private barHeight = 11;
    private barFontSize = 9;
    private barFontColor = "#000000";

    // Header lane visibility (defaults overridden by format panel)
    private showYearLane = true;
    private showQuarterLane = true;
    private showMonthLane = true;

    // Per-section text styling (defaults overridden by format panel). Anything
    // not covered by these four groups (e.g. tooltip text, the empty-state
    // message) still falls back to the global bgColor/textColor pair above.
    private labelFontFamily = "Arial, sans-serif";
    private labelFontSize = 11.5;
    private labelColor = "#F6F6F6";

    private legendFontFamily = "Arial, sans-serif";
    private legendFontSize = 10.5;
    private legendColor = "#F6F6F6";

    private filterFontFamily = "Arial, sans-serif";
    private filterFontSize = 11;
    private filterColor = "#F6F6F6";

    private headerFontFamily = "Arial, sans-serif";
    private headerFontSize = 10.5;
    private headerColor = "#F6F6F6";

    // Colors — every value below is a real swatch from the ACE Visual Design Guide
    // (no invented tints). Dark Blue (#193661) is the true documented brand dark,
    // used as the canvas and chrome (toolbar/header/label column all match the
    // canvas so there's no mismatched seam); Light Blue/Dark-Blue-tints/Grey are
    // applied as low-alpha accents rather than fabricated new colors.
    private bgColor = "#193661"; // Dark Blue (DB)
    private textColor = "#F6F6F6"; // Light Grey (LG) Lighter 80%
    // locationBgColor is derived each applySettings() call from the user-picked
    // "Location Row Color" hex, rendered at fixed 30% alpha so it stays a soft
    // tinted band rather than a harsh solid block regardless of hue chosen.
    private locationBgColor = "rgba(80,166,211,0.30)";
    private readonly gridColor = "rgba(193,212,239,0.35)"; // Dark Blue (DB) Lighter 80% at 35% alpha
    private todayColor = "#EF5350"; // functional status marker; defaults red, overridable

    // Auto-assigned series colors (before any manual override in "Series Colors")
    // cycle through the brand's documented accent/neutral tones instead of Power
    // BI's generic report-theme palette, so bars/diamonds are on-brand by default.
    private readonly BRAND_SERIES_PALETTE = [
        "#50A6D3", // Light Blue (LB)
        "#84A8DF", // Dark Blue (DB) Lighter 60%
        "#A8AAAC", // Medium Grey (MG)
        "#B9DBED", // Light Blue (LB) Lighter 60%
        "#C1D4EF", // Dark Blue (DB) Lighter 80%
        "#E9E9E9", // Light Grey (LG)
        "#DCEDF6", // Light Blue (LB) Lighter 80%
        "#DCDDDE"  // Medium Grey (MG) Lighter 60%
    ];

    constructor(options: VisualConstructorOptions) {
        this.events = options.host.eventService;
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.formattingSettingsService = new FormattingSettingsService();
        this.container = options.element;
        this.container.style.overflow = "hidden";
        this.container.style.fontFamily = "Arial, sans-serif";
        this.container.style.fontSize = "12px";
    }

    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);
        try {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                options.dataViews?.[0]
            );
            this.applySettings();

            if (!options.dataViews?.[0]?.table) {
                this.clearEl(this.container);
                const msg = document.createElement("div");
                msg.style.cssText = `color:${this.textColor};padding:24px;font-size:13px;background:${this.bgColor};height:100%;`;
                msg.textContent = "Map data fields to this visual: Location, Project Name, milestone dates, and start/finish date pairs.";
                this.container.appendChild(msg);
                this.events.renderingFinished(options);
                return;
            }

            this.allRows = this.parseRows(options.dataViews[0]);
            this.formattingSettings.cards = [this.formattingSettings.ganttConfig, this.buildSeriesColorCard()];
            this.render(options.viewport);
            this.events.renderingFinished(options);
        } catch (err) {
            console.error("Gantt render error:", err);
            this.events.renderingFailed(options, String(err));
        }
    }

    private hexToRgba(hex: string, alpha: number): string {
        const clean = hex.replace("#", "");
        const full = clean.length === 3 ? clean.split("").map(c => c + c).join("") : clean;
        const n = parseInt(full, 16);
        if (isNaN(n)) return `rgba(80,166,211,${alpha})`;
        const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
        return `rgba(${r},${g},${b},${alpha})`;
    }

    private applySettings(): void {
        if (!this.formattingSettings) return;
        const s = this.formattingSettings.ganttConfig;
        this.bgColor = s.bgColor.value.value || this.bgColor;
        this.textColor = s.textColor.value.value || this.textColor;
        this.barHeight = s.barHeight.value ?? this.barHeight;
        this.barFontSize = s.barFontSize.value ?? this.barFontSize;
        this.barFontColor = s.barFontColor.value.value || this.barFontColor;
        this.todayColor = s.todayColor.value.value || this.todayColor;
        this.locationBgColor = this.hexToRgba(s.locationColor.value.value || "#50A6D3", 0.30);
        this.showYearLane = s.showYearLane.value;
        this.showQuarterLane = s.showQuarterLane.value;
        this.showMonthLane = s.showMonthLane.value;

        const lt = this.formattingSettings.labelText;
        this.labelFontFamily = lt.fontFamily.value || this.labelFontFamily;
        this.labelFontSize = lt.fontSize.value ?? this.labelFontSize;
        this.labelColor = lt.color.value.value || this.labelColor;

        const lg = this.formattingSettings.legendText;
        this.legendFontFamily = lg.fontFamily.value || this.legendFontFamily;
        this.legendFontSize = lg.fontSize.value ?? this.legendFontSize;
        this.legendColor = lg.color.value.value || this.legendColor;

        const ft = this.formattingSettings.filterText;
        this.filterFontFamily = ft.fontFamily.value || this.filterFontFamily;
        this.filterFontSize = ft.fontSize.value ?? this.filterFontSize;
        this.filterColor = ft.color.value.value || this.filterColor;

        const ht = this.formattingSettings.headerText;
        this.headerFontFamily = ht.fontFamily.value || this.headerFontFamily;
        this.headerFontSize = ht.fontSize.value ?? this.headerFontSize;
        this.headerColor = ht.color.value.value || this.headerColor;
    }

    // FY starts Oct 1: Oct-Dec of calendar year Y belong to FY(Y+1).
    private fiscalYearOf(d: Date): number {
        return d.getMonth() >= this.FISCAL_YEAR_START_MONTH ? d.getFullYear() + 1 : d.getFullYear();
    }

    // Fiscal quarters: Q1 = Oct-Dec, Q2 = Jan-Mar, Q3 = Apr-Jun, Q4 = Jul-Sep.
    private fiscalQuarterOf(d: Date): number {
        const shifted = (d.getMonth() + (12 - this.FISCAL_YEAR_START_MONTH)) % 12;
        return Math.floor(shifted / 3) + 1;
    }

    private fiscalYearBounds(fy: number): { from: Date; to: Date } {
        return { from: new Date(fy - 1, this.FISCAL_YEAR_START_MONTH, 1), to: new Date(fy, this.FISCAL_YEAR_START_MONTH, 1) };
    }

    private fiscalQuarterBounds(fy: number, q: 1 | 2 | 3 | 4): { from: Date; to: Date } {
        const startMonth = (this.FISCAL_YEAR_START_MONTH + (q - 1) * 3) % 12;
        // Q1 starts in the prior calendar year; Q2-Q4 start in the FY's own calendar year.
        const startYear = q === 1 ? fy - 1 : fy;
        const endYear = startMonth + 3 > 11 ? startYear + 1 : startYear;
        return { from: new Date(startYear, startMonth, 1), to: new Date(endYear, (startMonth + 3) % 12, 1) };
    }

    // Dispatchers used by rendering/quick-jump code so the same call sites work
    // whether the live viewer toggle is set to Fiscal Year or plain Calendar Year.
    private periodYearOf(d: Date): number {
        return this.useFiscalYear ? this.fiscalYearOf(d) : d.getFullYear();
    }

    private periodQuarterOf(d: Date): number {
        return this.useFiscalYear ? this.fiscalQuarterOf(d) : Math.floor(d.getMonth() / 3) + 1;
    }

    private periodYearBounds(y: number): { from: Date; to: Date } {
        return this.useFiscalYear ? this.fiscalYearBounds(y) : { from: new Date(y, 0, 1), to: new Date(y + 1, 0, 1) };
    }

    private periodQuarterBounds(y: number, q: 1 | 2 | 3 | 4): { from: Date; to: Date } {
        if (this.useFiscalYear) return this.fiscalQuarterBounds(y, q);
        const startMonth = (q - 1) * 3;
        return { from: new Date(y, startMonth, 1), to: new Date(y, startMonth + 3, 1) };
    }

    private periodLabel(y: number): string {
        return this.useFiscalYear ? `FY${y}` : `${y}`;
    }

    private thisWeekBounds(): { from: Date; to: Date } {
        const now = new Date();
        const start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
        const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 7);
        return { from: start, to: end };
    }

    // Per-field override (set via the "Series Colors" section of the Format pane)
    // takes priority over the automatically assigned brand color. `seriesIndex` is
    // a running count across ALL bar + milestone series combined, so every bound
    // field gets a distinct brand tone (not just each bar or each milestone alone).
    private resolveColor(col: powerbi.DataViewMetadataColumn, seriesIndex: number): string {
        const override = (col.objects as Record<string, Record<string, { solid?: { color?: string } }>> | undefined)
            ?.dataColors?.fill?.solid?.color;
        return override || this.BRAND_SERIES_PALETTE[seriesIndex % this.BRAND_SERIES_PALETTE.length];
    }

    // Builds one Color Picker slice per currently-bound bar/milestone field, so the
    // Format pane's "Series Colors" section always matches whatever's on the visual.
    // Each slice is bound to that specific column via a measure-scoped selector, so
    // Power BI persists the override on that column (not tied to row/position).
    private buildSeriesColorCard(): fs.SimpleCard {
        const slices: fs.Slice[] = [];
        const addSlice = (def: SeriesDef) => {
            const selector = this.host.createSelectionIdBuilder()
                .withMeasure(def.queryName)
                .createSelectionId()
                .getSelector();
            slices.push(new fs.ColorPicker({
                name: "fill",
                displayName: def.name,
                value: { value: def.color },
                selector
            }));
        };
        this.barDefs.forEach(addSlice);
        this.milestoneDefs.forEach(addSlice);

        const card = new fs.SimpleCard();
        card.name = "dataColors";
        card.displayName = "Series Colors";
        card.slices = slices;
        return card;
    }

    // Right-click a bar/milestone to open Power BI's native context menu, offering
    // "Drill through" to any target page whose drillthrough field matches a value
    // present on this row. Chaining withMeasure() after withTable() narrows the
    // selection to (this row, this specific column) instead of the whole row, so a
    // project with values in multiple systems (e.g. both P6 and Maximo) still only
    // offers the drill-through relevant to the specific bar/milestone clicked —
    // when no drill-through field is bound for that series, this falls back to
    // plain row identity.
    private showSeriesContextMenu(ev: MouseEvent, rawRowIndex: number, fieldQueryName: string | undefined): void {
        if (!this.dataTable) return;
        const builder = this.host.createSelectionIdBuilder().withTable(this.dataTable, rawRowIndex);
        if (fieldQueryName) builder.withMeasure(fieldQueryName);
        const selectionId = builder.createSelectionId();
        this.selectionManager.showContextMenu(selectionId, { x: ev.clientX, y: ev.clientY });
    }

    private parseSingleValue(val: powerbi.PrimitiveValue): Date | null {
        if (val === null || val === undefined || val === "") return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        if (typeof val === "number") {
            if (!Number.isFinite(val)) return null;
            if (val > 1e10) {
                // Unix millisecond timestamp
                const d = new Date(val);
                return isNaN(d.getTime()) ? null : d;
            }
            if (Number.isInteger(val) && val >= 1900 && val <= 2100) {
                // Year-only integer from a date hierarchy — use Jan 1 of that year
                return new Date(val, 0, 1);
            }
            if (val >= 1) {
                // OLE Automation date (days since Dec 30, 1899 UTC)
                const d = new Date(-2209161600000 + val * 86400000);
                return isNaN(d.getTime()) ? null : d;
            }
            return null;
        }
        const d = new Date(String(val));
        return isNaN(d.getTime()) ? null : d;
    }

    // NOTE: milestoneDate / barStart / barFinish are shared roles that can carry
    // several independent user-chosen fields at once (that's what makes them dynamic).
    // Because of that we can no longer tell apart "one field expanded into a Year/
    // Quarter/Month/Day hierarchy" from "several distinct fields sharing a role" —
    // so each bound column here is read as a single plain date value. This requires
    // Power BI's Auto Date/Time option to stay OFF (File > Options > Data Load),
    // exactly as already required to fix the earlier date-hierarchy bugs.
    private parseRows(dataView: DataView): GanttRow[] {
        const table = dataView.table;
        this.dataTable = table || null;
        if (!table?.rows?.length) {
            this.milestoneDefs = [];
            this.barDefs = [];
            this.barDrillQueryNames = [];
            this.msDrillQueryNames = [];
            return [];
        }

        const columns = table.columns || [];

        const colMap: Record<string, number> = {};
        columns.forEach((col, i) => {
            Object.keys(col.roles || {}).forEach(role => {
                if (colMap[role] === undefined) colMap[role] = i;
            });
        });

        const getString = (row: powerbi.DataViewTableRow, role: string): string => {
            const idx = colMap[role];
            if (idx === undefined) return "";
            const v = row[idx];
            return v !== null && v !== undefined ? String(v).trim() : "";
        };

        const milestoneCols: { idx: number; name: string }[] = [];
        const barStartCols: { idx: number; name: string }[] = [];
        const barFinishCols: { idx: number; name: string }[] = [];
        const tooltipCols: { idx: number; name: string }[] = [];
        const seriesTooltipCols: { idx: number; name: string }[] = [];
        const drillCols: { idx: number; name: string }[] = [];
        columns.forEach((col, i) => {
            const name = col.displayName || col.queryName || `Field ${i}`;
            if (col.roles?.["milestoneDate"]) milestoneCols.push({ idx: i, name });
            if (col.roles?.["barStart"]) barStartCols.push({ idx: i, name });
            if (col.roles?.["barFinish"]) barFinishCols.push({ idx: i, name });
            if (col.roles?.["tooltipFields"]) tooltipCols.push({ idx: i, name });
            if (col.roles?.["seriesTooltipFields"]) seriesTooltipCols.push({ idx: i, name });
            if (col.roles?.["seriesDrillthroughField"]) drillCols.push({ idx: i, name });
        });

        // Bar starts/finishes are paired positionally: 1st Start field with 1st Finish
        // field, 2nd with 2nd, etc. Add them to the field wells in matching order.
        const barCount = Math.min(barStartCols.length, barFinishCols.length);

        // Rename a field via right-click > "Rename for this visual" in Power BI to
        // control exactly what shows up in the legend and tooltips. A per-field color
        // set in the Format pane ("Series Colors") persists on that column's `objects`
        // and takes priority over the automatically assigned brand color. Bars are
        // indexed first, milestones continue the count, so every series (bar or
        // milestone) gets a distinct brand tone.
        this.barDefs = barStartCols.slice(0, barCount).map((c, i) => ({
            name: c.name,
            queryName: columns[c.idx].queryName || c.name,
            color: this.resolveColor(columns[c.idx], i)
        }));
        this.milestoneDefs = milestoneCols.map((c, i) => ({
            name: c.name,
            queryName: columns[c.idx].queryName || c.name,
            color: this.resolveColor(columns[c.idx], barCount + i)
        }));

        const formatCellValue = (v: powerbi.PrimitiveValue): string | null => {
            if (v === null || v === undefined || v === "") return null;
            if (v instanceof Date) return this.fmtFull(v);
            if (typeof v === "number") return v.toLocaleString(undefined, { maximumFractionDigits: 2 });
            return String(v).trim() || null;
        };

        const getExtraFields = (row: powerbi.DataViewTableRow): TooltipField[] => {
            const out: TooltipField[] = [];
            tooltipCols.forEach(({ idx, name }) => {
                const text = formatCellValue(row[idx]);
                if (text) out.push({ name, value: text });
            });
            return out;
        };

        // A field is attached to exactly one bar or milestone series by name: rename it
        // (right-click > Rename for this visual) so it STARTS WITH that series' exact
        // name. If several series names match, the longest (most specific) match wins;
        // if none match, the field is dropped rather than applied to every series. Used
        // for both "Series-Specific Tooltip Fields" and the "Series Drill-through Field".
        const matchSeries = (name: string): { type: "bar" | "milestone"; index: number } | null => {
            const lname = name.toLowerCase();
            let bestType: "bar" | "milestone" | null = null;
            let bestIndex = -1;
            let bestLen = -1;
            this.barDefs.forEach((def, i) => {
                if (def.name && lname.startsWith(def.name.toLowerCase()) && def.name.length > bestLen) {
                    bestType = "bar"; bestIndex = i; bestLen = def.name.length;
                }
            });
            this.milestoneDefs.forEach((def, i) => {
                if (def.name && lname.startsWith(def.name.toLowerCase()) && def.name.length > bestLen) {
                    bestType = "milestone"; bestIndex = i; bestLen = def.name.length;
                }
            });
            return bestType ? { type: bestType, index: bestIndex } : null;
        };

        const barSeriesTooltipCols: { idx: number; name: string }[][] = this.barDefs.map(() => []);
        const msSeriesTooltipCols: { idx: number; name: string }[][] = this.milestoneDefs.map(() => []);
        seriesTooltipCols.forEach(col => {
            const match = matchSeries(col.name);
            if (match?.type === "bar") barSeriesTooltipCols[match.index].push(col);
            else if (match?.type === "milestone") msSeriesTooltipCols[match.index].push(col);
        });

        // One drill-through field per series (last match wins if more than one is bound
        // to the same series name — an edge case the user controls via naming).
        const barDrillColIdx: (number | undefined)[] = this.barDefs.map(() => undefined);
        const msDrillColIdx: (number | undefined)[] = this.milestoneDefs.map(() => undefined);
        drillCols.forEach(col => {
            const match = matchSeries(col.name);
            if (match?.type === "bar") barDrillColIdx[match.index] = col.idx;
            else if (match?.type === "milestone") msDrillColIdx[match.index] = col.idx;
        });
        this.barDrillQueryNames = barDrillColIdx.map(idx => idx !== undefined ? (columns[idx].queryName || columns[idx].displayName) : undefined);
        this.msDrillQueryNames = msDrillColIdx.map(idx => idx !== undefined ? (columns[idx].queryName || columns[idx].displayName) : undefined);

        const getSeriesExtra = (row: powerbi.DataViewTableRow, cols: { idx: number; name: string }[]): TooltipField[] => {
            const out: TooltipField[] = [];
            cols.forEach(({ idx, name }) => {
                const text = formatCellValue(row[idx]);
                if (text) out.push({ name, value: text });
            });
            return out;
        };

        const rows: GanttRow[] = [];
        table.rows.forEach((row, rawRowIndex) => {
            const location = getString(row, "location");
            const projectName = getString(row, "projectName");
            if (!location && !projectName) return;

            const milestones: MilestoneValue[] = milestoneCols.map((c, i) => ({
                date: this.parseSingleValue(row[c.idx]),
                extra: getSeriesExtra(row, msSeriesTooltipCols[i])
            }));
            const bars: BarValue[] = [];
            for (let i = 0; i < barCount; i++) {
                bars.push({
                    start: this.parseSingleValue(row[barStartCols[i].idx]),
                    finish: this.parseSingleValue(row[barFinishCols[i].idx]),
                    extra: getSeriesExtra(row, barSeriesTooltipCols[i])
                });
            }

            const hasBar = bars.some(b => b.start && b.finish);
            const hasMilestone = milestones.some(m => m.date !== null);
            if (!hasBar && !hasMilestone) return;

            rows.push({ location, projectName, milestones, bars, extraFields: getExtraFields(row), rawRowIndex });
        });

        return rows;
    }

    private groupByLocation(rows: GanttRow[]): LocationGroup[] {
        const map = new Map<string, GanttRow[]>();
        rows.forEach(r => {
            const key = r.location || "(No Location)";
            if (!map.has(key)) map.set(key, []);
            map.get(key).push(r);
        });
        const groups: LocationGroup[] = [];
        map.forEach((grpRows, name) => {
            grpRows.sort((a, b) => (a.projectName || "").localeCompare(b.projectName || ""));
            groups.push({ name, rows: grpRows });
        });
        groups.sort((a, b) => a.name.localeCompare(b.name));
        return groups;
    }

    // Fallback used only when the current fiscal/calendar year has no data at all
    // (see defaultDateRange) — spans the actual min-to-max extent of the data,
    // so a portfolio with no "current year" work still opens to something useful
    // instead of a blank chart.
    private dataExtentRange(rows: GanttRow[]): { from: Date; to: Date } {
        const dates: Date[] = [];
        rows.forEach(r => {
            r.milestones.forEach(m => { if (m.date) dates.push(m.date); });
            r.bars.forEach(b => {
                if (b.start) dates.push(b.start);
                if (b.finish) dates.push(b.finish);
            });
        });
        if (!dates.length) {
            const now = new Date();
            return {
                from: new Date(now.getFullYear(), now.getMonth() - 3, 1),
                to: new Date(now.getFullYear(), now.getMonth() + 9, 1)
            };
        }
        const minMs = Math.min(...dates.map(d => d.getTime()));
        const maxMs = Math.max(...dates.map(d => d.getTime()));
        const minD = new Date(minMs);
        const maxD = new Date(maxMs);
        return {
            from: new Date(minD.getFullYear(), minD.getMonth() - 1, 1),
            to: new Date(maxD.getFullYear(), maxD.getMonth() + 2, 1)
        };
    }

    // Default view on first load / after Reset: the current fiscal or calendar
    // year (matching the live FY/CY toggle), not the full min-to-max data span —
    // with multi-year messy data (some projects years old, some a decade out),
    // showing everything would be unreadable. Falls back to the actual data
    // extent only if the current period has no data at all, so it's never blank.
    private defaultDateRange(): { from: Date; to: Date } {
        const range = this.periodYearBounds(this.periodYearOf(new Date()));
        const hasDataInRange = this.allRows.some(r =>
            r.milestones.some(m => m.date && m.date >= range.from && m.date < range.to) ||
            r.bars.some(b =>
                (b.start && b.start >= range.from && b.start < range.to) ||
                (b.finish && b.finish >= range.from && b.finish < range.to))
        );
        return hasDataInRange ? range : this.dataExtentRange(this.allRows);
    }

    private fmtShort(d: Date): string {
        return d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
    }

    private fmtFull(d: Date): string {
        return d.toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" });
    }

    // Total vertical space a project row needs given the current dynamic series counts:
    // a lane for milestone diamonds stacked above a lane of stacked bars, centered in the row.
    private computeRowMetrics(): { rowHeight: number; contentH: number; milestoneLaneH: number; laneGap: number; barsLaneH: number } {
        const barCount = this.barDefs.length;
        const msCount = this.milestoneDefs.length;
        const milestoneLaneH = msCount > 0 ? this.MILESTONE_RADIUS * 2 : 0;
        const barsLaneH = barCount > 0 ? barCount * this.barHeight + Math.max(0, barCount - 1) * this.BAR_GAP : 0;
        const laneGap = (milestoneLaneH > 0 && barsLaneH > 0) ? this.LANE_GAP : 0;
        const contentH = milestoneLaneH + laneGap + barsLaneH;
        const rowHeight = Math.max(this.MIN_ROW_HEIGHT, contentH + this.ROW_V_PADDING);
        return { rowHeight, contentH, milestoneLaneH, laneGap, barsLaneH };
    }

    private measureCtx: CanvasRenderingContext2D | null = null;

    // Greedy word-wrap line count, mirroring the browser's own line-breaking so the
    // row height we reserve matches what "white-space:normal" will actually render.
    private wrapLineCount(text: string, maxWidth: number, fontSize: number, fontFamily: string): number {
        if (!text) return 1;
        if (!this.measureCtx) {
            this.measureCtx = document.createElement("canvas").getContext("2d");
        }
        const ctx = this.measureCtx;
        if (!ctx) return 1;
        ctx.font = `${fontSize}px ${fontFamily}`;
        const words = text.split(/\s+/).filter(Boolean);
        if (words.length === 0) return 1;
        let lines = 1;
        let lineText = words[0];
        for (let i = 1; i < words.length; i++) {
            const candidate = `${lineText} ${words[i]}`;
            if (ctx.measureText(candidate).width > maxWidth) {
                lines++;
                lineText = words[i];
            } else {
                lineText = candidate;
            }
        }
        return lines;
    }

    // Header lanes (Year/Quarter/Month) can each be hidden via the Format pane, so
    // the header's total height and each lane's vertical offset are computed fresh
    // per render instead of being a fixed constant.
    private computeHeaderMetrics(): { headerHeight: number; yearH: number; quarterH: number; monthH: number; quarterY: number; monthY: number } {
        const yearH = this.showYearLane ? this.FY_LANE_HEIGHT : 0;
        const quarterH = this.showQuarterLane ? this.QUARTER_LANE_HEIGHT : 0;
        const monthH = this.showMonthLane ? this.MONTH_LANE_HEIGHT : 0;
        return { headerHeight: yearH + quarterH + monthH, yearH, quarterH, monthH, quarterY: yearH, monthY: yearH + quarterH };
    }

    private render(viewport: IViewport): void {
        this.clearEl(this.container);

        const W = viewport.width;
        const H = viewport.height;

        const groups = this.groupByLocation(this.allRows);
        const displayRows: DisplayRow[] = [];
        groups.forEach(g => {
            displayRows.push({ type: "location", name: g.name });
            g.rows.forEach(r => displayRows.push({ type: "project", data: r }));
        });

        const { from: defaultFrom, to: defaultTo } = this.defaultDateRange();
        const dateFrom = this.userDateFrom || defaultFrom;
        const dateTo = this.userDateTo || defaultTo;

        const { rowHeight, contentH, milestoneLaneH, laneGap } = this.computeRowMetrics();
        const { headerHeight, quarterY: quarterLaneY, monthY: monthLaneY } = this.computeHeaderMetrics();

        const timelineWidth = Math.max(W - this.LABEL_WIDTH, 500);

        // Long location/project names wrap onto extra lines instead of being cut off,
        // so each row's height is its own max of the bars/milestone lane height and
        // whatever the wrapped label needs — most rows stay at the uniform base height.
        const PROJECT_LABEL_WIDTH = this.LABEL_WIDTH - 43; // 22px+6px padding, 15px marker+gap
        const LOCATION_LABEL_WIDTH = this.LABEL_WIDTH - 50; // 20px padding, "▶  " prefix, letter-spacing
        const labelLineHeight = Math.ceil(this.labelFontSize * 1.3);
        const rowHeights: number[] = displayRows.map(row => {
            const [text, maxWidth] = row.type === "location"
                ? [`▶  ${row.name}`, LOCATION_LABEL_WIDTH]
                : [row.data.projectName || row.data.location, PROJECT_LABEL_WIDTH];
            const lines = this.wrapLineCount(text, maxWidth, this.labelFontSize, this.labelFontFamily);
            const neededLabelHeight = lines * labelLineHeight + this.ROW_V_PADDING;
            return Math.max(rowHeight, neededLabelHeight);
        });
        const rowY: number[] = [];
        {
            let acc = 0;
            rowHeights.forEach(h => { rowY.push(acc); acc += h; });
        }
        const totalBodyHeight = rowHeights.reduce((sum, h) => sum + h, 0);

        // ── Outer wrapper ────────────────────────────────────────────────
        const outer = document.createElement("div");
        outer.className = "gantt-outer";
        outer.style.cssText = [
            `width:${W}px`, `height:${H}px`,
            "display:flex", "flex-direction:column", "overflow:hidden",
            `background:${this.bgColor}`, `color:${this.textColor}`,
            "position:relative"
        ].join(";");
        this.container.appendChild(outer);

        // ── Tooltip ──────────────────────────────────────────────────────
        const tooltip = document.createElement("div");
        tooltip.style.cssText = [
            "position:absolute", "display:none",
            "background:rgba(8,18,32,0.97)",
            `color:${this.textColor}`,
            "border:1px solid #2a5080",
            "border-radius:6px", "padding:10px 14px",
            "font-size:11.5px", "pointer-events:none", "z-index:100",
            "max-width:300px", "line-height:1.7",
            "box-shadow:0 4px 24px rgba(0,0,0,0.7)"
        ].join(";");
        outer.appendChild(tooltip);

        const showTip = (nodes: Node[], ev: MouseEvent) => {
            this.clearEl(tooltip);
            nodes.forEach(n => tooltip.appendChild(n));
            tooltip.style.display = "block";
            moveTip(ev);
        };
        const moveTip = (ev: MouseEvent) => {
            const rect = outer.getBoundingClientRect();
            let x = ev.clientX - rect.left + 14;
            let y = ev.clientY - rect.top - 10;
            if (x + 300 > W) x = ev.clientX - rect.left - 314;
            if (y + 180 > H) y = ev.clientY - rect.top - 190;
            tooltip.style.left = `${x}px`;
            tooltip.style.top = `${y}px`;
        };
        const hideTip = () => { tooltip.style.display = "none"; };

        // ── Controls bar ─────────────────────────────────────────────────
        const controls = document.createElement("div");
        controls.style.cssText = [
            `min-height:${this.CONTROLS_HEIGHT}px`, "flex-shrink:0",
            "display:flex", "align-items:center", "gap:10px",
            "padding:6px 10px", `background:${this.bgColor}`,
            `border-bottom:1px solid ${this.gridColor}`,
            "flex-wrap:wrap"
        ].join(";");

        // Legend chips — built dynamically from whatever bar/milestone fields are bound.
        const legend = document.createElement("div");
        legend.style.cssText = "display:flex;align-items:center;gap:14px;flex:1;flex-wrap:wrap;";
        const legendItems: Array<{ color: string; label: string; isBar: boolean }> = [
            ...this.barDefs.map(b => ({ color: b.color, label: b.name, isBar: true })),
            ...this.milestoneDefs.map(m => ({ color: m.color, label: m.name, isBar: false }))
        ];
        legendItems.forEach(item => {
            const chip = document.createElement("span");
            chip.style.cssText = `display:flex;align-items:center;gap:4px;font-size:${this.legendFontSize}px;color:${this.legendColor};font-family:${this.legendFontFamily};`;
            const swatch = document.createElement("span");
            if (item.isBar) {
                swatch.style.cssText = `display:inline-block;width:18px;height:9px;background:${item.color};border-radius:2px;flex-shrink:0;`;
            } else {
                swatch.style.cssText = `display:inline-block;width:10px;height:10px;transform:rotate(45deg);background:${item.color};flex-shrink:0;`;
            }
            chip.appendChild(swatch);
            chip.appendChild(document.createTextNode(item.label));
            legend.appendChild(chip);
        });
        controls.appendChild(legend);

        const inputStyle = [
            `background:#112840`, `color:${this.filterColor}`,
            "border:1px solid #2a5080", "border-radius:4px",
            "padding:2px 6px", `font-size:${this.filterFontSize}px`, "outline:none",
            `font-family:${this.filterFontFamily}`
        ].join(";");

        const mkLabel = (txt: string, inputEl: HTMLInputElement) => {
            const lbl = document.createElement("label");
            lbl.style.cssText = `display:flex;align-items:center;gap:5px;font-size:${this.filterFontSize}px;white-space:nowrap;color:${this.filterColor};font-family:${this.filterFontFamily};`;
            lbl.textContent = txt;
            lbl.appendChild(inputEl);
            return lbl;
        };

        const fmtForInput = (d: Date) => d.toISOString().split("T")[0];

        const fromInput = document.createElement("input");
        fromInput.type = "date";
        fromInput.value = fmtForInput(dateFrom);
        fromInput.style.cssText = inputStyle;

        const toInput = document.createElement("input");
        toInput.type = "date";
        toInput.value = fmtForInput(dateTo);
        toInput.style.cssText = inputStyle;

        const btnStyle = [
            "background:#112840", `color:${this.filterColor}`,
            "border:1px solid #2a5080", "border-radius:4px",
            "padding:3px 10px", `font-size:${this.filterFontSize}px`, "cursor:pointer",
            `font-family:${this.filterFontFamily}`
        ].join(";");

        const resetBtn = document.createElement("button");
        resetBtn.textContent = "Reset";
        resetBtn.style.cssText = btnStyle;

        // ── Quick view: jump straight to a fiscal/calendar year or quarter, or the
        // current week, instead of picking exact dates.
        const quickWrap = document.createElement("div");
        quickWrap.style.cssText = "display:flex;align-items:center;gap:6px;flex-wrap:wrap;";

        const yearModeBtn = document.createElement("button");
        yearModeBtn.textContent = this.useFiscalYear ? "View: Fiscal Year" : "View: Calendar Year";
        yearModeBtn.style.cssText = btnStyle;
        yearModeBtn.addEventListener("click", () => {
            this.useFiscalYear = !this.useFiscalYear;
            this.render(viewport);
        });

        const fyInput = document.createElement("input");
        fyInput.type = "number";
        fyInput.style.cssText = inputStyle + "width:64px;";
        // dateTo is always the exclusive first-instant of the NEXT period, so deriving the
        // displayed year from it would show one year ahead of what's actually on screen.
        // dateFrom is the inclusive start of the current period, so it always resolves correctly.
        fyInput.value = String(this.periodYearOf(this.userDateFrom || dateFrom));

        const quarterSelect = document.createElement("select");
        quarterSelect.style.cssText = inputStyle;
        [["all", "Full Year"], ["1", "Q1"], ["2", "Q2"], ["3", "Q3"], ["4", "Q4"]].forEach(([val, label]) => {
            const opt = document.createElement("option");
            opt.value = val;
            opt.textContent = label;
            quarterSelect.appendChild(opt);
        });

        const goFyBtn = document.createElement("button");
        goFyBtn.textContent = "Go";
        goFyBtn.style.cssText = btnStyle;
        goFyBtn.addEventListener("click", () => {
            const y = parseInt(fyInput.value, 10);
            if (!y || y < 1900 || y > 2200) return;
            const q = quarterSelect.value;
            const range = q === "all"
                ? this.periodYearBounds(y)
                : this.periodQuarterBounds(y, Number(q) as 1 | 2 | 3 | 4);
            this.userDateFrom = range.from;
            this.userDateTo = range.to;
            this.render(viewport);
        });

        const weekBtn = document.createElement("button");
        weekBtn.textContent = "This Week";
        weekBtn.style.cssText = btnStyle;
        weekBtn.addEventListener("click", () => {
            const range = this.thisWeekBounds();
            this.userDateFrom = range.from;
            this.userDateTo = range.to;
            this.render(viewport);
        });

        quickWrap.appendChild(yearModeBtn);
        quickWrap.appendChild(mkLabel(this.useFiscalYear ? "FY:" : "Year:", fyInput));
        quickWrap.appendChild(quarterSelect);
        quickWrap.appendChild(goFyBtn);
        quickWrap.appendChild(weekBtn);
        controls.appendChild(quickWrap);

        controls.appendChild(mkLabel("From:", fromInput));
        controls.appendChild(mkLabel("To:", toInput));
        controls.appendChild(resetBtn);
        outer.appendChild(controls);

        // Native date inputs fire "change" on every keystroke once all segments
        // have a value (not just on blur/Enter), so typing a new year digit-by-digit
        // briefly produces incomplete years (e.g. "0002"). Debounce the commit so
        // only the settled value after a typing pause triggers a re-render, and
        // reject out-of-range years so a transient partial value is ignored.
        let dateDebounceTimer: number | undefined;
        const isPlausibleYear = (d: Date) => d.getFullYear() >= 1900 && d.getFullYear() <= 2200;
        const onDateChange = () => {
            if (dateDebounceTimer !== undefined) {
                window.clearTimeout(dateDebounceTimer);
            }
            dateDebounceTimer = window.setTimeout(() => {
                const f = fromInput.valueAsDate;
                const t = toInput.valueAsDate;
                if (f && isPlausibleYear(f)) this.userDateFrom = f;
                if (t && isPlausibleYear(t)) this.userDateTo = t;
                this.render(viewport);
            }, 600);
        };
        fromInput.addEventListener("change", onDateChange);
        toInput.addEventListener("change", onDateChange);
        resetBtn.addEventListener("click", () => {
            this.userDateFrom = null;
            this.userDateTo = null;
            this.render(viewport);
        });

        // ── Header row (sticky month labels) ────────────────────────────
        const headerRow = document.createElement("div");
        headerRow.style.cssText = [
            "display:flex", "flex-shrink:0",
            `height:${headerHeight}px`,
            `background:${this.bgColor}`,
            `border-bottom:2px solid ${this.gridColor}`,
            "z-index:10", "overflow:hidden"
        ].join(";");

        const labelHeader = document.createElement("div");
        labelHeader.style.cssText = [
            `width:${this.LABEL_WIDTH}px`, "flex-shrink:0",
            "display:flex", "align-items:center",
            "padding:0 12px", "font-weight:600", "font-size:12px",
            `border-right:2px solid ${this.gridColor}`,
            "letter-spacing:0.03em"
        ].join(";");
        labelHeader.textContent = "Location / Project";
        headerRow.appendChild(labelHeader);

        const monthWrapper = document.createElement("div");
        monthWrapper.style.cssText = "flex:1;overflow:hidden;position:relative;";
        const monthSvgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        monthSvgEl.style.cssText = `width:${timelineWidth}px;height:${headerHeight}px;display:block;`;
        monthWrapper.appendChild(monthSvgEl);
        headerRow.appendChild(monthWrapper);
        outer.appendChild(headerRow);

        // ── Scrollable body ──────────────────────────────────────────────
        const bodyScroll = document.createElement("div");
        bodyScroll.style.cssText = "flex:1;overflow-y:auto;overflow-x:auto;display:flex;";
        outer.appendChild(bodyScroll);

        // Sync month header with horizontal scroll
        bodyScroll.addEventListener("scroll", () => {
            monthSvgEl.style.transform = `translateX(-${bodyScroll.scrollLeft}px)`;
        });

        // ── Label column (sticky left) ───────────────────────────────────
        const labelCol = document.createElement("div");
        labelCol.style.cssText = [
            `width:${this.LABEL_WIDTH}px`, "flex-shrink:0",
            "position:sticky", "left:0", "z-index:5",
            `background:${this.bgColor}`,
            `border-right:2px solid ${this.gridColor}`
        ].join(";");
        bodyScroll.appendChild(labelCol);

        // ── Gantt SVG area ───────────────────────────────────────────────
        const ganttWrapper = document.createElement("div");
        ganttWrapper.style.cssText = "flex:1;position:relative;flex-shrink:0;";
        const ganttSvgEl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        ganttSvgEl.style.cssText = `width:${timelineWidth}px;height:${totalBodyHeight}px;display:block;`;
        ganttWrapper.appendChild(ganttSvgEl);
        bodyScroll.appendChild(ganttWrapper);

        // ── D3 selections ────────────────────────────────────────────────
        const svg = d3.select(ganttSvgEl);
        const monthSvg = d3.select(monthSvgEl);

        // ── Time scale ───────────────────────────────────────────────────
        const xScale = d3.scaleTime()
            .domain([dateFrom, dateTo])
            .range([0, timelineWidth]);

        // ── Month grid lines and alternating bands ────────────────────────
        const months = d3.timeMonth.range(
            d3.timeMonth.floor(dateFrom),
            d3.timeMonth.ceil(dateTo)
        );

        months.forEach((m, i) => {
            const x1 = xScale(m);
            const x2 = i + 1 < months.length ? xScale(months[i + 1]) : timelineWidth;
            svg.append("rect")
                .attr("x", x1).attr("y", 0)
                .attr("width", x2 - x1).attr("height", totalBodyHeight)
                .attr("fill", i % 2 === 0 ? "rgba(255,255,255,0.018)" : "rgba(0,0,0,0)");
            svg.append("line")
                .attr("x1", x1).attr("y1", 0)
                .attr("x2", x1).attr("y2", totalBodyHeight)
                .attr("stroke", this.gridColor)
                .attr("stroke-width", 0.5)
                .attr("stroke-opacity", 0.7);
        });

        // Today marker
        const today = new Date();
        if (today >= dateFrom && today <= dateTo) {
            const tx = xScale(today);
            svg.append("line")
                .attr("x1", tx).attr("y1", 0)
                .attr("x2", tx).attr("y2", totalBodyHeight)
                .attr("stroke", this.todayColor)
                .attr("stroke-width", 1.5)
                .attr("stroke-dasharray", "4,3")
                .attr("stroke-opacity", 0.85);
        }

        // ── Header labels: Year / Quarter / Month lanes (each independently
        // toggleable, and Year/Quarter switch between fiscal and calendar framing) ──
        monthSvg.append("rect")
            .attr("x", 0).attr("y", 0)
            .attr("width", timelineWidth).attr("height", headerHeight)
            .attr("fill", this.bgColor);

        interface Span { label: string; x1: number; x2: number; }
        const yearSpans: Span[] = [];
        const qSpans: Span[] = [];
        months.forEach((m, i) => {
            const x1 = xScale(m);
            const x2 = i + 1 < months.length ? xScale(months[i + 1]) : timelineWidth;
            const yearLabel = this.periodLabel(this.periodYearOf(m));
            const qLabel = `Q${this.periodQuarterOf(m)} ${yearLabel}`;

            const lastYear = yearSpans[yearSpans.length - 1];
            if (lastYear && lastYear.label === yearLabel) lastYear.x2 = x2;
            else yearSpans.push({ label: yearLabel, x1, x2 });

            const lastQ = qSpans[qSpans.length - 1];
            if (lastQ && lastQ.label === qLabel) lastQ.x2 = x2;
            else qSpans.push({ label: qLabel, x1, x2 });
        });

        const drawSpanLane = (spans: Span[], laneY: number, laneH: number, fontWeight: string, opacity: number) => {
            spans.forEach(sp => {
                monthSvg.append("line")
                    .attr("x1", sp.x1).attr("y1", laneY)
                    .attr("x2", sp.x1).attr("y2", laneY + laneH)
                    .attr("stroke", this.gridColor).attr("stroke-width", 0.6);
                const w = sp.x2 - sp.x1;
                if (w > 26) {
                    monthSvg.append("text")
                        .attr("x", sp.x1 + w / 2)
                        .attr("y", laneY + laneH / 2 + 4)
                        .attr("text-anchor", "middle")
                        .attr("fill", this.headerColor)
                        .attr("font-size", `${this.headerFontSize}px`)
                        .attr("font-family", this.headerFontFamily)
                        .attr("font-weight", fontWeight)
                        .attr("opacity", opacity)
                        .text(sp.label);
                }
            });
            monthSvg.append("line")
                .attr("x1", 0).attr("y1", laneY + laneH)
                .attr("x2", timelineWidth).attr("y2", laneY + laneH)
                .attr("stroke", this.gridColor).attr("stroke-width", 0.6);
        };

        if (this.showYearLane) drawSpanLane(yearSpans, 0, this.FY_LANE_HEIGHT, "700", 1);
        if (this.showQuarterLane) drawSpanLane(qSpans, quarterLaneY, this.QUARTER_LANE_HEIGHT, "600", 0.85);

        if (this.showMonthLane) months.forEach(m => {
            const x = xScale(m);
            monthSvg.append("line")
                .attr("x1", x).attr("y1", monthLaneY)
                .attr("x2", x).attr("y2", headerHeight)
                .attr("stroke", this.gridColor).attr("stroke-width", 0.5);
            monthSvg.append("text")
                .attr("x", x + 5)
                .attr("y", monthLaneY + this.MONTH_LANE_HEIGHT / 2 + 4)
                .attr("fill", this.headerColor)
                .attr("font-size", `${this.headerFontSize}px`)
                .attr("font-family", this.headerFontFamily)
                .attr("font-weight", "600")
                .text(d3.timeFormat("%b %Y")(m));
        });

        // ── Row rendering ─────────────────────────────────────────────────
        displayRows.forEach((row, idx) => {
            const y = rowY[idx];
            const thisRowHeight = rowHeights[idx];
            const isEvenRow = idx % 2 === 0;

            if (row.type === "location") {
                // ── Location header ─────────────────────────────────────
                const div = document.createElement("div");
                div.style.cssText = [
                    `height:${thisRowHeight}px`,
                    "display:flex", "align-items:center",
                    "padding:0 10px", "font-weight:700", `font-size:${this.labelFontSize}px`,
                    "letter-spacing:0.06em", "text-transform:uppercase",
                    `background:${this.locationBgColor}`,
                    `border-bottom:1px solid ${this.gridColor}`,
                    `color:${this.labelColor}`, `font-family:${this.labelFontFamily}`, "cursor:default",
                    "white-space:normal", "overflow-wrap:break-word", "line-height:1.3"
                ].join(";");
                div.textContent = `▶  ${row.name}`;
                labelCol.appendChild(div);

                svg.append("rect")
                    .attr("x", 0).attr("y", y)
                    .attr("width", timelineWidth).attr("height", thisRowHeight)
                    .attr("fill", this.locationBgColor);
                svg.append("line")
                    .attr("x1", 0).attr("y1", y + thisRowHeight)
                    .attr("x2", timelineWidth).attr("y2", y + thisRowHeight)
                    .attr("stroke", this.gridColor).attr("stroke-width", 0.6);

            } else {
                // ── Project row ──────────────────────────────────────────
                const d = row.data;
                const rowBg = isEvenRow ? "rgba(255,255,255,0.025)" : "rgba(0,0,0,0)";

                // Label div
                const div = document.createElement("div");
                div.style.cssText = [
                    `height:${thisRowHeight}px`,
                    "display:flex", "align-items:center",
                    "padding:0 6px 0 22px",
                    `font-size:${this.labelFontSize}px`, `font-family:${this.labelFontFamily}`,
                    `color:${this.labelColor}`,
                    `border-bottom:1px solid ${this.gridColor}`,
                    "cursor:default", `background:${rowBg}`
                ].join(";");
                div.title = `${d.location} › ${d.projectName}`;

                const marker = document.createElement("span");
                marker.style.cssText = `color:${this.labelColor};opacity:0.45;margin-right:5px;font-size:10px;flex-shrink:0;align-self:flex-start;margin-top:2px;`;
                marker.textContent = "□";
                div.appendChild(marker);

                const nameSpan = document.createElement("span");
                nameSpan.style.cssText = "white-space:normal;overflow-wrap:break-word;word-break:break-word;line-height:1.3;";
                nameSpan.textContent = d.projectName || d.location;
                div.appendChild(nameSpan);
                labelCol.appendChild(div);

                // Row background stripe on SVG
                svg.append("rect")
                    .attr("x", 0).attr("y", y)
                    .attr("width", timelineWidth).attr("height", thisRowHeight)
                    .attr("fill", rowBg);
                svg.append("line")
                    .attr("x1", 0).attr("y1", y + thisRowHeight)
                    .attr("x2", timelineWidth).attr("y2", y + thisRowHeight)
                    .attr("stroke", this.gridColor).attr("stroke-width", 0.3);

                // Milestone lane centered above the bars lane, whole block vertically centered in the row
                // (using thisRowHeight, not the uniform base, so wrapped-label rows still center the bars)
                const contentTop = y + (thisRowHeight - contentH) / 2;
                const milestoneCenterY = contentTop + this.MILESTONE_RADIUS;
                const barsTop = contentTop + milestoneLaneH + laneGap;

                // ── Bars (stacked, one lane per dynamic bar series) ──────
                d.bars.forEach((bv, i) => {
                    if (!bv.start || !bv.finish || bv.start > bv.finish) return;
                    const def = this.barDefs[i];
                    if (!def) return;
                    const barY = barsTop + i * (this.barHeight + this.BAR_GAP);
                    const x1 = xScale(bv.start);
                    const x2 = xScale(bv.finish);
                    const clampX1 = Math.max(0, x1);
                    const clampX2 = Math.min(timelineWidth, x2);
                    if (clampX2 <= clampX1) return;
                    const bw = clampX2 - clampX1;

                    const bar = svg.append("rect")
                        .attr("x", clampX1).attr("y", barY)
                        .attr("width", bw).attr("height", this.barHeight)
                        .attr("fill", def.color)
                        .attr("rx", 3).attr("ry", 3)
                        .style("cursor", "pointer");

                    if (bw > 90) {
                        svg.append("text")
                            .attr("x", clampX1 + bw / 2)
                            .attr("y", barY + this.barHeight / 2 + this.barFontSize / 2 - 1)
                            .attr("text-anchor", "middle")
                            .attr("font-size", `${this.barFontSize}px`)
                            .attr("fill", this.barFontColor)
                            .attr("font-family", "Arial, sans-serif")
                            .attr("pointer-events", "none")
                            .attr("font-weight", "600")
                            .text(`${def.name}: ${this.fmtShort(bv.start)} - ${this.fmtShort(bv.finish)}`);
                    }

                    bar.on("mouseenter", (ev: MouseEvent) => {
                        bar.attr("fill-opacity", 0.85);
                        showTip(this.buildBarTip(d, i), ev);
                    })
                    .on("mousemove", (ev: MouseEvent) => moveTip(ev))
                    .on("mouseleave", () => { bar.attr("fill-opacity", 1); hideTip(); })
                    .on("contextmenu", (ev: MouseEvent) => {
                        ev.preventDefault();
                        this.showSeriesContextMenu(ev, d.rawRowIndex, this.barDrillQueryNames[i]);
                    });
                });

                // ── Milestone diamonds (one lane, N diamonds across the timeline) ──
                d.milestones.forEach((ms, i) => {
                    const def = this.milestoneDefs[i];
                    if (!ms.date || !def || ms.date < dateFrom || ms.date > dateTo) return;
                    const mx = xScale(ms.date);
                    const r = this.MILESTONE_RADIUS;
                    const pts = `${mx},${milestoneCenterY - r} ${mx + r},${milestoneCenterY} ${mx},${milestoneCenterY + r} ${mx - r},${milestoneCenterY}`;

                    const diamond = svg.append("polygon")
                        .attr("points", pts)
                        .attr("fill", def.color)
                        .attr("stroke", "rgba(255,255,255,0.4)")
                        .attr("stroke-width", 0.8)
                        .style("cursor", "pointer");

                    diamond.on("mouseenter", (ev: MouseEvent) => {
                        diamond.attr("stroke-width", 2).attr("stroke", "white");
                        showTip(this.buildMilestoneTip(d, i), ev);
                    })
                    .on("mousemove", (ev: MouseEvent) => moveTip(ev))
                    .on("mouseleave", () => {
                        diamond.attr("stroke-width", 0.8).attr("stroke", "rgba(255,255,255,0.4)");
                        hideTip();
                    })
                    .on("contextmenu", (ev: MouseEvent) => {
                        ev.preventDefault();
                        this.showSeriesContextMenu(ev, d.rawRowIndex, this.msDrillQueryNames[i]);
                    });
                });
            }
        });
    }

    private tipRow(label: string, value: string, color?: string): HTMLElement {
        const row = document.createElement("div");
        row.style.cssText = "display:flex;gap:6px;align-items:baseline;";
        if (color) {
            const dot = document.createElement("span");
            dot.style.cssText = `color:${color};flex-shrink:0;`;
            dot.textContent = "◆";
            row.appendChild(dot);
        }
        if (label) {
            const lbl = document.createElement("span");
            lbl.style.cssText = "opacity:0.75;white-space:nowrap;";
            lbl.textContent = label;
            row.appendChild(lbl);
        }
        const val = document.createElement("span");
        val.textContent = value;
        row.appendChild(val);
        return row;
    }

    private appendExtraFieldNodes(nodes: Node[], fields: TooltipField[]): void {
        if (!fields.length) return;
        nodes.push(this.mkHR());
        fields.forEach(f => nodes.push(this.tipRow(`${f.name}:`, f.value)));
    }

    private buildTipHeader(d: GanttRow): Node[] {
        const nodes: Node[] = [];
        const title = document.createElement("strong");
        title.style.cssText = "font-size:12.5px;display:block;";
        title.textContent = d.projectName;
        nodes.push(title);

        const loc = document.createElement("span");
        loc.style.cssText = "opacity:0.75;font-size:11px;display:block;margin-bottom:4px;";
        loc.textContent = d.location;
        nodes.push(loc);
        nodes.push(this.mkHR());
        return nodes;
    }

    // Tooltip for a single bar series only — deliberately excludes other bars'
    // and milestones' data so unrelated metrics don't clutter an unrelated tooltip.
    private buildBarTip(d: GanttRow, barIndex: number): Node[] {
        const nodes = this.buildTipHeader(d);
        const def = this.barDefs[barIndex];
        const bv = d.bars[barIndex];

        if (def && bv?.start && bv?.finish) {
            const hdr = document.createElement("span");
            hdr.style.cssText = `color:${def.color};font-weight:600;display:block;`;
            hdr.textContent = def.name;
            nodes.push(hdr);
            nodes.push(this.tipRow("Start:", this.fmtFull(bv.start)));
            nodes.push(this.tipRow("Finish:", this.fmtFull(bv.finish)));
            const days = Math.round((bv.finish.getTime() - bv.start.getTime()) / 86400000);
            nodes.push(this.tipRow("Duration:", `${days} days`));
        }

        this.appendExtraFieldNodes(nodes, bv?.extra ?? []);
        this.appendExtraFieldNodes(nodes, d.extraFields);
        return nodes;
    }

    // Tooltip for a single milestone series only — same isolation as buildBarTip.
    private buildMilestoneTip(d: GanttRow, msIndex: number): Node[] {
        const nodes = this.buildTipHeader(d);
        const def = this.milestoneDefs[msIndex];
        const ms = d.milestones[msIndex];

        if (def && ms?.date) {
            const msHdr = document.createElement("span");
            msHdr.style.cssText = `color:${def.color};font-weight:600;display:block;`;
            msHdr.textContent = `◆ ${def.name}`;
            nodes.push(msHdr);
            nodes.push(this.tipRow("Date:", this.fmtFull(ms.date)));
        }

        this.appendExtraFieldNodes(nodes, ms?.extra ?? []);
        this.appendExtraFieldNodes(nodes, d.extraFields);
        return nodes;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
