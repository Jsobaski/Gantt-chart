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
import DataView = powerbi.DataView;
import IViewport = powerbi.IViewport;

import { VisualFormattingSettingsModel } from "./settings";

interface SeriesDef {
    name: string;
    queryName: string;
    color: string;
}

interface BarValue {
    start: Date | null;
    finish: Date | null;
}

interface GanttRow {
    location: string;
    projectName: string;
    // Aligned by index with this.milestoneDefs
    milestones: (Date | null)[];
    // Aligned by index with this.barDefs
    bars: BarValue[];
    extraFields: { name: string; value: string }[];
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
    private readonly HEADER_HEIGHT = 38;
    private readonly CONTROLS_HEIGHT = 44;
    private readonly BAR_HEIGHT = 11;
    private readonly BAR_GAP = 2;
    private readonly MILESTONE_RADIUS = 7;
    private readonly LANE_GAP = 5;
    private readonly ROW_V_PADDING = 16;

    // Colors (defaults overridden by format panel)
    private bgColor = "#0d1b2a";
    private textColor = "#d0e4f7";
    private readonly headerBgColor = "#0a1728";
    private readonly locationBgColor = "#112840";
    private readonly gridColor = "#1a3a5c";
    private readonly todayColor = "#ef5350";

    constructor(options: VisualConstructorOptions) {
        this.events = options.host.eventService;
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.container = options.element;
        this.container.style.overflow = "hidden";
        this.container.style.fontFamily = "Segoe UI, Arial, sans-serif";
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

    private applySettings(): void {
        if (!this.formattingSettings) return;
        const s = this.formattingSettings.ganttConfig;
        this.bgColor = s.bgColor.value.value || this.bgColor;
        this.textColor = s.textColor.value.value || this.textColor;
    }

    // Per-field override (set via the "Series Colors" section of the Format pane)
    // takes priority over the automatically assigned palette color.
    private resolveColor(col: powerbi.DataViewMetadataColumn, fallbackKey: string): string {
        const override = (col.objects as Record<string, Record<string, { solid?: { color?: string } }>> | undefined)
            ?.dataColors?.fill?.solid?.color;
        return override || this.host.colorPalette.getColor(fallbackKey).value;
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
        if (!table?.rows?.length) {
            this.milestoneDefs = [];
            this.barDefs = [];
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
        columns.forEach((col, i) => {
            const name = col.displayName || col.queryName || `Field ${i}`;
            if (col.roles?.["milestoneDate"]) milestoneCols.push({ idx: i, name });
            if (col.roles?.["barStart"]) barStartCols.push({ idx: i, name });
            if (col.roles?.["barFinish"]) barFinishCols.push({ idx: i, name });
            if (col.roles?.["tooltipFields"]) tooltipCols.push({ idx: i, name });
        });

        // Bar starts/finishes are paired positionally: 1st Start field with 1st Finish
        // field, 2nd with 2nd, etc. Add them to the field wells in matching order.
        const barCount = Math.min(barStartCols.length, barFinishCols.length);

        // Rename a field via right-click > "Rename for this visual" in Power BI to
        // control exactly what shows up in the legend and tooltips. A per-field color
        // set in the Format pane ("Series Colors") persists on that column's `objects`
        // and takes priority over the automatically assigned palette color.
        this.milestoneDefs = milestoneCols.map(c => ({
            name: c.name,
            queryName: columns[c.idx].queryName || c.name,
            color: this.resolveColor(columns[c.idx], `milestone_${c.name}`)
        }));
        this.barDefs = barStartCols.slice(0, barCount).map(c => ({
            name: c.name,
            queryName: columns[c.idx].queryName || c.name,
            color: this.resolveColor(columns[c.idx], `bar_${c.name}`)
        }));

        const getExtraFields = (row: powerbi.DataViewTableRow): { name: string; value: string }[] => {
            const out: { name: string; value: string }[] = [];
            tooltipCols.forEach(({ idx, name }) => {
                const v = row[idx];
                if (v === null || v === undefined || v === "") return;
                let text: string;
                if (v instanceof Date) text = this.fmtFull(v);
                else if (typeof v === "number") text = v.toLocaleString(undefined, { maximumFractionDigits: 2 });
                else text = String(v).trim();
                if (text) out.push({ name, value: text });
            });
            return out;
        };

        const rows: GanttRow[] = [];
        table.rows.forEach(row => {
            const location = getString(row, "location");
            const projectName = getString(row, "projectName");
            if (!location && !projectName) return;

            const milestones = milestoneCols.map(c => this.parseSingleValue(row[c.idx]));
            const bars: BarValue[] = [];
            for (let i = 0; i < barCount; i++) {
                bars.push({
                    start: this.parseSingleValue(row[barStartCols[i].idx]),
                    finish: this.parseSingleValue(row[barFinishCols[i].idx])
                });
            }

            const hasBar = bars.some(b => b.start && b.finish);
            const hasMilestone = milestones.some(m => m !== null);
            if (!hasBar && !hasMilestone) return;

            rows.push({ location, projectName, milestones, bars, extraFields: getExtraFields(row) });
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

    private computeDateRange(rows: GanttRow[]): { from: Date; to: Date } {
        const dates: Date[] = [];
        rows.forEach(r => {
            r.milestones.forEach(d => { if (d) dates.push(d); });
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
        const barsLaneH = barCount > 0 ? barCount * this.BAR_HEIGHT + Math.max(0, barCount - 1) * this.BAR_GAP : 0;
        const laneGap = (milestoneLaneH > 0 && barsLaneH > 0) ? this.LANE_GAP : 0;
        const contentH = milestoneLaneH + laneGap + barsLaneH;
        const rowHeight = Math.max(this.MIN_ROW_HEIGHT, contentH + this.ROW_V_PADDING);
        return { rowHeight, contentH, milestoneLaneH, laneGap, barsLaneH };
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

        const { from: dataFrom, to: dataTo } = this.computeDateRange(this.allRows);
        const dateFrom = this.userDateFrom || dataFrom;
        const dateTo = this.userDateTo || dataTo;

        const { rowHeight, contentH, milestoneLaneH, laneGap } = this.computeRowMetrics();

        const timelineWidth = Math.max(W - this.LABEL_WIDTH, 500);
        const totalBodyHeight = displayRows.length * rowHeight;

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
            `height:${this.CONTROLS_HEIGHT}px`, "flex-shrink:0",
            "display:flex", "align-items:center", "gap:10px",
            "padding:0 10px", `background:${this.headerBgColor}`,
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
            chip.style.cssText = "display:flex;align-items:center;gap:4px;font-size:10.5px;";
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
            `background:#112840`, `color:${this.textColor}`,
            "border:1px solid #2a5080", "border-radius:4px",
            "padding:2px 6px", "font-size:11px", "outline:none"
        ].join(";");

        const mkLabel = (txt: string, inputEl: HTMLInputElement) => {
            const lbl = document.createElement("label");
            lbl.style.cssText = "display:flex;align-items:center;gap:5px;font-size:11px;white-space:nowrap;";
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

        const resetBtn = document.createElement("button");
        resetBtn.textContent = "Reset";
        resetBtn.style.cssText = [
            "background:#112840", `color:${this.textColor}`,
            "border:1px solid #2a5080", "border-radius:4px",
            "padding:3px 10px", "font-size:11px", "cursor:pointer"
        ].join(";");

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
            `height:${this.HEADER_HEIGHT}px`,
            `background:${this.headerBgColor}`,
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
        monthSvgEl.style.cssText = `width:${timelineWidth}px;height:${this.HEADER_HEIGHT}px;display:block;`;
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

        // ── Month header labels ───────────────────────────────────────────
        monthSvg.append("rect")
            .attr("x", 0).attr("y", 0)
            .attr("width", timelineWidth).attr("height", this.HEADER_HEIGHT)
            .attr("fill", this.headerBgColor);

        months.forEach(m => {
            const x = xScale(m);
            monthSvg.append("line")
                .attr("x1", x).attr("y1", 0)
                .attr("x2", x).attr("y2", this.HEADER_HEIGHT)
                .attr("stroke", this.gridColor).attr("stroke-width", 0.5);
            monthSvg.append("text")
                .attr("x", x + 5)
                .attr("y", this.HEADER_HEIGHT / 2 + 5)
                .attr("fill", this.textColor)
                .attr("font-size", "11px")
                .attr("font-family", "Segoe UI, Arial, sans-serif")
                .attr("font-weight", "600")
                .text(d3.timeFormat("%b %Y")(m));
        });

        // ── Row rendering ─────────────────────────────────────────────────
        displayRows.forEach((row, idx) => {
            const y = idx * rowHeight;
            const isEvenRow = idx % 2 === 0;

            if (row.type === "location") {
                // ── Location header ─────────────────────────────────────
                const div = document.createElement("div");
                div.style.cssText = [
                    `height:${rowHeight}px`,
                    "display:flex", "align-items:center",
                    "padding:0 10px", "font-weight:700", "font-size:12px",
                    "letter-spacing:0.06em", "text-transform:uppercase",
                    `background:${this.locationBgColor}`,
                    `border-bottom:1px solid ${this.gridColor}`,
                    `color:${this.textColor}`, "cursor:default"
                ].join(";");
                div.textContent = `▶  ${row.name}`;
                labelCol.appendChild(div);

                svg.append("rect")
                    .attr("x", 0).attr("y", y)
                    .attr("width", timelineWidth).attr("height", rowHeight)
                    .attr("fill", this.locationBgColor).attr("fill-opacity", 0.45);
                svg.append("line")
                    .attr("x1", 0).attr("y1", y + rowHeight)
                    .attr("x2", timelineWidth).attr("y2", y + rowHeight)
                    .attr("stroke", this.gridColor).attr("stroke-width", 0.6);

            } else {
                // ── Project row ──────────────────────────────────────────
                const d = row.data;
                const rowBg = isEvenRow ? "rgba(255,255,255,0.025)" : "rgba(0,0,0,0)";

                // Label div
                const div = document.createElement("div");
                div.style.cssText = [
                    `height:${rowHeight}px`,
                    "display:flex", "align-items:center",
                    "padding:0 6px 0 22px",
                    "font-size:11.5px",
                    `border-bottom:1px solid ${this.gridColor}`,
                    "cursor:default", `background:${rowBg}`,
                    "white-space:nowrap", "overflow:hidden"
                ].join(";");
                div.title = `${d.location} › ${d.projectName}`;

                const marker = document.createElement("span");
                marker.style.cssText = `color:${this.textColor};opacity:0.45;margin-right:5px;font-size:10px;flex-shrink:0;`;
                marker.textContent = "□";
                div.appendChild(marker);

                const nameSpan = document.createElement("span");
                nameSpan.style.cssText = "overflow:hidden;text-overflow:ellipsis;";
                nameSpan.textContent = d.projectName || d.location;
                div.appendChild(nameSpan);
                labelCol.appendChild(div);

                // Row background stripe on SVG
                svg.append("rect")
                    .attr("x", 0).attr("y", y)
                    .attr("width", timelineWidth).attr("height", rowHeight)
                    .attr("fill", rowBg);
                svg.append("line")
                    .attr("x1", 0).attr("y1", y + rowHeight)
                    .attr("x2", timelineWidth).attr("y2", y + rowHeight)
                    .attr("stroke", this.gridColor).attr("stroke-width", 0.3);

                // Milestone lane centered above the bars lane, whole block vertically centered in the row
                const contentTop = y + (rowHeight - contentH) / 2;
                const milestoneCenterY = contentTop + this.MILESTONE_RADIUS;
                const barsTop = contentTop + milestoneLaneH + laneGap;

                // ── Bars (stacked, one lane per dynamic bar series) ──────
                d.bars.forEach((bv, i) => {
                    if (!bv.start || !bv.finish || bv.start > bv.finish) return;
                    const def = this.barDefs[i];
                    if (!def) return;
                    const barY = barsTop + i * (this.BAR_HEIGHT + this.BAR_GAP);
                    const x1 = xScale(bv.start);
                    const x2 = xScale(bv.finish);
                    const clampX1 = Math.max(0, x1);
                    const clampX2 = Math.min(timelineWidth, x2);
                    if (clampX2 <= clampX1) return;
                    const bw = clampX2 - clampX1;

                    const bar = svg.append("rect")
                        .attr("x", clampX1).attr("y", barY)
                        .attr("width", bw).attr("height", this.BAR_HEIGHT)
                        .attr("fill", def.color)
                        .attr("rx", 3).attr("ry", 3)
                        .style("cursor", "pointer");

                    if (bw > 90) {
                        svg.append("text")
                            .attr("x", clampX1 + bw / 2)
                            .attr("y", barY + this.BAR_HEIGHT / 2 + 4)
                            .attr("text-anchor", "middle")
                            .attr("font-size", "9px")
                            .attr("fill", "#041020")
                            .attr("font-family", "Segoe UI, Arial, sans-serif")
                            .attr("pointer-events", "none")
                            .attr("font-weight", "600")
                            .text(`${def.name}: ${this.fmtShort(bv.start)} - ${this.fmtShort(bv.finish)}`);
                    }

                    bar.on("mouseenter", (ev: MouseEvent) => {
                        bar.attr("fill-opacity", 0.85);
                        showTip(this.buildBarTip(d, i), ev);
                    })
                    .on("mousemove", (ev: MouseEvent) => moveTip(ev))
                    .on("mouseleave", () => { bar.attr("fill-opacity", 1); hideTip(); });
                });

                // ── Milestone diamonds (one lane, N diamonds across the timeline) ──
                d.milestones.forEach((date, i) => {
                    const def = this.milestoneDefs[i];
                    if (!date || !def || date < dateFrom || date > dateTo) return;
                    const mx = xScale(date);
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

    private appendExtraFieldNodes(nodes: Node[], d: GanttRow): void {
        if (!d.extraFields.length) return;
        nodes.push(this.mkHR());
        d.extraFields.forEach(f => nodes.push(this.tipRow(`${f.name}:`, f.value)));
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

        this.appendExtraFieldNodes(nodes, d);
        return nodes;
    }

    // Tooltip for a single milestone series only — same isolation as buildBarTip.
    private buildMilestoneTip(d: GanttRow, msIndex: number): Node[] {
        const nodes = this.buildTipHeader(d);
        const def = this.milestoneDefs[msIndex];
        const date = d.milestones[msIndex];

        if (def && date) {
            const msHdr = document.createElement("span");
            msHdr.style.cssText = `color:${def.color};font-weight:600;display:block;`;
            msHdr.textContent = `◆ ${def.name}`;
            nodes.push(msHdr);
            nodes.push(this.tipRow("Date:", this.fmtFull(date)));
        }

        this.appendExtraFieldNodes(nodes, d);
        return nodes;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
