"use strict";

import "./../style/visual.less";
import * as d3 from "d3";
import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import IViewport = powerbi.IViewport;

import { VisualFormattingSettingsModel } from "./settings";

interface GanttRow {
    location: string;
    projectName: string;
    cappExecutionStart: Date | null;
    cappFundingFYDate: Date | null;
    cappPlanningStart: Date | null;
    maximoStart: Date | null;
    maximoFinish: Date | null;
    p6Start: Date | null;
    p6Finish: Date | null;
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
    private readonly ROW_HEIGHT = 40;
    private readonly HEADER_HEIGHT = 38;
    private readonly CONTROLS_HEIGHT = 44;
    private readonly BAR_HEIGHT = 11;
    private readonly BAR_GAP = 2;
    private readonly MILESTONE_RADIUS = 7;

    // Colors (defaults overridden by format panel)
    private p6Color = "#4dabf7";
    private maximoColor = "#26c6da";
    private cappExecColor = "#ab47bc";
    private cappFundingColor = "#66bb6a";
    private cappPlanningColor = "#ffa726";
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
                msg.textContent = "Map data fields to this visual: Location, Project Name, date fields.";
                this.container.appendChild(msg);
                this.events.renderingFinished(options);
                return;
            }

            this.allRows = this.parseRows(options.dataViews[0]);
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
        this.p6Color = s.p6BarColor.value.value || this.p6Color;
        this.maximoColor = s.maximoBarColor.value.value || this.maximoColor;
        this.cappExecColor = s.cappExecColor.value.value || this.cappExecColor;
        this.cappFundingColor = s.cappFundingColor.value.value || this.cappFundingColor;
        this.cappPlanningColor = s.cappPlanningColor.value.value || this.cappPlanningColor;
        this.bgColor = s.bgColor.value.value || this.bgColor;
        this.textColor = s.textColor.value.value || this.textColor;
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

    // Handles both raw date columns and Power BI date hierarchies (Year/Quarter/Month/Day).
    // Power BI auto date/time always produces exactly 4 columns per date field in this order:
    //   index 0 = Year, index 1 = Quarter, index 2 = Month, index 3 = Day
    // Column display names vary by locale/version so we use positional matching as primary
    // and keyword matching as a cross-check.
    private getDateForRole(
        row: powerbi.DataViewTableRow,
        columns: powerbi.DataViewMetadataColumn[],
        role: string
    ): Date | null {
        const matches: Array<{idx: number; name: string}> = [];
        (columns || []).forEach((col, idx) => {
            if (col.roles?.[role]) {
                matches.push({ idx, name: (col.displayName || col.queryName || "").toLowerCase() });
            }
        });

        if (matches.length === 0) return null;
        if (matches.length === 1) return this.parseSingleValue(row[matches[0].idx]);

        // Multiple columns = date hierarchy.
        // Strategy A: name-based (works when columns are named "Year", "Month", "Day" etc.)
        let year: number | null = null;
        let month: number | null = null;
        let day: number | null = null;

        for (const { idx, name } of matches) {
            const v = row[idx];
            if (v === null || v === undefined) continue;
            const n = Number(v);
            if (isNaN(n) || !Number.isFinite(n)) continue;
            if (name.includes("year") && !name.includes("quarter")) year = n;
            else if (name.includes("month")) month = n;
            else if (name.includes("day") && !name.includes("week")) day = n;
        }

        // Strategy B: positional fallback for non-English or custom column names.
        // Power BI hierarchy order is always: [Year, Quarter, Month, Day]
        const numAt = (i: number): number | null => {
            if (i >= matches.length) return null;
            const n = Number(row[matches[i].idx]);
            return isNaN(n) || !Number.isFinite(n) ? null : n;
        };

        if (year === null) {
            const n = numAt(0);
            if (n !== null && n >= 1900 && n <= 2200) year = n;
        }
        if (month === null) {
            // index 2 = Month (1-12)
            const n = numAt(2);
            if (n !== null && n >= 1 && n <= 12) month = n;
        }
        if (day === null) {
            // index 3 = Day (1-31)
            const n = numAt(3);
            if (n !== null && n >= 1 && n <= 31) day = n;
        }

        if (year !== null) {
            const date = new Date(year, month !== null ? month - 1 : 0, day ?? 1);
            return isNaN(date.getTime()) ? null : date;
        }

        return this.parseSingleValue(row[matches[0].idx]);
    }

    private parseRows(dataView: DataView): GanttRow[] {
        const table = dataView.table;
        if (!table?.rows?.length) return [];

        const columns = table.columns || [];

        // Text fields: still use first-column-per-role map
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

        const getDate = (row: powerbi.DataViewTableRow, role: string): Date | null =>
            this.getDateForRole(row, columns, role);

        const rows: GanttRow[] = [];
        table.rows.forEach(row => {
            const location = getString(row, "location");
            const projectName = getString(row, "projectName");
            if (!location && !projectName) return;

            const cappExecutionStart = getDate(row, "cappExecutionStart");
            const cappFundingFYDate = getDate(row, "cappFundingFYDate");
            const cappPlanningStart = getDate(row, "cappPlanningStart");
            const maximoStart = getDate(row, "maximoStart");
            const maximoFinish = getDate(row, "maximoFinish");
            const p6Start = getDate(row, "p6Start");
            const p6Finish = getDate(row, "p6Finish");

            const hasBar = (maximoStart && maximoFinish) || (p6Start && p6Finish);
            const hasMilestone = cappExecutionStart || cappFundingFYDate || cappPlanningStart;
            if (!hasBar && !hasMilestone) return;

            rows.push({
                location, projectName,
                cappExecutionStart, cappFundingFYDate, cappPlanningStart,
                maximoStart, maximoFinish, p6Start, p6Finish
            });
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
        map.forEach((grpRows, name) => groups.push({ name, rows: grpRows }));
        return groups;
    }

    private computeDateRange(rows: GanttRow[]): { from: Date; to: Date } {
        const dates: Date[] = [];
        rows.forEach(r => {
            [r.cappExecutionStart, r.cappFundingFYDate, r.cappPlanningStart,
                r.maximoStart, r.maximoFinish, r.p6Start, r.p6Finish]
                .forEach(d => { if (d) dates.push(d); });
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

    private esc(s: string): string {
        return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
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

        const timelineWidth = Math.max(W - this.LABEL_WIDTH, 500);
        const totalBodyHeight = displayRows.length * this.ROW_HEIGHT;

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

        // Legend chips
        const legend = document.createElement("div");
        legend.style.cssText = "display:flex;align-items:center;gap:14px;flex:1;";
        const legendItems: Array<{ color: string; label: string; isBar: boolean }> = [
            { color: this.p6Color, label: "P6", isBar: true },
            { color: this.maximoColor, label: "Maximo", isBar: true },
            { color: this.cappPlanningColor, label: "Planning Start", isBar: false },
            { color: this.cappFundingColor, label: "Funding FY", isBar: false },
            { color: this.cappExecColor, label: "Execution Start", isBar: false }
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

        const onDateChange = () => {
            const f = fromInput.valueAsDate;
            const t = toInput.valueAsDate;
            if (f) this.userDateFrom = f;
            if (t) this.userDateTo = t;
            this.render(viewport);
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
            const y = idx * this.ROW_HEIGHT;
            const isEvenRow = idx % 2 === 0;

            if (row.type === "location") {
                // ── Location header ─────────────────────────────────────
                const div = document.createElement("div");
                div.style.cssText = [
                    `height:${this.ROW_HEIGHT}px`,
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
                    .attr("width", timelineWidth).attr("height", this.ROW_HEIGHT)
                    .attr("fill", this.locationBgColor).attr("fill-opacity", 0.45);
                svg.append("line")
                    .attr("x1", 0).attr("y1", y + this.ROW_HEIGHT)
                    .attr("x2", timelineWidth).attr("y2", y + this.ROW_HEIGHT)
                    .attr("stroke", this.gridColor).attr("stroke-width", 0.6);

            } else {
                // ── Project row ──────────────────────────────────────────
                const d = row.data;
                const rowBg = isEvenRow ? "rgba(255,255,255,0.025)" : "rgba(0,0,0,0)";

                // Label div
                const div = document.createElement("div");
                div.style.cssText = [
                    `height:${this.ROW_HEIGHT}px`,
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
                    .attr("width", timelineWidth).attr("height", this.ROW_HEIGHT)
                    .attr("fill", rowBg);
                svg.append("line")
                    .attr("x1", 0).attr("y1", y + this.ROW_HEIGHT)
                    .attr("x2", timelineWidth).attr("y2", y + this.ROW_HEIGHT)
                    .attr("stroke", this.gridColor).attr("stroke-width", 0.3);

                const midY = y + this.ROW_HEIGHT / 2;
                const p6BarY = midY - this.BAR_GAP - this.BAR_HEIGHT;
                const maxBarY = midY + this.BAR_GAP;

                // ── P6 bar ───────────────────────────────────────────────
                if (d.p6Start && d.p6Finish && d.p6Start <= d.p6Finish) {
                    const x1 = xScale(d.p6Start);
                    const x2 = xScale(d.p6Finish);
                    const clampX1 = Math.max(0, x1);
                    const clampX2 = Math.min(timelineWidth, x2);
                    if (clampX2 > clampX1) {
                        const bw = clampX2 - clampX1;
                        const bar = svg.append("rect")
                            .attr("x", clampX1).attr("y", p6BarY)
                            .attr("width", bw).attr("height", this.BAR_HEIGHT)
                            .attr("fill", this.p6Color)
                            .attr("rx", 3).attr("ry", 3)
                            .style("cursor", "pointer");

                        if (bw > 90) {
                            svg.append("text")
                                .attr("x", clampX1 + bw / 2)
                                .attr("y", p6BarY + this.BAR_HEIGHT / 2 + 4)
                                .attr("text-anchor", "middle")
                                .attr("font-size", "9px")
                                .attr("fill", "#041020")
                                .attr("font-family", "Segoe UI, Arial, sans-serif")
                                .attr("pointer-events", "none")
                                .attr("font-weight", "600")
                                .text(`P6: ${this.fmtShort(d.p6Start)} - ${this.fmtShort(d.p6Finish)}`);
                        }

                        bar.on("mouseenter", (ev: MouseEvent) => {
                            bar.attr("fill-opacity", 0.85);
                            showTip(this.buildBarTip(d, "p6"), ev);
                        })
                        .on("mousemove", (ev: MouseEvent) => moveTip(ev))
                        .on("mouseleave", () => { bar.attr("fill-opacity", 1); hideTip(); });
                    }
                }

                // ── Maximo bar ───────────────────────────────────────────
                if (d.maximoStart && d.maximoFinish && d.maximoStart <= d.maximoFinish) {
                    const x1 = xScale(d.maximoStart);
                    const x2 = xScale(d.maximoFinish);
                    const clampX1 = Math.max(0, x1);
                    const clampX2 = Math.min(timelineWidth, x2);
                    if (clampX2 > clampX1) {
                        const bw = clampX2 - clampX1;
                        const bar = svg.append("rect")
                            .attr("x", clampX1).attr("y", maxBarY)
                            .attr("width", bw).attr("height", this.BAR_HEIGHT)
                            .attr("fill", this.maximoColor)
                            .attr("rx", 3).attr("ry", 3)
                            .style("cursor", "pointer");

                        if (bw > 90) {
                            svg.append("text")
                                .attr("x", clampX1 + bw / 2)
                                .attr("y", maxBarY + this.BAR_HEIGHT / 2 + 4)
                                .attr("text-anchor", "middle")
                                .attr("font-size", "9px")
                                .attr("fill", "#041020")
                                .attr("font-family", "Segoe UI, Arial, sans-serif")
                                .attr("pointer-events", "none")
                                .attr("font-weight", "600")
                                .text(`Maximo: ${this.fmtShort(d.maximoStart)} - ${this.fmtShort(d.maximoFinish)}`);
                        }

                        bar.on("mouseenter", (ev: MouseEvent) => {
                            bar.attr("fill-opacity", 0.85);
                            showTip(this.buildBarTip(d, "maximo"), ev);
                        })
                        .on("mousemove", (ev: MouseEvent) => moveTip(ev))
                        .on("mouseleave", () => { bar.attr("fill-opacity", 1); hideTip(); });
                    }
                }

                // ── Milestone diamonds ───────────────────────────────────
                const milestones = [
                    { date: d.cappPlanningStart, color: this.cappPlanningColor, label: "CAPP Planning Start" },
                    { date: d.cappFundingFYDate, color: this.cappFundingColor, label: "CAPP Funding FY Date" },
                    { date: d.cappExecutionStart, color: this.cappExecColor, label: "CAPP Execution Start" }
                ];

                milestones.forEach(ms => {
                    if (!ms.date || ms.date < dateFrom || ms.date > dateTo) return;
                    const mx = xScale(ms.date);
                    const r = this.MILESTONE_RADIUS;
                    const pts = `${mx},${midY - r} ${mx + r},${midY} ${mx},${midY + r} ${mx - r},${midY}`;

                    const diamond = svg.append("polygon")
                        .attr("points", pts)
                        .attr("fill", ms.color)
                        .attr("stroke", "rgba(255,255,255,0.4)")
                        .attr("stroke-width", 0.8)
                        .style("cursor", "pointer");

                    diamond.on("mouseenter", (ev: MouseEvent) => {
                        diamond.attr("stroke-width", 2).attr("stroke", "white");
                        showTip(this.buildMilestoneTip(d, ms), ev);
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

    private buildBarTip(d: GanttRow, barType: "p6" | "maximo"): Node[] {
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

        if (barType === "p6" && d.p6Start && d.p6Finish) {
            const hdr = document.createElement("span");
            hdr.style.cssText = `color:${this.p6Color};font-weight:600;display:block;`;
            hdr.textContent = "P6 Schedule";
            nodes.push(hdr);
            nodes.push(this.tipRow("Start:", this.fmtFull(d.p6Start)));
            nodes.push(this.tipRow("Finish:", this.fmtFull(d.p6Finish)));
            const days = Math.round((d.p6Finish.getTime() - d.p6Start.getTime()) / 86400000);
            nodes.push(this.tipRow("Duration:", `${days} days`));
        } else if (barType === "maximo" && d.maximoStart && d.maximoFinish) {
            const hdr = document.createElement("span");
            hdr.style.cssText = `color:${this.maximoColor};font-weight:600;display:block;`;
            hdr.textContent = "Maximo Schedule";
            nodes.push(hdr);
            nodes.push(this.tipRow("Start:", this.fmtFull(d.maximoStart)));
            nodes.push(this.tipRow("Finish:", this.fmtFull(d.maximoFinish)));
            const days = Math.round((d.maximoFinish.getTime() - d.maximoStart.getTime()) / 86400000);
            nodes.push(this.tipRow("Duration:", `${days} days`));
        }

        nodes.push(this.mkHR());
        if (d.cappPlanningStart) nodes.push(this.tipRow("Planning Start:", this.fmtFull(d.cappPlanningStart), this.cappPlanningColor));
        if (d.cappFundingFYDate) nodes.push(this.tipRow("Funding FY:", this.fmtFull(d.cappFundingFYDate), this.cappFundingColor));
        if (d.cappExecutionStart) nodes.push(this.tipRow("Execution Start:", this.fmtFull(d.cappExecutionStart), this.cappExecColor));
        return nodes;
    }

    private buildMilestoneTip(d: GanttRow, ms: { date: Date; color: string; label: string }): Node[] {
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

        const msHdr = document.createElement("span");
        msHdr.style.cssText = `color:${ms.color};font-weight:600;display:block;`;
        msHdr.textContent = `◆ ${ms.label}`;
        nodes.push(msHdr);
        nodes.push(this.tipRow("Date:", this.fmtFull(ms.date)));

        if (d.p6Start && d.p6Finish) nodes.push(this.tipRow("P6:", `${this.fmtShort(d.p6Start)} → ${this.fmtShort(d.p6Finish)}`));
        if (d.maximoStart && d.maximoFinish) nodes.push(this.tipRow("Maximo:", `${this.fmtShort(d.maximoStart)} → ${this.fmtShort(d.maximoFinish)}`));
        return nodes;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
