import { safeLink } from "../model/hyperlinks";
import { formatValue, numberColor } from "../model/number-format";
import { colName, inRange, MAX_COL, MAX_ROW, Range } from "../model/ooxml/refs";
import type { CellInfo, Sheet } from "../model/sheet";
import type { Table } from "../model/table";
import type { ListOption, ValidationRule } from "../model/validation";
import { editText } from "../model/values";
import type { TableLocation, Workbook } from "../model/workbook";
import type { CellData, CellStyle, GridData, RowData, RowKind } from "../view/grid-data";
import type { AreaRange, BlockConfig } from "./block-config";

const MAX_ROWS = 5000;
const MAX_DISPLAY_CELLS = 30_000;
function rowLimit(columns: number) { return Math.max(1, Math.min(MAX_ROWS, Math.floor(MAX_DISPLAY_CELLS / Math.max(columns, 1)))); }

// Colours of number formats like [Red], in Obsidian theme colours.
const NUMBER_COLORS: Record<string, string> = {
	red: "var(--text-error)",
	green: "var(--color-green)",
	blue: "var(--color-blue)",
	magenta: "var(--color-pink)",
	cyan: "var(--color-cyan)",
	yellow: "var(--color-yellow)",
};

/** Turns the part of the workbook a code block points at into data for the grid view. */
export function buildGrid(wb: Workbook, cfg: BlockConfig, readOnly = false, includeEntryColumn = false): GridData {
	const target = cfg.target;
	const grid =
		target.kind === "table" ? tableGrid(wb, wb.table(target.table)) : sheetGrid(wb, wb.sheet(target.sheet), target.range, target.columns);
	grid.tooltip = cfg.file;
	// Without headings a sheet loses its letters and row numbers; a table keeps its column names.
	grid.rowNumbers = cfg.headings;
	grid.columnHeader = cfg.headings || target.kind === "table";
	// Tables fit the page; sheets keep Excel's column widths and scroll sideways if wider.
	grid.fit = target.kind === "table";
	grid.readOnly = readOnly;
	if (target.kind === "sheet") {
		grid.sheets = wb.sheets.map((sheet) => sheet.name);
		grid.activeSheet = wb.sheet(target.sheet).name;
		grid.structureEditable = !readOnly;
		grid.frozen = wb.sheet(target.sheet).frozenPane();
		// An empty column beside an unrestricted sheet is the entry point for new data.
		const lastColumn = wb.sheet(target.sheet).displayRange()?.right ?? 1;
		if (includeEntryColumn && !readOnly && !target.range && !target.columns && lastColumn < MAX_COL && lastColumn < 256) {
			const col = lastColumn + 1;
			grid.columnNumbers!.push(col); grid.columns.push(colName(col)); grid.columnWidths!.push(8.43);
			for (const row of grid.rows) row.cells.push({ row: Number(row.label), col, text: "", editText: "", editable: true });
		}
	}
	if (readOnly) {
		grid.rows = grid.rows.filter((row) => row.kind !== "new");
		for (const row of grid.rows) for (const cell of row.cells) cell.editable = false;
	}
	if (grid.rows.some((row) => row.cells.some((cell) => cell.formula))) grid.note = [grid.note, "公式显示 Excel 上次保存的结果；修改数据后，请在 Excel 中重新计算。"].filter(Boolean).join(" ");
	return grid;
}

/** An Excel table in saved row order, excluding rows hidden by Excel filters. */
function tableGrid(wb: Workbook, location: TableLocation): GridData {
	const { table, sheet } = location;
	const ref = table.ref;
	const data = table.dataRange;
	const columns = range(ref.left, Math.min(ref.right, ref.left + 255));
	const limit = rowLimit(columns.length);
	const blocker = wb.appendBlocker(location);

	const rows: { row: number; kind: RowKind }[] = [];
	let r = data.top;
	for (; r <= data.bottom && rows.length < limit; r++) if (!sheet.isRowHidden(r)) rows.push({ row: r, kind: "data" });
	const complete = r > data.bottom;
	for (let r = data.bottom + 1; r <= ref.bottom; r++) if (!sheet.isRowHidden(r)) rows.push({ row: r, kind: "totals" });
	if (!blocker && complete) rows.push({ row: ref.bottom + 1, kind: "new" });

	const notes = [];
	if (!complete) notes.push(`Showing the first ${limit} visible data rows (preview cell budget).`);
	if (ref.right - ref.left + 1 > columns.length) notes.push("Showing the first 256 table columns. Use a sheet range to view other columns.");
	if (blocker) notes.push(`New rows can't be added: ${blocker}`);

	return {
		title: `${table.name} · ${sheet.name}`,
		tooltip: "",
		columns: table.columns.slice(0, columns.length).map((c) => c.name),
		columnWidths: columns.map((c) => columnWidth(sheet, c)),
		rows: new RowBuilder(wb, sheet, columns, rows, table).build(),
		note: notes.join(" ") || undefined,
	};
}

/**
 * A worksheet laid out like Excel shows it: its widths, merges and formatting,
 * without hidden rows/columns. `pick` lists the columns to show instead, in order
 * (shown even if hidden in Excel, since they were asked for by name).
 */
function sheetGrid(wb: Workbook, sheet: Sheet, area: AreaRange | undefined, pick: number[] | undefined): GridData {
	const used = sheet.displayRange();
	const top = area?.top ?? 1;
	const left = area?.left ?? 1;
	const right = area?.right ?? Math.min(Math.max(used?.right ?? 1, left), left + 255);
	const usedBottom = area?.bottom ?? Math.max(used?.bottom ?? 0, top - 1);

	const columns = (pick ?? range(left, right).filter((c) => !sheet.column(c).hidden)).slice(0, 256);
	const limit = rowLimit(columns.length + 1);
	const rows: { row: number; kind: RowKind }[] = [];
	let r = top;
	for (; r <= usedBottom && rows.length < limit; r++) {
		if (!sheet.isRowHidden(r)) rows.push({ row: r, kind: "data" });
	}
	const complete = r > usedBottom;
	// An open-ended area gets an empty row at the bottom for adding entries.
	if (area?.bottom === undefined && complete && usedBottom < MAX_ROW) rows.push({ row: usedBottom + 1, kind: "new" });

	return {
		title: sheet.name,
		tooltip: "",
		columns: columns.map(colName),
		columnNumbers: [...columns],
		columnWidths: columns.map((c) => columnWidth(sheet, c)),
		rows: new RowBuilder(wb, sheet, columns, rows).build(),
		note: [
			complete ? "" : `Showing the first ${limit} visible rows (preview cell budget). Add a range: line to show a different part.`,
			(pick && pick.length > 256) || (!area && (used?.right ?? 1) > right) || right - left + 1 > 256 ? "Showing the first 256 columns. Use range: or cols: to show other columns." : "",
		].filter(Boolean).join(" ") || undefined,
	};
}

/** Excel's width, or the default one when it's 0 (as hidden columns sometimes store). */
function columnWidth(sheet: Sheet, col: number): number {
	return sheet.column(col).width || 8.43;
}

function range(from: number, to: number): number[] {
	const out = [];
	for (let i = from; i <= to; i++) out.push(i);
	return out;
}

/** Builds the cells of the shown rows and columns, working out each validation rule only once. */
class RowBuilder {
	private merges: Range[];
	private shownRows: Set<number>;
	private shownCols: Set<number>;
	/** Merged cells can only span when columns are shown left to right. */
	private canSpan: boolean;
	private ruleInfo = new Map<ValidationRule, { options?: ListOption[]; help?: string }>();

	constructor(
		private wb: Workbook,
		private sheet: Sheet,
		private columns: number[],
		private rows: { row: number; kind: RowKind }[],
		/** Set in table mode; in sheet mode each cell looks up its own table. */
		private table?: Table,
	) {
		this.merges = sheet.mergedRanges();
		this.shownRows = new Set(rows.map((r) => r.row));
		this.shownCols = new Set(columns);
		this.canSpan = columns.every((c, i) => i === 0 || c > columns[i - 1]);
	}

	build(): RowData[] {
		return this.rows.map(({ row, kind }) => ({ label: String(row), kind, cells: this.columns.map((col) => this.cell(row, col)) }));
	}

	private cell(r: number, c: number): CellData {
		const merge = this.merges.find((m) => inRange(m, r, c));
		if (merge && !this.canSpan) {
			// Reordered columns: show the value in the merge's first cell only.
			if (merge.top === r && merge.left === c) return { ...this.content(r, c), mergeRange: merge };
			return { row: r, col: c, text: "", editText: "", editable: false, mergeRange: merge, hint: "Part of a merged cell" };
		}
		if (merge) {
			// The merged block is drawn by its first visible cell (normally its top-left).
			const first = this.firstShown(merge);
			if (first.row !== r || first.col !== c) return { row: r, col: c, text: "", editText: "", editable: false, covered: true, mergeRange: merge };
			return { ...this.content(merge.top, merge.left), span: this.spanOf(merge), mergeRange: merge };
		}
		return this.content(r, c);
	}

	private firstShown(m: Range): { row: number; col: number } {
		let row = m.top;
		while (row < m.bottom && !this.shownRows.has(row)) row++;
		let col = m.left;
		while (col < m.right && !this.shownCols.has(col)) col++;
		return { row, col };
	}

	private spanOf(m: Range): { cols: number; rows: number } {
		let cols = 0;
		let rows = 0;
		for (let c = m.left; c <= m.right; c++) if (this.shownCols.has(c)) cols++;
		for (let r = m.top; r <= m.bottom; r++) if (this.shownRows.has(r)) rows++;
		return { cols: Math.max(cols, 1), rows: Math.max(rows, 1) };
	}

	private content(r: number, c: number): CellData {
		const info = this.sheet.cell(r, c);
		const header = this.sheet.isTableHeader(r, c);
		const pending = !!info.formula && info.value === null;
		const table = this.table ?? this.wb.tableAt(this.sheet, r, c)?.table;
		const rule = header ? undefined : this.rule(r, c, table);
		return {
			row: r,
			col: c,
			text: formatValue(info.value, info.format),
			link: this.sheet.hyperlink(r, c) ?? (typeof info.value === "string" ? safeLink(info.value) : undefined),
			editText: editText(info.value, info.formula),
			editable: !header,
			hasContent: this.sheet.hasContent(r, c),
			tableCell: !!table && inRange(table.ref, r, c),
			numeric: typeof info.value === "number",
			formula: !!info.formula,
			pending,
			options: rule?.options,
			help: rule?.help,
			hint: header ? "Table header: rename columns in Excel" : info.formula ? `=${info.formula} · 显示 Excel 上次计算结果；修改数据后需在 Excel 中重新计算` : rule?.help,
			tableHeader: header || undefined,
			style: styleOf(info),
		};
	}

	private rule(r: number, c: number, table: Table | undefined) {
		const rule = this.wb.validationAt(this.sheet, r, c, table);
		if (!rule) return undefined;
		let info = this.ruleInfo.get(rule);
		if (!info) {
			const options = rule.showDropdown ? this.wb.listOptions(this.sheet, rule) : [];
			info = {
				options: options.length ? options : undefined,
				help: this.wb.validationHelp(this.sheet, rule) || undefined,
			};
			this.ruleInfo.set(rule, info);
		}
		return info;
	}
}

function styleOf(info: CellInfo): CellStyle | undefined {
	const look = info.look;
	const formatColor = typeof info.value === "number" ? numberColor(info.value, info.format) : undefined;
	const color = (formatColor && NUMBER_COLORS[formatColor]) ?? look.color;
	const style: CellStyle = {};
	if (look.bold) style.bold = true;
	if (look.italic) style.italic = true;
	if (look.underline) style.underline = true;
	if (look.strike) style.strike = true;
	if (color) style.color = color;
	if (look.fill) style.background = look.fill;
	if (look.align) style.align = look.align;
	if (look.wrap) style.wrap = true;
	return Object.keys(style).length ? style : undefined;
}
