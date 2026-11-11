import { safeLink } from "./hyperlinks";
import { shiftFormula } from "./ooxml/formula";
import type { Package } from "./ooxml/package";
import { formatRange, formatRef, inRange, MAX_COL, MAX_ROW, parseRange, parseRef, Range } from "./ooxml/refs";
import type { CellFormat, CellStylePatch, Styles } from "./ooxml/styles";
import { child, createChild, elements, removeChildren, richText, setText, textOf } from "./ooxml/xml";
import type { Table } from "./table";
import { readValidations, ValidationRule } from "./validation";
import { CellInput, CellValue, dateToSerial, serialToDate, validateInput } from "./values";

export interface CellInfo {
	value: CellValue;
	formula?: string;
	/** Excel number format code, e.g. `#,##0.00 "€"`. */
	format: string;
	/** Font, fill and alignment. */
	look: CellFormat;
}

export interface ColumnInfo {
	/** In Excel's unit: roughly the number of characters that fit. */
	width: number;
	hidden: boolean;
	style?: number;
}

// Excel's width of a column nobody resized (Calibri 11).
const DEFAULT_COLUMN_WIDTH = 8.43;

/** Workbook-wide data a sheet needs to read and write cells. */
export interface SheetContext {
	pkg: Package;
	styles: Styles;
	sharedStrings: string[];
	date1904: boolean;
}

interface CellEntry {
	row: number;
	col: number;
	el: Element;
}

interface SharedMaster {
	formula: string;
	row: number;
	col: number;
}

const key = (row: number, col: number) => `${row}:${col}`;

/** A worksheet (xl/worksheets/sheetN.xml). Cell XML is indexed on first use. */
export class Sheet {
	private rows = new Map<number, Element>();
	private cells = new Map<string, CellEntry>();
	private sharedMasters = new Map<string, SharedMaster>();
	private indexed = false;
	private rules?: ValidationRule[];
	private columnList?: { min: number; max: number; info: ColumnInfo }[];
	private defaultWidth = DEFAULT_COLUMN_WIDTH;

	constructor(
		readonly name: string,
		private path: string,
		private ctx: SheetContext,
		readonly tables: Table[],
	) {}

	private writeBlockedRanges?: Range[];
	private links?: { range: Range; target: string }[];

	private get root(): Element {
		return this.ctx.pkg.xml(this.path)!.documentElement;
	}

	get partPath(): string { return this.path; }

	/** Frozen leading rows/columns; normal split panes are not frozen. */
	frozenPane(): { rows: number; columns: number } {
		const views = child(this.root, "sheetViews");
		const view = views && child(views, "sheetView");
		const pane = view && child(view, "pane");
		if (!pane || !["frozen", "frozenSplit"].includes(pane.getAttribute("state") ?? "")) return { rows: 0, columns: 0 };
		return { rows: Math.max(0, Number(pane.getAttribute("ySplit") ?? 0)), columns: Math.max(0, Number(pane.getAttribute("xSplit") ?? 0)) };
	}

	/** Materialize shared formulas before coordinates move, then drop XML indexes. */
	prepareStructureEdit() {
		this.index();
		for (const si of [...this.sharedMasters.keys()]) this.unshare(si);
	}

	invalidate() {
		this.rows.clear(); this.cells.clear(); this.sharedMasters.clear();
		this.indexed = false; this.rules = undefined; this.columnList = undefined; this.writeBlockedRanges = undefined; this.links = undefined;
	}

	private get sheetData(): Element {
		return child(this.root, "sheetData") ?? createChild(this.root, "sheetData");
	}

	private index() {
		if (this.indexed) return;
		this.indexed = true;
		let rowNum = 0;
		for (const rowEl of elements(this.sheetData, "row")) {
			const r = rowEl.getAttribute("r");
			rowNum = r ? Number(r) : rowNum + 1;
			rowEl.setAttribute("r", String(rowNum));
			this.rows.set(rowNum, rowEl);

			let colNum = 0;
			for (const el of elements(rowEl, "c")) {
				const ref = el.getAttribute("r");
				colNum = ref ? parseRef(ref).col : colNum + 1;
				el.setAttribute("r", formatRef(rowNum, colNum));
				this.cells.set(key(rowNum, colNum), { row: rowNum, col: colNum, el });

				const f = child(el, "f");
				const si = f?.getAttribute("si");
				if (f && si !== null && si !== undefined && f.getAttribute("t") === "shared" && textOf(f)) {
					this.sharedMasters.set(si, { formula: textOf(f), row: rowNum, col: colNum });
				}
			}
		}
	}

	hyperlink(row: number, col: number): string | undefined {
		if (!this.links) {
			this.links = [];
			const split = this.partPath.lastIndexOf("/");
			const rels = this.ctx.pkg.xml(`${this.partPath.slice(0, split)}/_rels/${this.partPath.slice(split + 1)}.rels`);
			const targets = new Map((rels ? elements(rels.documentElement, "Relationship") : []).filter((rel) => rel.getAttribute("TargetMode") === "External").map((rel) => [rel.getAttribute("Id"), rel.getAttribute("Target") ?? ""]));
			const links = child(this.root, "hyperlinks");
			for (const link of links ? elements(links, "hyperlink") : []) {
				const target = safeLink(targets.get(link.getAttribute("r:id")) ?? "");
				if (target && link.getAttribute("ref")) this.links.push({ range: parseRange(link.getAttribute("ref")!), target });
			}
		}
		return this.links.find((link) => inRange(link.range, row, col))?.target;
	}

	cell(row: number, col: number): CellInfo {
		this.index();
		const el = this.cells.get(key(row, col))?.el;
		const style = this.styleOf(el, row, col);
		return {
			value: el ? this.readValue(el, style) : null,
			formula: el ? this.readFormula(el, row, col) : undefined,
			format: this.ctx.styles.formatCode(style),
			look: this.ctx.styles.cellFormat(style),
		};
	}

	/** A cell's own style, else its row's, else its column's (how empty cells get formatted). */
	private styleOf(el: Element | undefined, row: number, col: number): number {
		const own = el?.getAttribute("s");
		if (own) return Number(own);
		const rowEl = this.rows.get(row);
		if (rowEl?.getAttribute("customFormat") === "1" && rowEl.getAttribute("s")) return Number(rowEl.getAttribute("s"));
		return this.column(col).style ?? 0;
	}

	column(col: number): ColumnInfo {
		if (!this.columnList) {
			const format = child(this.root, "sheetFormatPr");
			const base = format?.getAttribute("baseColWidth");
			this.defaultWidth = Number(format?.getAttribute("defaultColWidth") ?? (base ? Number(base) + 0.71 : DEFAULT_COLUMN_WIDTH));
			const cols = child(this.root, "cols");
			this.columnList = (cols ? elements(cols, "col") : []).map((c) => ({
				min: Number(c.getAttribute("min")),
				max: Number(c.getAttribute("max")),
				info: {
					width: c.getAttribute("width") ? Number(c.getAttribute("width")) : this.defaultWidth,
					hidden: c.getAttribute("hidden") === "1" || c.getAttribute("hidden") === "true",
					style: c.getAttribute("style") ? Number(c.getAttribute("style")) : undefined,
				},
			}));
		}
		return this.columnList.find((c) => col >= c.min && col <= c.max)?.info ?? { width: this.defaultWidth, hidden: false };
	}

	isRowHidden(row: number): boolean {
		this.index();
		const hidden = this.rows.get(row)?.getAttribute("hidden");
		return hidden === "1" || hidden === "true";
	}

	/** Whether the cell holds anything (a value or a formula), ignoring formatting. */
	hasContent(row: number, col: number): boolean {
		this.index();
		const el = this.cells.get(key(row, col))?.el;
		return !!el && (!!child(el, "v") || !!child(el, "f") || !!child(el, "is"));
	}

	/** From A1 to the furthest cell with content, or null for an empty sheet. */
	usedRange(): Range | null {
		this.index();
		let bottom = 0;
		let right = 0;
		for (const { row, col } of this.cells.values()) {
			if (!this.hasContent(row, col)) continue;
			bottom = Math.max(bottom, row);
			right = Math.max(right, col);
		}
		return bottom ? { top: 1, left: 1, bottom, right } : null;
	}

	/** Includes explicitly inserted blank rows/columns, unlike the content range. */
	displayRange(): Range | null {
		const ref = child(this.root, "dimension")?.getAttribute("ref");
		const used = this.usedRange();
		if (!ref) return used;
		const dimension = parseRange(ref);
		return { top: 1, left: 1, bottom: Math.max(used?.bottom ?? 1, dimension.bottom), right: Math.max(used?.right ?? 1, dimension.right) };
	}

	mergedRanges(): Range[] {
		const list = child(this.root, "mergeCells");
		return (list ? elements(list, "mergeCell") : []).map((m) => parseRange(m.getAttribute("ref") ?? "A1"));
	}

	private readValue(el: Element, style: number): CellValue {
		const v = child(el, "v");
		const raw = v ? textOf(v) : undefined;
		switch (el.getAttribute("t")) {
			case "s":
				return raw === undefined ? null : (this.ctx.sharedStrings[Number(raw)] ?? "");
			case "inlineStr": {
				const is = child(el, "is");
				return is ? richText(is) : "";
			}
			case "str":
				return raw ?? "";
			case "b":
				return raw === undefined ? null : raw === "1";
			case "e":
				return raw ? { error: raw } : null;
			case "d":
				return raw ? new Date(raw) : null;
			default: {
				if (raw === undefined || raw === "") return null;
				const n = Number(raw);
				return this.ctx.styles.isDate(style) ? serialToDate(n, this.ctx.date1904) : n;
			}
		}
	}

	private readFormula(el: Element, row: number, col: number): string | undefined {
		const f = child(el, "f");
		if (!f) return undefined;
		const own = textOf(f);
		if (own) return own;
		// A shared-formula child cell: derive its formula from the group's first cell.
		const master = this.sharedMasters.get(f.getAttribute("si") ?? "");
		return master ? shiftFormula(master.formula, row - master.row, col - master.col) : undefined;
	}

	validations(): ValidationRule[] {
		return (this.rules ??= readValidations(this.root));
	}

	validationAt(row: number, col: number): ValidationRule | undefined {
		return this.validations().find((rule) => rule.ranges.some((r) => inRange(r, row, col)));
	}

	/** Grows rules that end on the row above down to `row`, as Excel does when a table grows. */
	extendValidations(row: number, left: number, right: number) {
		for (const rule of this.validations()) {
			let changed = false;
			const ranges = rule.ranges.map((r) => {
				if (r.bottom !== row - 1 || r.left > right || r.right < left) return r;
				changed = true;
				return { ...r, bottom: row };
			});
			if (changed) {
				rule.setRanges(ranges);
				this.markDirty();
			}
		}
	}

	isTableHeader(row: number, col: number): boolean {
		return this.tables.some((t) => t.headerRows > 0 && row === t.ref.top && inRange(t.ref, row, col));
	}

	/** Refuse cell writes whose semantics cannot be preserved safely. */
	assertWritable(row: number, col: number) {
		if (child(this.root, "sheetProtection")) throw new Error("工作表已受保护，请先在 Excel 中解除保护。");
		this.index();
		if (!this.writeBlockedRanges) {
			this.writeBlockedRanges = [];
			for (const entry of this.cells.values()) {
				const f = child(entry.el, "f");
				if (f && ["array", "dataTable"].includes(f.getAttribute("t") ?? "")) this.writeBlockedRanges.push(parseRange(f.getAttribute("ref") || formatRef(entry.row, entry.col)));
			}
		}
		if (this.writeBlockedRanges.some((range) => inRange(range, row, col))) throw new Error("数组公式或数据表区域请在 Excel 中编辑。");
	}

	formatCell(row: number, col: number, patch: CellStylePatch) {
		if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) throw new Error("单元格超出工作表范围。");
		this.assertWritable(row, col);
		const existing = this.cells.get(key(row, col))?.el;
		const original = this.styleOf(existing, row, col);
		const next = this.ctx.styles.formatStyle(original, patch);
		if (next === original) return;
		(existing ?? this.createCell(row, col)).setAttribute("s", String(next));
		this.growDimension(row, col);
		this.markDirty();
	}

	setCell(row: number, col: number, input: CellInput) {
		validateInput(input);
		this.assertWritable(row, col);
		if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) throw new Error(`Cell ${formatRef(row, col)} is outside the sheet.`);
		this.index();
		const k = key(row, col);
		let el = this.cells.get(k)?.el;

		const f = el && child(el, "f");
		if (f?.getAttribute("t") === "shared" && f.getAttribute("ref")) this.unshare(f.getAttribute("si") ?? "");

		if (input.kind === "empty") {
			if (!el) return;
			clearContent(el);
			if (!el.hasAttribute("s")) {
				el.parentNode?.removeChild(el);
				this.cells.delete(k);
			}
			this.markDirty();
			return;
		}

		el ??= this.createCell(row, col);
		clearContent(el);
		switch (input.kind) {
			case "boolean":
				el.setAttribute("t", "b");
				setText(createChild(el, "v"), input.value ? "1" : "0");
				break;
			case "number":
				setText(createChild(el, "v"), String(input.value));
				break;
			case "text": {
				el.setAttribute("t", "inlineStr");
				const t = createChild(createChild(el, "is"), "t");
				setText(t, input.value);
				if (/\r|\n/.test(input.value)) {
					const style = this.styleOf(el, row, col);
					if (!this.ctx.styles.cellFormat(style).wrap) el.setAttribute("s", String(this.ctx.styles.formatStyle(style, { wrap: true })));
				}
				if (/^\s|\s$/.test(input.value)) t.setAttribute("xml:space", "preserve");
				break;
			}
			case "date":
				el.setAttribute("s", String(this.ctx.styles.dateStyle(Number(el.getAttribute("s") ?? 0))));
				setText(createChild(el, "v"), String(dateToSerial(input.value, this.ctx.date1904)));
				break;
			case "formula":
				setText(createChild(el, "f"), input.formula);
				break;
		}
		this.growDimension(row, col);
		this.markDirty();
	}

	/** Creates empty, formatted cells across a row so it looks like the row above. */
	prepareRow(row: number, left: number, right: number) {
		this.index();
		for (let col = left; col <= right; col++) {
			if (!this.cells.has(key(row, col))) this.createCell(row, col);
		}
		this.growDimension(row, right);
		this.markDirty();
	}

	private createCell(row: number, col: number): Element {
		let rowEl = this.rows.get(row);
		if (!rowEl) {
			const next = elements(this.sheetData, "row").find((r) => Number(r.getAttribute("r")) > row) ?? null;
			rowEl = createChild(this.sheetData, "row", next);
			rowEl.setAttribute("r", String(row));
			this.rows.set(row, rowEl);
		}
		// "spans" is only a loading hint and would be wrong once the row gains cells.
		rowEl.removeAttribute("spans");

		const nextCell = elements(rowEl, "c").find((c) => parseRef(c.getAttribute("r")!).col > col) ?? null;
		const el = createChild(rowEl, "c", nextCell);
		el.setAttribute("r", formatRef(row, col));

		// New rows keep the formatting of the row above (dates, currency, ...), except below a header.
		const above = this.cells.get(key(row - 1, col))?.el.getAttribute("s");
		if (above && above !== "0" && !this.isTableHeader(row - 1, col)) el.setAttribute("s", above);

		this.cells.set(key(row, col), { row, col, el });
		return el;
	}

	/** Gives every cell of a shared-formula group its own formula, so its first cell can change. */
	private unshare(si: string) {
		const master = this.sharedMasters.get(si);
		if (!master) return;
		for (const { row, col, el } of this.cells.values()) {
			const f = child(el, "f");
			if (!f || f.getAttribute("t") !== "shared" || f.getAttribute("si") !== si) continue;
			setText(f, shiftFormula(master.formula, row - master.row, col - master.col));
			f.removeAttribute("t");
			f.removeAttribute("si");
			f.removeAttribute("ref");
		}
		this.sharedMasters.delete(si);
	}

	private growDimension(row: number, col: number) {
		const dim = child(this.root, "dimension");
		const ref = dim?.getAttribute("ref");
		if (!dim || !ref) return;
		const r = parseRange(ref);
		const grown = {
			top: Math.min(r.top, row),
			left: Math.min(r.left, col),
			bottom: Math.max(r.bottom, row),
			right: Math.max(r.right, col),
		};
		dim.setAttribute("ref", formatRange(grown));
	}

	private markDirty() {
		this.ctx.pkg.markDirty(this.path);
	}
}

function clearContent(el: Element) {
	for (const attr of ["t", "cm", "vm"]) el.removeAttribute(attr);
	removeChildren(el);
}
