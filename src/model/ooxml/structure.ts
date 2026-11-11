import type { Sheet } from "../sheet";
import type { Package } from "./package";
import { colName, formatRef, MAX_COL, MAX_ROW, parseRange, parseRef } from "./refs";
import { moveIndex, moveInterval, moveReference, rewriteReferences, StructureEdit, validateEdit } from "./structure-refs";
import { child, createChild, elements, setText, textOf } from "./xml";

function descendants(root: Element): Element[] {
	return [root, ...Array.from(root.getElementsByTagName("*"))];
}

function selected(index: number, edit: StructureEdit): boolean {
	return index >= edit.start && index < edit.start + edit.count;
}

/** Reject unsupported dependencies before changing any XML or writing a file. */
function preflight(pkg: Package, sheets: Sheet[], target: Sheet, edit: StructureEdit, workbookPath: string) {
	validateEdit(edit);
	const root = pkg.xml(target.partPath)!.documentElement;
	if (child(root, "sheetProtection")) throw new Error("此工作表受 Excel 保护，请先在 Excel 中解除保护再修改行列。");
	if (edit.kind === "clear") return;
	for (const name of ["drawing", "legacyDrawing", "legacyDrawingHF", "pivotTableParts", "oleObjects", "controls", "protectedRanges", "rowBreaks", "colBreaks", "scenarios", "customSheetViews"]) {
		if (child(root, name)) throw new Error(`此工作表包含 ${name}，暂不支持调整其行列。请在 Excel 中操作。`);
	}
	if (pkg.rels(workbookPath).some((r) => /\/(externalLink|pivotCacheDefinition)$/.test(r.type))) {
		throw new Error("工作簿包含外部链接或数据透视缓存，暂不支持结构编辑。请在 Excel 中操作。");
	}
	for (const sheet of sheets) {
		const doc = pkg.xml(sheet.partPath)!;
		if (pkg.rels(sheet.partPath).some((r) => /\/(comments|threadedComment|queryTable)$/.test(r.type))) {
			throw new Error("工作簿包含批注或查询表，暂不支持结构编辑。请在 Excel 中操作。");
		}
		if (child(doc.documentElement, "drawing") || child(doc.documentElement, "pivotTableParts")) {
			throw new Error("工作簿包含图表/绘图或数据透视表，暂不支持结构编辑。请在 Excel 中操作。");
		}
		for (const el of descendants(doc.documentElement)) {
			if (el.localName === "f" && ["array", "dataTable"].includes(el.getAttribute("t") ?? "")) {
				throw new Error("工作簿包含数组/数据表公式，暂不支持结构编辑。请在 Excel 中操作。");
			}
			if (["f", "formula", "formula1", "formula2"].includes(el.localName)) checkFormula(textOf(el));
		}
		for (const table of sheet.tables) {
			for (const el of descendants(pkg.xml(table.path)!.documentElement)) {
				if (/Formula$/.test(el.localName)) checkFormula(textOf(el));
			}
		}
	}
	const names = child(pkg.xml(workbookPath)!.documentElement, "definedNames");
	if (pkg.rels(workbookPath).some((r) => /\/chartsheet$/.test(r.type))) throw new Error("工作簿包含图表工作表，暂不支持结构编辑。请在 Excel 中操作。");
	for (const name of names ? elements(names, "definedName") : []) checkFormula(textOf(name));

	// Inserting must not push even formatted/hidden cells or ranges past Excel's edge.
	if (edit.kind === "insert") {
		const limit = edit.axis === "row" ? MAX_ROW : MAX_COL;
		for (const el of descendants(root)) {
			if (el.localName === "c" && moveReference(el.getAttribute("r") ?? "", edit) === "#REF!") throw new Error("插入将超出 Excel 行列上限。");
			for (const attr of ["ref", "sqref"]) {
				const value = el.getAttribute(attr);
				if (value && value.split(/\s+/).some((r) => moveReference(r, edit) === "#REF!")) throw new Error("插入将超出 Excel 行列上限。");
			}
			if ((edit.axis === "row" && el.localName === "row" && Number(el.getAttribute("r")) + edit.count > limit && Number(el.getAttribute("r")) >= edit.start)
				|| (edit.axis === "column" && el.localName === "col" && Number(el.getAttribute("max")) + edit.count > limit && Number(el.getAttribute("max")) >= edit.start)) {
				throw new Error("插入将超出 Excel 行列上限。");
			}
		}
	}
	for (const table of target.tables) {
		const r = table.ref;
		if (edit.kind === "delete" && edit.axis === "row") {
			if (selected(r.top, edit) || (table.totalsRows && selected(r.bottom, edit))) throw new Error("暂不能删除 Excel 表格对象的标题行或汇总行。");
			if (edit.start <= table.dataRange.top && edit.start + edit.count - 1 >= table.dataRange.bottom) throw new Error("Excel 表格至少保留一行数据。");
		}
		if (edit.kind === "delete" && edit.axis === "column" && edit.start <= r.right && edit.start + edit.count > r.left) {
			if (edit.start <= r.left && edit.start + edit.count > r.right) throw new Error("不能删除 Excel 表格对象的全部列。");
			// Complex structured references can name deleted columns without A1 refs.
			const deleted = table.columns.filter((_, i) => selected(r.left + i, edit)).map((c) => c.name);
			for (const sheet of sheets) {
				const docs = [pkg.xml(sheet.partPath)!, ...sheet.tables.map((t) => pkg.xml(t.path)!)];
				for (const doc of docs) for (const el of descendants(doc.documentElement)) {
					if (["f", "formula", "formula1", "formula2", "calculatedColumnFormula", "totalsRowFormula"].includes(el.localName)
						&& textOf(el).includes("[") && deleted.some((name) => textOf(el).includes(`[${name}]`))) {
						throw new Error("要删除的表格列被结构化公式引用，请先在 Excel 中处理这些引用。");
					}
				}
			}
			for (const name of names ? elements(names, "definedName") : []) {
				if (deleted.some((column) => textOf(name).includes(`[${column}]`))) throw new Error("要删除的表格列被命名范围引用，请在 Excel 中操作。");
			}
		}
	}
}

function checkFormula(formula: string) {
	// INDIRECT encodes references inside strings; external and 3D refs need a
	// different resolver. Refuse instead of silently keeping stale coordinates.
	const bare = formula.replace(/\bINDIRECT\s*\(\s*"[^"\[]+\[[^"\]]+\]"\s*\)/gi, "0").replace(/"(?:[^"]|"")*"/g, '""');
	if (/\b(?:INDIRECT|OFFSET)\s*\(/i.test(bare) || /(?:\[[^\]]+\][^!]*!|(?:'[^']*:[^']*'|[\w\p{L}]+:[\w\p{L}]+)!)/u.test(bare)) {
		throw new Error("工作簿包含 INDIRECT/OFFSET、外部或三维引用，暂不支持结构编辑。请在 Excel 中操作。");
	}
}

/** Whole-worksheet insert/delete/clear with dependent references updated together. */
export function editStructure(pkg: Package, workbookPath: string, sheets: Sheet[], target: Sheet, edit: StructureEdit) {
	preflight(pkg, sheets, target, edit, workbookPath);
	const root = pkg.xml(target.partPath)!.documentElement;
	if (edit.kind === "clear") {
		const data = child(root, "sheetData"); if (!data) return;
		for (const row of elements(data, "row")) for (const c of elements(row, "c")) {
			const ref = parseRef(c.getAttribute("r")!);
			if (selected(edit.axis === "row" ? ref.row : ref.col, edit) && !target.isTableHeader(ref.row, ref.col)) target.setCell(ref.row, ref.col, { kind: "empty" });
		}
		return;
	}
	for (const sheet of sheets) sheet.prepareStructureEdit();
	const frozen = target.frozenPane();
	const sheetData = child(root, "sheetData")!;
	// Save styles from the preceding line before moving coordinates.
	const templateRow = elements(sheetData, "row").find((r) => Number(r.getAttribute("r")) === Math.max(1, edit.start - 1));
	const templateCells = templateRow ? elements(templateRow, "c").map((c) => ({ col: parseRef(c.getAttribute("r")!).col, style: c.getAttribute("s") })) : [];
	for (const row of elements(sheetData, "row")) {
		const oldRow = Number(row.getAttribute("r"));
		const nextRow = edit.axis === "row" ? moveIndex(oldRow, edit) : oldRow;
		if (nextRow === null) { row.parentNode!.removeChild(row); continue; }
		row.setAttribute("r", String(nextRow)); row.removeAttribute("spans");
		for (const c of elements(row, "c")) {
			const ref = parseRef(c.getAttribute("r")!);
			const col = edit.axis === "column" ? moveIndex(ref.col, edit) : ref.col;
			if (col === null) c.parentNode!.removeChild(c);
			else c.setAttribute("r", formatRef(nextRow, col));
		}
	}
	if (edit.kind === "insert" && edit.axis === "row") {
		const next = elements(sheetData, "row").find((r) => Number(r.getAttribute("r")) >= edit.start) ?? null;
		for (let i = 0; i < edit.count; i++) {
			const row = createChild(sheetData, "row", next); row.setAttribute("r", String(edit.start + i));
			for (const attr of ["s", "customFormat", "ht", "customHeight"]) {
				const v = templateRow?.getAttribute(attr); if (v !== null && v !== undefined) row.setAttribute(attr, v);
			}
			for (const cell of templateCells) if (cell.style && !target.isTableHeader(Math.max(1, edit.start - 1), cell.col)) {
				const c = createChild(row, "c"); c.setAttribute("r", formatRef(edit.start + i, cell.col)); c.setAttribute("s", cell.style);
			}
		}
	}
	moveColumns(root, edit);
	moveRanges(root, edit);
	if (edit.kind === "insert") {
		const dim = child(root, "dimension") ?? createChild(root, "dimension", child(root, "sheetViews") ?? child(root, "sheetFormatPr") ?? child(root, "sheetData") ?? null);
		const r = parseRange(dim.getAttribute("ref") ?? "A1");
		if (edit.axis === "row") r.bottom = Math.max(r.bottom, edit.start + edit.count - 1);
		else r.right = Math.max(r.right, edit.start + edit.count - 1);
		dim.setAttribute("ref", `${formatRef(r.top, r.left)}:${formatRef(r.bottom, r.right)}`);
	}
	for (const table of target.tables) moveTable(pkg.xml(table.path)!.documentElement, root, edit);
	// All formula owners (including other worksheets and names) refer to the
	// same original coordinates and must be rewritten exactly once.
	for (const sheet of sheets) {
		for (const path of [sheet.partPath, ...sheet.tables.map((t) => t.path)]) {
			for (const el of descendants(pkg.xml(path)!.documentElement)) {
				const location = el.localName === "hyperlink" ? el.getAttribute("location") : null;
				if (location) {
					const next = rewriteReferences(location, sheet.name, target.name, edit);
					if (next !== location) { el.setAttribute("location", next); pkg.markDirty(path); }
				}
				if (["f", "formula", "formula1", "formula2", "calculatedColumnFormula", "totalsRowFormula"].includes(el.localName) && !elements(el).length) {
					const old = textOf(el), next = rewriteReferences(old, sheet.name, target.name, edit);
					if (old !== next) { setText(el, next); pkg.markDirty(path); }
				}
			}
		}
	}
	const names = child(pkg.xml(workbookPath)!.documentElement, "definedNames");
	const sheetList = child(pkg.xml(workbookPath)!.documentElement, "sheets");
	const sheetNames = sheetList ? elements(sheetList, "sheet").map((s) => s.getAttribute("name") ?? "") : [];
	for (const name of names ? elements(names, "definedName") : []) {
		const scope = name.getAttribute("localSheetId");
		setText(name, rewriteReferences(textOf(name), scope === null ? undefined : sheetNames[Number(scope)], target.name, edit));
	}
	if (names) pkg.markDirty(workbookPath);
	for (const pane of descendants(root).filter((e) => e.localName === "pane" && ["frozen", "frozenSplit"].includes(e.getAttribute("state") ?? ""))) {
		let count = edit.axis === "row" ? frozen.rows : frozen.columns;
		if (edit.start <= count) count += edit.kind === "insert" ? edit.count : -Math.min(edit.count, count - edit.start + 1);
		const attr = edit.axis === "row" ? "ySplit" : "xSplit";
		if (count) pane.setAttribute(attr, String(count)); else pane.removeAttribute(attr);
		if (!Number(pane.getAttribute("xSplit")) && !Number(pane.getAttribute("ySplit"))) pane.parentNode!.removeChild(pane);
	}
	for (const sheet of sheets) {
		sheet.invalidate();
		if (sheet === target) { pkg.markDirty(sheet.partPath); for (const table of sheet.tables) pkg.markDirty(table.path); }
	}
}

function moveColumns(root: Element, edit: StructureEdit) {
	if (edit.axis !== "column") return;
	const cols = child(root, "cols"); if (!cols) return;
	const template = elements(cols, "col").find((c) => Number(c.getAttribute("min")) <= Math.max(1, edit.start - 1) && Number(c.getAttribute("max")) >= Math.max(1, edit.start - 1))?.cloneNode(true) as Element | undefined;
	for (const col of elements(cols, "col")) {
		const next = moveInterval(Number(col.getAttribute("min")), Number(col.getAttribute("max")), edit);
		if (!next) col.parentNode!.removeChild(col);
		else { col.setAttribute("min", String(next[0])); col.setAttribute("max", String(next[1])); }
	}
	if (edit.kind === "insert" && template && !elements(cols, "col").some((c) => Number(c.getAttribute("min")) <= edit.start && Number(c.getAttribute("max")) >= edit.start)) {
		template.setAttribute("min", String(edit.start)); template.setAttribute("max", String(edit.start + edit.count - 1)); template.removeAttribute("hidden");
		cols.insertBefore(template, elements(cols, "col").find((c) => Number(c.getAttribute("min")) > edit.start) ?? null);
	}
	for (const row of elements(child(root, "sheetData")!, "row")) if (edit.kind === "insert") {
		const source = elements(row, "c").find((c) => parseRef(c.getAttribute("r")!).col === Math.max(1, edit.start - 1));
		const style = source?.getAttribute("s");
		if (!style) continue;
		const next = elements(row, "c").find((c) => parseRef(c.getAttribute("r")!).col >= edit.start) ?? null;
		for (let i = 0; i < edit.count; i++) {
			const c = createChild(row, "c", next); c.setAttribute("r", formatRef(Number(row.getAttribute("r")), edit.start + i)); c.setAttribute("s", style);
		}
	}
}

function moveRanges(root: Element, edit: StructureEdit) {
	// autoFilter column ids are relative to its original left edge.
	for (const filter of descendants(root).filter((e) => e.localName === "autoFilter")) {
		const ref = filter.getAttribute("ref");
		if (!ref || edit.axis !== "column") continue;
		const left = parseRange(ref).left;
		const newRef = moveReference(ref, edit);
		if (newRef === "#REF!") continue;
		const newLeft = parseRange(newRef).left;
		for (const col of elements(filter, "filterColumn")) {
			const next = moveIndex(left + Number(col.getAttribute("colId")), edit);
			if (next === null) filter.removeChild(col); else col.setAttribute("colId", String(next - newLeft));
		}
	}
	for (const el of descendants(root)) {
		if (!el.parentNode) continue;
		for (const attr of ["ref", "sqref", "activeCell", "topLeftCell"]) {
			if (el.localName === "c" && attr === "ref") continue;
			const value = el.getAttribute(attr); if (!value) continue;
			const next = value.split(/\s+/).map((ref) => moveReference(ref, edit)).filter((ref) => ref !== "#REF!").join(" ");
			if (next) el.setAttribute(attr, next);
			else if (["pane", "selection", "sheetView"].includes(el.localName)) {
				const ref = parseRef(value.split(/\s+/)[0]);
				if (edit.axis === "row") ref.row = moveIndex(ref.row, edit) ?? edit.start;
				else ref.col = moveIndex(ref.col, edit) ?? edit.start;
				el.setAttribute(attr, formatRef(ref.row, ref.col));
			}
			else if (el.localName === "dimension") el.setAttribute(attr, "A1");
			else el.parentNode?.removeChild(el);
		}
		if (el.localName === "sqref" && !el.hasAttribute("sqref")) {
			const next = textOf(el).split(/\s+/).map((r) => moveReference(r, edit)).filter((r) => r !== "#REF!").join(" ");
			if (next) setText(el, next); else el.parentNode?.parentNode?.removeChild(el.parentNode);
		}
	}
	for (const el of descendants(root)) {
		if (["mergeCells", "dataValidations"].includes(el.localName)) el.setAttribute("count", String(elements(el).length));
	}
}

function moveTable(root: Element, sheetRoot: Element, edit: StructureEdit) {
	const before = parseRange(root.getAttribute("ref")!);
	const list = child(root, "tableColumns")!;
	if (edit.axis === "column") {
		if (edit.kind === "insert" && edit.start > before.left && edit.start <= before.right) {
			const existing = elements(list, "tableColumn");
			const names = new Set(existing.map((e) => e.getAttribute("name")));
			let id = Math.max(0, ...existing.map((e) => Number(e.getAttribute("id"))));
			for (let i = 0; i < edit.count; i++) {
				let n = 1; while (names.has(`Column${n}`)) n++;
				const name = `Column${n}`; names.add(name);
				const c = createChild(list, "tableColumn", existing[edit.start - before.left] ?? null);
				c.setAttribute("id", String(++id)); c.setAttribute("name", name);
				if (Number(root.getAttribute("headerRowCount") ?? 1)) {
					const row = elements(child(sheetRoot, "sheetData")!, "row").find((r) => Number(r.getAttribute("r")) === before.top)!;
					const ref = formatRef(before.top, edit.start + i);
					let cell = elements(row, "c").find((c) => c.getAttribute("r") === ref);
					if (!cell) cell = createChild(row, "c", elements(row, "c").find((c) => parseRef(c.getAttribute("r")!).col > edit.start + i) ?? null);
					cell.setAttribute("r", ref); cell.setAttribute("t", "inlineStr"); setText(createChild(createChild(cell, "is"), "t"), name);
				}
			}
		} else if (edit.kind === "delete") {
			elements(list, "tableColumn").forEach((col, i) => { if (selected(before.left + i, edit)) list.removeChild(col); });
		}
		list.setAttribute("count", String(elements(list).length));
	}
	moveRanges(root, edit);
}
