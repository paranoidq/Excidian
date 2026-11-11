import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { Package } from "../src/model/ooxml/package";
import { moveReference, rewriteReferences, StructureEdit } from "../src/model/ooxml/structure-refs";
import { child, createChild, elements } from "../src/model/ooxml/xml";
import { Workbook } from "../src/model/workbook";
import { restoreWorkbook, updateWorkbook, workbookHistory } from "../src/model/workbook-store";
import type { StoredFile } from "../src/model/storage";
import { financeWorkbook } from "./fixture";

const insert: StructureEdit = { axis: "row", kind: "insert", start: 3, count: 2 };
const remove: StructureEdit = { axis: "row", kind: "delete", start: 3, count: 2 };

describe("structural references", () => {
	it("moves absolute and relative refs, contracts ranges and emits REF for deleted cells", () => {
		assert.equal(moveReference("$A$3:B9", insert), "$A$5:B11");
		assert.equal(moveReference("A1:A9", remove), "A1:A7");
		assert.equal(moveReference("A3:A4", remove), "#REF!");
		assert.equal(moveReference("A4:A8", remove), "A3:A6");
		assert.equal(moveReference("$3:$9", insert), "$5:$11");
		assert.equal(moveReference("$C:$F", { ...insert, axis: "column" }), "$E:$H");
	});
	it("resolves cross-sheet refs and leaves strings, functions, names and structured refs alone", () => {
		const f = `SUM($A$3:A8)+'销售 表'!B3+Other!C3+LOG10(A3)+"A3"+Table1[A3]+A3_name`;
		assert.equal(rewriteReferences(f, "销售 表", "销售 表", insert), `SUM($A$5:A10)+'销售 表'!B5+Other!C3+LOG10(A5)+"A3"+Table1[A3]+A3_name`);
		assert.equal(rewriteReferences("A3+Sheet1!A3", "Other", "Sheet1", remove), "A3+Sheet1!#REF!");
	});
});

describe("worksheet structure", () => {
	it("inserts multiple data rows, adjusts shared formulas and tables, and survives save/load", async () => {
		let wb = await Workbook.load(await financeWorkbook());
		wb.sheet("Riepilogo").setCell(8, 1, { kind: "formula", formula: "'2026'!$C$3+'2026'!C2" });
		wb.editStructure("2026", insert);
		wb = await Workbook.load(await wb.save());
		assert.equal(wb.sheet("2026").cell(5, 3).value, -45);
		assert.equal(wb.sheet("2026").cell(5, 5).formula, "SUM($C$2:C5)");
		assert.equal(wb.sheet("2026").cell(3, 5).formula, "SUM($C$2:C3)");
		assert.equal(wb.table("Movimenti").table.ref.bottom, 5);
		assert.equal(wb.sheet("Riepilogo").cell(8, 1).formula, "'2026'!$C$5+'2026'!C2");
		assert.equal(wb.sheet("2026").validationAt(5, 3)?.type, "decimal");
	});

	it("deletes a data row, preserves styles and shrinks dependent ranges", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		wb.editStructure("2026", { axis: "row", kind: "delete", start: 2, count: 1 });
		assert.equal(wb.sheet("2026").cell(2, 3).value, -45);
		assert.equal(wb.sheet("2026").cell(2, 5).formula, "SUM($C$2:C2)");
		assert.equal(wb.table("Movimenti").table.ref.bottom, 2);
		assert.match(wb.sheet("2026").cell(2, 3).format, /€/);
	});

	it("inserts/deletes worksheet columns and updates table column metadata and names", async () => {
		let wb = await Workbook.load(await financeWorkbook());
		wb.editStructure("2026", { axis: "column", kind: "insert", start: 3, count: 2 });
		assert.equal(wb.table("Movimenti").table.ref.right, 7);
		assert.deepEqual(wb.table("Movimenti").table.columns.map((c) => c.name), ["Mese", "Data", "Column1", "Column2", "Importo", "Categoria", "Saldo"]);
		assert.equal(wb.sheet("2026").cell(1, 3).value, "Column1");
		assert.equal(wb.sheet("2026").cell(2, 5).value, 1500);
		assert.equal(wb.sheet("2026").cell(2, 7).formula, "SUM($E$2:E2)");
		wb = await Workbook.load(await wb.save());
		wb.editStructure("2026", { axis: "column", kind: "delete", start: 3, count: 2 });
		assert.equal(wb.table("Movimenti").table.columns.length, 5);
		assert.equal(wb.sheet("2026").cell(2, 3).value, 1500);
	});

	it("moves merges, hidden columns, frozen boundaries and cross-sheet validation sources", async () => {
		const pkg = await Package.load(await financeWorkbook());
		const root = pkg.xml("xl/worksheets/sheet1.xml")!.documentElement;
		const views = createChild(root, "sheetViews", child(root, "cols")!);
		const pane = createChild(createChild(views, "sheetView"), "pane");
		pane.setAttribute("state", "frozen"); pane.setAttribute("xSplit", "2"); pane.setAttribute("ySplit", "3"); pane.setAttribute("topLeftCell", "C4");
		pkg.markDirty("xl/worksheets/sheet1.xml");
		let wb = await Workbook.load(await pkg.toBytes());
		assert.deepEqual(wb.sheet().frozenPane(), { rows: 3, columns: 2 });
		wb.editStructure("Riepilogo", { axis: "row", kind: "insert", start: 2, count: 1 });
		assert.deepEqual(wb.sheet().frozenPane(), { rows: 4, columns: 2 });
		assert.equal(wb.sheet().mergedRanges()[0].top, 4);
		wb.editStructure("Riepilogo", { axis: "column", kind: "delete", start: 1, count: 1 });
		assert.deepEqual(wb.sheet().frozenPane(), { rows: 4, columns: 1 });
		assert.equal(wb.sheet().column(2).hidden, true);
		wb = await Workbook.load(await wb.save());
		assert.deepEqual(wb.sheet().frozenPane(), { rows: 4, columns: 1 });
		wb.editStructure("Liste", { axis: "row", kind: "insert", start: 3, count: 1 });
		assert.equal(wb.sheet("2026").validationAt(2, 4)?.formula1, "Liste!$A$2:$A$7");
	});

	it("clears content without shifting cells, discarding formatting or renaming table headers", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		wb.editStructure("2026", { axis: "column", kind: "clear", start: 3, count: 1 });
		assert.equal(wb.sheet("2026").cell(1, 3).value, "Importo");
		assert.equal(wb.sheet("2026").cell(2, 3).value, null);
		assert.equal(wb.sheet("2026").cell(3, 4).value, "Spesa");
		assert.match(wb.sheet("2026").cell(2, 3).format, /€/);
	});

	it("blocks unsupported or destructive table boundaries before mutating XML", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		assert.throws(() => wb.editStructure("2026", { axis: "row", kind: "delete", start: 1, count: 1 }), /标题行/);
		assert.throws(() => wb.editStructure("2026", { axis: "column", kind: "delete", start: 3, count: 1 }), /结构化公式/);
		assert.equal(wb.sheet("2026").cell(2, 3).value, 1500);
	});

	it("rejects overflow even when the workbook dimension is stale", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		wb.sheet().setCell(20, 16384, { kind: "text", value: "edge" });
		assert.throws(() => wb.editStructure("Riepilogo", { axis: "column", kind: "insert", start: 1, count: 1 }), /上限/);
		assert.equal(wb.sheet().cell(20, 16384).value, "edge");
	});

	it("preserves blank inserted extent and internal hyperlinks after a round trip", async () => {
		const pkg = await Package.load(await financeWorkbook());
		const root = pkg.xml("xl/worksheets/sheet1.xml")!.documentElement;
		const link = createChild(createChild(root, "hyperlinks"), "hyperlink");
		link.setAttribute("ref", "A1"); link.setAttribute("location", "'2026'!C3");
		pkg.markDirty("xl/worksheets/sheet1.xml");
		let wb = await Workbook.load(await pkg.toBytes());
		wb.editStructure("2026", { axis: "row", kind: "insert", start: 3, count: 1 });
		wb.editStructure("Riepilogo", { axis: "column", kind: "insert", start: 5, count: 1 });
		const saved = await wb.save();
		wb = await Workbook.load(saved);
		assert.equal(wb.sheet().displayRange()!.right, 5);
		const again = await Package.load(saved);
		assert.equal(elements(child(again.xml("xl/worksheets/sheet1.xml")!.documentElement, "hyperlinks")!, "hyperlink")[0].getAttribute("location"), "'2026'!C4");
	});

	it("does not write anything when a workbook has unsupported drawing dependencies", async () => {
		const pkg = await Package.load(await financeWorkbook());
		createChild(pkg.xml("xl/worksheets/sheet1.xml")!.documentElement, "drawing");
		pkg.markDirty("xl/worksheets/sheet1.xml");
		const bytes = await pkg.toBytes(); let writes = 0;
		const file: StoredFile = { id: "blocked-test", name: "blocked.xlsx", displayPath: "blocked.xlsx", read: async () => bytes,
			write: async () => { writes++; }, writeSibling: async () => { writes++; }, siblings: async () => [], watch: () => () => {} };
		await assert.rejects(updateWorkbook(file, (wb) => wb.editStructure("2026", insert)), /图表/);
		assert.equal(writes, 0);
	});
});

describe("workbook undo/redo", () => {
	it("restores exact bytes and refuses to overwrite a later external edit", async () => {
		let bytes = await financeWorkbook();
		const original = bytes.slice();
		const file: StoredFile = {
			id: "history-test", name: "demo.xlsx", displayPath: "demo.xlsx", read: async () => bytes,
			write: async (data) => { bytes = data; }, writeSibling: async () => {}, siblings: async () => [], watch: () => () => {},
		};
		await updateWorkbook(file, (wb) => wb.editStructure("2026", insert));
		assert.deepEqual(workbookHistory(file), { undo: true, redo: false });
		const edited = bytes.slice();
		await restoreWorkbook(file, "undo"); assert.deepEqual(bytes, original);
		await restoreWorkbook(file, "redo"); assert.deepEqual(bytes, edited);
		bytes = await financeWorkbook({ totalsRow: true });
		await assert.rejects(restoreWorkbook(file, "undo"), /文件已被/);
		assert.deepEqual(workbookHistory(file), { undo: false, redo: false });
	});
});
