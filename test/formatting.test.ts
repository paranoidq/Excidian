import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { Package } from "../src/model/ooxml/package";
import { child, createChild, elements } from "../src/model/ooxml/xml";
import { Workbook } from "../src/model/workbook";
import { buildGrid } from "../src/controller/grid-builder";
import { parseBlockConfig } from "../src/controller/block-config";
import { financeWorkbook, readEntry, richTextWorkbook } from "./fixture";

describe("saved spreadsheet presentation", () => {
	it("keeps the saved row order and excludes Excel-hidden rows in both sheet and table views", async () => {
		const path = "xl/worksheets/sheet2.xml";
		const pkg = await Package.load(await financeWorkbook());
		const root = pkg.xml(path)!.documentElement;
		const table = pkg.xml("xl/tables/table1.xml")!.documentElement;
		// Sort-state metadata is descriptive. Preview uses the already persisted values/order.
		createChild(child(table, "autoFilter")!, "sortState").setAttribute("ref", "C2:C3");
		pkg.markDirty("xl/tables/table1.xml");
		let wb = await Workbook.load(await pkg.toBytes());
		const cfg = parseBlockConfig("file: demo.xlsx\ntable: Movimenti");
		assert.deepEqual(buildGrid(wb, cfg, true).rows.map((row) => [row.label, row.cells[2].editText]), [["2", "1500"], ["3", "-45"]]);
		const row2 = elements(child(root, "sheetData")!, "row").find((row) => row.getAttribute("r") === "2")!;
		row2.setAttribute("hidden", "1"); pkg.markDirty(path);
		wb = await Workbook.load(await pkg.toBytes());
		assert.deepEqual(buildGrid(wb, cfg, true).rows.map((row) => row.label), ["3"]);
		const sheet = buildGrid(wb, parseBlockConfig("file: demo.xlsx\nsheet: 2026"), true);
		assert.deepEqual(sheet.rows.map((row) => row.label), ["1", "3"]);
	});
});

describe("character formatting in spreadsheet text", () => {
	it("keeps shared and inline runs separate from the cell font and preserves the plain text", async () => {
		const wb = await Workbook.load(await richTextWorkbook());
		const config = parseBlockConfig("file: demo.xlsx\nsheet: Liste\nrange: A8:B8");
		const [shared, inline] = buildGrid(wb, config, true).rows[0].cells;
		for (const cell of [shared, inline]) {
			assert.equal(cell.text, "标题\n普通正文 蓝色下划线继承整格样式");
			assert.equal(cell.editText, cell.text);
			assert.equal(cell.style?.wrap, true);
			assert.deepEqual(cell.runs?.map(({ text, font }) => [text, font?.color, !!font?.bold, !!font?.underline]), [
				["标题", "#FF0000", true, false],
				["\n普通正文 ", undefined, false, false],
				["蓝色下划线", "#0000FF", false, true],
				["继承整格样式", "#D00000", true, true],
			]);
		}
	});

	it("applies a whole-cell font change without affecting other cells sharing the string", async () => {
		const original = await richTextWorkbook();
		const wb = await Workbook.load(original);
		const sheet = wb.sheet("Liste");
		const before = sheet.cell(8, 3);
		sheet.formatCell(8, 1, { bold: false, color: "#008000" });
		const saved = await wb.save(false);
		const updated = (await Workbook.load(saved)).sheet("Liste");
		assert.equal(updated.cell(8, 1).value, before.value);
		assert.deepEqual(updated.cell(8, 1).runs?.slice(0, 3).map((run) => run.font), [
			{ color: "#008000" }, { color: "#008000" }, { color: "#008000", underline: true },
		]);
		assert.equal(updated.cell(8, 1).look.bold, undefined);
		assert.equal(updated.cell(8, 1).look.color, "#008000");
		assert.deepEqual(updated.cell(8, 3), before);
		assert.equal(readEntry(saved, "xl/sharedStrings.xml"), readEntry(original, "xl/sharedStrings.xml"));
		assert.match(readEntry(saved, "xl/worksheets/sheet3.xml")!, /<rFont val="Calibri"/);
	});
});

describe("writing cell formats", () => {
	it("round-trips font flags, colors and fills without changing cell values, number formats or formula caches", async () => {
		const original = await financeWorkbook();
		const wb = await Workbook.load(original);
		const sheet = wb.sheet("2026");
		const before = sheet.cell(2, 3);
		const formula = sheet.cell(2, 5);
		const patch = { bold: true, italic: true, underline: true, strike: true, fill: "#FFF2CC", color: "#C00000" };
		sheet.formatCell(2, 3, patch);
		sheet.formatCell(2, 5, patch);
		const bytes = await wb.save(false);
		const reread = (await Workbook.load(bytes)).sheet("2026");
		assert.equal(reread.cell(2, 3).value, before.value);
		assert.equal(reread.cell(2, 3).format, before.format);
		assert.deepEqual(reread.cell(2, 3).look, patch);
		assert.equal(reread.cell(2, 5).value, formula.value);
		assert.equal(reread.cell(2, 5).formula, formula.formula);
		for (const part of ["xl/calcChain.xml", "xl/workbook.xml", "xl/tables/table1.xml", "xl/sharedStrings.xml"]) assert.equal(readEntry(bytes, part), readEntry(original, part));
	});
	it("reuses style records, preserves alignment and supports clearing individual formats", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		const sheet = wb.sheet();
		sheet.formatCell(3, 1, { bold: false, fill: "#FFFFFF", color: "#000000", underline: true });
		sheet.formatCell(3, 2, { bold: false, fill: "#FFFFFF", color: "#000000", underline: true });
		const first = await wb.save(false);
		const after = (await Workbook.load(first)).sheet();
		assert.deepEqual(after.cell(3, 1).look, { underline: true, color: "#000000", fill: "#FFFFFF", align: "center", wrap: true });
		sheet.formatCell(3, 1, { bold: false, fill: "#FFFFFF", color: "#000000", underline: true });
		assert.equal(readEntry(await wb.save(false), "xl/styles.xml"), readEntry(first, "xl/styles.xml"));
		sheet.formatCell(3, 1, { fill: null, color: null, underline: false });
		assert.deepEqual((await Workbook.load(await wb.save(false))).sheet().cell(3, 1).look, { align: "center", wrap: true });
	});
	it("creates a valid style part when the workbook has none and preserves multiline text", async () => {
		const pkg = await Package.load(await financeWorkbook());
		pkg.removePart("xl/styles.xml", "xl/workbook.xml");
		for (const path of ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"]) {
			for (const row of elements(child(pkg.xml(path)!.documentElement, "sheetData")!, "row")) for (const cell of elements(row, "c")) cell.removeAttribute("s");
			pkg.markDirty(path);
		}
		const wb = await Workbook.load(await pkg.toBytes());
		wb.sheet("Liste").setCell(10, 1, { kind: "text", value: "第一行\n第二行\n" });
		wb.sheet("Liste").formatCell(10, 1, { bold: true, fill: "#FFFF00" });
		const saved = await wb.save();
		const again = (await Workbook.load(saved)).sheet("Liste").cell(10, 1);
		assert.equal(again.value, "第一行\n第二行\n");
		assert.equal(again.look.wrap, true);
		assert.equal(again.look.bold, true);
		assert.equal(again.look.fill, "#FFFF00");
		assert.match(readEntry(saved, "[Content_Types].xml")!, /PartName="\/xl\/styles.xml"/);
		assert.match(readEntry(saved, "xl/_rels/workbook.xml.rels")!, /\/styles" Target="styles.xml"/);
	});
});
