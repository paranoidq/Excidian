import assert from "node:assert/strict";
import { it } from "node:test";
import { parseBlockConfig } from "../src/controller/block-config";
import { buildGrid } from "../src/controller/grid-builder";
import { Package } from "../src/model/ooxml/package";
import { parseRange } from "../src/model/ooxml/refs";
import { child, createChild, elements } from "../src/model/ooxml/xml";
import { Workbook } from "../src/model/workbook";
import { financeWorkbook } from "./fixture";

it("round-trips merge and unmerge with the original values and cell formats", async () => {
	const wb = await Workbook.load(await financeWorkbook());
	const sheet = wb.sheet();
	const range = parseRange("A4:B5");
	sheet.setCell(4, 1, { kind: "text", value: "Merged heading" });
	sheet.formatCell(5, 2, { fill: "#FFF2CC" });
	sheet.mergeCells(range);
	const reopened = await Workbook.load(await wb.save(false));
	assert.deepEqual(reopened.sheet().mergedRanges(), [parseRange("A3:B3"), range]);
	const grid = buildGrid(reopened, parseBlockConfig("file: demo.xlsx"), false);
	assert.deepEqual(grid.rows[3].cells[0].span, { rows: 2, cols: 2 });
	assert.deepEqual(grid.rows[4].cells[1].mergeRange, range);
	assert.equal(grid.rows[4].cells[1].covered, true);
	reopened.sheet().unmergeCells([range]);
	const after = (await Workbook.load(await reopened.save(false))).sheet();
	assert.deepEqual(after.mergedRanges(), [parseRange("A3:B3")]);
	assert.equal(after.cell(4, 1).value, "Merged heading");
	assert.equal(after.cell(5, 2).look.fill, "#FFF2CC");
});

it("requires permission to clear non-anchor content and retains its formatting", async () => {
	const sheet = (await Workbook.load(await financeWorkbook())).sheet();
	const range = parseRange("A1:B1");
	const previous = sheet.cell(1, 2);
	assert.throws(() => sheet.mergeCells(range), /确认/);
	assert.deepEqual(sheet.cell(1, 2), previous);
	sheet.mergeCells(range, true);
	sheet.unmergeCells([range]);
	assert.equal(sheet.cell(1, 1).value, "Totale");
	assert.equal(sheet.cell(1, 2).value, null);
	assert.equal(sheet.cell(1, 2).formula, undefined);
	assert.equal(sheet.cell(1, 2).format, previous.format);
});

it("rejects hidden gaps, intersecting merges, Excel tables and stale unmerge selections", async () => {
	const wb = await Workbook.load(await financeWorkbook());
	const sheet = wb.sheet();
	assert.throws(() => sheet.mergeCells(parseRange("A2:A3")), /已合并/);
	assert.throws(() => sheet.mergeCells(parseRange("B2:C2")), /隐藏列/);
	assert.throws(() => sheet.mergeCells(parseRange("A5:A6")), /隐藏行/);
	assert.throws(() => wb.sheet("2026").mergeCells(parseRange("A2:B2")), /Excel 表格/);
	assert.throws(() => sheet.unmergeCells([parseRange("A3:B3"), parseRange("A4:B4")]), /已变化/);
	assert.deepEqual(sheet.mergedRanges(), [parseRange("A3:B3")]);
});

it("writes mergeCells in worksheet schema order and handles a missing dimension", async () => {
	const path = "xl/worksheets/sheet1.xml";
	const pkg = await Package.load(await financeWorkbook());
	const root = pkg.xml(path)!.documentElement;
	root.removeChild(child(root, "dimension")!);
	root.removeChild(child(root, "mergeCells")!);
	const filter = createChild(root, "autoFilter", child(root, "dataValidations")!);
	filter.setAttribute("ref", "A1:B4");
	pkg.markDirty(path);
	const wb = await Workbook.load(await pkg.toBytes());
	const range = parseRange("A9:B10");
	wb.sheet().mergeCells(range);
	const bytes = await wb.save(false);
	const serialized = (await Package.load(bytes)).xml(path)!.documentElement;
	const names = elements(serialized).map((el) => el.localName);
	assert.equal(names.indexOf("mergeCells"), names.indexOf("autoFilter") + 1);
	assert.equal(child(serialized, "mergeCells")!.getAttribute("count"), "1");
	const reopened = await Workbook.load(bytes);
	assert.equal(reopened.sheet().displayRange()?.bottom, 10);
	reopened.sheet().unmergeCells([range]);
	const final = (await Package.load(await reopened.save(false))).xml(path)!.documentElement;
	assert.equal(child(final, "mergeCells"), undefined);
});
