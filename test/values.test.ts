import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { dateToSerial, editText, parseInput, serialToDate } from "../src/model/values";
import { Workbook } from "../src/model/workbook";
import { Package } from "../src/model/ooxml/package";
import { child, createChild } from "../src/model/ooxml/xml";
import { financeWorkbook } from "./fixture";

describe("data fidelity", () => {
	it("preserves identifiers, decimal precision and intentional whitespace", () => {
		for (const value of ["00123", "-00123", "1234567890123456", "0.1234567890123456", " padded ", "  "]) assert.deepEqual(parseInput(value), { kind: "text", value });
		assert.deepEqual(parseInput("123456789012345"), { kind: "number", value: 123456789012345 });
		assert.deepEqual(parseInput("-12,50"), { kind: "number", value: -12.5 });
		assert.deepEqual(parseInput("0.05"), { kind: "number", value: 0.05 });
		assert.deepEqual(parseInput("'123"), { kind: "text", value: "123" });
	});
	it("respects Text formats and existing strings; booleans round-trip as booleans", async () => {
		assert.deepEqual(parseInput("=1+1", { format: "@" }), { kind: "text", value: "=1+1" });
		for (const text of ["123", "2026-11-10", "TRUE"]) assert.deepEqual(parseInput(text, { value: "text" }), { kind: "text", value: text });
		const wb = await Workbook.load(await financeWorkbook());
		wb.sheet("Liste").setCell(10, 1, parseInput("FALSE"));
		wb.sheet("Liste").setCell(11, 1, parseInput("00123"));
		wb.sheet("Liste").setCell(12, 1, parseInput(" padded "));
		const again = await Workbook.load(await wb.save());
		assert.equal(again.sheet("Liste").cell(10, 1).value, false);
		assert.equal(again.sheet("Liste").cell(11, 1).value, "00123");
		assert.equal(again.sheet("Liste").cell(12, 1).value, " padded ");
		assert.deepEqual(parseInput(editText(false, undefined), { value: false }), { kind: "boolean", value: false });
	});
	it("retains date times and rejects impossible dates or time components", () => {
		const date = new Date("2026-11-10T12:34:56.789Z");
		assert.deepEqual(parseInput(editText(date, undefined)), { kind: "date", value: date });
		for (const epoch of [false, true]) assert.equal(serialToDate(dateToSerial(date, epoch), epoch).getTime(), date.getTime());
		for (const value of ["2026-02-31", "2026-11-10 24:00", "2026-11-10 10:60"]) assert.deepEqual(parseInput(value), { kind: "text", value });
	});
	it("blocks protected and array-formula cells before writing", async () => {
		const pkg = await Package.load(await financeWorkbook());
		const root = pkg.xml("xl/worksheets/sheet3.xml")!.documentElement;
		createChild(root, "sheetProtection").setAttribute("sheet", "1"); pkg.markDirty("xl/worksheets/sheet3.xml");
		let wb = await Workbook.load(await pkg.toBytes());
		assert.throws(() => wb.sheet("Liste").setCell(2, 1, parseInput("change")), /受保护/);
		root.removeChild(child(root, "sheetProtection")!);
		const c = createChild(createChild(child(root, "sheetData")!, "row"), "c"); c.setAttribute("r", "D2");
		const f = createChild(c, "f"); f.setAttribute("t", "array"); f.setAttribute("ref", "D2:E4");
		pkg.markDirty("xl/worksheets/sheet3.xml"); wb = await Workbook.load(await pkg.toBytes());
		assert.throws(() => wb.sheet("Liste").setCell(4, 5, parseInput("change")), /数组公式/);
		assert.doesNotThrow(() => wb.sheet("Liste").setCell(5, 5, parseInput("safe")));
	});
});

describe("links", () => {
	it("exposes web links but never executable or local-file schemes", async () => {
		const { safeLink } = await import("../src/model/hyperlinks");
		assert.equal(safeLink("https://obsidian.md"), "https://obsidian.md");
		assert.equal(safeLink("mailto:user@example.com"), "mailto:user@example.com");
		for (const link of ["javascript:alert(1)", "file:///etc/passwd", "data:text/html,x", "https://", " https://example.com", "https://exa\nmple.com"]) assert.equal(safeLink(link), undefined);
	});
});

it("rejects non-storable cell values without changing the cell", async () => {
	const wb = await Workbook.load(await financeWorkbook()); const sheet = wb.sheet("Liste"); const original = sheet.cell(2, 1).value;
	for (const input of [{ kind: "text", value: "x".repeat(32768) }, { kind: "text", value: "bad\u0000text" }, { kind: "number", value: Infinity }] as const) assert.throws(() => sheet.setCell(2, 1, input));
	assert.equal(sheet.cell(2, 1).value, original);
});
