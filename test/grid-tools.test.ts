import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { clipboardText, parseClipboard, pasteChanges } from "../src/view/grid-tools";
import type { GridData } from "../src/view/grid-data";

function grid(): GridData {
	return { title: "demo", tooltip: "", columns: ["A", "B"], structureEditable: true, rows: [
		{ label: "1", kind: "data", cells: [{ row: 1, col: 1, text: "Name", editText: "Name", editable: true }, { row: 1, col: 2, text: "Amount", editText: "Amount", editable: true }] },
		...[ ["beta", 10], ["Alpha", 2], ["alpha", 2], ["", null] ].map(([name, value], i) => ({ label: String(i + 2), kind: "data" as const, cells: [{ row: i + 2, col: 1, text: String(name), editText: String(name), editable: true }, { row: i + 2, col: 2, text: value === null ? "" : String(value), editText: String(value ?? ""), editable: true }] })),
		{ label: "6", kind: "new", cells: [{ row: 6, col: 1, text: "", editText: "", editable: true }, { row: 6, col: 2, text: "", editText: "", editable: true }] },
	] };
}

describe("clipboard and rectangular selection", () => {
	it("round-trips quoted tabs, multiline cells, spaces, blanks and quotes", () => {
		const values = [["00123", "a\tb", 'he said "yes"'], [" padded ", "two\nlines", ""]];
		assert.deepEqual(parseClipboard(clipboardText(values) + "\r\n"), values);
		assert.deepEqual(parseClipboard(""), [[""]]);
		assert.deepEqual(parseClipboard("a\tb\n\t\n"), [["a", "b"], ["", ""]]);
	});
	it("rejects ragged, malformed and excessive data", () => {
		assert.throws(() => parseClipboard("a\tb\nc"), /矩形/);
		assert.throws(() => parseClipboard('"unfinished'), /引号/);
		assert.throws(() => parseClipboard(Array(10002).fill("x").join("\t")), /10,000/);
	});
	it("maps paste to worksheet coordinates, supports fill, and rejects overflow atomically", () => {
		const data = grid();
		const result = pasteChanges(data, { r: 1, c: 0 }, [["00123", "2"], ["00345", "4"]]);
		assert.deepEqual(result.map((change) => [change.cell.row, change.cell.col, change.text]), [[2, 1, "00123"], [2, 2, "2"], [3, 1, "00345"], [3, 2, "4"]]);
		assert.equal(pasteChanges(data, { r: 0, c: 0 }, [[""]], { anchor: { r: 2, c: 1 }, end: { r: 1, c: 0 } }).length, 4);
		assert.throws(() => pasteChanges(data, { r: 5, c: 1 }, [["a", "b"]]), /超出/);
	});
	it("rejects merged, readonly, hidden and reordered targets", () => {
		const data = grid();
		data.rows[1].cells[1].span = { rows: 2, cols: 1 };
		assert.throws(() => pasteChanges(data, { r: 1, c: 0 }, [["a", "b"]]), /合并/);
		delete data.rows[1].cells[1].span;
		data.rows[1].cells[1].editable = false;
		assert.throws(() => pasteChanges(data, { r: 1, c: 0 }, [["a", "b"]]), /只读/);
		data.rows[1].cells[0].row = 4;
		assert.throws(() => pasteChanges(data, { r: 1, c: 0 }, [["a"], ["b"]]), /连续/);
	});
});
