import assert from "node:assert/strict";
import { it } from "node:test";
import { updateWorkbook, workbookHistory } from "../src/model/workbook-store";
import { Workbook } from "../src/model/workbook";
import type { StoredFile } from "../src/model/storage";
import { financeWorkbook } from "./fixture";

it("refuses a save if an external writer changes the workbook during backup", async () => {
	let bytes = await financeWorkbook();
	const external = await Workbook.load(bytes); external.sheet("Liste").setCell(2, 1, { kind: "text", value: "external" });
	const externalBytes = await external.save(); let writes = 0;
	const file: StoredFile = {
		id: `race-${Math.random()}`, name: "demo.xlsx", displayPath: "demo.xlsx",
		read: async () => bytes, write: async (value) => { writes++; bytes = value; },
		writeSibling: async () => { bytes = externalBytes; }, siblings: async () => [], watch: () => () => {},
	};
	await assert.rejects(updateWorkbook(file, (wb) => wb.sheet("Liste").setCell(2, 1, { kind: "text", value: "ours" })), /保存前已被其他程序修改/);
	assert.equal(writes, 0);
	assert.equal((await Workbook.load(bytes)).sheet("Liste").cell(2, 1).value, "external");
	assert.equal(workbookHistory(file).undo, false);
});
