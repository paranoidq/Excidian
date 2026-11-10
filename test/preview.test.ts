import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { parseBlockConfig } from "../src/controller/block-config";
import { buildGrid } from "../src/controller/grid-builder";
import { DeviceSettings } from "../src/model/device-settings";
import { Workbook } from "../src/model/workbook";
import { financeWorkbook } from "./fixture";

describe("workbook previews", () => {
	it("lists every worksheet and resolves the initial sheet by name or position", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		const cfg = parseBlockConfig("file: demo.xlsx");
		const first = buildGrid(wb, cfg, true);
		assert.deepEqual(first.sheets, ["Riepilogo", "2026", "Liste"]);
		assert.equal(first.activeSheet, "Riepilogo");
		assert.equal(buildGrid(wb, parseBlockConfig("file: demo.xlsx\nsheet: 2"), true).activeSheet, "2026");
		const other = buildGrid(wb, { ...cfg, target: { kind: "sheet", sheet: "Liste" } }, true);
		assert.equal(other.activeSheet, "Liste");
		assert.equal(other.rows[0].cells[0].text, "Tipo");
	});

	it("makes every sheet/table cell read-only and removes the append row, then restores editable cells", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		for (const target of ["sheet: 2026", "table: Movimenti"]) {
			const cfg = parseBlockConfig(`file: demo.xlsx\n${target}`);
			const read = buildGrid(wb, cfg, true);
			assert.equal(read.readOnly, true);
			assert.ok(read.rows.every((row) => row.kind !== "new" && row.cells.every((cell) => !cell.editable)));
			const edit = buildGrid(wb, cfg, false);
			assert.ok(edit.rows.some((row) => row.kind === "new"));
			assert.ok(edit.rows.some((row) => row.cells.some((cell) => cell.editable)));
		}
	});

	it("parses explicit mode overrides and rejects typos", () => {
		assert.equal(parseBlockConfig("file: demo.xlsx").mode, undefined);
		assert.equal(parseBlockConfig("file: demo.xlsx\nmode: readonly").mode, "readonly");
		assert.equal(parseBlockConfig("file: demo.xlsx\nmode: edit").mode, "edit");
		assert.throws(() => parseBlockConfig("file: demo.xlsx\nmode: edti"), /Invalid `mode:`/);
	});

	it("defaults to read-only and persists the configurable default without losing other settings", () => {
		const values = new Map<string, unknown>();
		const store = { load: (key: string) => values.get(key), save: (key: string, value: unknown) => { values.set(key, value); } };
		const settings = new DeviceSettings(store);
		assert.equal(settings.defaultReadOnly, true);
		settings.oneDriveRoot = "/OneDrive";
		settings.defaultReadOnly = false;
		const again = new DeviceSettings(store);
		assert.equal(again.defaultReadOnly, false);
		assert.equal(again.oneDriveRoot, "/OneDrive");
		again.defaultReadOnly = true;
		assert.equal(settings.defaultReadOnly, true);
	});
});
