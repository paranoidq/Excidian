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

	it("offers an entry column only in unrestricted editable sheet previews", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		const cfg = parseBlockConfig("file: demo.xlsx\nsheet: 2026");
		const grid = buildGrid(wb, cfg, false, true);
		assert.equal(grid.columns[grid.columns.length - 1], "F");
		assert.equal(grid.columnNumbers![grid.columnNumbers!.length - 1], 6);
		assert.equal(grid.rows[1].cells[5].editable, true);
		assert.equal(buildGrid(wb, cfg, true, true).columns.length, 5);
		assert.equal(buildGrid(wb, parseBlockConfig("file: demo.xlsx\nsheet: 2026\ncols: A:C"), false, true).columns.length, 3);
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

it("bounds wide worksheet previews instead of constructing a million DOM cells", async () => {
	const { Package } = await import("../src/model/ooxml/package");
	const { child } = await import("../src/model/ooxml/xml");
	const pkg = await Package.load(await financeWorkbook());
	child(pkg.xml("xl/worksheets/sheet1.xml")!.documentElement, "dimension")!.setAttribute("ref", "A1:IV5000");
	pkg.markDirty("xl/worksheets/sheet1.xml");
	const grid = buildGrid(await Workbook.load(await pkg.toBytes()), parseBlockConfig("file: demo.xlsx"), true, true);
	assert.ok(grid.rows.length * grid.columns.length <= 30_000);
	assert.ok(grid.rows.length < 5000);
	assert.match(grid.note!, /preview cell budget/);
});
