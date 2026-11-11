import assert from "node:assert/strict";
import { it } from "node:test";
import { NodeFile } from "../src/platform/desktop/node-file";

it("opens the exact workbook path using Electron's default-application API and propagates errors", async () => {
	const prior = (globalThis as any).window;
	const paths: string[] = []; let error = "";
	(globalThis as any).window = { require: (module: string) => {
		assert.equal(module, "electron");
		return { shell: { openPath: async (path: string) => { paths.push(path); return error; } } };
	} };
	try {
		const path = "/tmp/测试 文件 'with quotes'.xlsx";
		const file = new NodeFile(path);
		await file.openLocal(); assert.deepEqual(paths, [path]);
		error = "Failed to open path";
		await assert.rejects(file.openLocal(), /Failed to open path/);
	} finally { (globalThis as any).window = prior; }
});
