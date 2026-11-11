import { backupName, isConflictCopyName, isExcelOwnerName, SiblingFile, StoredFile } from "./storage";
import { Workbook } from "./workbook";

// Writes to the same file are chained so two quick edits never race each other.
const writeQueues = new Map<string, Promise<unknown>>();
// Files already backed up during this Obsidian session.
const backedUp = new Set<string>();
interface HistoryEntry { before: Uint8Array; after: Uint8Array }
interface History { undo: HistoryEntry[]; redo: HistoryEntry[] }
const histories = new Map<string, History>();
const MAX_HISTORY_BYTES = 32 * 1024 * 1024;

function same(a: Uint8Array, b: Uint8Array): boolean {
	return a.length === b.length && a.every((value, i) => value === b[i]);
}

function enqueue<T>(file: StoredFile, action: () => Promise<T>): Promise<T> {
	const next = (writeQueues.get(file.id) ?? Promise.resolve()).catch(() => undefined).then(action);
	writeQueues.set(file.id, next);
	return next;
}

export function workbookHistory(file: StoredFile): { undo: boolean; redo: boolean } {
	const history = histories.get(file.id);
	return { undo: !!history?.undo.length, redo: !!history?.redo.length };
}

/** Snapshot restore is serialized with saves and refuses to overwrite external edits. */
export function restoreWorkbook(file: StoredFile, direction: "undo" | "redo"): Promise<Workbook> {
	return enqueue(file, async () => {
		const history = histories.get(file.id);
		const from = history?.[direction];
		const entry = from?.[from.length - 1];
		if (!entry || !history) throw new Error(direction === "undo" ? "没有可撤销的修改。" : "没有可重做的修改。");
		const current = await file.read();
		if (!same(current, direction === "undo" ? entry.after : entry.before)) {
			histories.delete(file.id);
			throw new Error("文件已被其他程序或预览修改，已停止撤销以保留最新内容。请刷新。");
		}
		const bytes = direction === "undo" ? entry.before : entry.after;
		const wb = await Workbook.load(bytes);
		if (!same(await file.read(), current)) throw new Error("文件在撤销前已变化，请刷新后重试。");
		await file.write(bytes);
		from!.pop(); history[direction === "undo" ? "redo" : "undo"].push(entry);
		return wb;
	});
}

export async function readWorkbook(file: StoredFile): Promise<Workbook> {
	return Workbook.load(await file.read());
}

/**
 * Loads the file fresh, applies `change`, and saves it. Re-reading right before
 * writing keeps changes made in Excel (or synced from another device) meanwhile.
 * The first save of a session keeps a backup of the original next to the file.
 * Resolves to the updated workbook.
 */
export function updateWorkbook(file: StoredFile, change: (wb: Workbook) => void, options?: { recalculate?: boolean }): Promise<Workbook> {
	return enqueue(file, async () => {
			const original = await file.read();
			const wb = await Workbook.load(original);
			change(wb);
			const bytes = await wb.save(options?.recalculate);
			if (!backedUp.has(file.id)) {
				await file.writeSibling(backupName(file.name), original);
				backedUp.add(file.id);
			}
			// A sync client/Excel may have changed the file while XML or backup was prepared.
			if (!same(await file.read(), original)) throw new Error("文件在保存前已被其他程序修改，本次修改未写入。请刷新后重试。");
			await file.write(bytes);
			const history = histories.get(file.id) ?? { undo: [], redo: [] };
			const last = history.undo[history.undo.length - 1];
			if ((last && !same(last.after, original)) || (!last && history.redo.length && !same(history.redo[history.redo.length - 1].before, original))) history.undo = [];
			history.redo = [];
			history.undo.push({ before: original.slice(), after: bytes.slice() });
			while (history.undo.length > 20 || history.undo.reduce((sum, e) => sum + e.before.length + e.after.length, 0) > MAX_HISTORY_BYTES) history.undo.shift();
			histories.delete(file.id); histories.set(file.id, history);
			// Bound total history memory across workbooks too (least recently saved first).
			while ([...histories.values()].reduce((sum, h) => sum + [...h.undo, ...h.redo].reduce((n, e) => n + e.before.length + e.after.length, 0), 0) > MAX_HISTORY_BYTES) histories.delete(histories.keys().next().value!);
			return wb;
		});
}

/** Sync-conflict copies of the file in its folder, newest first. */
export async function findConflictCopies(file: StoredFile): Promise<SiblingFile[]> {
	const siblings = await file.siblings();
	return siblings.filter((s) => isConflictCopyName(file.name, s.name)).sort((a, b) => b.modified.getTime() - a.modified.getTime());
}

/** Excel keeps a "~$name.xlsx" owner file next to a workbook while it is open. */
export async function isOpenInExcel(file: StoredFile): Promise<boolean> {
	return (await file.siblings()).some((s) => isExcelOwnerName(file.name, s.name));
}
