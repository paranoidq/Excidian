import type { CellData, GridData } from "./grid-data";

export interface CellPoint { r: number; c: number }
export interface CellRange { anchor: CellPoint; end: CellPoint }
export function bounds(range: CellRange) {
	return { top: Math.min(range.anchor.r, range.end.r), bottom: Math.max(range.anchor.r, range.end.r), left: Math.min(range.anchor.c, range.end.c), right: Math.max(range.anchor.c, range.end.c) };
}

/** TSV, including Excel's quoted tabs, quotes and multi-line cells. No HTML is read. */
export function parseClipboard(text: string): string[][] {
	if (text.length > 2_000_000) throw new Error("粘贴内容过大，请分批粘贴（每次最多 10,000 个单元格）。");
	const rows: string[][] = [[]]; let value = "", quoted = false, ended = false;
	for (let i = 0; i < text.length; i++) {
		const ch = text[i]; ended = false;
		if (quoted) {
			if (ch === '"') { if (text[i + 1] === '"') { value += '"'; i++; } else quoted = false; }
			else value += ch;
		} else if (ch === '"' && value === "") quoted = true;
		else if (ch === "\t" || ch === "\r" || ch === "\n") {
			rows[rows.length - 1].push(value); value = "";
			if (ch !== "\t") {
				if (ch === "\r" && text[i + 1] === "\n") i++;
				rows.push([]); ended = true;
			}
		} else value += ch;
	}
	if (quoted) throw new Error("剪贴板中的引号未闭合，请检查复制的区域。");
	if (ended) rows.pop(); else rows[rows.length - 1].push(value);
	const width = rows[0]?.length ?? 0;
	if (rows.some((row) => row.length !== width)) throw new Error("粘贴内容必须是完整的矩形区域。");
	if (rows.length * width > 10_000) throw new Error("每次最多粘贴 10,000 个单元格。");
	return rows;
}

export function clipboardText(rows: string[][]): string {
	return rows.map((row) => row.map((value) => /[\t\r\n"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value).join("\t")).join("\r\n");
}

export interface CellChange { cell: CellData; text: string }
export function pasteChanges(grid: GridData, point: CellPoint, values: string[][], selection?: CellRange): CellChange[] {
	const area = selection && bounds(selection);
	const fill = values.length === 1 && values[0].length === 1 && area;
	const height = fill ? area.bottom - area.top + 1 : values.length;
	const width = fill ? area.right - area.left + 1 : values[0].length;
	const start = fill ? { r: area.top, c: area.left } : point;
	if (height * width > 10_000) throw new Error("每次最多修改 10,000 个单元格。");
	const changes: CellChange[] = [];
	for (let r = 0; r < height; r++) for (let c = 0; c < width; c++) {
		const cell = grid.rows[start.r + r]?.cells[start.c + c];
		if (!cell) throw new Error("粘贴超出当前显示范围，请先插入足够的行列。");
		if (!cell.editable || cell.covered || cell.span) throw new Error("粘贴区域含只读、标题或合并单元格，请选择普通数据区域。");
		const priorRow = r > 0 && grid.rows[start.r + r - 1]?.cells[start.c + c];
		const priorCol = c > 0 && grid.rows[start.r + r]?.cells[start.c + c - 1];
		if ((priorRow && cell.row !== priorRow.row + 1) || (priorCol && cell.col !== priorCol.col + 1)) throw new Error("粘贴区域包含隐藏或重排的行列，请选择连续区域。");
		changes.push({ cell, text: fill ? values[0][0] : values[r][c] });
	}
	return changes;
}
