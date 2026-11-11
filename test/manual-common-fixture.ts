/** Generate a disposable workbook for Obsidian checks; never replaces an existing file. */
import { existsSync, writeFileSync } from "node:fs";
import { strToU8, unzipSync, zipSync } from "fflate";
import { financeWorkbook } from "./fixture";

async function main() {
	const path = process.argv[2];
	if (!path || existsSync(path)) throw new Error("Provide a new fixture path; existing files are never overwritten.");
	const files = unzipSync(await financeWorkbook());
	const inline = (ref: string, text: string, style = 0) => `<c r="${ref}" t="inlineStr" s="${style}"><is><t xml:space="preserve">${text}</t></is></c>`;
	const headers = ["编号", "名称", "金额", "完成", "日期", "备注"];
	const rows = [`<row r="1">${headers.map((text, i) => inline(`${String.fromCharCode(65 + i)}1`, text, 5)).join("")}</row>`];
	for (let r = 2; r <= 21; r++) rows.push(`<row r="${r}">${inline(`A${r}`, String(r - 1).padStart(5, "0"), 6)}${inline(`B${r}`, `项目 ${r - 1}`)}<c r="C${r}" s="3"><v>${22 - r}</v></c><c r="D${r}" t="b"><v>${r % 2}</v></c><c r="E${r}" s="2"><v>46336</v></c>${inline(`F${r}`, r === 4 ? "这是一段较长的备注，可开启自动换行查看。\n第二行备注保留换行。" : r === 2 ? "https://obsidian.md" : r === 3 ? "  保留首尾空格  " : "备注")}</row>`);
	files["xl/worksheets/sheet1.xml"] = strToU8(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:F21"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" state="frozen" topLeftCell="B2"/></sheetView></sheetViews><cols><col min="1" max="2" width="16" customWidth="1"/><col min="3" max="5" width="14" customWidth="1"/><col min="6" max="6" width="24" customWidth="1"/></cols><sheetData>${rows.join("")}</sheetData></worksheet>`);
	const xml = new TextDecoder().decode(files["xl/styles.xml"]).replace('cellXfs count="6"', 'cellXfs count="7"').replace('</cellXfs>', '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>');
	files["xl/styles.xml"] = strToU8(xml);
	writeFileSync(path, zipSync(files));
	console.log(path);
}
void main();
