import type { CellValue } from "../values";
import type { CellFont, Styles } from "./styles";
import { child, elements, textOf } from "./xml";

export interface TextRun {
	text: string;
	/** An absent run font inherits the cell font. */
	font?: CellFont;
}

export interface RichText {
	text: string;
	runs: TextRun[];
	source: Element;
}

export function isRichText(value: CellValue | RichText): value is RichText {
	return value !== null && typeof value === "object" && "runs" in value;
}

/** Reads a shared or inline string. Phonetic annotations are not cell text. */
export function readCellText(el: Element, styles: Styles): string | RichText {
	const first = child(el, "t");
	if (first && !child(el, "r")) return textOf(first);
	const runs: TextRun[] = [];
	for (const part of elements(el)) {
		if (part.localName === "t") runs.push({ text: textOf(part) });
		if (part.localName !== "r") continue;
		const properties = child(part, "rPr");
		runs.push({ text: textOf(child(part, "t")), font: properties && styles.readFont(properties) });
	}
	return runs.length ? { text: runs.map((run) => run.text).join(""), runs, source: el } : "";
}
