import { contrastText, resolveColor } from "./colors";
import type { Package } from "./package";
import { child, createChild, elements } from "./xml";

/** How a cell looks. Colours are "#RRGGBB"; unset means "use the Obsidian theme". */
export interface CellFormat {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	color?: string;
	fill?: string;
	align?: "left" | "center" | "right" | "justify";
	wrap?: boolean;
}

export interface CellStylePatch {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	color?: string | null;
	fill?: string | null;
	wrap?: boolean;
}

export const FONT_TOGGLES = { bold: "b", italic: "i", underline: "u", strike: "strike" } as const;
export type FontToggle = keyof typeof FONT_TOGGLES;

export function defaultStylesXml(namespace: string): string {
	return `<styleSheet xmlns="${namespace}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}

function styleShape(el: Element): unknown[] {
	return [el.namespaceURI, el.localName, Array.from(el.attributes).map((attr) => [attr.name, attr.value]).sort(), el.textContent?.trim(), elements(el).map(styleShape)];
}

const styleKey = (el: Element): string => JSON.stringify(styleShape(el));

const ALIGNMENTS: Record<string, CellFormat["align"]> = {
	left: "left",
	fill: "left",
	center: "center",
	centerContinuous: "center",
	distributed: "center",
	right: "right",
	justify: "justify",
};

/** <b/>, <i/>… count as on unless val="0"/"false". */
function flag(el: Element | undefined): boolean {
	if (!el) return false;
	const val = el.getAttribute("val");
	return val !== "0" && val !== "false" && val !== "none";
}

/** Excel's short-date format: shown in the user's locale, so it gets special handling. */
export const LOCALE_SHORT_DATE = "mm-dd-yy";

const BUILTIN_FORMATS: Record<number, string> = {
	0: "General",
	1: "0",
	2: "0.00",
	3: "#,##0",
	4: "#,##0.00",
	9: "0%",
	10: "0.00%",
	11: "0.00E+00",
	14: LOCALE_SHORT_DATE,
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[Red](#,##0.00)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mm:ss.0",
	48: "##0.0E+0",
	49: "@",
};

function isBuiltinDate(id: number): boolean {
	return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
}

/** True when a format code shows dates or times (has d, m, y, h or s outside literals). */
export function isDateFormat(code: string): boolean {
	const bare = code
		.split(";")[0]
		.replace(/"[^"]*"/g, "")
		.replace(/\\.|[_*]./g, "")
		.replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, "");
	return /[dmyhs]/i.test(bare) && !/^general$/i.test(bare.trim());
}

/** Reads cell styles and reuses matching font, fill and format records on writes. */
export class Styles {
	private xfs: Element[] = [];
	private fonts: Element[] = [];
	private fills: Element[] = [];
	private customFormats = new Map<number, string>();
	private dateVariants = new Map<number, number>();
	private formats = new Map<number, CellFormat>();
	private interned = new Map<Element[], Map<string, number>>();
	private variants = new Map<string, number>();

	constructor(
		private pkg: Package,
		private path: string | undefined,
		private theme: string[],
		private createStylePart: () => string,
	) { this.readLists(); }

	private readLists() {
		const root = this.path ? this.pkg.xml(this.path)?.documentElement : undefined;
		if (!root) return;
		const numFmts = child(root, "numFmts");
		for (const f of numFmts ? elements(numFmts, "numFmt") : []) {
			this.customFormats.set(Number(f.getAttribute("numFmtId")), f.getAttribute("formatCode") ?? "General");
		}
		const list = (name: string, item: string) => {
			const el = child(root, name);
			return el ? elements(el, item) : [];
		};
		this.xfs = list("cellXfs", "xf");
		this.fonts = list("fonts", "font");
		this.fills = list("fills", "fill");
	}

	/** Font, fill and alignment of a style, ready to display. */
	cellFormat(styleIndex: number): CellFormat {
		let format = this.formats.get(styleIndex);
		if (format) return format;
		format = {};
		const xf = this.xfs[styleIndex];
		if (xf) {
			const fontId = Number(xf.getAttribute("fontId") ?? 0);
			const font = this.fonts[fontId];
			if (font) {
				if (flag(child(font, "b"))) format.bold = true;
				if (flag(child(font, "i"))) format.italic = true;
				if (flag(child(font, "u"))) format.underline = true;
				if (flag(child(font, "strike"))) format.strike = true;
				const color = resolveColor(child(font, "color"), this.theme, "font");
				if (color && (!color.isDefault || (fontId !== 0 && child(font, "color")?.hasAttribute("rgb")))) format.color = color.hex;
			}

			const pattern = this.fills[Number(xf.getAttribute("fillId") ?? 0)];
			const patternFill = pattern && child(pattern, "patternFill");
			if (patternFill?.getAttribute("patternType") === "solid") {
				const fill = resolveColor(child(patternFill, "fgColor"), this.theme, "fill");
				if (fill && (!fill.isDefault || child(patternFill, "fgColor")?.hasAttribute("rgb"))) {
					format.fill = fill.hex;
					// Excel's default black text would vanish on a dark fill, and theme text on a light one.
					format.color ??= contrastText(fill.hex);
				}
			}

			const alignment = child(xf, "alignment");
			const align = ALIGNMENTS[alignment?.getAttribute("horizontal") ?? ""];
			if (align) format.align = align;
			const wrap = alignment?.getAttribute("wrapText");
			if (wrap === "1" || wrap === "true") format.wrap = true;
		}
		this.formats.set(styleIndex, format);
		return format;
	}

	private formatId(styleIndex: number): number {
		return Number(this.xfs[styleIndex]?.getAttribute("numFmtId") ?? 0);
	}

	formatCode(styleIndex: number): string {
		const id = this.formatId(styleIndex);
		return this.customFormats.get(id) ?? BUILTIN_FORMATS[id] ?? "General";
	}

	isDate(styleIndex: number): boolean {
		const id = this.formatId(styleIndex);
		const custom = this.customFormats.get(id);
		return custom !== undefined ? isDateFormat(custom) : isBuiltinDate(id);
	}

	/** Clone the effective style and touch only the requested properties. */
	formatStyle(styleIndex: number, patch: CellStylePatch): number {
		const key = JSON.stringify([styleIndex, patch]);
		const cached = this.variants.get(key);
		if (cached !== undefined) return cached;
		for (const color of [patch.color, patch.fill]) {
			if (color != null && !/^#[0-9a-f]{6}$/i.test(color)) throw new Error("无效的单元格颜色。");
		}
		if (!this.path) { this.path = this.createStylePart(); this.readLists(); }
		const root = this.pkg.xml(this.path)?.documentElement;
		const cellXfs = root && child(root, "cellXfs");
		const source = this.xfs[styleIndex] ?? this.xfs[0];
		if (!root || !cellXfs || !source) throw new Error("工作簿的样式记录不完整，请先在 Excel 中保存。");
		const xf = source.cloneNode(true) as Element;

		if (patch.color !== undefined || Object.keys(FONT_TOGGLES).some((name) => patch[name as FontToggle] !== undefined)) {
			const fonts = child(root, "fonts");
			const original = this.fonts[Number(source.getAttribute("fontId") ?? 0)];
			if (!fonts || !original) throw new Error("工作簿的字体记录不完整。");
			const font = original.cloneNode(true) as Element;
			for (const [name, tag] of Object.entries(FONT_TOGGLES)) {
				const enabled = patch[name as FontToggle], el = child(font, tag);
				if (enabled === undefined || enabled === flag(el)) continue;
				if (el) font.removeChild(el);
				if (enabled) createChild(font, tag);
			}
			if (patch.color !== undefined) {
				const old = child(font, "color");
				if (old) font.removeChild(old);
				if (patch.color) createChild(font, "color").setAttribute("rgb", `FF${patch.color.slice(1).toUpperCase()}`);
			}
			xf.setAttribute("fontId", String(this.intern(fonts, this.fonts, font)));
			xf.setAttribute("applyFont", "1");
		}
		if (patch.fill !== undefined) {
			const fills = child(root, "fills");
			if (!fills || !this.fills[0]) throw new Error("工作簿的填充记录不完整。");
			let id = 0;
			if (patch.fill) {
				const fill = fills.ownerDocument.createElementNS(fills.namespaceURI, fills.prefix ? `${fills.prefix}:fill` : "fill");
				const pattern = createChild(fill, "patternFill"); pattern.setAttribute("patternType", "solid");
				createChild(pattern, "fgColor").setAttribute("rgb", `FF${patch.fill.slice(1).toUpperCase()}`);
				createChild(pattern, "bgColor").setAttribute("indexed", "64");
				id = this.intern(fills, this.fills, fill);
			}
			xf.setAttribute("fillId", String(id)); xf.setAttribute("applyFill", "1");
		}
		if (patch.wrap !== undefined) {
			const alignment = child(xf, "alignment") ?? createChild(xf, "alignment", xf.firstChild);
			alignment.setAttribute("wrapText", patch.wrap ? "1" : "0"); xf.setAttribute("applyAlignment", "1");
		}
		const index = this.intern(cellXfs, this.xfs, xf);
		this.variants.set(key, index);
		return index;
	}

	private intern(parent: Element, records: Element[], record: Element): number {
		let index = this.interned.get(records);
		if (!index) { index = new Map(records.map((el, i) => [styleKey(el), i])); this.interned.set(records, index); }
		const key = styleKey(record), known = index.get(key);
		if (known !== undefined) return known;
		if (records.length >= 65_000) throw new Error("工作簿样式数量已达到上限。");
		const next = records.length;
		parent.appendChild(record); records.push(record); index.set(key, next);
		parent.setAttribute("count", String(records.length)); this.pkg.markDirty(this.path!);
		return next;
	}

	/** A style like `styleIndex` but showing a date; appended to the workbook if needed. */
	dateStyle(styleIndex: number): number {
		if (this.isDate(styleIndex)) return styleIndex;
		const known = this.dateVariants.get(styleIndex);
		if (known !== undefined) return known;

		const root = this.path ? this.pkg.xml(this.path)?.documentElement : undefined;
		const cellXfs = root && child(root, "cellXfs");
		const source = this.xfs[styleIndex] ?? this.xfs[0];
		if (!cellXfs || !source) return styleIndex;

		const xf = source.cloneNode(true) as Element;
		xf.setAttribute("numFmtId", "14");
		xf.setAttribute("applyNumberFormat", "1");
		const index = this.intern(cellXfs, this.xfs, xf);
		this.dateVariants.set(styleIndex, index);
		return index;
	}
}
