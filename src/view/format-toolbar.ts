import { setIcon } from "obsidian";
import type { CellStylePatch, FontToggle } from "../model/ooxml/styles";
import type { CellData } from "./grid-data";

const FONT_BUTTONS: [FontToggle, string, string][] = [
	["bold", "bold", "加粗"], ["italic", "italic", "斜体"],
	["underline", "underline", "下划线"], ["strike", "strikethrough", "删除线"],
];
const COLORS = [
	"#000000", "#FFFFFF", "#808080", "#D9E1F2", "#FCE4D6", "#FFF2CC", "#E2F0D9", "#DDEBF7",
	"#C00000", "#FF0000", "#FFC000", "#FFFF00", "#70AD47", "#00B0F0", "#4472C4", "#7030A0",
];

export function fontState(cell: CellData, property: FontToggle): boolean | "mixed" {
	if (!cell.runs) return !!cell.style?.[property];
	const runs = cell.runs.filter((run) => run.text);
	const active = runs.filter((run) => run.font?.[property]).length;
	return active === 0 ? false : active === runs.length ? true : "mixed";
}

export class FormatToolbar {
	private el: HTMLElement;
	private toggles = new Map<FontToggle, HTMLButtonElement>();
	private palettes: HTMLElement[] = [];

	constructor(parent: HTMLElement, onToggle: (property: FontToggle) => void, onPatch: (patch: CellStylePatch) => void) {
		this.el = parent.createDiv({ cls: "excidian-format", attr: { role: "group", "aria-label": "单元格格式" } });
		const button = (parent: HTMLElement, label: string, icon: string) => {
			const el = parent.createEl("button", { cls: "clickable-icon", attr: { type: "button", title: label, "aria-label": label } });
			setIcon(el, icon);
			return el;
		};
		for (const [property, icon, label] of FONT_BUTTONS) {
			const control = button(this.el, label, icon);
			if (property === "bold" || property === "italic") {
				const key = property === "bold" ? "B" : "I";
				control.title = `${label}（Ctrl/⌘+${key}）`;
				control.setAttribute("aria-keyshortcuts", `Control+${key} Meta+${key}`);
			}
			control.onclick = () => onToggle(property);
			this.toggles.set(property, control);
		}
		for (const [property, label, icon, initial] of [["fill", "填充颜色", "paint-bucket", "#FFFF00"], ["color", "文字颜色", "baseline", "#FF0000"]] as const) {
			const group = this.el.createDiv("excidian-format-color");
			const apply = button(group, label, icon);
			apply.addClass("excidian-format-current-color");
			const choose = button(group, `选择${label}`, "chevron-down");
			choose.addClass("excidian-format-chevron");
			const palette = group.createDiv({ cls: "excidian-format-palette", attr: { role: "group", "aria-label": label } });
			this.palettes.push(palette);
			palette.toggle(false);
			let selected: string | null = initial;
			apply.style.setProperty("--excidian-swatch", selected);
			const pick = (color: string | null) => {
				selected = color;
				apply.style.setProperty("--excidian-swatch", color ?? "transparent");
				this.closePalettes();
				onPatch({ [property]: color });
			};
			apply.onclick = () => pick(selected);
			choose.setAttribute("aria-expanded", "false");
			choose.onclick = () => {
				const show = palette.style.display === "none";
				this.closePalettes();
				palette.toggle(show);
				choose.setAttribute("aria-expanded", String(show));
			};
			const swatches = palette.createDiv("excidian-format-swatches");
			for (const color of COLORS) {
				const swatch = swatches.createEl("button", { cls: "excidian-format-swatch", attr: { type: "button", title: color, "aria-label": `${label} ${color}` } });
				swatch.style.backgroundColor = color;
				swatch.onclick = () => pick(color);
			}
			palette.createEl("button", { text: property === "fill" ? "无填充" : "自动颜色", attr: { type: "button" } }).onclick = () => pick(null);
			palette.addEventListener("keydown", (e) => {
				if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.closePalettes(); choose.focus(); }
			});
		}
	}

	closePalettes() {
		for (const palette of this.palettes) {
			palette.toggle(false);
			palette.previousElementSibling?.setAttribute("aria-expanded", "false");
		}
	}

	update(visible: boolean, cells: CellData[]) {
		this.el.toggle(visible);
		this.el.querySelectorAll("button").forEach((button) => { button.disabled = !cells.length; });
		for (const [property, button] of this.toggles) {
			const states = cells.map((cell) => fontState(cell, property));
			const pressed = states.every((state) => state === false) ? false : states.every((state) => state === true) ? true : "mixed";
			button.setAttribute("aria-pressed", String(pressed));
		}
		if (!cells.length) this.closePalettes();
	}
}
