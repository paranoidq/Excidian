import { colName, colNumber, MAX_COL, MAX_ROW } from "./refs";

export type Axis = "row" | "column";
export interface StructureEdit {
	axis: Axis;
	kind: "insert" | "delete" | "clear";
	start: number;
	count: number;
}

export function validateEdit(edit: StructureEdit) {
	const limit = edit.axis === "row" ? MAX_ROW : MAX_COL;
	if (!["row", "column"].includes(edit.axis) || !["insert", "delete", "clear"].includes(edit.kind)
		|| !Number.isInteger(edit.start) || !Number.isInteger(edit.count) || edit.start < 1 || edit.count < 1
		|| edit.start + edit.count - 1 > limit || edit.count >= limit) throw new Error("行列范围无效。");
}

export function moveIndex(n: number, edit: StructureEdit): number | null {
	if (edit.kind === "clear" || n < edit.start) return n;
	if (edit.kind === "insert") return n + edit.count;
	return n < edit.start + edit.count ? null : n - edit.count;
}

/** Range deletion contracts the surviving interval; insertion inside expands it. */
export function moveInterval(from: number, to: number, edit: StructureEdit): [number, number] | null {
	if (from > to) { const next = moveInterval(to, from, edit); return next ? [next[1], next[0]] : null; }
	if (edit.kind !== "delete") return [moveIndex(from, edit)!, moveIndex(to, edit)!];
	const end = edit.start + edit.count - 1;
	if (to < edit.start) return [from, to];
	if (from > end) return [from - edit.count, to - edit.count];
	if (from >= edit.start && to <= end) return null;
	return [from < edit.start ? from : edit.start, to > end ? to - edit.count : edit.start - 1];
}

const CELL = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})$/i;
const COLUMN = /^(\$?)([A-Z]{1,3})$/i;
const ROW = /^(\$?)([1-9]\d{0,6})$/;

/** A1 ranges, whole-column and whole-row references; preserves absolute markers. */
export function moveReference(ref: string, edit: StructureEdit): string {
	const [a, b] = ref.split(":");
	const ma = a.match(CELL), mb = (b ?? a).match(CELL);
	if (ma && mb) {
		if ([ma, mb].some((m) => colNumber(m[2]) > MAX_COL || Number(m[4]) > MAX_ROW)) return ref;
		const rows = edit.axis === "row" ? moveInterval(Number(ma[4]), Number(mb[4]), edit) : [Number(ma[4]), Number(mb[4])];
		const cols = edit.axis === "column" ? moveInterval(colNumber(ma[2]), colNumber(mb[2]), edit) : [colNumber(ma[2]), colNumber(mb[2])];
		if (!rows || !cols || Math.max(...rows) > MAX_ROW || Math.max(...cols) > MAX_COL) return "#REF!";
		const start = `${ma[1]}${colName(cols[0])}${ma[3]}${rows[0]}`;
		return b ? `${start}:${mb[1]}${colName(cols[1])}${mb[3]}${rows[1]}` : start;
	}
	if (b) {
		const pattern = edit.axis === "row" ? ROW : COLUMN;
		const x = a.match(pattern), y = b.match(pattern);
		if (x && y) {
			const number = (s: string) => edit.axis === "row" ? Number(s) : colNumber(s);
			const limit = edit.axis === "row" ? MAX_ROW : MAX_COL;
			if (number(x[2]) > limit || number(y[2]) > limit) return ref;
			const interval = moveInterval(number(x[2]), number(y[2]), edit);
			if (!interval || Math.max(...interval) > limit) return "#REF!";
			const format = (n: number) => edit.axis === "row" ? String(n) : colName(n);
			return `${x[1]}${format(interval[0])}:${y[1]}${format(interval[1])}`;
		}
	}
	return ref;
}

const TOKEN = /^(?:(('(?:[^']|'')+'|[\p{L}\p{N}_\\][\p{L}\p{N}_.\\]*)!))?(\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[1-9]\d{0,6}:\$?[1-9]\d{0,6})/iu;
const NAME_CHAR = /[\p{L}\p{N}_.\\]/u;

/** Structural moves adjust absolute references too, unlike copying a formula. */
export function rewriteReferences(formula: string, ownerSheet: string | undefined, targetSheet: string, edit: StructureEdit): string {
	let result = "";
	for (let i = 0; i < formula.length;) {
		// Excel string literals and nested structured references are opaque tokens.
		if (formula[i] === '"' || formula[i] === "[") {
			const start = i;
			if (formula[i] === '"') {
				i++;
				while (i < formula.length) {
					if (formula[i++] !== '"') continue;
					if (formula[i] === '"') i++; else break;
				}
			} else {
				let depth = 0;
				do { if (formula[i] === "[") depth++; if (formula[i] === "]") depth--; i++; } while (i < formula.length && depth);
			}
			result += formula.slice(start, i);
			continue;
		}
		const match = (i === 0 || !NAME_CHAR.test(formula[i - 1])) ? formula.slice(i).match(TOKEN) : null;
		const after = match ? formula[i + match[0].length] : undefined;
		if (!match || (after && (NAME_CHAR.test(after) || after === "(" || after === "[" || after === "!"))) {
			result += formula[i++]; continue;
		}
		const qualifier = match[2];
		const sheet = qualifier ? (qualifier.startsWith("'") ? qualifier.slice(1, -1).replace(/''/g, "'") : qualifier) : ownerSheet;
		result += sheet?.toLowerCase() === targetSheet.toLowerCase() ? (match[1] ?? "") + moveReference(match[3], edit) : match[0];
		i += match[0].length;
	}
	return result;
}
