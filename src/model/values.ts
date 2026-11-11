export interface ErrorValue {
	error: string;
}

/** A cell's value as read from the file. Dates are UTC `Date`s. */
export type CellValue = number | string | boolean | Date | ErrorValue | null;

/** What can be written into a cell. */
export type CellInput =
	| { kind: "empty" }
	| { kind: "number"; value: number }
	| { kind: "boolean"; value: boolean }
	| { kind: "text"; value: string }
	| { kind: "date"; value: Date }
	| { kind: "formula"; formula: string };

const DAY_MS = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

export function serialToDate(serial: number, date1904: boolean): Date {
	const ms = (date1904 ? EPOCH_1904 : EPOCH_1900) + serial * DAY_MS;
	return new Date(Math.round(ms));
}

export function dateToSerial(date: Date, date1904: boolean): number {
	return (date.getTime() - (date1904 ? EPOCH_1904 : EPOCH_1900)) / DAY_MS;
}

/** Context keeps identifiers and cells formatted as Text from changing type. */
export interface InputContext { value?: CellValue; format?: string }

export function parseInput(text: string, context: InputContext = {}): CellInput {
	if (text === "") return { kind: "empty" };
	if (text.startsWith("'")) return { kind: "text", value: text.slice(1) };
	if (context.format === "@") return { kind: "text", value: text };
	if (text.startsWith("=") && text.length > 1) return { kind: "formula", formula: text.slice(1) };
	// Existing text remains text, including numeric labels, dates and TRUE/FALSE.
	if (typeof context.value === "string" || text !== text.trim()) return { kind: "text", value: text };
	if (/^(TRUE|FALSE)$/i.test(text)) return { kind: "boolean", value: text.toUpperCase() === "TRUE" };
	if (/^[+-]?\d+(?:[.,]\d+)?$/.test(text)) {
		const unsigned = text.replace(/^[+-]/, "");
		const significant = unsigned.replace(/[.,]/g, "").replace(/^0+/, "");
		// Excel itself only guarantees 15 significant decimal digits.
		if (!/^0\d/.test(unsigned) && significant.length <= 15 && Number.isFinite(Number(text.replace(",", ".")))) {
			return { kind: "number", value: Number(text.replace(",", ".")) };
		}
		return { kind: "text", value: text };
	}
	const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/);
	if (iso) {
		const input = dateInput(Number(iso[1]), Number(iso[2]), Number(iso[3]));
		const h = Number(iso[4] ?? 0), m = Number(iso[5] ?? 0), sec = Number(iso[6] ?? 0);
		if (input?.kind === "date" && h < 24 && m < 60 && sec < 60) {
			input.value.setUTCHours(h, m, sec, Number((iso[7] ?? "0").padEnd(3, "0")));
			return input;
		}
	}
	const dmy = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
	if (dmy) return dateInput(Number(dmy[3]), Number(dmy[2]), Number(dmy[1])) ?? { kind: "text", value: text };
	return { kind: "text", value: text };
}

function dateInput(year: number, month: number, day: number): CellInput | undefined {
	const date = new Date(0);
	date.setUTCFullYear(year, month - 1, day);
	// Rejects impossible dates such as 31/02/2026 instead of rolling them over.
	if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return undefined;
	return { kind: "date", value: date };
}

/** The text put in the editor when a cell is edited. */
export function editText(value: CellValue, formula: string | undefined): string {
	if (formula) return `=${formula}`;
	if (value === null) return "";
	if (value instanceof Date) {
		const iso = value.toISOString();
		return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso.slice(0, -1).replace("T", " ");
	}
	if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
	if (typeof value === "object") return value.error;
	return String(value);
}

/** Reject values that Excel/XML cannot store before the package is changed. */
export function validateInput(input: CellInput) {
	if (input.kind === "text" || input.kind === "formula") {
		const text = input.kind === "text" ? input.value : input.formula;
		if (text.length > (input.kind === "text" ? 32767 : 8192)) throw new Error("单元格文本或公式超过 Excel 长度限制。");
		if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(text)) throw new Error("内容包含 Excel 不支持的控制字符。");
	}
	if (input.kind === "number" && !Number.isFinite(input.value)) throw new Error("无法保存非有限数字。");
	if (input.kind === "date" && !Number.isFinite(input.value.getTime())) throw new Error("无效日期。");
}
