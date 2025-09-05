import { clsx, type ClassValue } from "clsx"
import { twMerge } from "tailwind-merge"
import * as XLSX from "xlsx"

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}


/**
 * Expand cell merges in a worksheet, filling in missing cells with a reference to the top-left cell of the merge.
 *
 * This is a modified version of `XLSX.utils.sheet_add_json` that doesn't alter the original worksheet.
 *
 * @param ws The worksheet to modify.
 */
export function expandMerges(ws: XLSX.WorkSheet) {
  const merges = (ws["!merges"] || []) as XLSX.Range[];
  for (const m of merges) {
    const topLeft = XLSX.utils.encode_cell(m.s);
    const v = (ws as any)[topLeft]?.v;
    if (v === undefined) continue;

    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (!(ws as any)[addr]) (ws as any)[addr] = { t: "s", v }; // fill missing
      }
    }
  }
}

/**
   * Returns true if the given value has content. For strings, this means
   * non-whitespace. For other types, this means the value is not null or
   * undefined.
   * @param v the value to check
   * @returns true if the value has content, false otherwise
   */
export const hasContent = (v: unknown) => {
  if (v == null) return false;
  if (typeof v === "string") return v.trim().length > 0; // whitespace == empty
  return true;
}

/**
 * Returns true if the given cell has content. For strings, this means
 * non-whitespace. For other types, this means the value is not null or
 * undefined.
 * @param {XLSX.CellObject | undefined} cell the cell to check
 * @returns {boolean} true if the cell has content, false otherwise
 */
export const isCellNonEmpty = (cell: XLSX.CellObject | undefined) => {
  if (!cell || cell.v == null) return false;
  if (typeof cell.v === "string") return cell.v.trim().length > 0;
  return true;
}

/**
 * Returns a set of row numbers in the given worksheet that contain at least
 * one non-empty cell. The row numbers are 1-indexed, i.e., the first row is 1.
 * @param {XLSX.WorkSheet} ws the worksheet to scan
 * @returns {Set<number>} the set of row numbers with non-empty cells
 */
export function getNonEmptyRowSet(ws: XLSX.WorkSheet): Set<number> {
  const keep = new Set<number>();
  const ref = ws["!ref"];
  if (!ref) return keep;
  const range = XLSX.utils.decode_range(ref);
  for (let r = range.s.r; r <= range.e.r; r++) {
    let hasAny = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (isCellNonEmpty((ws as any)[addr])) { hasAny = true; break; }
    }
    if (hasAny) keep.add(r + 1);
  }
  return keep;
}