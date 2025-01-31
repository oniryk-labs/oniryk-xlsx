/**
 * Converts a zero-based column index to an Excel column name
 * Generates column names in the format A, B, C, ..., Z, AA, AB, etc.
 *
 * @param index - Zero-based index of the column (0 = A, 1 = B, etc.)
 * @returns Excel-style column name (e.g., 'A' for 0, 'B' for 1, 'AA' for 26)
 */
export function columnName(index: number): string {
  let column = '';
  while (index >= 0) {
    column = String.fromCharCode(65 + (index % 26)) + column;
    index = Math.floor(index / 26) - 1;
  }
  return column;
}
