import { Package } from './packing.js';
import type SharedStrings from './shared-strings.js';
import type Sheet from './sheet.js';

/**
 * Creates an Excel XLSX file from the provided sheet and shared strings
 * Convenience function that wraps the Package class functionality
 *
 * @param sheet - Sheet instance containing the worksheet data
 * @param strings - SharedStrings instance managing text content
 * @returns Promise that resolves to the XLSX file as a Buffer
 *
 * @example
 * const strings = new SharedStrings();
 * const sheet = new Sheet(strings);
 * sheet.addRow(['Name', 'Age']);
 * sheet.addRow(['John', 25]);
 * const buffer = await build(sheet, strings);
 */
export async function build(sheet: Sheet, strings: SharedStrings): Promise<Buffer> {
  const pack = new Package(sheet, strings);
  return await pack.build();
}
