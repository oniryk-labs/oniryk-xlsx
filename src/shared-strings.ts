import { randomUUID } from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { escapeXML, finishStream, promiseWrite } from './utils';

/**
 * Manages shared strings in an Excel workbook
 * Handles deduplication and XML generation for the shared strings table
 * Used to optimize file size by storing repeated strings only once
 */
export default class SharedStrings {
  /** Array of unique strings in the workbook */
  private strings: string[] = [];

  /** Map of strings to their indices for quick lookup */
  private stringIndexMap = new Map<string, number>();

  /** Temporary file path for the shared strings XML */
  private file: string;

  /**
   * Creates a new SharedStrings instance
   * Initializes a temporary file for XML generation
   */
  constructor() {
    this.file = path.join(os.tmpdir(), `xlsx-strings-${randomUUID()}.xml`);
  }

  /**
   * Adds a string to the shared strings table
   * If the string already exists, returns its index
   * If the string is new, adds it and returns the new index
   *
   * @param str - String to add to the shared strings table
   * @returns Index of the string in the shared strings table
   */
  public add(str: string): number {
    if (this.stringIndexMap.has(str)) {
      return this.stringIndexMap.get(str)!;
    }

    const index = this.strings.push(str) - 1;
    this.stringIndexMap.set(str, index);

    return index;
  }

  /**
   * Generates the shared strings XML file
   * Creates an XML file containing all unique strings in the workbook
   * Processes strings in chunks to manage memory usage
   *
   * @returns Promise that resolves to the path of the generated XML file
   */
  public async generateSharedStringsXML(): Promise<string> {
    const ws = fs.createWriteStream(this.file);
    const write = promiseWrite(ws);

    await write(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`,
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            count="${this.strings.length}" uniqueCount="${this.strings.length}">`
    );

    const acc = [];

    for (const str of this.strings) {
      acc.push(`<si><t>${escapeXML(str)}</t></si>`);

      if (acc.length % 5000 === 0) {
        await write(acc.join(''));
        acc.length = 0;
      }
    }

    await write(acc.join(''), '</sst>');
    await finishStream(ws);

    return this.file;
  }

  /**
   * Gets the total number of unique strings in the table
   * @returns Number of unique strings
   */
  public size(): number {
    return this.strings.length;
  }

  /**
   * Clears all strings and mappings from memory
   * Should be called when the shared strings table is no longer needed
   */
  public destroy(): void {
    this.strings.length = 0;
    this.stringIndexMap.clear();
  }
}
