import { randomUUID } from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { columnName } from './column.js';
import type SharedStrings from './shared-strings.js';
import {
  convertToExcelDate as date,
  finishStream,
  isValidDate,
  PromisedWriter,
  promiseWrite,
} from './utils.js';

/** Represents a cell value in an Excel sheet */
export type Cell = string | number | Date | null;

/** Represents a row of cells in an Excel sheet */
export type Row = Cell[];

/** Function type for mapping column properties to XML string */
export type ColMapper = ([k, w]: [number, number]) => string;

/**
 * Maps column properties to their XML representation
 * @param k - Column index (0-based)
 * @param w - Column width
 * @returns XML string representing column properties
 */
const mapper: ColMapper = ([k, w]) => {
  return `<col min="${k + 1}" max="${k + 1}" width="${w}" customWidth="1"/>`;
};

/**
 * Represents an Excel worksheet
 * Handles the creation and management of worksheet data including rows, columns, and cell formatting
 */
export default class Sheet {
  /** Map of pre-calculated column names (A-ZZ) */
  private static readonly COLUMNS: Map<number, string> = new Map();

  /** Number of rows to process in each chunk when generating XML */
  private static readonly CHUNK_SIZE = 2500;

  /** Storage for worksheet rows */
  private rows: Row[] = [];

  /** Reference to shared strings table */
  private sharedStrings: SharedStrings;

  /** Temporary file path for the worksheet XML */
  private file: string;

  /** Map of column indices to their widths */
  private columnWidths = new Map<number, number>();

  /** Initialize column names map */
  static {
    for (let i = 0, ZZ = 702; i < ZZ; i += 1) {
      Sheet.COLUMNS.set(i, columnName(i));
    }
  }

  /**
   * Creates a new Sheet instance
   * @param sharedStrings - SharedStrings instance for managing string deduplication
   */
  constructor(sharedStrings: SharedStrings) {
    this.sharedStrings = sharedStrings;
    this.file = path.join(os.tmpdir(), `xlsx-${randomUUID()}.xml`);
  }

  /**
   * Adds a single row to the worksheet
   * @param row - Array of cell values
   */
  addRow(row: Row) {
    this.rows.push(row);
  }

  /**
   * Adds multiple rows to the worksheet
   * @param rows - Array of rows to add
   */
  addRows(rows: Row[]) {
    this.rows = [...this.rows, ...rows];
  }

  /**
   * Gets the total number of rows in the worksheet
   * @returns Number of rows
   */
  rowsCount(): number {
    return this.rows.length;
  }

  /**
   * Sets the width for a single column or multiple columns
   * @param index - Column index or array of [index, width] pairs
   * @param width - Column width (when setting single column)
   */
  setColumWidth(index: number, width: number): void;
  setColumWidth(sizes: [number, number][]): void;
  setColumWidth(indexOrSizes: number | [number, number][], width?: number): void {
    if (Array.isArray(indexOrSizes)) {
      for (const [index, width] of indexOrSizes) {
        this.columnWidths.set(index, width);
      }
    } else if (typeof indexOrSizes === 'number' && typeof width === 'number') {
      this.columnWidths.set(indexOrSizes, width);
    }
  }

  /**
   * Gets the Excel column name for a given index
   * @param index - Zero-based column index
   * @returns Excel column name (e.g., 'A', 'B', 'AA')
   * @private
   */
  private getExcelColumn(index: number): string {
    return Sheet.COLUMNS.get(index) || columnName(index);
  }

  /**
   * Writes a chunk of rows to the worksheet XML
   * @param write - Promise-based writer function
   * @param rows - Array of rows to write
   * @param startIndex - Starting row index for this chunk
   * @returns Promise that resolves when chunk is written
   * @private
   */
  private async writeChunk(
    write: PromisedWriter,
    rows: Row[],
    startIndex: number
  ): Promise<void> {
    const add = (str: string) => this.sharedStrings.add(str);
    const chunks: string[] = [];

    for (let i = 0, rlen = rows.length; i < rlen; i += 1) {
      const cols = rows[i];
      const row = (startIndex + i + 1).toString();

      chunks.push(`<row r="${row}">`);

      for (let ci = 0, clen = cols.length; ci < clen; ci += 1) {
        const cell = cols[ci];
        const ref = `${this.getExcelColumn(ci)}${row}`;

        if (cell === null || cell === undefined) {
          chunks.push(`<c r="${ref}"/>`);
          continue;
        }

        switch (typeof cell) {
          case 'string':
            chunks.push(`<c r="${ref}" t="s"><v>${add(cell)}</v></c>`);
            continue;
          case 'number':
            chunks.push(`<c r="${ref}"><v>${cell}</v></c>`);
            continue;
          case 'object':
            if (isValidDate(cell)) {
              chunks.push(`<c r="${ref}" s="1"><v>${date(cell)}</v></c>`);
            } else {
              chunks.push(`<c r="${ref}" t="s"><v>${add(String(cell))}</v></c>`);
            }
            continue;
          default:
            chunks.push(`<c r="${ref}" t="s"><v>${add(String(cell))}</v></c>`);
        }
      }

      chunks.push('</row>');
    }

    const content = chunks.join('');
    chunks.length = 0;

    await write(content);
  }

  /**
   * Generates the complete worksheet XML file
   * @returns Promise that resolves to the path of the generated XML file
   */
  async generateSheetXML(): Promise<string> {
    const ws = fs.createWriteStream(this.file);
    const write = promiseWrite(ws);

    await write(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    );

    if (this.columnWidths.size > 0) {
      await write(
        '<cols>',
        [...this.columnWidths.entries()].map(mapper).join(''),
        '</cols>'
      );
    }

    await write('<sheetData>');

    for (let i = 0; i < this.rows.length; i += Sheet.CHUNK_SIZE) {
      const chunk = this.rows.slice(i, i + Sheet.CHUNK_SIZE);
      await this.writeChunk(write, chunk, i);
    }

    await write('</sheetData></worksheet>');
    await finishStream(ws);

    return this.file;
  }
}
