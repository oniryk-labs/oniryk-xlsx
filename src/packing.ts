import AdmZip from 'adm-zip';
import fs from 'fs';
import path from 'path';
import SharedStrings from './shared-strings.js';
import Sheet from './sheet.js';
import getBaseRels from './templates/base-rels.js';
import getContentTypes from './templates/content-type.js';
import generateStyles from './templates/styles.js';
import { getWorkbookRels, getWorkbookXML } from './templates/workbook.js';

/**
 * Manages the creation of Excel XLSX packages
 * Handles the assembly of various XML components into a final ZIP archive
 * following the Office Open XML SpreadsheetML format
 */
export class Package {
  /** ZIP archive instance for creating the final XLSX file */
  private zip: AdmZip;

  /** SharedStrings instance for managing text content */
  private sharedStrings: SharedStrings;

  /** Sheet instance containing the worksheet data */
  private sheet: Sheet;

  /** List of temporary files that need cleanup */
  private disposable: string[] = [];

  /**
   * Creates a new Package instance
   * @param sheet - Sheet instance containing worksheet data
   * @param sharedStrings - SharedStrings instance for text content management
   */
  constructor(sheet: Sheet, sharedStrings: SharedStrings) {
    this.zip = new AdmZip();
    this.sharedStrings = sharedStrings;
    this.sheet = sheet;
  }

  /**
   * Adds string content to the ZIP archive
   * @param path - Path within the ZIP archive
   * @param content - String content to add
   * @private
   */
  private add(path: string, content: string) {
    this.zip.addFile(path, Buffer.from(content, 'utf8'));
  }

  /**
   * Adds a local file to the ZIP archive
   * Tracks the file for later cleanup
   * @param file - Target path within the ZIP archive
   * @param localpath - Local filesystem path of the file to add
   * @private
   */
  private addFile(file: string, localpath: string) {
    const zipName = path.basename(file);
    const zipPath = path.dirname(file);
    this.zip.addLocalFile(localpath, zipPath, zipName);
    this.disposable.push(localpath);
  }

  /**
   * Adds required Excel template files to the ZIP archive
   * Includes relationships, workbook, styles, and content type definitions
   * @private
   */
  private addMockFiles() {
    this.add('_rels/.rels', getBaseRels());
    this.add('xl/workbook.xml', getWorkbookXML());
    this.add('xl/_rels/workbook.xml.rels', getWorkbookRels());
    this.add('xl/styles.xml', generateStyles());
    this.add('[Content_Types].xml', getContentTypes());
  }

  /**
   * Builds the final XLSX package
   * Assembles all components, generates required XML files,
   * and creates the ZIP archive
   * @returns Promise that resolves to the XLSX file as a Buffer
   */
  public async build(): Promise<Buffer> {
    this.addMockFiles();
    const file = await this.sheet.generateSheetXML();
    this.addFile('xl/worksheets/sheet1.xml', file);

    if (this.sharedStrings.size() > 0) {
      const content = await this.sharedStrings.generateSharedStringsXML();
      this.addFile('xl/sharedStrings.xml', content);
    }

    const buffer = this.zip.toBuffer();
    this.dispose();
    return buffer;
  }

  /**
   * @deprecated Use build() instead
   * Legacy method for package creation
   * @returns Promise that resolves to the XLSX file as a Buffer
   */
  public async pack(): Promise<Buffer> {
    console.warn('deprecated: use build() instead');
    return await this.build();
  }

  /**
   * Cleans up temporary files created during package assembly
   * @private
   */
  private dispose() {
    for (const file of this.disposable) {
      fs.unlink(file, () => {});
    }
  }
}
