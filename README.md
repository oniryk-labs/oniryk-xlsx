# @oniryk/xlsx

<p align="center">
  <a href="https://www.npmjs.com/package/@oniryk/xlsx">
      <img src="https://img.shields.io/npm/v/@oniryk/xlsx.svg?style=for-the-badge" alt="npm version" />
  </a>
  <a href="https://www.npmjs.com/package/@oniryk/xlsx">
    <img src="https://img.shields.io/npm/dt/@oniryk/xlsx.svg?style=for-the-badge" alt="npm total downloads" />
  </a>
  <a href="https://www.npmjs.com/package/@oniryk/xlsx">
    <img src="https://img.shields.io/npm/dm/@oniryk/xlsx.svg?style=for-the-badge" alt="npm monthly downloads" />
  </a>
  <a href="https://www.npmjs.com/package/@oniryk/xlsx">
    <img src="https://img.shields.io/npm/l/@oniryk/xlsx.svg?style=for-the-badge" alt="npm license" />
  </a>
</p>

A lightweight, efficient TypeScript library for generating single-sheet Excel XLSX files with support for large datasets

## Features

- Optimized for large datasets
- Support for dates, numbers, and text content
- Custom column widths
- TypeScript types included

## Installation

```bash
npm install @oniryk/xlsx
```

## Quick Start

```typescript
import { SharedStrings, Sheet, build } from '@oniryk/xlsx';

// Create instances for string management and worksheet
const strings = new SharedStrings();
const sheet = new Sheet(strings);

// Add headers
sheet.addRow(['Name', 'Age', 'Date']);

// Add data
sheet.addRow(['John Doe', 25, new Date('2024-01-31')]);
sheet.addRow(['Jane Smith', 30, new Date('2024-02-15')]);

// Generate the Excel file
const buffer = await build(sheet, strings);

// Save to file or send as response
await fs.writeFile('output.xlsx', buffer);
```

## API Reference

### SharedStrings

Manages string deduplication across the workbook:

```typescript
const strings = new SharedStrings();
```

Methods:
- `add(str: string): number` - Adds a string to the shared strings table
- `size(): number` - Gets the total number of unique strings
- `destroy(): void` - Cleans up resources

### Sheet

Handles worksheet data and formatting:

```typescript
const sheet = new Sheet(sharedStrings);
```

Methods:
- `addRow(row: Row): void` - Adds a single row
- `addRows(rows: Row[]): void` - Adds multiple rows
- `rowsCount(): number` - Gets total row count
- `setColumWidth(index: number, width: number): void` - Sets width for a single column
- `setColumWidth(sizes: [number, number][]): void` - Sets widths for multiple columns

### Types

```typescript
type Cell = string | number | Date | null;
type Row = Cell[];
```

## Column Width Configuration

You can customize column widths either individually or in bulk:

```typescript
// Set single column width
sheet.setColumWidth(0, 15); // Set column A to width 15

// Set multiple column widths
sheet.setColumWidth([
  [0, 15], // Column A: width 15
  [1, 20], // Column B: width 20
  [2, 10]  // Column C: width 10
]);
```

## Date Handling

Dates are automatically converted to Excel's internal format and styled appropriately:

```typescript
sheet.addRow(['Date', new Date('2024-01-31')]);
```

## Limitations

Current version limitations:

1. **Single Sheet Only**: The library currently only supports generating Excel files with a single worksheet. Multi-sheet support is planned for future releases.
2. **Basic Styling**: Only basic cell formatting is supported (dates and numbers). Advanced styling features like colors, borders, and fonts are not yet implemented.

## Contributing

Contributions are welcome! Please ensure:

1. TypeScript types are maintained
2. Documentation is updated
3. Code follows the existing style

## License

ISC License
