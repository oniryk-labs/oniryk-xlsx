import { createReadStream, unlink, WriteStream } from 'fs';

/** XML entity mapping for special characters */
const entities = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
};

/**
 * Escapes special characters in a string for XML content
 * Converts &, <, >, ", and ' to their XML entity equivalents
 *
 * @param value - Value to escape
 * @returns Escaped string if input is string, otherwise returns original value
 *
 * @example
 * escapeXML('Hello < World') // returns 'Hello &lt; World'
 */
export function escapeXML(value: any) {
  if (typeof value !== 'string') return value;

  return value.replace(/[\x00-\x1F\x7F]/g, '?').replace(/[&<>"'\\]/g, (char: string) => {
    return entities[char as keyof typeof entities];
  });
}

/** Excel epoch date (December 30, 1899) used for date conversions */
const EPOCH = new Date(1899, 11, 30, 0, 0, 0).getTime();

/** Milliseconds in a day */
const DAY_MS = 24 * 60 * 60 * 1000;

/**
 * Converts a JavaScript Date to Excel serial number format
 * Excel uses days since December 30, 1899 as its internal date representation
 *
 * @param date - JavaScript Date object to convert
 * @returns Number of days since Excel epoch, with fractional part for time
 */
export function convertToExcelDate(date: Date) {
  return Math.round(((date.getTime() - EPOCH) / DAY_MS) * 100000) / 100000;
}

/**
 * Formats a date as YYYY-MM-DD
 *
 * @param date - Date to format
 * @returns Date string in YYYY-MM-DD format
 *
 * @example
 * simpleDateFormat(new Date('2024-01-31')) // returns '2024-01-31'
 */
export function simpleDateFormat(date: Date) {
  return date.toISOString().slice(0, 10);
}

/**
 * Type guard to check if a value is a valid Date object
 *
 * @param value - Value to check
 * @returns True if value is a Date instance
 */
export function isValidDate(value: any): value is Date {
  return value instanceof Date;
}

/**
 * Creates a read stream that automatically deletes the file when finished
 * Useful for temporary file handling
 *
 * @param filePath - Path to the file to stream
 * @returns ReadStream that will delete the source file when complete
 */
export function destructiveStream(filePath: string) {
  const stream = createReadStream(filePath);

  stream.on('end', () => {
    stream.destroy();
    unlink(filePath, (err) => {
      if (err) console.error('Error deleting file:', err);
    });
  });

  stream.on('error', (error) => {
    stream.destroy();
    console.error('Stream error:', error);
  });

  return stream;
}

/**
 * Simple performance measurement utility
 *
 * @param label - Label to identify the performance measurement
 * @returns Object with end function to log elapsed time
 *
 * @example
 * const timer = perf('operation');
 * // ... do something ...
 * timer.end(); // logs: "operation: XXXms"
 */
export function perf(label: string) {
  const start = performance.now();
  return {
    end: () => {
      console.log(`${label}: ${(performance.now() - start).toFixed(2)}ms`);
    },
  };
}

/** Function type for promised write operations */
export type PromisedWriter = (...content: string[]) => Promise<void>;

/**
 * Creates a promise-based writer function for a WriteStream
 * Allows for easier async/await usage of stream writes
 *
 * @param stream - WriteStream to wrap
 * @returns Promise-based write function
 *
 * @example
 * const writer = promiseWrite(writeStream);
 * await writer('content1', 'content2');
 */
export function promiseWrite(stream: WriteStream): PromisedWriter {
  return (...content: string[]) => {
    return new Promise<void>((resolve, reject) => {
      stream.write(content.join(''), (err) => (err ? reject(err) : resolve()));
    });
  };
}

/**
 * Safely finishes a write stream
 * Returns a promise that resolves when the stream is fully closed
 *
 * @param stream - WriteStream to finish
 * @returns Promise that resolves when stream is finished
 */
export function finishStream(stream: WriteStream): Promise<void> {
  return new Promise((resolve, reject) => {
    stream.on('finish', () => {
      stream.destroy();
      resolve();
    });
    stream.on('error', reject);
    stream.end();
  });
}
