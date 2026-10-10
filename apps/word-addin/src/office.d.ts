/**
 * Type declarations for Office.js / Word API globals.
 * These are loaded from CDN in index.html, so they are available at runtime
 * but not discoverable by TypeScript without explicit declarations.
 */

declare namespace Word {
  function run<T>(callback: (context: Word.RequestContext) => Promise<T>): Promise<T>;

  interface RequestContext {
    document: Word.Document;
    sync(): Promise<void>;
  }

  interface Document {
    body: Word.Body;
    getSelection(): Word.Range;
  }

  interface Body extends Word.Range {
    insertText(text: string, location: Word.InsertLocation): Word.Range;
    insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
    paragraphs: Word.ParagraphCollection;
    /** Collection of all top-level tables in the document body.
     *  Use `body.tables.items` to enumerate tables; `table.values` and
     *  `table.rowCount`/`columnCount` to inspect contents. */
    tables: Word.TableCollection;
  }

  interface Range {
    text: string;
    font: Word.Font;
    style: string;
    insertText(text: string, location: Word.InsertLocation): Word.Range;
    insertHtml(html: string, location: Word.InsertLocation): Word.Range;
    insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
    insertTable(
      rowCount: number,
      columnCount: number,
      insertLocation: InsertLocation,
      values?: string[][]
    ): Word.Table;
    getRange(location?: string): Word.Range;
    insertInlinePictureFromBase64(base64Image: string, insertLocation: Word.InsertLocation): unknown;
    load(option: string | string[]): void;
    search(searchText: string, options?: Word.SearchOptions): Word.SearchRangeCollection;
    paragraphs: Word.ParagraphCollection;
    getFirst(): Word.Range;
    /** Deletes the range content from the document (WordApi 1.3+). */
    delete(): void;
    /** Returns HTML representation of the range; access .value after context.sync() */
    getHtml(preferredFragmentOnly?: boolean): { value: string };
  }

  interface Table {
    style: string;
    /** Built-in style name (e.g. "TableGrid"). Setting this is the supported
     *  way to apply a built-in Word table style. */
    styleBuiltIn: string;
    /** 2-D string matrix of cell contents. Each row is a string[], each
     *  column a string. Empty cells appear as "". Must be .load()-ed and
     *  context.sync()-ed before reading. */
    values: string[][];
    /** Number of rows in the table (must be loaded). */
    rowCount: number;
    /** Number of columns in the table (must be loaded). */
    columnCount: number;
    /** Deletes the entire table from the document (WordApi 1.3+). */
    delete(): void;
    /** Returns the range that spans the whole table (body-level), useful for
     *  re-locating the table inside the document. */
    getRange(rangeLocation?: string): Word.Range;
    /** Returns the parent body that owns this table (for table-level ops). */
    load(option: string | string[]): void;
  }

  interface TableCollection {
    items: Word.Table[];
    load(option: string | string[]): void;
    getFirst(): Word.Table;
    getLast(): Word.Table;
  }

  interface Font {
    name: string;
    size: number;
    color: string;
    bold: boolean | string;
    italic: boolean | string;
    underline: string;
    load(option: string | string[]): void;
  }

  interface Paragraph {
    text: string;
    style: string;
    isListItem: boolean;
    font: Word.Font;
    insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation): Word.Paragraph;
    getRange(rangeLocation: string): Word.Range;
    delete(): void;
    load(option: string | string[]): void;
  }

  interface ParagraphCollection {
    items: Word.Paragraph[];
    load(option: string | string[]): void;
    getFirst(): Word.Paragraph;
    getLast(): Word.Paragraph;
  }

  interface SearchOptions {
    matchCase?: boolean;
    matchWholeWord?: boolean;
    matchPrefix?: boolean;
    matchSuffix?: boolean;
  }

  interface SearchRangeCollection {
    items: Word.Range[];
    load(option: string | string[]): void;
  }

  type InsertLocation =
    | "Start"
    | "End"
    | "Before"
    | "After"
    | "Replace";
}

declare namespace Office {
  function onReady(callback?: (info: { host: string; platform: string }) => void): Promise<{ host: string; platform: string }>;
}

declare const Office: {
  onReady(callback?: (info?: { host: string; platform: string }) => void): Promise<{ host: string; platform: string }>;
};