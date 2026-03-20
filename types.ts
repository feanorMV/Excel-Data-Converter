export type FileType = 'ITEM_MASTER' | 'ITEM_MASTER_V2' | 'ITEM_MASTER_UPDATED' | 'FACTS' | 'STORE_ITEMS' | 'STORE' | 'STOCK' | 'PRICE' | 'UNKNOWN';

export interface StatusUpdate {
    message: string;
    status: 'processing' | 'success' | 'error';
    progress?: number; // 0 to 100
}

export type StatusUpdateCallback = (update: StatusUpdate) => void;

// This represents a row after reading from Excel
// FIX: Added `Date` to the union type to allow for Date objects when parsing
// Excel files. This resolves a type error in `services/excelProcessor.ts` where
// an `instanceof Date` check was performed on a value of this type.
export type ExcelRow = Record<string, string | number | Date | null>;

export interface CsvFile {
    name: string;
    content: string;
    rowCount: number;
}

export interface CsvGenerationOptions {
    [key: string]: boolean | string | Record<string, string> | undefined;
    delimiter?: string;
    columnMapping?: Record<string, string>;
}