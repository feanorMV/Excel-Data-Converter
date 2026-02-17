
export type FileType = 'ITEM_MASTER' | 'ITEM_MASTER_V2' | 'FACTS' | 'STORE_ITEMS' | 'STORE' | 'STOCK' | 'PRICE' | 'UNKNOWN';

export interface StatusUpdate {
    message: string;
    status: 'processing' | 'success' | 'error';
}

export type StatusUpdateCallback = (update: StatusUpdate) => void;

// This represents a row after reading from Excel
export type ExcelRow = Record<string, string | number | null>;

export interface CsvFile {
    name: string;
    content: string;
}
