
import React, { useState, useCallback, useRef } from 'react';
import ReactDOM from 'react-dom/client';
import { FolderUp, Download, RefreshCw, AlertTriangle, CheckCircle, Loader2, FileCheck2, FileText, XCircle } from 'lucide-react';

// ==================================================================================
// MERGED FROM types.ts
// ==================================================================================
type FileType = 'ITEM_MASTER' | 'ITEM_MASTER_V2' | 'FACTS' | 'STORE_ITEMS' | 'STORE' | 'STOCK' | 'PRICE' | 'UNKNOWN';

interface StatusUpdate {
    message: string;
    status: 'processing' | 'success' | 'error';
}

type StatusUpdateCallback = (update: StatusUpdate) => void;

type ExcelRow = Record<string, string | number | Date | null>;

interface CsvFile {
    name: string;
    content: string;
}

type CsvGenerationOptions = Record<string, boolean>;


// ==================================================================================
// MERGED FROM services/excelProcessor.ts
// ==================================================================================
declare var XLSX: any;

// --- UTILITY FUNCTIONS ---

function getTodayDateString(): string {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
}

function arrayToCsv(data: Record<string, any>[], columns: string[]): string {
    const header = columns.join(',') + '\n';
    const rows = data.map(row => {
        return columns.map(col => {
            let value = row[col];
            if (value === null || typeof value === 'undefined') {
                return '';
            }
            value = String(value);
            if (value.includes('"') || value.includes(',')) {
                return `"${value.replace(/"/g, '""')}"`;
            }
            return value;
        }).join(',');
    }).join('\n');
    return header + rows;
}

function excelSerialDateToJSDate(serial: number): Date {
    // 25569 is the number of days from 1900-01-01 to 1970-01-01 (epoch).
    const utcMilliseconds = (serial - 25569) * 86400 * 1000;
    const utcDate = new Date(utcMilliseconds);
    // Create a new Date object in the local timezone using the UTC date parts.
    // This correctly translates the calendar date without timezone shifts.
    return new Date(utcDate.getUTCFullYear(), utcDate.getUTCMonth(), utcDate.getUTCDate());
}


function formatDateToYYYYMMDD(date: Date): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// --- FILE TYPE DETECTION ---

const FILE_TYPE_DEFINITIONS = {
    STORE: { keywords: ["Store UID*"] },
    STORE_ITEMS: { keywords: ["Store UID*", "Product UID*", "In assortment?", "Purchase price"] },
    FACTS: { keywords: ["Product UID*", "Store UID*", "Date*"] },
    ITEM_MASTER_V2: { keywords: ["UID*", "Product name*", "Manufacturer UID"] },
    ITEM_MASTER: { keywords: ["UID*", "Product name*", "Barcode", "Manufacturer"] },
    STOCK: { keywords: ['StoreID', 'ItemUID', 'Quantity'] },
    PRICE: { keywords: ['ItemUID', 'PriceList', 'Price'] },
};


const detectFileType = (workbook: any): { type: FileType; sheetName: string | null } => {
    const sheetNames = workbook.SheetNames;

    // The order of checks matters. More specific checks should come first.
    const typeCheckOrder: FileType[] = ['ITEM_MASTER_V2', 'ITEM_MASTER', 'STORE_ITEMS', 'FACTS', 'STORE', 'STOCK', 'PRICE'];

    for (const type of typeCheckOrder) {
        const definition = FILE_TYPE_DEFINITIONS[type as keyof typeof FILE_TYPE_DEFINITIONS];
        if (!definition) continue;

        for (const sheetName of sheetNames) {
            const sheet = workbook.Sheets[sheetName];
            if (sheet) {
                const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
                // Search first 20 rows for keywords
                for (let i = 0; i < Math.min(20, jsonData.length); i++) {
                    const row = jsonData[i];
                    if (row && row.some(cell => typeof cell === 'string' && definition.keywords.every(kw => row.join('|').includes(kw)))) {
                        return { type: type as FileType, sheetName: sheetName }; // Found a match
                    }
                }
            }
        }
    }
    return { type: 'UNKNOWN', sheetName: null }; // No match in any sheet
};

// --- SPECIALIZED PROCESSORS ---

/**
 * Processes a "Stores" file.
 */
function processStoresFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback): CsvFile[] {
    updateStatus({ message: `Processing Stores file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });

    let headerRowIndex = -1;
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && FILE_TYPE_DEFINITIONS.STORE.keywords.some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error('Could not find a valid header row in the Stores sheet.');
    
    const headers = rawData[headerRowIndex].map(h => String(h || '').trim());
    const dataRows = rawData.slice(headerRowIndex + 1);

    const storesData = dataRows.map(row => {
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });
        
        return {
            store_uid: rowObject['Store UID*'],
            name: rowObject['Store name*'],
            region: rowObject['Region'],
            group_name: rowObject['Group name'],
            floor_space: parseInt(String(rowObject['Square']), 10) || 0,
            in_shelf: parseInt(String(rowObject['In Shelf?']), 10) || 0,
            licence_start_date: '2023-01-01',
            is_deleted: parseInt(String(rowObject['To delete']), 10) || 0,
        };
    }).filter(row => row.store_uid);

    if (storesData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Stores file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `stores_${dateStr}.csv`,
        content: arrayToCsv(storesData, ['store_uid', 'name', 'region', 'group_name', 'floor_space', 'in_shelf', 'licence_start_date', 'is_deleted'])
    }];

    updateStatus({ message: 'Stores processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes a "Store Items" file containing assortment, pricing, and supplier data.
 */
function processStoreItemsFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions): CsvFile[] {
    updateStatus({ message: `Processing Store Items file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });

    let headerRowIndex = -1;
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && FILE_TYPE_DEFINITIONS.STORE_ITEMS.keywords.some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error('Could not find a valid header row in the Items sheet.');
    
    const headers = rawData[headerRowIndex].map(h => String(h || '').trim());
    const dataRows = rawData.slice(headerRowIndex + 1);

    const itemsData: any[] = [];
    const suppliersMap = new Map<string, any>();

    dataRows.forEach(row => {
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });

        const storeUid = rowObject['Store UID*'];
        const itemUid = rowObject['Product UID*'];

        if (storeUid && itemUid) {
            itemsData.push({
                store_uid: storeUid,
                item_uid: itemUid,
                is_active_planogram: parseInt(String(rowObject['In assortment?']), 10) || 0,
                purchase_price: parseFloat(String(rowObject['Purchase price'] || '0').replace(',', '.')) || null,
                retail_price: parseFloat(String(rowObject['Sale price'] || '0').replace(',', '.')) || null,
                external_supplier_uid: rowObject['Supplier UID']
            });

            const supplierUid = rowObject['Supplier UID'];
            if (supplierUid && !suppliersMap.has(String(supplierUid))) {
                suppliersMap.set(String(supplierUid), {
                    supplier_uid: supplierUid,
                    name: rowObject['Supplier'],
                    is_deleted: 0
                });
            }
        }
    });
    
    if (itemsData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Items file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];

    csvs.push({
        name: `items_${dateStr}.csv`,
        content: arrayToCsv(itemsData, ['item_uid', 'store_uid', 'is_active_planogram', 'purchase_price', 'retail_price', 'external_supplier_uid'])
    });

    if ((options.suppliers ?? true) && suppliersMap.size > 0) {
        csvs.push({
            name: `suppliers_${dateStr}.csv`,
            content: arrayToCsv(Array.from(suppliersMap.values()), ['supplier_uid', 'name', 'is_deleted'])
        });
    }
    
    updateStatus({ message: 'Store Items processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes a "Facts" file containing sales and stock data.
 */
function processFactsFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback): CsvFile[] {
    updateStatus({ message: `Processing Facts file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });

    let headerRowIndex = -1;
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && FILE_TYPE_DEFINITIONS.FACTS.keywords.some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error('Could not find a valid header row in the Facts sheet.');

    const headers = rawData[headerRowIndex].map(h => String(h || '').trim());
    const dataRows = rawData.slice(headerRowIndex + 1);

    const factsData = dataRows.map(row => {
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });

        let formattedDate: string | null = null;
        const dateValue = rowObject['Date*'];

        if (dateValue) {
            let jsDate: Date | null = null;
            if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
                // Path 1: Already a valid Date object (from cellDates:true)
                jsDate = dateValue;
            } else if (typeof dateValue === 'number' && dateValue > 1) {
                // Path 2: Excel serial number
                jsDate = excelSerialDateToJSDate(dateValue);
            } else {
                // Path 3: Try to parse from a string
                const d = new Date(String(dateValue));
                if (!isNaN(d.getTime())) {
                    jsDate = d;
                }
            }
            
            if (jsDate) {
                formattedDate = formatDateToYYYYMMDD(jsDate);
            }
        }

        return {
            item_uid: rowObject['Product UID*'],
            store_uid: rowObject['Store UID*'],
            date: formattedDate,
            stock: parseFloat(String(rowObject['Stock'] || '0').replace(',', '.')) || null,
            sold_qty: parseFloat(String(rowObject['Out sale'] || '0').replace(',', '.')) || null,
            revenue: parseFloat(String(rowObject['Revenue'] || '0').replace(',', '.')) || null,
            cogs: parseFloat(String(rowObject['COGS'] || '0').replace(',', '.')) || null,
        };
    }).filter(row => row.item_uid && row.store_uid && row.date);
    
    if (factsData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Facts file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `facts_${dateStr}.csv`,
        content: arrayToCsv(factsData, ["item_uid", "store_uid", "date", "stock", "sold_qty", "revenue", "cogs"])
    }];
    
    updateStatus({ message: 'Facts processing complete.', status: 'success' });
    return csvs;
}


/**
 * Processes the original "Item Master" file format (V1).
 */
function processItemMasterFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions): CsvFile[] {
    updateStatus({ message: `Processing Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);
    
    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });

    let headerRowIndex = -1;
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && FILE_TYPE_DEFINITIONS.ITEM_MASTER.keywords.some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error('Could not find header row in Masteritems file.');

    const headers = rawData[headerRowIndex].map(h => String(h || '').trim());
    let dataStartIndex = -1;
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
        if(rawData[i] && rawData[i].some(cell => cell !== null && String(cell).trim() !== '')) {
            dataStartIndex = i;
            break;
        }
    }
    
    const dataRows = dataStartIndex !== -1 ? rawData.slice(dataStartIndex) : [];
    const df_template: ExcelRow[] = dataRows.map(row => {
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });
        return rowObject;
    });

    const masteritemsData: any[] = [], barcodesData: any[] = [], dimensionsData: any[] = [];
    const seenBrands = new Set<string>(), seenManufacturers = new Set<string>(), seenErpCategories = new Map<string, any>(), seenUids = new Set<string>();
    const levelMapping = { 1: 'Segment', 2: 'Family', 3: 'Class', 4: 'Brick' };

    df_template.forEach(row => {
        const uidValue = row['UID*'];
        if (uidValue === null || uidValue === undefined) return;
        const uid = String(uidValue).trim();
        if (uid === '') return;
        
        let main_unit_uid = row['Main Unit UID'];
        if (main_unit_uid === null || main_unit_uid === undefined || String(main_unit_uid).trim() === '') {
            const unitName = row['Unit'];
            if (unitName && String(unitName).trim()) {
                main_unit_uid = `${uid}_${String(unitName).trim()}`;
            } else {
                main_unit_uid = `${uid}_01`; // Fallback to a numbered UID
            }
        }

        if (!seenUids.has(uid)) {
            masteritemsData.push({ item_uid: uid, name: row['Product name*'], manufacturer_uid: row['Manufacturer'], brand_uid: row['Brand'], is_fractional: row['Is fractional?'] ? parseInt(String(row['Is fractional?']), 10) || 0 : 0, additional_1: row['Segment Description'], additional_2: row['Family Description'], additional_3: row['Class Description'], additional_4: row['Brick Description'], main_unit_uid: main_unit_uid, erp_category_uid: row['Brick Code'], });
            seenUids.add(uid);
        }
        if (row['Barcode']) barcodesData.push({ item_uid: uid, barcode: row['Barcode'], is_main: 1 });
        if (row['Brand']) seenBrands.add(String(row['Brand']));
        if (row['Unit']) dimensionsData.push({ item_uid: uid, unit_name: row['Unit'], width: parseFloat(String(row['Width'] || '').replace(',', '.')) || 0, height: parseFloat(String(row['Height'] || '').replace(',', '.')) || 0, depth: parseFloat(String(row['Length'] || '').replace(',', '.')) || 0, netweight: parseFloat(String(row['Netweight'] || '').replace(',', '.')) || null, volume: parseFloat(String(row['Volume'] || '').replace(',', '.')) || null, dimension_uid: main_unit_uid, coef: 1, is_deleted: 0, });
        Object.values(levelMapping).forEach((levelName, index) => {
            const uidCol = `${levelName} Code`, nameCol = `${levelName} Description`, catUid = row[uidCol];
            if (catUid && !seenErpCategories.has(String(catUid))) {
                const parentLevelName = index > 0 ? levelMapping[index as keyof typeof levelMapping] : null, parentUidCol = parentLevelName ? `${parentLevelName} Code` : null;
                seenErpCategories.set(String(catUid), { erp_category_uid: catUid, name: row[nameCol], parent_category_uid: parentUidCol ? row[parentUidCol] : null, });
            }
        });
        if (row['Manufacturer']) seenManufacturers.add(String(row['Manufacturer']));
    });

    if (masteritemsData.length === 0) {
        updateStatus({ message: 'No valid item data found in Masteritems file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    csvs.push({ name: `masteritems_${dateStr}.csv`, content: arrayToCsv(masteritemsData, ['item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional', 'additional_1', 'additional_2', 'additional_3', 'additional_4', 'main_unit_uid', 'erp_category_uid']) });
    
    if((options.barcodes ?? true) && barcodesData.length > 0) csvs.push({ name: `barcodes_${dateStr}.csv`, content: arrayToCsv(barcodesData, ['item_uid', 'barcode', 'is_main']) });
    if((options.brands ?? true) && seenBrands.size > 0) csvs.push({ name: `brands_${dateStr}.csv`, content: arrayToCsv([...seenBrands].map(b => ({ brand_uid: b, name: b, is_deleted: 0 })), ['brand_uid', 'name', 'is_deleted']) });
    if((options.dimensions ?? true) && dimensionsData.length > 0) csvs.push({ name: `dimensions_${dateStr}.csv`, content: arrayToCsv(dimensionsData, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'netweight', 'volume', 'dimension_uid', 'coef', 'is_deleted']) });
    if ((options.erpcategories ?? true) && seenErpCategories.size > 0) csvs.push({ name: `erpcategories_${dateStr}.csv`, content: arrayToCsv([...seenErpCategories.values()], ['erp_category_uid', 'name', 'parent_category_uid']) });
    if((options.manufacturers ?? true) && seenManufacturers.size > 0) csvs.push({ name: `manufacturers_${dateStr}.csv`, content: arrayToCsv([...seenManufacturers].map(m => ({ manufacturer_uid: m, name: m, is_deleted: 0 })), ['manufacturer_uid', 'name', 'is_deleted']) });
    
    updateStatus({ message: 'Masteritems processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes the new, complex "Item Master" file format (V2).
 */
function processItemMasterV2File(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions): CsvFile[] {
    updateStatus({ message: `Processing new Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });

    let headerRowIndex = -1;
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && FILE_TYPE_DEFINITIONS.ITEM_MASTER_V2.keywords.some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error('Could not find header row in the new Masteritems file.');
    
    const headers = rawData[headerRowIndex].map(h => String(h || '').trim());
    const dataRows = rawData.slice(headerRowIndex + 1);
    
    const df_template: ExcelRow[] = dataRows.map(row => {
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });
        return rowObject;
    }).filter(row => row['UID*'] !== null && row['UID*'] !== undefined && String(row['UID*']).trim() !== '');

    if (df_template.length === 0) {
        updateStatus({ message: 'No valid item data found in new Masteritems file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const masteritemsData: any[] = [], barcodesData: any[] = [], dimensionsData: any[] = [];
    const brandsMap = new Map<string, any>(), manufacturersMap = new Map<string, any>(), erpCategoriesMap = new Map<string, any>();

    df_template.forEach(row => {
        const item_uid = String(row['UID*']);
        
        let erpCategoryUid: string | number | Date | null = null;
        for (let i = 6; i >= 1; i--) {
            const catUidCol = `Category level ${i} UID`;
            const val = row[catUidCol];
            if (val !== null && val !== undefined && String(val).trim() !== '') {
                erpCategoryUid = val;
                break;
            }
        }
        if (erpCategoryUid === null) {
            for (let i = 6; i >= 1; i--) {
                const catNameCol = `Category level ${i}`;
                const val = row[catNameCol];
                if (val !== null && val !== undefined && String(val).trim() !== '') {
                    erpCategoryUid = val;
                    break;
                }
            }
        }

        let main_unit_uid = row['Main Unit UID'] || row['Main unit UID'];
        if (main_unit_uid === null || main_unit_uid === undefined || String(main_unit_uid).trim() === '') {
            const unitName = row['Unit'];
            if (unitName && String(unitName).trim()) {
                main_unit_uid = `${item_uid}_${String(unitName).trim()}`;
            } else {
                main_unit_uid = `${item_uid}_01`;
            }
        }

        const brandUidValue = row['Brand UID'];
        const brandNameValue = row['Brand'];
        let effectiveBrandUid: string | number | Date | null = null;
        if (brandUidValue !== null && brandUidValue !== undefined && String(brandUidValue).trim() !== '') {
            effectiveBrandUid = brandUidValue;
        } else if (brandNameValue !== null && brandNameValue !== undefined && String(brandNameValue).trim() !== '') {
            effectiveBrandUid = brandNameValue;
        }

        const manufUidValue = row['Manufacturer UID'];
        const manufNameValue = row['Manufacturer'];
        let effectiveManufUid: string | number | Date | null = null;
        if (manufUidValue !== null && manufUidValue !== undefined && String(manufUidValue).trim() !== '') {
            effectiveManufUid = manufUidValue;
        } else if (manufNameValue !== null && manufNameValue !== undefined && String(manufNameValue).trim() !== '') {
            effectiveManufUid = manufNameValue;
        }

        masteritemsData.push({
            item_uid: item_uid, name: row['Product name*'], manufacturer_uid: effectiveManufUid, brand_uid: effectiveBrandUid,
            is_fractional: parseInt(String(row['Is fractional?']), 10) || 0,
            main_unit_uid: main_unit_uid, is_deleted: parseInt(String(row['To delete']), 10) || 0,
            additional_1: parseFloat(String(row['Add 1']).replace(',', '.')) || null,
            additional_2: row['Add 2'], additional_3: row['Add 3'], additional_4: row['Add 4'],
            additional_5: parseFloat(String(row['Add 5']).replace(',', '.')) || null,
            additional_6: parseFloat(String(row['Add 6']).replace(',', '.')) || null,
            additional_7: row['Add 7'],
            additional_8: parseFloat(String(row['Add 8']).replace(',', '.')) || null,
            additional_9: parseFloat(String(row['Add 9']).replace(',', '.')) || null,
            additional_10: parseFloat(String(row['Add 10']).replace(',', '.')) || null,
            additional_11: parseFloat(String(row['Add 11']).replace(',', '.')) || null,
            additional_12: parseFloat(String(row['Add 12']).replace(',', '.')) || null,
            additional_13: parseFloat(String(row['Add 13']).replace(',', '.')) || null,
            additional_14: parseFloat(String(row['Add 14']).replace(',', '.')) || null,
            additional_15: parseFloat(String(row['Add 15']).replace(',', '.')) || null,
            additional_16: parseFloat(String(row['Add 16']).replace(',', '.')) || null,
            additional_17: parseFloat(String(row['Add 17']).replace(',', '.')) || null,
            additional_18: parseFloat(String(row['Add 18']).replace(',', '.')) || null,
            additional_19: parseFloat(String(row['Add 19']).replace(',', '.')) || null,
            additional_20: parseFloat(String(row['Add 20']).replace(',', '.')) || null,
            erp_category_uid: erpCategoryUid
        });

        if (row['Barcode']) barcodesData.push({ item_uid: item_uid, barcode: row['Barcode'], is_main: 1 });

        if (effectiveBrandUid !== null) {
            const key = String(effectiveBrandUid);
            if (!brandsMap.has(key)) {
                brandsMap.set(key, { 
                    brand_uid: effectiveBrandUid, 
                    name: brandNameValue || effectiveBrandUid, 
                    is_deleted: 0 
                });
            }
        }

        if (row['Unit']) {
            dimensionsData.push({
                item_uid: item_uid, unit_name: row['Unit'],
                width: parseFloat((row['Width (cm, in)'] || '').toString().replace(',', '.')) || null,
                height: parseFloat((row['Height (cm, in)'] || '').toString().replace(',', '.')) || null,
                depth: parseFloat((row['Depth (cm, in)'] || '').toString().replace(',', '.')) || null,
                coef: 1, is_deleted: 0, dimension_uid: main_unit_uid
            });
        }
        
        let previousLevelEffectiveUid: string | number | Date | null = null;
        for (let level = 1; level <= 6; level++) {
            const uidCol = `Category level ${level} UID`;
            const nameCol = `Category level ${level}`;
            
            const currentUidValue = row[uidCol];
            const currentNameValue = row[nameCol];

            let currentLevelEffectiveUid: string | number | Date | null = null;
            
            if (currentUidValue !== null && currentUidValue !== undefined && String(currentUidValue).trim() !== '') {
                currentLevelEffectiveUid = currentUidValue;
            } else if (currentNameValue !== null && currentNameValue !== undefined && String(currentNameValue).trim() !== '') {
                currentLevelEffectiveUid = currentNameValue;
            }

            if (currentLevelEffectiveUid !== null) {
                const key = String(currentLevelEffectiveUid);
                const existingEntry = erpCategoriesMap.get(key);
                
                if (!existingEntry || (existingEntry.name === null && currentNameValue !== null)) {
                    erpCategoriesMap.set(key, {
                        erp_category_uid: currentLevelEffectiveUid,
                        name: currentNameValue,
                        parent_category_uid: previousLevelEffectiveUid
                    });
                }
                previousLevelEffectiveUid = currentLevelEffectiveUid;
            }
        }
        
        if (effectiveManufUid !== null) {
            const key = String(effectiveManufUid);
            if (!manufacturersMap.has(key)) {
                manufacturersMap.set(key, {
                    manufacturer_uid: effectiveManufUid,
                    name: manufNameValue || effectiveManufUid,
                    is_deleted: 0
                });
            }
        }
    });

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    const masteritems_cols = [
        'item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional', 'additional_1', 'additional_2', 'additional_3', 'additional_4', 'additional_5', 'additional_6', 'additional_7', 'additional_8', 'additional_9', 'additional_10', 'additional_11', 'additional_12', 'additional_13', 'additional_14', 'additional_15', 'additional_16', 'additional_17', 'additional_18', 'additional_19', 'additional_20', 'main_unit_uid', 'erp_category_uid', 'is_deleted'
    ];
    csvs.push({ name: `masteritems_${dateStr}.csv`, content: arrayToCsv(masteritemsData, masteritems_cols) });

    if((options.barcodes ?? true) && barcodesData.length > 0) csvs.push({ name: `barcodes_${dateStr}.csv`, content: arrayToCsv(barcodesData, ['item_uid', 'barcode', 'is_main']) });
    if((options.brands ?? true) && brandsMap.size > 0) csvs.push({ name: `brands_${dateStr}.csv`, content: arrayToCsv(Array.from(brandsMap.values()), ['brand_uid', 'name', 'is_deleted']) });
    if((options.dimensions ?? true) && dimensionsData.length > 0) csvs.push({ name: `dimensions_${dateStr}.csv`, content: arrayToCsv(dimensionsData, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'coef', 'is_deleted', 'dimension_uid']) });
    if ((options.erpcategories ?? true) && erpCategoriesMap.size > 0) csvs.push({ name: `erpcategories_${dateStr}.csv`, content: arrayToCsv(Array.from(erpCategoriesMap.values()), ['erp_category_uid', 'name', 'parent_category_uid']) });
    if((options.manufacturers ?? true) && manufacturersMap.size > 0) csvs.push({ name: `manufacturers_${dateStr}.csv`, content: arrayToCsv(Array.from(manufacturersMap.values()), ['manufacturer_uid', 'name', 'is_deleted']) });
    
    updateStatus({ message: 'New Masteritems processing complete.', status: 'success' });
    return csvs;
}

// --- MAIN DISPATCHER ---

const generateCsvsFromExcel = async (
    file: File,
    updateStatus: StatusUpdateCallback,
    options: CsvGenerationOptions = {}
): Promise<{ csvFiles: CsvFile[]; detectedType: FileType }> => {
    
    updateStatus({ message: 'Reading and analyzing Excel file...', status: 'processing' });
    const data = await file.arrayBuffer();
    // Use `cellDates: true` to ensure the library parses dates into JS Date objects
    const workbook = XLSX.read(data, { cellDates: true });
    
    const { type: detectedType, sheetName } = detectFileType(workbook);
    
    if (detectedType === 'UNKNOWN' || !sheetName) {
        throw new Error('Unknown file type. The file does not match any known templates.');
    }
    
    let csvFiles: CsvFile[] = [];

    switch (detectedType) {
        case 'STORE':
            csvFiles = processStoresFile(workbook, sheetName, updateStatus);
            break;
        case 'STORE_ITEMS':
            csvFiles = processStoreItemsFile(workbook, sheetName, updateStatus, options);
            break;
        case 'ITEM_MASTER':
            csvFiles = processItemMasterFile(workbook, sheetName, updateStatus, options);
            break;
        case 'ITEM_MASTER_V2':
            csvFiles = processItemMasterV2File(workbook, sheetName, updateStatus, options);
            break;
        case 'FACTS':
            csvFiles = processFactsFile(workbook, sheetName, updateStatus);
            break;
        case 'STOCK':
        case 'PRICE':
            // Placeholder for other file types
            throw new Error(`Processing for "${detectedType}" files is not yet implemented.`);
    }

    return { csvFiles, detectedType };
};


// ==================================================================================
// MERGED FROM App.tsx
// ==================================================================================
// Fix for non-standard directory attributes on input element
declare module 'react' {
    interface InputHTMLAttributes<T> {
        webkitdirectory?: string;
        directory?: string;
    }
}

declare var JSZip: any;

type ProcessingState = 'idle' | 'processing' | 'success' | 'error';
interface FileInfo {
    file: File;
    type: FileType;
    status: 'pending' | 'processing' | 'success' | 'error';
    error?: string;
}

interface LogEntry extends StatusUpdate {
    // Using file index to associate log entries with a specific file
    fileIndex: number;
}

const CsvOptionsConfig: Record<string, string[]> = {
  ITEM_MASTER: ['barcodes', 'brands', 'dimensions', 'erpcategories', 'manufacturers'],
  ITEM_MASTER_V2: ['barcodes', 'brands', 'dimensions', 'erpcategories', 'manufacturers'],
  STORE_ITEMS: ['suppliers'],
};

const App: React.FC = () => {
    const [fileInfos, setFileInfos] = useState<FileInfo[]>([]);
    const [archiveName, setArchiveName] = useState<string>('data_export');
    const [processingState, setProcessingState] = useState<ProcessingState>('idle');
    const [statusUpdates, setStatusUpdates] = useState<LogEntry[]>([]);
    const [errorMessage, setErrorMessage] = useState<string>('');
    const [zipUrl, setZipUrl] = useState<string | null>(null);
    const [zipFileName, setZipFileName] = useState<string>('');
    const [csvOptions, setCsvOptions] = useState<CsvGenerationOptions>({});
    const fileInputRef = useRef<HTMLInputElement>(null);

    const resetState = () => {
        setProcessingState('idle');
        setStatusUpdates([]);
        setErrorMessage('');
        setZipUrl(null);
        setCsvOptions({});
        if (zipUrl) {
            URL.revokeObjectURL(zipUrl);
        }
    };

    const handleFilesChange = async (fileList: FileList | null) => {
        if (!fileList) return;
        
        setFileInfos([]);
        resetState();
        
        const validFiles = Array.from(fileList).filter(f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls'));
        
        if (validFiles.length === 0) {
            setErrorMessage('No valid Excel files (.xlsx, .xls) found in the selected directory.');
            setProcessingState('error');
            return;
        }

        const initialFileInfos: FileInfo[] = validFiles.map(file => ({
            file,
            type: 'UNKNOWN',
            status: 'pending'
        }));
        setFileInfos(initialFileInfos);

        // Asynchronously detect file types
        const detectionPromises = validFiles.map(async (file, index) => {
            try {
                // Pass empty callback and no options for detection phase
                const { detectedType } = await generateCsvsFromExcel(file, () => {});
                return { index, type: detectedType };
            } catch (error) {
                return { index, type: 'UNKNOWN' as 'UNKNOWN', error: error instanceof Error ? error.message : "Detection failed" };
            }
        });
        
        const results = await Promise.all(detectionPromises);
        
        setFileInfos(currentInfos => {
            const newInfos = [...currentInfos];
            results.forEach(result => {
                if (result) {
                    newInfos[result.index].type = result.type;
                    if(result.error) {
                       newInfos[result.index].status = 'error';
                       newInfos[result.index].error = result.error;
                    }
                }
            });
            return newInfos;
        });
    };

    const handleProcess = useCallback(async () => {
        if (fileInfos.length === 0) return;

        setProcessingState('processing');
        setStatusUpdates([]);
        setErrorMessage('');
        setZipUrl(null);
        
        const zip = new JSZip();
        const allGeneratedCsvs: CsvFile[] = [];
        let hasProcessedAnyFile = false;
        let localErrorMessages = '';

        for (let i = 0; i < fileInfos.length; i++) {
            const info = fileInfos[i];
            if (info.type === 'UNKNOWN' || info.status === 'error') {
                 setStatusUpdates(prev => [...prev, { message: `Skipping invalid file: ${info.file.name}`, status: 'error', fileIndex: i }]);
                 continue;
            }

            setFileInfos(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'processing' } : f));
            
            const updateCallback = (update: StatusUpdate) => {
                setStatusUpdates(prev => [...prev, { ...update, message: `[${info.file.name}] ${update.message}`, fileIndex: i }]);
            };

            try {
                const { csvFiles } = await generateCsvsFromExcel(info.file, updateCallback, csvOptions);
                csvFiles.forEach(csv => allGeneratedCsvs.push(csv));
                hasProcessedAnyFile = true;
                setFileInfos(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'success' } : f));
                 // Update all 'processing' logs for this file to 'success'
                setStatusUpdates(prev => prev.map(log =>
                    (log.fileIndex === i && log.status === 'processing') ? { ...log, status: 'success' } : log
                ));

            } catch (error) {
                const message = error instanceof Error ? error.message : 'An unknown error occurred.';
                localErrorMessages += `Error in ${info.file.name}: ${message}\n`;
                setStatusUpdates(prev => [...prev, { message: `Failed to process ${info.file.name}: ${message}`, status: 'error', fileIndex: i }]);
                setFileInfos(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'error', error: message } : f));
                // Update all 'processing' logs for this file to 'error'
                setStatusUpdates(prev => prev.map(log =>
                    (log.fileIndex === i && log.status === 'processing') ? { ...log, status: 'error' } : log
                ));
            }
        }
        
        setErrorMessage(localErrorMessages.trim());

        if (allGeneratedCsvs.length > 0) {
            try {
                allGeneratedCsvs.forEach(csv => {
                    zip.file(csv.name, csv.content);
                });

                const finalZipName = `${archiveName.trim() || 'data_export'}_${getTodayDateString()}.zip`;
                const zipBlob = await zip.generateAsync({ type: 'blob' });
                const url = URL.createObjectURL(zipBlob);

                setZipUrl(url);
                setZipFileName(finalZipName);
                setProcessingState('success');
                setStatusUpdates(prev => [...prev, { message: 'All valid files processed and bundled into a ZIP archive.', status: 'success', fileIndex: -1 }]);
            } catch (zipError) {
                const message = zipError instanceof Error ? zipError.message : 'An unknown error occurred while creating the ZIP file.';
                setErrorMessage(prev => (prev ? `${prev}\n${message}` : message));
                setProcessingState('error');
            }
        } else if (hasProcessedAnyFile && !localErrorMessages) {
            setProcessingState('success');
            setStatusUpdates(prev => [...prev, { message: 'Processing complete. All valid files were processed but contained no data to export.', status: 'success', fileIndex: -1 }]);
        } else {
            setProcessingState('error');
            if (!localErrorMessages) {
                setErrorMessage("No valid files were processed successfully.");
            }
        }

    }, [fileInfos, archiveName, csvOptions]);

    const handleReset = () => {
        setFileInfos([]);
        resetState();
        if (fileInputRef.current) {
            fileInputRef.current.value = '';
        }
    };

    const handleCsvOptionChange = (option: string) => {
        setCsvOptions(prev => ({
            ...prev,
            [option]: !(prev[option] ?? true),
        }));
    };
    
    const StatusIcon = ({ status }: { status: 'processing' | 'success' | 'error' | 'pending' }) => {
        switch (status) {
            case 'processing': return <Loader2 className="animate-spin text-blue-500 w-5 h-5 mr-3 shrink-0" />;
            case 'success': return <CheckCircle className="text-green-500 w-5 h-5 mr-3 shrink-0" />;
            case 'error': return <XCircle className="text-red-500 w-5 h-5 mr-3 shrink-0" />;
            default: return <FileText className="text-gray-400 w-5 h-5 mr-3 shrink-0" />;
        }
    };

    const formatFileTypeName = (typeName: FileType) => {
        if (typeName === 'UNKNOWN') return 'Unknown';
        if (typeName === 'ITEM_MASTER') return 'Masteritems';
        if (typeName === 'ITEM_MASTER_V2') return 'Masteritems (New Format)';
        if (typeName === 'FACTS') return 'Facts Data';
        if (typeName === 'STORE_ITEMS') return 'Store Items';
        if (typeName === 'STORE') return 'Stores';
        return typeName.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
    };

    const formatOptionName = (name: string) => {
        if (name === 'erpcategories') return 'ERP Categories';
        return name.charAt(0).toUpperCase() + name.slice(1);
    };

    const hasValidFiles = fileInfos.some(f => f.type !== 'UNKNOWN' && f.status !== 'error');

    const uniqueOptions = [...new Set(
        fileInfos
            .map(f => f.type)
            .filter(t => t in CsvOptionsConfig)
            .flatMap(t => CsvOptionsConfig[t as keyof typeof CsvOptionsConfig])
    )];

    return (
        <div className="min-h-screen bg-gray-50 text-gray-800 flex items-center justify-center p-4 transition-colors duration-300">
            <div className="w-full max-w-3xl mx-auto">
                <header className="text-center mb-8">
                    <h1 className="text-4xl md:text-5xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-primary to-secondary">Excel Data Converter</h1>
                    <p className="mt-3 text-lg text-gray-600">Select a folder to convert all valid template files into a structured CSV archive.</p>
                </header>

                <main className="bg-white rounded-2xl shadow-2xl p-6 md:p-8 space-y-6">
                    {fileInfos.length === 0 ? (
                         <div
                            className="border-2 border-dashed border-gray-300 rounded-lg p-12 text-center cursor-pointer hover:border-primary transition-all duration-300 group"
                            onClick={() => fileInputRef.current?.click()}
                        >
                            <FolderUp className="mx-auto h-12 w-12 text-gray-400 group-hover:text-primary transition-colors duration-300" />
                            <p className="mt-2 font-semibold text-primary">Click to select a folder</p>
                            <p className="text-sm text-gray-500">All .xlsx and .xls files will be processed</p>
                            <input
                                type="file"
                                ref={fileInputRef}
                                onChange={(e) => handleFilesChange(e.target.files)}
                                className="hidden"
                                webkitdirectory=""
                                directory=""
                            />
                        </div>
                    ) : (
                        <div className="space-y-3">
                             <h3 className="text-lg font-semibold border-b border-gray-200 pb-2">Detected Files:</h3>
                             <div className="max-h-60 overflow-y-auto space-y-2 pr-2">
                                {fileInfos.map((info, index) => (
                                    <div key={index} className={`flex items-center p-2 rounded-md ${info.status === 'error' ? 'bg-red-50' : 'bg-gray-100'}`}>
                                        <StatusIcon status={info.status} />
                                        <div className="flex-grow">
                                            <p className="font-medium text-sm">{info.file.name}</p>
                                            <p className={`text-xs ${info.status === 'error' ? 'text-red-600' : 'text-gray-500'}`}>
                                               {info.status === 'error' ? info.error : `Type: ${formatFileTypeName(info.type)}`}
                                            </p>
                                        </div>
                                    </div>
                                ))}
                             </div>
                        </div>
                    )}
                    
                    {uniqueOptions.length > 0 && (
                        <div className="space-y-3 pt-4 border-t border-gray-200">
                            <h3 className="text-lg font-semibold">CSV Generation Options</h3>
                            <p className="text-sm text-gray-500">Select which supplementary CSV files to generate.</p>
                            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-6 gap-y-3 pt-2">
                                {uniqueOptions.map(option => (
                                    <label key={option} className="flex items-center space-x-3 cursor-pointer">
                                        <input
                                            type="checkbox"
                                            checked={csvOptions[option] ?? true}
                                            onChange={() => handleCsvOptionChange(option)}
                                            disabled={processingState === 'processing'}
                                            className="h-4 w-4 rounded border-gray-300 text-primary focus:ring-primary-focus transition disabled:opacity-50"
                                        />
                                        <span className="text-sm font-medium text-gray-700">{formatOptionName(option)}</span>
                                    </label>
                                ))}
                            </div>
                        </div>
                    )}

                    {fileInfos.length > 0 && (
                        <div>
                            <label htmlFor="archive-name" className="block text-sm font-medium text-gray-700 mb-1">
                                ZIP Archive Name
                            </label>
                            <input
                                type="text"
                                id="archive-name"
                                value={archiveName}
                                onChange={(e) => setArchiveName(e.target.value)}
                                className="block w-full px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm placeholder-gray-400 focus:outline-none focus:ring-primary-focus focus:border-primary-focus sm:text-sm"
                                placeholder="e.g., data_export"
                                disabled={processingState === 'processing'}
                            />
                        </div>
                    )}
                    
                    {processingState !== 'processing' && errorMessage && (
                        <div className="bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded-md" role="alert">
                            <div className="flex">
                                <AlertTriangle className="w-5 h-5 mr-3 shrink-0"/>
                                <div>
                                    <p className="font-bold">Processing Errors</p>
                                    <p className="whitespace-pre-wrap">{errorMessage}</p>
                                </div>
                            </div>
                        </div>
                    )}

                    <div className="flex flex-col sm:flex-row gap-4">
                        <button
                            onClick={handleProcess}
                            disabled={!hasValidFiles || processingState === 'processing'}
                            className="w-full inline-flex justify-center items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-primary hover:bg-primary-hover focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-focus disabled:bg-gray-400 disabled:cursor-not-allowed transition-all duration-300"
                        >
                            {processingState === 'processing' ? <Loader2 className="animate-spin -ml-1 mr-3 h-5 w-5"/> : <CheckCircle className="-ml-1 mr-3 h-5 w-5" />}
                            {processingState === 'processing' ? 'Processing...' : 'Start Processing'}
                        </button>

                        <button
                            onClick={handleReset}
                            className="w-full sm:w-auto inline-flex justify-center items-center px-6 py-3 border border-gray-300 text-base font-medium rounded-md shadow-sm text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-focus transition-all duration-300"
                        >
                           <RefreshCw className="mr-3 h-5 w-5"/> Reset
                        </button>
                    </div>

                    {(statusUpdates.length > 0) && (
                        <div className="space-y-2 pt-4 border-t border-gray-200">
                             <h3 className="text-lg font-semibold">Processing Log:</h3>
                            <ul className="space-y-2 text-sm max-h-40 overflow-y-auto pr-2">
                                {statusUpdates.map((update, index) => (
                                    <li key={index} className="flex items-start p-2 bg-gray-50 rounded-md">
                                        <StatusIcon status={update.status} />
                                        <span className="flex-grow">{update.message}</span>
                                    </li>
                                ))}
                            </ul>
                        </div>
                    )}

                    {processingState === 'success' && zipUrl && (
                        <div className="pt-4 border-t border-gray-200">
                            <a
                                href={zipUrl}
                                download={zipFileName}
                                className="w-full inline-flex justify-center items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-secondary hover:bg-green-600 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500 transition-all duration-300"
                            >
                                <Download className="-ml-1 mr-3 h-5 w-5"/>
                                Download ZIP File ({zipFileName})
                            </a>
                        </div>
                    )}
                </main>
                 <footer className="text-center mt-8">
                    <p className="text-sm text-gray-500">
                        For Internal Use Only
                    </p>
                </footer>
            </div>
        </div>
    );
};


// ==================================================================================
// RENDER THE APP
// ==================================================================================
const rootElement = document.getElementById('root');
if (rootElement) {
    const root = ReactDOM.createRoot(rootElement);
    root.render(<App />);
}
