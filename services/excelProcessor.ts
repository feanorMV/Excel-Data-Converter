
import * as XLSX from 'xlsx';
import type { StatusUpdateCallback, ExcelRow, CsvFile, FileType, CsvGenerationOptions } from '../types.ts';

const yieldToUI = () => new Promise(resolve => {
    if (typeof requestAnimationFrame !== 'undefined') {
        requestAnimationFrame(() => setTimeout(resolve, 0));
    } else {
        setTimeout(resolve, 0);
    }
});

// --- UTILITY FUNCTIONS ---

function getTodayDateString(): string {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
}

async function arrayToCsv(data: Record<string, any>[], columns: string[], selectedColumns: Record<string, boolean> | null = null, onProgress?: (progress: number) => void, signal?: AbortSignal, delimiter: string = ',', columnMapping: Record<string, string> = {}): Promise<string> {
    const finalColumns = selectedColumns ? columns.filter(col => selectedColumns[col] ?? true) : columns;
    if (finalColumns.length === 0) return '';

    const header = finalColumns.map(col => {
        const mappedCol = columnMapping[col] || col;
        if (mappedCol.includes('"') || mappedCol.includes(delimiter)) {
            return `"${mappedCol.replace(/"/g, '""')}"`;
        }
        return mappedCol;
    }).join(delimiter) + '\n';
    let rows = '';
    
    for (let i = 0; i < data.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const row = data[i];
        const rowContent = finalColumns.map(col => {
            let value = row[col];
            if (value === null || typeof value === 'undefined') {
                return '';
            }
            value = String(value);
            if (value.includes('"') || value.includes(delimiter)) {
                return `"${value.replace(/"/g, '""')}"`;
            }
            return value;
        }).join(delimiter);
        
        rows += rowContent + '\n';
        
        if (i % 2000 === 0) {
            await yieldToUI();
            if (onProgress) onProgress(Math.round((i / data.length) * 100));
        }
    }
    
    if (onProgress) onProgress(100);
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

const ADDITIONAL_HEADERS_KEYWORDS = Array.from({ length: 20 }, (_, i) => [`Add ${i + 1}`, `additional_${i + 1}`]).flat();

function getIsDeletedValue(row: any): number {
    const val = row['Deleted'] ?? row['Is Deleted'] ?? row['Is deleted'] ?? row['To delete'] ?? row['To Delete'] ?? row['is_deleted'];
    if (val !== undefined && val !== null && String(val).trim() !== '') {
        if (typeof val === 'boolean') return val ? 1 : 0;
        const strVal = String(val).toLowerCase().trim();
        if (strVal === 'true' || strVal === 'yes') return 1;
        if (strVal === 'false' || strVal === 'no') return 0;
        const parsed = parseInt(strVal, 10);
        return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
}

function getHeadersAndDataStart(worksheet: any, keywords: string[], skipRowsAfterHeader: number = 0): { headers: string[], headerRowIndex: number, dataStartIndex: number } {
    let rawData: any[][] = [];
    if (worksheet['!ref']) {
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        const detectionRange = { s: { r: 0, c: 0 }, e: { r: Math.min(range.e.r, 100), c: range.e.c } };
        rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false, range: detectionRange });
    } else {
        rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    }

    let headerRowIndex = -1;
    const searchKeywords = [...keywords, ...ADDITIONAL_HEADERS_KEYWORDS];
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && searchKeywords.some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    
    if (headerRowIndex === -1) return { headers: [], headerRowIndex: -1, dataStartIndex: -1 };

    const headers = rawData[headerRowIndex].map(h => String(h || '').trim());
    let dataStartIndex = headerRowIndex + 1 + skipRowsAfterHeader;

    if (skipRowsAfterHeader === 0) {
        // Find first non-empty row if not explicitly skipping
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
            if (rawData[i] && rawData[i].some(cell => cell !== null && String(cell).trim() !== '')) {
                dataStartIndex = i;
                break;
            }
        }
    }

    return { headers, headerRowIndex, dataStartIndex };
}

// --- FILE TYPE DETECTION ---

const FILE_TYPE_DEFINITIONS = {
    STORE: { keywords: ["Store UID*"] },
    STORE_ITEMS: { keywords: ["Store UID*", "Product UID*", "In assortment?", "Purchase price"] },
    FACTS: { keywords: ["Product UID*", "Store UID*", "Date*"] },
    ITEM_MASTER_UPDATED: { keywords: ["UID*", "Product name*", "Manufacturer", "Brand", "Is fractional?", "Segment Description", "Brick Code"] },
    ITEM_MASTER_V2: { keywords: ["UID*", "Product name*", "Category level 1"] },
    ITEM_MASTER: { keywords: ["UID*", "Product name*", "Barcode", "Manufacturer"] },
    STOCK: { keywords: ['StoreID', 'ItemUID', 'Quantity'] },
    PRICE: { keywords: ['ItemUID', 'PriceList', 'Price'] },
};


const detectFileType = (workbook: any): { type: FileType; sheetName: string | null } => {
    const sheetNames = workbook.SheetNames;
    const typeCheckOrder: FileType[] = ['ITEM_MASTER_UPDATED', 'ITEM_MASTER_V2', 'ITEM_MASTER', 'STORE_ITEMS', 'FACTS', 'STORE', 'STOCK', 'PRICE'];
    
    // Cache for sheet data to avoid redundant sheet_to_json calls
    const sheetDataCache = new Map<string, any[][]>();

    for (const sheetName of sheetNames) {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) continue;
        
        let jsonData: any[][] = [];
        if (sheet['!ref']) {
            const range = XLSX.utils.decode_range(sheet['!ref']);
            const detectionRange = { s: { r: 0, c: 0 }, e: { r: Math.min(range.e.r, 100), c: range.e.c } };
            jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false, range: detectionRange });
        } else {
            jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false });
        }
        sheetDataCache.set(sheetName, jsonData);
    }

    for (const type of typeCheckOrder) {
        const definition = FILE_TYPE_DEFINITIONS[type as keyof typeof FILE_TYPE_DEFINITIONS];
        if (!definition) continue;

        for (const sheetName of sheetNames) {
            const jsonData = sheetDataCache.get(sheetName);
            if (jsonData) {
                const searchLimit = Math.min(50, jsonData.length);
                for (let i = 0; i < searchLimit; i++) {
                    const row = jsonData[i];
                    if (row && definition.keywords.every(kw => row.some(cell => typeof cell === 'string' && cell.includes(kw)))) {
                        return { type: type as FileType, sheetName: sheetName };
                    }
                }
            }
        }
    }
    return { type: 'UNKNOWN', sheetName: null };
};

// --- SPECIALIZED PROCESSORS ---

/**
 * Processes a "Stores" file.
 */
async function processStoresFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Stores file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find a valid header row in the Stores sheet.');
    
    // Read data directly into objects
    const df_template: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    const storesData: any[] = [];
    for (let i = 0; i < df_template.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const rowObject = df_template[i];
        
        storesData.push({
            store_uid: rowObject['Store UID*'],
            name: rowObject['Store name*'],
            region: rowObject['Region'],
            group_name: rowObject['Group name'],
            floor_space: parseInt(String(rowObject['Square']), 10) || 0,
            in_shelf: parseInt(String(rowObject['In Shelf?']), 10) || 0,
            licence_start_date: '2023-01-01',
            is_deleted: getIsDeletedValue(rowObject),
        });

        if (i % 1000 === 0) {
            await yieldToUI();
            updateStatus({ message: 'Parsing store data...', status: 'processing', progress: Math.round((i / df_template.length) * 50) });
        }
    }

    const filteredStoresData = storesData.filter(row => row.store_uid);

    if (filteredStoresData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Stores file, skipping CSV generation.', status: 'success', progress: 100 });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `stores_${dateStr}.csv`,
        rowCount: filteredStoresData.length,
        content: await arrayToCsv(
            filteredStoresData, 
            ['store_uid', 'name', 'region', 'group_name', 'floor_space', 'in_shelf', 'licence_start_date', 'is_deleted'], 
            selectedColumns,
            (p) => updateStatus({ message: 'Generating CSV...', status: 'processing', progress: 50 + Math.round(p / 2) }),
            signal,
            options.delimiter,
            options.columnMapping
        )
    }];

    updateStatus({ message: 'Stores processing complete.', status: 'success', progress: 100 });
    return csvs;
}

/**
 * Processes a "Store Items" file containing assortment, pricing, and supplier data.
 */
async function processStoreItemsFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Store Items file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find a valid header row in the Items sheet.');
    
    const df_template: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    const itemsData: any[] = [];
    const suppliersMap = new Map<string, any>();

    for (let i = 0; i < df_template.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const rowObject = df_template[i];

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

        if (i % 1000 === 0) {
            await yieldToUI();
            updateStatus({ message: 'Parsing store items data...', status: 'processing', progress: Math.round((i / df_template.length) * 50) });
        }
    }
    
    if (itemsData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Items file, skipping CSV generation.', status: 'success', progress: 100 });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];

    csvs.push({
        name: `items_${dateStr}.csv`,
        rowCount: itemsData.length,
        content: await arrayToCsv(
            itemsData, 
            ['item_uid', 'store_uid', 'is_active_planogram', 'purchase_price', 'retail_price', 'external_supplier_uid'], 
            selectedColumns,
            (p) => updateStatus({ message: 'Generating Items CSV...', status: 'processing', progress: 50 + Math.round(p / 4) }),
            signal,
            options.delimiter,
            options.columnMapping
        )
    });

    if ((options.suppliers ?? true) && suppliersMap.size > 0) {
        csvs.push({
            name: `suppliers_${dateStr}.csv`,
            rowCount: suppliersMap.size,
            content: await arrayToCsv(
                Array.from(suppliersMap.values()), 
                ['supplier_uid', 'name', 'is_deleted'], 
                selectedColumns,
                (p) => updateStatus({ message: 'Generating Suppliers CSV...', status: 'processing', progress: 75 + Math.round(p / 4) }),
                signal,
                options.delimiter,
                options.columnMapping
            )
        });
    }
    
    updateStatus({ message: 'Store Items processing complete.', status: 'success', progress: 100 });
    return csvs;
}

/**
 * Processes a "Facts" file containing sales and stock data.
 */
async function processFactsFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Facts file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find a valid header row in the Facts sheet.');

    const df_template: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    const factsData: any[] = [];
    for (let i = 0; i < df_template.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const rowObject = df_template[i];

        let formattedDate: string | null = null;
        const dateValue = rowObject['Date*'];

        if (dateValue) {
            let jsDate: Date | null = null;
            if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
                jsDate = dateValue;
            } else if (typeof dateValue === 'number' && dateValue > 1) {
                jsDate = excelSerialDateToJSDate(dateValue);
            } else {
                const d = new Date(String(dateValue));
                if (!isNaN(d.getTime())) {
                    jsDate = d;
                }
            }
            
            if (jsDate) {
                formattedDate = formatDateToYYYYMMDD(jsDate);
            }
        }

        factsData.push({
            item_uid: rowObject['Product UID*'],
            store_uid: rowObject['Store UID*'],
            date: formattedDate,
            stock: parseFloat(String(rowObject['Stock'] || '0').replace(',', '.')) || null,
            sold_qty: parseFloat(String(rowObject['Out sale'] || '0').replace(',', '.')) || null,
            revenue: parseFloat(String(rowObject['Revenue'] || '0').replace(',', '.')) || null,
            cogs: parseFloat(String(rowObject['COGS'] || '0').replace(',', '.')) || null,
        });

        if (i % 1000 === 0) {
            await yieldToUI();
            updateStatus({ message: 'Parsing facts data...', status: 'processing', progress: Math.round((i / df_template.length) * 50) });
        }
    }

    const filteredFactsData = factsData.filter(row => row.item_uid && row.store_uid && row.date);
    
    if (filteredFactsData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Facts file, skipping CSV generation.', status: 'success', progress: 100 });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `facts_${dateStr}.csv`,
        rowCount: filteredFactsData.length,
        content: await arrayToCsv(
            filteredFactsData, 
            ["item_uid", "store_uid", "date", "stock", "sold_qty", "revenue", "cogs"], 
            selectedColumns,
            (p) => updateStatus({ message: 'Generating Facts CSV...', status: 'processing', progress: 50 + Math.round(p / 2) }),
            signal,
            options.delimiter,
            options.columnMapping
        )
    }];
    
    updateStatus({ message: 'Facts processing complete.', status: 'success', progress: 100 });
    return csvs;
}


/**
 * Processes the original "Item Master" file format (V1).
 */
async function processItemMasterFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);
    
    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find header row in Masteritems file.');

    const df_template: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    const masteritemsData: any[] = [], barcodesData: any[] = [], dimensionsData: any[] = [];
    const seenBrands = new Set<string>(), seenManufacturers = new Set<string>(), seenErpCategories = new Map<string, any>(), seenUids = new Set<string>();
    const levelMapping = { 1: 'Segment', 2: 'Family', 3: 'Class', 4: 'Brick' };

    for (let i = 0; i < df_template.length; i++) {
        const row = df_template[i];
        const uidValue = row['UID*'];
        if (uidValue !== null && uidValue !== undefined) {
            const uid = String(uidValue).trim();
            if (uid !== '') {
                let main_unit_uid = row['Main Unit UID'];
                if (main_unit_uid === null || main_unit_uid === undefined || String(main_unit_uid).trim() === '') {
                    const unitName = row['Unit'];
                    if (unitName && String(unitName).trim()) {
                        main_unit_uid = `${uid}_${String(unitName).trim()}`;
                    } else {
                        main_unit_uid = `${uid}_01`;
                    }
                }

                if (!seenUids.has(uid)) {
                    const getAdd = (num: number) => row[`Add ${num}`] ?? row[`additional_${num}`];
                    masteritemsData.push({ 
                        item_uid: uid, 
                        name: row['Product name*'], 
                        manufacturer_uid: row['Manufacturer'], 
                        brand_uid: row['Brand'], 
                        is_fractional: row['Is fractional?'] ? parseInt(String(row['Is fractional?']), 10) || 0 : 0, 
                        is_deleted: getIsDeletedValue(row),
                        additional_1: row['Segment Description'] ?? getAdd(1), 
                        additional_2: row['Family Description'] ?? getAdd(2), 
                        additional_3: row['Class Description'] ?? getAdd(3), 
                        additional_4: row['Brick Description'] ?? getAdd(4), 
                        additional_5: getAdd(5), additional_6: getAdd(6), additional_7: getAdd(7), additional_8: getAdd(8),
                        additional_9: getAdd(9), additional_10: getAdd(10), additional_11: getAdd(11), additional_12: getAdd(12),
                        additional_13: getAdd(13), additional_14: getAdd(14), additional_15: getAdd(15), additional_16: getAdd(16),
                        additional_17: getAdd(17), additional_18: getAdd(18), additional_19: getAdd(19), additional_20: getAdd(20),
                        main_unit_uid: main_unit_uid, 
                        erp_category_uid: row['Brick Code'], 
                    });
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
            }
        }
        if (i % 1000 === 0) {
            await yieldToUI();
            updateStatus({ message: 'Processing item details...', status: 'processing', progress: 25 + Math.round((i / df_template.length) * 25) });
        }
    }

    if (masteritemsData.length === 0) {
        updateStatus({ message: 'No valid item data found in Masteritems file, skipping CSV generation.', status: 'success', progress: 100 });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    csvs.push({ 
        name: `masteritems_${dateStr}.csv`, 
        rowCount: masteritemsData.length,
        content: await arrayToCsv(
            masteritemsData, 
            ['item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional', 'is_deleted', 'additional_1', 'additional_2', 'additional_3', 'additional_4', 'additional_5', 'additional_6', 'additional_7', 'additional_8', 'additional_9', 'additional_10', 'additional_11', 'additional_12', 'additional_13', 'additional_14', 'additional_15', 'additional_16', 'additional_17', 'additional_18', 'additional_19', 'additional_20', 'main_unit_uid', 'erp_category_uid'], 
            selectedColumns,
            (p) => updateStatus({ message: 'Generating Masteritems CSV...', status: 'processing', progress: 50 + Math.round(p / 10) }),
            signal,
            options.delimiter,
            options.columnMapping
        ) 
    });
    
    if((options.barcodes ?? true) && barcodesData.length > 0) csvs.push({ name: `barcodes_${dateStr}.csv`, rowCount: barcodesData.length, content: await arrayToCsv(barcodesData, ['item_uid', 'barcode', 'is_main'], null, (p) => updateStatus({ message: 'Generating Barcodes CSV...', status: 'processing', progress: 60 + Math.round(p / 10) }), signal, options.delimiter, options.columnMapping) });
    if((options.brands ?? true) && seenBrands.size > 0) csvs.push({ name: `brands_${dateStr}.csv`, rowCount: seenBrands.size, content: await arrayToCsv([...seenBrands].map(b => ({ brand_uid: b, name: b, is_deleted: 0 })), ['brand_uid', 'name', 'is_deleted'], selectedColumns, (p) => updateStatus({ message: 'Generating Brands CSV...', status: 'processing', progress: 70 + Math.round(p / 10) }), signal, options.delimiter, options.columnMapping) });
    if((options.dimensions ?? true) && dimensionsData.length > 0) csvs.push({ name: `dimensions_${dateStr}.csv`, rowCount: dimensionsData.length, content: await arrayToCsv(dimensionsData, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'netweight', 'volume', 'dimension_uid', 'coef', 'is_deleted'], selectedColumns, (p) => updateStatus({ message: 'Generating Dimensions CSV...', status: 'processing', progress: 80 + Math.round(p / 10) }), signal, options.delimiter, options.columnMapping) });
    if ((options.erpcategories ?? true) && seenErpCategories.size > 0) csvs.push({ name: `erpcategories_${dateStr}.csv`, rowCount: seenErpCategories.size, content: await arrayToCsv([...seenErpCategories.values()], ['erp_category_uid', 'name', 'parent_category_uid'], selectedColumns, (p) => updateStatus({ message: 'Generating ERP Categories CSV...', status: 'processing', progress: 90 + Math.round(p / 5) }), signal, options.delimiter, options.columnMapping) });
    if((options.manufacturers ?? true) && seenManufacturers.size > 0) csvs.push({ name: `manufacturers_${dateStr}.csv`, rowCount: seenManufacturers.size, content: await arrayToCsv([...seenManufacturers].map(m => ({ manufacturer_uid: m, name: m, is_deleted: 0 })), ['manufacturer_uid', 'name', 'is_deleted'], selectedColumns, (p) => updateStatus({ message: 'Generating Manufacturers CSV...', status: 'processing', progress: 95 + Math.round(p / 5) }), signal, options.delimiter, options.columnMapping) });
    
    updateStatus({ message: 'Masteritems processing complete.', status: 'success', progress: 100 });
    return csvs;
}

/**
 * Processes the new, complex "Item Master" file format (V2).
 */
async function processItemMasterV2File(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing new Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find header row in the new Masteritems file.');
    
    const df_template_raw: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    const df_template: ExcelRow[] = [];
    for (let i = 0; i < df_template_raw.length; i++) {
        const rowObject = df_template_raw[i];
        if (rowObject['UID*'] !== null && rowObject['UID*'] !== undefined && String(rowObject['UID*']).trim() !== '') {
            df_template.push(rowObject);
        }
    }

    if (df_template.length === 0) {
        updateStatus({ message: 'No valid item data found in new Masteritems file, skipping CSV generation.', status: 'success', progress: 100 });
        return [];
    }

    const masteritemsData: any[] = [], barcodesData: any[] = [], dimensionsData: any[] = [];
    const brandsMap = new Map<string, any>(), manufacturersMap = new Map<string, any>(), erpCategoriesMap = new Map<string, any>();

    for (let i = 0; i < df_template.length; i++) {
        const row = df_template[i];
        const item_uid = String(row['UID*']);
        
        let erpCategoryUid: string | number | Date | null = null;
        for (let j = 6; j >= 1; j--) {
            const catUidCol = `Category level ${j} UID`;
            const catNameCol = `Category level ${j}`;
            const uidVal = row[catUidCol];
            const nameVal = row[catNameCol];
            
            if (uidVal !== null && uidVal !== undefined && String(uidVal).trim() !== '') {
                erpCategoryUid = uidVal;
                break;
            } else if (nameVal !== null && nameVal !== undefined && String(nameVal).trim() !== '') {
                erpCategoryUid = nameVal;
                break;
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

        const getAddVal = (row: any, num: number) => row[`Add ${num}`] ?? row[`additional_${num}`];
        const getFloatAddVal = (row: any, num: number) => {
            const val = getAddVal(row, num);
            return val !== null && val !== undefined && String(val).trim() !== '' ? (parseFloat(String(val).replace(',', '.')) || null) : null;
        };

        masteritemsData.push({
            item_uid: item_uid, name: row['Product name*'], manufacturer_uid: effectiveManufUid, brand_uid: effectiveBrandUid,
            is_fractional: parseInt(String(row['Is fractional?']), 10) || 0,
            main_unit_uid: main_unit_uid, is_deleted: getIsDeletedValue(row),
            additional_1: getFloatAddVal(row, 1),
            additional_2: getAddVal(row, 2),
            additional_3: getAddVal(row, 3),
            additional_4: getAddVal(row, 4),
            additional_5: getFloatAddVal(row, 5),
            additional_6: getFloatAddVal(row, 6),
            additional_7: getAddVal(row, 7),
            additional_8: getFloatAddVal(row, 8),
            additional_9: getFloatAddVal(row, 9),
            additional_10: getFloatAddVal(row, 10),
            additional_11: getFloatAddVal(row, 11),
            additional_12: getFloatAddVal(row, 12),
            additional_13: getFloatAddVal(row, 13),
            additional_14: getFloatAddVal(row, 14),
            additional_15: getFloatAddVal(row, 15),
            additional_16: getFloatAddVal(row, 16),
            additional_17: getFloatAddVal(row, 17),
            additional_18: getFloatAddVal(row, 18),
            additional_19: getFloatAddVal(row, 19),
            additional_20: getFloatAddVal(row, 20),
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

        if (i % 1000 === 0) await yieldToUI();
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    const masteritems_cols = [
        'item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional', 'additional_1', 'additional_2', 'additional_3', 'additional_4', 'additional_5', 'additional_6', 'additional_7', 'additional_8', 'additional_9', 'additional_10', 'additional_11', 'additional_12', 'additional_13', 'additional_14', 'additional_15', 'additional_16', 'additional_17', 'additional_18', 'additional_19', 'additional_20', 'main_unit_uid', 'erp_category_uid', 'is_deleted'
    ];
    csvs.push({ name: `masteritems_${dateStr}.csv`, rowCount: masteritemsData.length, content: await arrayToCsv(masteritemsData, masteritems_cols, selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });

    if((options.barcodes ?? true) && barcodesData.length > 0) csvs.push({ name: `barcodes_${dateStr}.csv`, rowCount: barcodesData.length, content: await arrayToCsv(barcodesData, ['item_uid', 'barcode', 'is_main'], null, undefined, signal, options.delimiter, options.columnMapping) });
    if((options.brands ?? true) && brandsMap.size > 0) csvs.push({ name: `brands_${dateStr}.csv`, rowCount: brandsMap.size, content: await arrayToCsv(Array.from(brandsMap.values()), ['brand_uid', 'name', 'is_deleted'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
    if((options.dimensions ?? true) && dimensionsData.length > 0) csvs.push({ name: `dimensions_${dateStr}.csv`, rowCount: dimensionsData.length, content: await arrayToCsv(dimensionsData, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'coef', 'is_deleted', 'dimension_uid'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
    if ((options.erpcategories ?? true) && erpCategoriesMap.size > 0) csvs.push({ name: `erpcategories_${dateStr}.csv`, rowCount: erpCategoriesMap.size, content: await arrayToCsv(Array.from(erpCategoriesMap.values()), ['erp_category_uid', 'name', 'parent_category_uid'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
    if((options.manufacturers ?? true) && manufacturersMap.size > 0) csvs.push({ name: `manufacturers_${dateStr}.csv`, rowCount: manufacturersMap.size, content: await arrayToCsv(Array.from(manufacturersMap.values()), ['manufacturer_uid', 'name', 'is_deleted'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
    
    updateStatus({ message: 'New Masteritems processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes the updated "Item Master" file format based on the provided Python script logic.
 */
async function processItemMasterUpdatedFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing updated Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find a valid header row in the updated Masteritems sheet.');

    const df_template: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    // 1. Masteritems CSV
    updateStatus({ message: 'Processing Masteritems...', status: 'processing' });
    const required_columns = ['UID*', 'Product name*', 'Manufacturer', 'Brand', 'Is fractional?', 'Segment Description', 'Family Description', 'Class Description', 'Brick Code', 'Brick Description', 'Main Unit UID'];

    // Check if all columns exist, and add empty ones if missing
    for (let i = 0; i < df_template.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const row = df_template[i];
        required_columns.forEach(col => {
            if (!(col in row)) {
                row[col] = null;
            }
        });
        if (i % 1000 === 0) await yieldToUI();
    }

    const masteritems_columns = {
        'UID*': 'item_uid',
        'Product name*': 'name',
        'Manufacturer': 'manufacturer_uid',
        'Brand': 'brand_uid',
        'Is fractional?': 'is_fractional',
        'Main Unit UID': 'main_unit_uid',
        'Brick Code': 'erp_category_uid',
        'Segment Description': 'additional_1',
        'Family Description': 'additional_2',
        'Class Description': 'additional_3',
        'Brick Description': 'additional_4'
    };

    let df_masteritems: any[] = [];
    for (let i = 0; i < df_template.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const row = df_template[i];
        const newRow: Record<string, any> = {};
        for (const key in masteritems_columns) {
            newRow[masteritems_columns[key as keyof typeof masteritems_columns]] = row[key];
        }
        
        const is_fractional_val = newRow['is_fractional'];
        const parsed_val = parseInt(String(is_fractional_val), 10);
        newRow['is_fractional'] = isNaN(parsed_val) ? 0 : parsed_val;

        newRow['is_deleted'] = getIsDeletedValue(row);

        // Extract additional_1 to additional_20 if present, fallback to Segment Description etc for 1-4
        for (let j = 1; j <= 20; j++) {
            const val = row[`Add ${j}`] ?? row[`additional_${j}`];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
                if ([2, 3, 4, 7].includes(j)) {
                    newRow[`additional_${j}`] = val; // text
                } else {
                    newRow[`additional_${j}`] = parseFloat(String(val).replace(',', '.')) || null; // float
                }
            } else if (j > 4 && !( `additional_${j}` in newRow )) {
                newRow[`additional_${j}`] = null;
            }
        }

        if (newRow.item_uid && newRow.name) {
            df_masteritems.push(newRow);
        }
        if (i % 1000 === 0) await yieldToUI();
    }
    
    // drop duplicates
    const seenMasterItems = new Set<string>();
    df_masteritems = df_masteritems.filter(row => {
        const id = String(row.item_uid);
        if (seenMasterItems.has(id)) {
            return false;
        }
        seenMasterItems.add(id);
        return true;
    });


    const column_order = [
        'item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional', 'is_deleted',
        'additional_1', 'additional_2', 'additional_3', 'additional_4',
        'additional_5', 'additional_6', 'additional_7', 'additional_8',
        'additional_9', 'additional_10', 'additional_11', 'additional_12',
        'additional_13', 'additional_14', 'additional_15', 'additional_16',
        'additional_17', 'additional_18', 'additional_19', 'additional_20',
        'main_unit_uid', 'erp_category_uid'
    ];

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    csvs.push({ name: `masteritems_${dateStr}.csv`, rowCount: df_masteritems.length, content: await arrayToCsv(df_masteritems, column_order, selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
    updateStatus({ message: 'Masteritems processed.', status: 'success' });

    // 2. Barcodes CSV
    if (options.barcodes ?? true) {
        updateStatus({ message: 'Processing Barcodes...', status: 'processing' });
        const df_barcodes: any[] = [];
        for (let i = 0; i < df_template.length; i++) {
            if (signal?.aborted) throw new Error('Processing cancelled by user');
            const row = df_template[i];
            if (row['UID*'] && row['Barcode']) {
                df_barcodes.push({
                    item_uid: row['UID*'],
                    barcode: row['Barcode'],
                    is_main: 1
                });
            }
            if (i % 1000 === 0) await yieldToUI();
        }
        if (df_barcodes.length > 0) {
            csvs.push({ name: `barcodes_${dateStr}.csv`, rowCount: df_barcodes.length, content: await arrayToCsv(df_barcodes, ['item_uid', 'barcode', 'is_main'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
        }
        updateStatus({ message: 'Barcodes processed.', status: 'success' });
    }

    // 3. Brands CSV
    if (options.brands ?? true) {
        updateStatus({ message: 'Processing Brands...', status: 'processing' });
        const brandsSet = new Set<string>();
        for (let i = 0; i < df_template.length; i++) {
            if (signal?.aborted) throw new Error('Processing cancelled by user');
            const brand = df_template[i]['Brand'];
            if (brand) brandsSet.add(String(brand));
            if (i % 1000 === 0) await yieldToUI();
        }
        const df_brands = Array.from(brandsSet).map(brand => ({
            brand_uid: brand,
            name: brand,
            is_deleted: 0
        }));
        if (df_brands.length > 0) {
            csvs.push({ name: `brands_${dateStr}.csv`, rowCount: df_brands.length, content: await arrayToCsv(df_brands, ['brand_uid', 'name', 'is_deleted'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
        }
        updateStatus({ message: 'Brands processed.', status: 'success' });
    }

    // 4. Dimensions CSV
    if (options.dimensions ?? true) {
        updateStatus({ message: 'Processing Dimensions...', status: 'processing' });
        const df_dimensions: any[] = [];
        for (let i = 0; i < df_template.length; i++) {
            if (signal?.aborted) throw new Error('Processing cancelled by user');
            const row = df_template[i];
            if (row['UID*'] && row['Unit']) {
                const newRow: Record<string, any> = {
                    item_uid: row['UID*'],
                    unit_name: row['Unit'],
                    width: row['Width'],
                    height: row['Height'],
                    depth: row['Length'],
                    netweight: row['Netweight'],
                    volume: row['Volume'],
                    dimension_uid: row['Main Unit UID'],
                    coef: 1,
                    is_deleted: 0
                };
                ['width', 'height', 'depth', 'netweight', 'volume'].forEach(col => {
                    if (newRow[col]) {
                        newRow[col] = parseFloat(String(newRow[col]).replace(',', '.'));
                        if (isNaN(newRow[col])) newRow[col] = null;
                    }
                });
                df_dimensions.push(newRow);
            }
            if (i % 1000 === 0) await yieldToUI();
        }
        if (df_dimensions.length > 0) {
            csvs.push({ name: `dimensions_${dateStr}.csv`, rowCount: df_dimensions.length, content: await arrayToCsv(df_dimensions, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'netweight', 'volume', 'dimension_uid', 'coef', 'is_deleted'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
        }
        updateStatus({ message: 'Dimensions processed.', status: 'success' });
    }

    // 5. ERP Categories CSV
    if (options.erpcategories ?? true) {
        updateStatus({ message: 'Processing ERP Categories...', status: 'processing' });
        const erp_category_list: any[] = [];
        const level_mapping: Record<number, string> = { 1: 'Segment', 2: 'Family', 3: 'Class', 4: 'Brick' };

        for (let level_num = 1; level_num <= 4; level_num++) {
            const level_name = level_mapping[level_num];
            const uid_col = `${level_name} Code`;
            const name_col = `${level_name} Description`;
            const parent_level_num = level_num - 1;
            const parent_uid_col = parent_level_num > 0 ? `${level_mapping[parent_level_num]} Code` : null;

            if (headersInfo.headers.includes(uid_col) && headersInfo.headers.includes(name_col)) {
                const seenCategories = new Set<string>();
                for (let i = 0; i < df_template.length; i++) {
                    if (signal?.aborted) throw new Error('Processing cancelled by user');
                    const row = df_template[i];
                    if (row[uid_col]) {
                        const catId = String(row[uid_col]);
                        if (!seenCategories.has(catId)) {
                            seenCategories.add(catId);
                            erp_category_list.push({
                                erp_category_uid: row[uid_col],
                                name: row[name_col],
                                parent_category_uid: parent_uid_col ? row[parent_uid_col] : null
                            });
                        }
                    }
                    if (i % 1000 === 0) await yieldToUI();
                }
            }
            await yieldToUI();
        }

        if (erp_category_list.length > 0) {
            csvs.push({ name: `erpcategories_${dateStr}.csv`, rowCount: erp_category_list.length, content: await arrayToCsv(erp_category_list, ['erp_category_uid', 'name', 'parent_category_uid'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
        }
        updateStatus({ message: 'ERP Categories processed.', status: 'success' });
    }

    // 6. Manufacturers CSV
    if (options.manufacturers ?? true) {
        updateStatus({ message: 'Processing Manufacturers...', status: 'processing' });
        const manufacturersSet = new Set<string>();
        for (let i = 0; i < df_template.length; i++) {
            if (signal?.aborted) throw new Error('Processing cancelled by user');
            const m = df_template[i]['Manufacturer'];
            if (m) manufacturersSet.add(String(m));
            if (i % 1000 === 0) await yieldToUI();
        }
        const df_manufacturers = Array.from(manufacturersSet).map(m => ({
            manufacturer_uid: m,
            name: m,
            is_deleted: 0
        }));
        if (df_manufacturers.length > 0) {
            csvs.push({ name: `manufacturers_${dateStr}.csv`, rowCount: df_manufacturers.length, content: await arrayToCsv(df_manufacturers, ['manufacturer_uid', 'name', 'is_deleted'], selectedColumns, undefined, signal, options.delimiter, options.columnMapping) });
        }
        updateStatus({ message: 'Manufacturers processed.', status: 'success' });
    }

    updateStatus({ message: 'Updated Masteritems processing complete.', status: 'success' });
    return csvs;
}

// --- MAIN DISPATCHER ---

async function processStockFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Stock file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find a valid header row in the Stock sheet.');

    const df_template: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    const stockData: any[] = [];
    for (let i = 0; i < df_template.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const rowObject = df_template[i];

        stockData.push({
            store_uid: rowObject['StoreID'],
            item_uid: rowObject['ItemUID'],
            stock: parseFloat(String(rowObject['Quantity'] || '0').replace(',', '.')) || null,
        });

        if (i % 1000 === 0) {
            await yieldToUI();
            updateStatus({ message: 'Parsing stock data...', status: 'processing', progress: Math.round((i / df_template.length) * 50) });
        }
    }

    const filteredStockData = stockData.filter(row => row.item_uid && row.store_uid);
    
    if (filteredStockData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Stock file, skipping CSV generation.', status: 'success', progress: 100 });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `stock_${dateStr}.csv`,
        rowCount: filteredStockData.length,
        content: await arrayToCsv(
            filteredStockData, 
            ["store_uid", "item_uid", "stock"], 
            selectedColumns,
            (p) => updateStatus({ message: 'Generating Stock CSV...', status: 'processing', progress: 50 + Math.round(p / 2) }),
            signal,
            options.delimiter,
            options.columnMapping
        )
    }];
    
    updateStatus({ message: 'Stock processing complete.', status: 'success', progress: 100 });
    return csvs;
}

async function processPriceFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null, headersInfo: { headers: string[], dataStartIndex: number }, signal?: AbortSignal): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Price file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    if (headersInfo.dataStartIndex === -1) throw new Error('Could not find a valid header row in the Price sheet.');

    const df_template: any[] = XLSX.utils.sheet_to_json(worksheet, { 
        header: headersInfo.headers, 
        range: headersInfo.dataStartIndex, 
        defval: null, 
        raw: false 
    });
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');

    const priceData: any[] = [];
    for (let i = 0; i < df_template.length; i++) {
        if (signal?.aborted) throw new Error('Processing cancelled by user');
        const rowObject = df_template[i];

        priceData.push({
            item_uid: rowObject['ItemUID'],
            price_list: rowObject['PriceList'],
            price: parseFloat(String(rowObject['Price'] || '0').replace(',', '.')) || null,
        });

        if (i % 1000 === 0) {
            await yieldToUI();
            updateStatus({ message: 'Parsing price data...', status: 'processing', progress: Math.round((i / df_template.length) * 50) });
        }
    }

    const filteredPriceData = priceData.filter(row => row.item_uid && row.price_list);
    
    if (filteredPriceData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Price file, skipping CSV generation.', status: 'success', progress: 100 });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `price_${dateStr}.csv`,
        rowCount: filteredPriceData.length,
        content: await arrayToCsv(
            filteredPriceData, 
            ["item_uid", "price_list", "price"], 
            selectedColumns,
            (p) => updateStatus({ message: 'Generating Price CSV...', status: 'processing', progress: 50 + Math.round(p / 2) }),
            signal,
            options.delimiter,
            options.columnMapping
        )
    }];
    
    updateStatus({ message: 'Price processing complete.', status: 'success', progress: 100 });
    return csvs;
}

export const generateCsvsFromExcel = async (
    file: File,
    updateStatus: StatusUpdateCallback,
    options: CsvGenerationOptions = {},
    selectedColumns: Record<string, boolean> | null = null,
    isDetectionOnly: boolean = false,
    signal?: AbortSignal
): Promise<{ csvFiles: CsvFile[]; detectedType: FileType; headers: string[] }> => {
    
    if (!isDetectionOnly) {
        updateStatus({ message: 'Reading and analyzing Excel file...', status: 'processing' });
    }
    const data = await file.arrayBuffer();
    if (signal?.aborted) throw new Error('Processing cancelled by user');
    // cellNF: true and cellText: true are crucial for preserving formatted values (like leading zeros)
    const readOptions: any = { cellDates: true, cellNF: true, cellText: true, dense: true };
    if (isDetectionOnly) {
        readOptions.sheetRows = 100; // Read a bit more for safer detection
    }
    const workbook = XLSX.read(data, readOptions);
    await yieldToUI();
    if (signal?.aborted) throw new Error('Processing cancelled by user');
    
    const { type: detectedType, sheetName } = detectFileType(workbook);
    
    if (detectedType === 'UNKNOWN' || !sheetName) {
        throw new Error('Unknown file type. The file does not match any known templates.');
    }

    const definition = FILE_TYPE_DEFINITIONS[detectedType as keyof typeof FILE_TYPE_DEFINITIONS];
    const skipRows = detectedType === 'ITEM_MASTER_UPDATED' ? 1 : 0;
    const headersInfo = definition ? getHeadersAndDataStart(workbook.Sheets[sheetName], definition.keywords, skipRows) : { headers: [], headerRowIndex: -1, dataStartIndex: -1 };
    const headers = headersInfo.headers;
    
    let csvFiles: CsvFile[] = [];

    switch (detectedType) {
        case 'STORE':
            csvFiles = await processStoresFile(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
        case 'STORE_ITEMS':
            csvFiles = await processStoreItemsFile(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
        case 'ITEM_MASTER':
            csvFiles = await processItemMasterFile(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
        case 'ITEM_MASTER_UPDATED':
            csvFiles = await processItemMasterUpdatedFile(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
        case 'ITEM_MASTER_V2':
            csvFiles = await processItemMasterV2File(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
        case 'FACTS':
            csvFiles = await processFactsFile(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
        case 'STOCK':
            csvFiles = await processStockFile(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
        case 'PRICE':
            csvFiles = await processPriceFile(workbook, sheetName, updateStatus, options, selectedColumns, headersInfo, signal);
            break;
    }

    return { csvFiles, detectedType, headers };
};