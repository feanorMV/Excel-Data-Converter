
import * as XLSX from 'xlsx';
import type { StatusUpdateCallback, ExcelRow, CsvFile, FileType, CsvGenerationOptions } from '../types.ts';

const yieldToUI = () => new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));

// --- UTILITY FUNCTIONS ---

function getTodayDateString(): string {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
}

function arrayToCsv(data: Record<string, any>[], columns: string[], selectedColumns: Record<string, boolean> | null = null): string {
    const finalColumns = selectedColumns ? columns.filter(col => selectedColumns[col] ?? true) : columns;
    if (finalColumns.length === 0) return '';

    const header = finalColumns.join(',') + '\n';
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

function getHeaders(worksheet: any, keywords: string[]): string[] {
    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    let headerRowIndex = -1;
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && keywords.some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) return [];
    return rawData[headerRowIndex].map(h => String(h || '').trim());
}

// --- FILE TYPE DETECTION ---

const FILE_TYPE_DEFINITIONS = {
    STORE: { keywords: ["Store UID*"] },
    STORE_ITEMS: { keywords: ["Store UID*", "Product UID*", "In assortment?", "Purchase price"] },
    FACTS: { keywords: ["Product UID*", "Store UID*", "Date*"] },
    ITEM_MASTER_UPDATED: { keywords: ["UID*", "Product name*", "Manufacturer", "Brand", "Is fractional?", "Segment Description", "Brick Code"] },
    ITEM_MASTER_V2: { keywords: ["UID*", "Product name*", "Manufacturer UID"] },
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
        
        // Read rows for detection. Since we use sheetRows: 50 in XLSX.read for detection, 
        // jsonData will already be limited to 50 rows if isDetectionOnly was true.
        const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
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
async function processStoresFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, selectedColumns: Record<string, boolean> | null): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Stores file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    await yieldToUI();

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

    const storesData: any[] = [];
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });
        
        storesData.push({
            store_uid: rowObject['Store UID*'],
            name: rowObject['Store name*'],
            region: rowObject['Region'],
            group_name: rowObject['Group name'],
            floor_space: parseInt(String(rowObject['Square']), 10) || 0,
            in_shelf: parseInt(String(rowObject['In Shelf?']), 10) || 0,
            licence_start_date: '2023-01-01',
            is_deleted: parseInt(String(rowObject['To delete']), 10) || 0,
        });

        if (i % 1000 === 0) await yieldToUI();
    }

    const filteredStoresData = storesData.filter(row => row.store_uid);

    if (filteredStoresData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Stores file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `stores_${dateStr}.csv`,
        content: arrayToCsv(filteredStoresData, ['store_uid', 'name', 'region', 'group_name', 'floor_space', 'in_shelf', 'licence_start_date', 'is_deleted'], selectedColumns)
    }];

    updateStatus({ message: 'Stores processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes a "Store Items" file containing assortment, pricing, and supplier data.
 */
async function processStoreItemsFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Store Items file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    await yieldToUI();

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

    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
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

        if (i % 1000 === 0) await yieldToUI();
    }
    
    if (itemsData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Items file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];

    csvs.push({
        name: `items_${dateStr}.csv`,
        content: arrayToCsv(itemsData, ['item_uid', 'store_uid', 'is_active_planogram', 'purchase_price', 'retail_price', 'external_supplier_uid'], selectedColumns)
    });

    if ((options.suppliers ?? true) && suppliersMap.size > 0) {
        csvs.push({
            name: `suppliers_${dateStr}.csv`,
            content: arrayToCsv(Array.from(suppliersMap.values()), ['supplier_uid', 'name', 'is_deleted'], selectedColumns)
        });
    }
    
    updateStatus({ message: 'Store Items processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes a "Facts" file containing sales and stock data.
 */
async function processFactsFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, selectedColumns: Record<string, boolean> | null): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Facts file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    await yieldToUI();

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

    const factsData: any[] = [];
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
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

        if (i % 1000 === 0) await yieldToUI();
    }

    const filteredFactsData = factsData.filter(row => row.item_uid && row.store_uid && row.date);
    
    if (filteredFactsData.length === 0) {
        updateStatus({ message: 'No valid data rows found in Facts file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [{
        name: `facts_${dateStr}.csv`,
        content: arrayToCsv(filteredFactsData, ["item_uid", "store_uid", "date", "stock", "sold_qty", "revenue", "cogs"], selectedColumns)
    }];
    
    updateStatus({ message: 'Facts processing complete.', status: 'success' });
    return csvs;
}


/**
 * Processes the original "Item Master" file format (V1).
 */
async function processItemMasterFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null): Promise<CsvFile[]> {
    updateStatus({ message: `Processing Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);
    
    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    await yieldToUI();

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
    const df_template: ExcelRow[] = [];
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });
        df_template.push(rowObject);
        if (i % 1000 === 0) await yieldToUI();
    }

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
            }
        }
        if (i % 1000 === 0) await yieldToUI();
    }

    if (masteritemsData.length === 0) {
        updateStatus({ message: 'No valid item data found in Masteritems file, skipping CSV generation.', status: 'success' });
        return [];
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    csvs.push({ name: `masteritems_${dateStr}.csv`, content: arrayToCsv(masteritemsData, ['item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional', 'additional_1', 'additional_2', 'additional_3', 'additional_4', 'main_unit_uid', 'erp_category_uid'], selectedColumns) });
    
    if((options.barcodes ?? true) && barcodesData.length > 0) csvs.push({ name: `barcodes_${dateStr}.csv`, content: arrayToCsv(barcodesData, ['item_uid', 'barcode', 'is_main']) });
    if((options.brands ?? true) && seenBrands.size > 0) csvs.push({ name: `brands_${dateStr}.csv`, content: arrayToCsv([...seenBrands].map(b => ({ brand_uid: b, name: b, is_deleted: 0 })), ['brand_uid', 'name', 'is_deleted'], selectedColumns) });
    if((options.dimensions ?? true) && dimensionsData.length > 0) csvs.push({ name: `dimensions_${dateStr}.csv`, content: arrayToCsv(dimensionsData, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'netweight', 'volume', 'dimension_uid', 'coef', 'is_deleted'], selectedColumns) });
    if ((options.erpcategories ?? true) && seenErpCategories.size > 0) csvs.push({ name: `erpcategories_${dateStr}.csv`, content: arrayToCsv([...seenErpCategories.values()], ['erp_category_uid', 'name', 'parent_category_uid'], selectedColumns) });
    if((options.manufacturers ?? true) && seenManufacturers.size > 0) csvs.push({ name: `manufacturers_${dateStr}.csv`, content: arrayToCsv([...seenManufacturers].map(m => ({ manufacturer_uid: m, name: m, is_deleted: 0 })), ['manufacturer_uid', 'name', 'is_deleted'], selectedColumns) });
    
    updateStatus({ message: 'Masteritems processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes the new, complex "Item Master" file format (V2).
 */
async function processItemMasterV2File(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null): Promise<CsvFile[]> {
    updateStatus({ message: `Processing new Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    await yieldToUI();

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
    
    const df_template: ExcelRow[] = [];
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });
        if (rowObject['UID*'] !== null && rowObject['UID*'] !== undefined && String(rowObject['UID*']).trim() !== '') {
            df_template.push(rowObject);
        }
        if (i % 1000 === 0) await yieldToUI();
    }

    if (df_template.length === 0) {
        updateStatus({ message: 'No valid item data found in new Masteritems file, skipping CSV generation.', status: 'success' });
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
            const val = row[catUidCol];
            if (val !== null && val !== undefined && String(val).trim() !== '') {
                erpCategoryUid = val;
                break;
            }
        }
        if (erpCategoryUid === null) {
            for (let j = 6; j >= 1; j--) {
                const catNameCol = `Category level ${j}`;
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

        if (i % 1000 === 0) await yieldToUI();
    }

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    const masteritems_cols = [
        'item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional', 'additional_1', 'additional_2', 'additional_3', 'additional_4', 'additional_5', 'additional_6', 'additional_7', 'additional_8', 'additional_9', 'additional_10', 'additional_11', 'additional_12', 'additional_13', 'additional_14', 'additional_15', 'additional_16', 'additional_17', 'additional_18', 'additional_19', 'additional_20', 'main_unit_uid', 'erp_category_uid', 'is_deleted'
    ];
    csvs.push({ name: `masteritems_${dateStr}.csv`, content: arrayToCsv(masteritemsData, masteritems_cols, selectedColumns) });

    if((options.barcodes ?? true) && barcodesData.length > 0) csvs.push({ name: `barcodes_${dateStr}.csv`, content: arrayToCsv(barcodesData, ['item_uid', 'barcode', 'is_main']) });
    if((options.brands ?? true) && brandsMap.size > 0) csvs.push({ name: `brands_${dateStr}.csv`, content: arrayToCsv(Array.from(brandsMap.values()), ['brand_uid', 'name', 'is_deleted'], selectedColumns) });
    if((options.dimensions ?? true) && dimensionsData.length > 0) csvs.push({ name: `dimensions_${dateStr}.csv`, content: arrayToCsv(dimensionsData, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'coef', 'is_deleted', 'dimension_uid'], selectedColumns) });
    if ((options.erpcategories ?? true) && erpCategoriesMap.size > 0) csvs.push({ name: `erpcategories_${dateStr}.csv`, content: arrayToCsv(Array.from(erpCategoriesMap.values()), ['erp_category_uid', 'name', 'parent_category_uid'], selectedColumns) });
    if((options.manufacturers ?? true) && manufacturersMap.size > 0) csvs.push({ name: `manufacturers_${dateStr}.csv`, content: arrayToCsv(Array.from(manufacturersMap.values()), ['manufacturer_uid', 'name', 'is_deleted'], selectedColumns) });
    
    updateStatus({ message: 'New Masteritems processing complete.', status: 'success' });
    return csvs;
}

/**
 * Processes the updated "Item Master" file format based on the provided Python script logic.
 */
async function processItemMasterUpdatedFile(workbook: any, sheetName: string, updateStatus: StatusUpdateCallback, options: CsvGenerationOptions, selectedColumns: Record<string, boolean> | null): Promise<CsvFile[]> {
    updateStatus({ message: `Processing updated Masteritems file from sheet "${sheetName}"...`, status: 'processing' });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });
    await yieldToUI();

    let headerRowIndex = -1;
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && ["UID*", "Product name*", "Barcode", "Manufacturer"].some(kw => cell.includes(kw)))) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error('Could not find a valid header row in the updated Masteritems sheet.');

    const headers = rawData[headerRowIndex].map(h => String(h || '').trim());
    const dataRows = rawData.slice(headerRowIndex + 2);

    const df_template: ExcelRow[] = [];
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const rowObject: ExcelRow = {};
        headers.forEach((header, index) => {
            if (header) {
                const value = row[index];
                rowObject[header] = typeof value === 'string' ? value.trim() : value;
            }
        });
        df_template.push(rowObject);
        if (i % 1000 === 0) await yieldToUI();
    }

    // 1. Masteritems CSV
    updateStatus({ message: 'Processing Masteritems...', status: 'processing' });
    const required_columns = ['UID*', 'Product name*', 'Manufacturer', 'Brand', 'Is fractional?', 'Segment Description', 'Family Description', 'Class Description', 'Brick Code', 'Brick Description', 'Main Unit UID'];

    // Check if all columns exist, and add empty ones if missing
    for (let i = 0; i < df_template.length; i++) {
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
        const row = df_template[i];
        const newRow: Record<string, any> = {};
        for (const key in masteritems_columns) {
            newRow[masteritems_columns[key as keyof typeof masteritems_columns]] = row[key];
        }
        
        const is_fractional_val = newRow['is_fractional'];
        const parsed_val = parseInt(String(is_fractional_val), 10);
        newRow['is_fractional'] = isNaN(parsed_val) ? 0 : parsed_val;

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
        'item_uid', 'name', 'manufacturer_uid', 'brand_uid', 'is_fractional',
        'additional_1', 'additional_2', 'additional_3', 'additional_4',
        'main_unit_uid', 'erp_category_uid'
    ];

    const dateStr = getTodayDateString();
    const csvs: CsvFile[] = [];
    csvs.push({ name: `masteritems_${dateStr}.csv`, content: arrayToCsv(df_masteritems, column_order, selectedColumns) });
    updateStatus({ message: 'Masteritems processed.', status: 'success' });

    // 2. Barcodes CSV
    if (options.barcodes ?? true) {
        updateStatus({ message: 'Processing Barcodes...', status: 'processing' });
        const df_barcodes: any[] = [];
        for (let i = 0; i < df_template.length; i++) {
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
            csvs.push({ name: `barcodes_${dateStr}.csv`, content: arrayToCsv(df_barcodes, ['item_uid', 'barcode', 'is_main'], selectedColumns) });
        }
        updateStatus({ message: 'Barcodes processed.', status: 'success' });
    }

    // 3. Brands CSV
    if (options.brands ?? true) {
        updateStatus({ message: 'Processing Brands...', status: 'processing' });
        const brandsSet = new Set<string>();
        for (let i = 0; i < df_template.length; i++) {
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
            csvs.push({ name: `brands_${dateStr}.csv`, content: arrayToCsv(df_brands, ['brand_uid', 'name', 'is_deleted'], selectedColumns) });
        }
        updateStatus({ message: 'Brands processed.', status: 'success' });
    }

    // 4. Dimensions CSV
    if (options.dimensions ?? true) {
        updateStatus({ message: 'Processing Dimensions...', status: 'processing' });
        const df_dimensions: any[] = [];
        for (let i = 0; i < df_template.length; i++) {
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
            csvs.push({ name: `dimensions_${dateStr}.csv`, content: arrayToCsv(df_dimensions, ['item_uid', 'unit_name', 'width', 'height', 'depth', 'netweight', 'volume', 'dimension_uid', 'coef', 'is_deleted'], selectedColumns) });
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

            if (headers.includes(uid_col) && headers.includes(name_col)) {
                const seenCategories = new Set<string>();
                for (let i = 0; i < df_template.length; i++) {
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
            csvs.push({ name: `erpcategories_${dateStr}.csv`, content: arrayToCsv(erp_category_list, ['erp_category_uid', 'name', 'parent_category_uid'], selectedColumns) });
        }
        updateStatus({ message: 'ERP Categories processed.', status: 'success' });
    }

    // 6. Manufacturers CSV
    if (options.manufacturers ?? true) {
        updateStatus({ message: 'Processing Manufacturers...', status: 'processing' });
        const manufacturersSet = new Set<string>();
        for (let i = 0; i < df_template.length; i++) {
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
            csvs.push({ name: `manufacturers_${dateStr}.csv`, content: arrayToCsv(df_manufacturers, ['manufacturer_uid', 'name', 'is_deleted'], selectedColumns) });
        }
        updateStatus({ message: 'Manufacturers processed.', status: 'success' });
    }

    updateStatus({ message: 'Updated Masteritems processing complete.', status: 'success' });
    return csvs;
}

// --- MAIN DISPATCHER ---

export const generateCsvsFromExcel = async (
    file: File,
    updateStatus: StatusUpdateCallback,
    options: CsvGenerationOptions = {},
    selectedColumns: Record<string, boolean> | null = null,
    isDetectionOnly: boolean = false
): Promise<{ csvFiles: CsvFile[]; detectedType: FileType; headers: string[] }> => {
    
    if (!isDetectionOnly) {
        updateStatus({ message: 'Reading and analyzing Excel file...', status: 'processing' });
    }
    const data = await file.arrayBuffer();
    // Use standard reading options without 'dense' mode which can sometimes cause issues with sheet_to_json
    const readOptions: any = { cellDates: true };
    if (isDetectionOnly) {
        readOptions.sheetRows = 100; // Read a bit more for safer detection
    }
    const workbook = XLSX.read(data, readOptions);
    await yieldToUI();
    
    const { type: detectedType, sheetName } = detectFileType(workbook);
    
    if (detectedType === 'UNKNOWN' || !sheetName) {
        throw new Error('Unknown file type. The file does not match any known templates.');
    }

    const definition = FILE_TYPE_DEFINITIONS[detectedType as keyof typeof FILE_TYPE_DEFINITIONS];
    const headers = definition ? getHeaders(workbook.Sheets[sheetName], definition.keywords) : [];
    
    let csvFiles: CsvFile[] = [];

    switch (detectedType) {
        case 'STORE':
            csvFiles = await processStoresFile(workbook, sheetName, updateStatus, selectedColumns);
            break;
        case 'STORE_ITEMS':
            csvFiles = await processStoreItemsFile(workbook, sheetName, updateStatus, options, selectedColumns);
            break;
        case 'ITEM_MASTER':
            csvFiles = await processItemMasterFile(workbook, sheetName, updateStatus, options, selectedColumns);
            break;
        case 'ITEM_MASTER_UPDATED':
            csvFiles = await processItemMasterUpdatedFile(workbook, sheetName, updateStatus, options, selectedColumns);
            break;
        case 'ITEM_MASTER_V2':
            csvFiles = await processItemMasterV2File(workbook, sheetName, updateStatus, options, selectedColumns);
            break;
        case 'FACTS':
            csvFiles = await processFactsFile(workbook, sheetName, updateStatus, selectedColumns);
            break;
        case 'STOCK':
        case 'PRICE':
            // Placeholder for other file types
            throw new Error(`Processing for "${detectedType}" files is not yet implemented.`);
    }

    return { csvFiles, detectedType, headers };
};