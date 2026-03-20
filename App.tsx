

import React, { useState, useCallback, useRef } from 'react';
import { FolderUp, Download, RefreshCw, AlertTriangle, CheckCircle, Loader2, FileText, XCircle, Settings } from 'lucide-react';
import JSZip from 'jszip';
import { generateCsvsFromExcel } from './services/excelProcessor.ts';
import type { StatusUpdate, FileType, CsvFile, CsvGenerationOptions } from './types.ts';

// Fix for non-standard directory attributes on input element
declare module 'react' {
    interface InputHTMLAttributes {
        webkitdirectory?: string;
        directory?: string;
    }
}

type ProcessingState = 'idle' | 'processing' | 'success' | 'error';
interface FileInfo {
    file: File;
    type: FileType;
    status: 'pending' | 'processing' | 'success' | 'error';
    error?: string;
    headers: string[];
    progress?: number;
}

interface LogEntry extends StatusUpdate {
    // Using file index to associate log entries with a specific file
    fileIndex: number;
}

const CsvOptionsConfig: Record<string, string[]> = {
  ITEM_MASTER: ['barcodes', 'brands', 'dimensions', 'erpcategories', 'manufacturers'],
  ITEM_MASTER_V2: ['barcodes', 'brands', 'dimensions', 'erpcategories', 'manufacturers'],
  ITEM_MASTER_UPDATED: ['barcodes', 'brands', 'dimensions', 'erpcategories', 'manufacturers'],
  STORE_ITEMS: ['suppliers'],
};

const getTodayDateString = (): string => {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
};

const ALL_OUTPUT_COLUMNS = [
    'store_uid', 'name', 'region', 'group_name', 'floor_space', 'in_shelf', 'licence_start_date', 'is_deleted',
    'item_uid', 'is_active_planogram', 'purchase_price', 'retail_price', 'external_supplier_uid',
    'supplier_uid', 'date', 'stock', 'sold_qty', 'revenue', 'cogs',
    'manufacturer_uid', 'brand_uid', 'is_fractional', 'additional_1', 'additional_2', 'additional_3', 'additional_4',
    'additional_5', 'additional_6', 'additional_7', 'additional_8', 'additional_9', 'additional_10', 'additional_11',
    'additional_12', 'additional_13', 'additional_14', 'additional_15', 'additional_16', 'additional_17', 'additional_18',
    'additional_19', 'additional_20', 'main_unit_uid', 'erp_category_uid', 'barcode', 'is_main',
    'unit_name', 'width', 'height', 'depth', 'netweight', 'volume', 'dimension_uid', 'coef', 'parent_category_uid'
];

export const App: React.FC = () => {
    const [fileInfos, setFileInfos] = useState<FileInfo[]>([]);
    const [archiveName, setArchiveName] = useState<string>('data_export');
    const [processingState, setProcessingState] = useState<ProcessingState>('idle');
    const [statusUpdates, setStatusUpdates] = useState<LogEntry[]>([]);
    const [errorMessage, setErrorMessage] = useState<string>('');
    const [zipUrl, setZipUrl] = useState<string | null>(null);
    const [zipFileName, setZipFileName] = useState<string>('');
    const [csvOptions, setCsvOptions] = useState<CsvGenerationOptions>({ delimiter: ',' });
    const [isAdvancedOptionsOpen, setIsAdvancedOptionsOpen] = useState(false);
    const [selectedColumns, setSelectedColumns] = useState<Record<string, boolean>>({});
    const [columnMapping, setColumnMapping] = useState<Record<string, string>>({});
    const [generatedFilesSummary, setGeneratedFilesSummary] = useState<{ name: string, rowCount: number }[]>([]);
    const fileInputRef = useRef<HTMLInputElement>(null);
    const abortControllerRef = useRef<AbortController | null>(null);

    const resetState = () => {
        setProcessingState('idle');
        setStatusUpdates([]);
        setErrorMessage('');
        setZipUrl(null);
        setCsvOptions({ delimiter: ',' });
        setColumnMapping({});
        setGeneratedFilesSummary([]);
        if (zipUrl) {
            URL.revokeObjectURL(zipUrl);
        }
        if (abortControllerRef.current) {
            abortControllerRef.current.abort();
            abortControllerRef.current = null;
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
            status: 'pending',
            headers: []
        }));
        setFileInfos(initialFileInfos);

        // Sequentially detect file types to prevent UI freezing
        const detectFiles = async () => {
            for (let i = 0; i < validFiles.length; i++) {
                const file = validFiles[i];
                
                // Allow UI to breathe and update
                await new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
                
                try {
                    // Use isDetectionOnly: true to read only the first 50 rows
                    const { detectedType, headers } = await generateCsvsFromExcel(file, () => {}, {}, null, true);
                    setFileInfos(currentInfos => {
                        const newInfos = [...currentInfos];
                        newInfos[i] = {
                            ...newInfos[i],
                            type: detectedType,
                            headers: headers || []
                        };
                        return newInfos;
                    });
                } catch (error) {
                    setFileInfos(currentInfos => {
                        const newInfos = [...currentInfos];
                        newInfos[i] = {
                            ...newInfos[i],
                            status: 'error',
                            error: error instanceof Error ? error.message : "Detection failed"
                        };
                        return newInfos;
                    });
                }
            }
        };

        detectFiles();
    };

    const handleCancel = useCallback(() => {
        if (abortControllerRef.current) {
            abortControllerRef.current.abort();
        }
    }, []);

    const handleProcess = useCallback(async () => {
        if (fileInfos.length === 0) return;

        setProcessingState('processing');
        setStatusUpdates([]);
        setErrorMessage('');
        setZipUrl(null);
        setGeneratedFilesSummary([]);
        
        abortControllerRef.current = new AbortController();
        
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
                if (update.progress !== undefined) {
                    setFileInfos(prev => prev.map((f, idx) => idx === i ? { ...f, progress: update.progress } : f));
                }
            };

            try {
                const { csvFiles } = await generateCsvsFromExcel(info.file, updateCallback, { ...csvOptions, columnMapping }, selectedColumns, false, abortControllerRef.current.signal);
                csvFiles.forEach(csv => allGeneratedCsvs.push(csv));
                hasProcessedAnyFile = true;
                setFileInfos(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'success', progress: 100 } : f));
                 // Update all 'processing' logs for this file to 'success'
                setStatusUpdates(prev => prev.map(log =>
                    (log.fileIndex === i && log.status === 'processing') ? { ...log, status: 'success', progress: 100 } : log
                ));

            } catch (error) {
                const message = error instanceof Error ? error.message : 'An unknown error occurred.';
                if (message === 'Processing cancelled by user') {
                    setStatusUpdates(prev => [...prev, { message: `Processing cancelled for ${info.file.name}`, status: 'error', fileIndex: i }]);
                    setFileInfos(prev => prev.map((f, idx) => idx === i ? { ...f, status: 'error', error: 'Cancelled' } : f));
                    break;
                }
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

        if (abortControllerRef.current?.signal.aborted) {
            setProcessingState('error');
            setErrorMessage(prev => prev ? `${prev}\nProcessing was cancelled.` : 'Processing was cancelled.');
            return;
        }

        if (allGeneratedCsvs.length > 0) {
            setGeneratedFilesSummary(allGeneratedCsvs.map(csv => ({ name: csv.name, rowCount: csv.rowCount })));
            // FIX: Wrap zip generation in a try-catch block to handle potential errors.
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

    const masterProgress = fileInfos.length > 0 
        ? Math.round(fileInfos.reduce((acc, curr) => acc + (curr.progress || (curr.status === 'success' ? 100 : 0)), 0) / fileInfos.length)
        : 0;

    return (
        <div className="min-h-screen bg-gray-50 text-gray-800 flex items-center justify-center p-4 transition-colors duration-300">
            <div className="w-full max-w-3xl mx-auto">
                    <header className="text-center mb-8">
                        <h1 className="text-4xl md:text-5xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-primary to-secondary">Excel Data Converter</h1>
                        <p className="mt-3 text-lg text-gray-600">Select a folder to convert all valid template files into a structured CSV archive.</p>
                        
                        {processingState === 'processing' && (
                            <div className="mt-6 max-w-md mx-auto">
                                <div className="flex justify-between mb-1 text-sm font-medium text-primary">
                                    <span>Overall Progress</span>
                                    <span>{masterProgress}%</span>
                                </div>
                                <div className="w-full bg-gray-200 rounded-full h-2.5 overflow-hidden">
                                    <div 
                                        className="bg-primary h-2.5 rounded-full transition-all duration-300 ease-out" 
                                        style={{ width: `${masterProgress}%` }}
                                    ></div>
                                </div>
                            </div>
                        )}
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
                            />
                        </div>
                    ) : (
                        <div className="space-y-3">
                             <h3 className="text-lg font-semibold border-b border-gray-200 pb-2">Detected Files:</h3>
                             <div className="max-h-60 overflow-y-auto space-y-2 pr-2">
                                {fileInfos.map((info, index) => (
                                    <div key={index} className={`flex flex-col p-3 rounded-md ${info.status === 'error' ? 'bg-red-50' : 'bg-gray-100'}`}>
                                        <div className="flex items-center">
                                            <StatusIcon status={info.status} />
                                            <div className="flex-grow">
                                                <p className="font-medium text-sm">{info.file.name}</p>
                                                <p className={`text-xs ${info.status === 'error' ? 'text-red-600' : 'text-gray-500'}`}>
                                                   {info.status === 'error' ? info.error : `Type: ${formatFileTypeName(info.type)}`}
                                                </p>
                                            </div>
                                            {info.status === 'processing' && info.progress !== undefined && (
                                                <span className="text-xs font-bold text-primary ml-2">{info.progress}%</span>
                                            )}
                                        </div>
                                        {info.status === 'processing' && info.progress !== undefined && (
                                            <div className="mt-2 w-full bg-gray-200 rounded-full h-1.5 overflow-hidden">
                                                <div 
                                                    className="bg-primary h-1.5 rounded-full transition-all duration-300 ease-out" 
                                                    style={{ width: `${info.progress}%` }}
                                                ></div>
                                            </div>
                                        )}
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
                        {processingState === 'processing' ? (
                            <button
                                onClick={handleCancel}
                                className="w-full inline-flex justify-center items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-red-600 hover:bg-red-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-red-500 transition-all duration-300"
                            >
                                <XCircle className="-ml-1 mr-3 h-5 w-5" />
                                Cancel Processing
                            </button>
                        ) : (
                            <button
                                onClick={handleProcess}
                                disabled={!hasValidFiles}
                                className="w-full inline-flex justify-center items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-primary hover:bg-primary-hover focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-focus disabled:bg-gray-400 disabled:cursor-not-allowed transition-all duration-300"
                            >
                                <CheckCircle className="-ml-1 mr-3 h-5 w-5" />
                                Start Processing
                            </button>
                        )}

                        <button
                            onClick={handleReset}
                            className="w-full sm:w-auto inline-flex justify-center items-center px-6 py-3 border border-gray-300 text-base font-medium rounded-md shadow-sm text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-focus transition-all duration-300"
                        >
                           <RefreshCw className="mr-3 h-5 w-5"/> Reset
                        </button>

                        <button
                            onClick={() => setIsAdvancedOptionsOpen(true)}
                            className="w-full sm:w-auto inline-flex justify-center items-center px-6 py-3 border border-gray-300 text-base font-medium rounded-md shadow-sm text-gray-700 bg-white hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary-focus transition-all duration-300"
                        >
                           <Settings className="mr-3 h-5 w-5"/> Advanced Options
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
                                className="w-full inline-flex justify-center items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-secondary hover:bg-green-600 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500 transition-all duration-300 mb-4"
                            >
                                <Download className="-ml-1 mr-3 h-5 w-5"/>
                                Download ZIP File ({zipFileName})
                            </a>

                            {generatedFilesSummary.length > 0 && (
                                <div className="mt-4 bg-gray-50 rounded-lg p-4 border border-gray-200">
                                    <h4 className="text-md font-semibold mb-3 text-gray-800">Generated Files Summary:</h4>
                                    <ul className="space-y-2 text-sm text-gray-600 max-h-60 overflow-y-auto pr-2">
                                        {generatedFilesSummary.map((file, idx) => (
                                            <li key={idx} className="flex justify-between items-center border-b border-gray-100 pb-2 last:border-0 last:pb-0">
                                                <span className="font-medium text-gray-700 truncate mr-4" title={file.name}>{file.name}</span>
                                                <span className="bg-white px-2 py-1 rounded shadow-sm text-xs font-bold text-primary whitespace-nowrap">
                                                    {file.rowCount.toLocaleString()} rows
                                                </span>
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                            )}
                        </div>
                    )}
                </main>
                 <footer className="text-center mt-8">
                    <p className="text-sm text-gray-500">
                        For Internal Use Only
                    </p>
                </footer>
            </div>

            {isAdvancedOptionsOpen && (
                <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4">
                    <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-2xl w-full">
                        <h2 className="text-2xl font-bold mb-4">Advanced Options</h2>
                        <div className="space-y-6">
                            <div>
                                <h3 className="text-lg font-semibold">CSV Delimiter</h3>
                                <p className="text-sm text-gray-500 mb-2">Select the delimiter to use for the generated CSV files.</p>
                                <select
                                    value={csvOptions.delimiter as string}
                                    onChange={(e) => setCsvOptions(prev => ({ ...prev, delimiter: e.target.value }))}
                                    className="block w-full max-w-xs px-3 py-2 bg-white border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-primary-focus focus:border-primary-focus sm:text-sm"
                                >
                                    <option value=",">Comma (,)</option>
                                    <option value=";">Semicolon (;)</option>
                                    <option value="\t">Tab</option>
                                    <option value="|">Pipe (|)</option>
                                </select>
                            </div>

                            <div>
                                <h3 className="text-lg font-semibold">Column Filtering & Mapping</h3>
                                <p className="text-sm text-gray-500 mb-2">Select which output columns to include and optionally rename them.</p>
                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-3 pt-2 max-h-96 overflow-y-auto pr-2">
                                    {ALL_OUTPUT_COLUMNS.map(header => (
                                        <div key={header} className="flex items-center space-x-3">
                                            <input
                                                type="checkbox"
                                                checked={selectedColumns[header] ?? true}
                                                onChange={() => setSelectedColumns(prev => ({ ...prev, [header]: !(prev[header] ?? true) }))}
                                                className="h-4 w-4 rounded border-gray-300 text-primary focus:ring-primary-focus transition"
                                            />
                                            <div className="flex flex-col flex-grow">
                                                <span className="text-sm font-medium text-gray-700">{header}</span>
                                                <input
                                                    type="text"
                                                    placeholder="Rename to..."
                                                    value={columnMapping[header] || ''}
                                                    onChange={(e) => setColumnMapping(prev => ({ ...prev, [header]: e.target.value }))}
                                                    disabled={!(selectedColumns[header] ?? true)}
                                                    className="mt-1 block w-full px-2 py-1 text-xs border border-gray-300 rounded shadow-sm focus:outline-none focus:ring-primary-focus focus:border-primary-focus disabled:bg-gray-100 disabled:text-gray-400"
                                                />
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </div>
                        <div className="mt-6 flex justify-end">
                            <button 
                                onClick={() => setIsAdvancedOptionsOpen(false)} 
                                className="px-6 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition-colors duration-300"
                            >
                                Close
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};
