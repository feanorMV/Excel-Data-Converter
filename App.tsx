
import React, { useState, useCallback, useRef } from 'react';
import { FolderUp, Download, RefreshCw, AlertTriangle, CheckCircle, Loader2, FileCheck2, FileText, XCircle } from 'lucide-react';
import { generateCsvsFromExcel } from './services/excelProcessor';
import type { StatusUpdate, FileType, CsvFile } from './types';

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


const getTodayDateString = (): string => {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
};

const App: React.FC = () => {
    const [fileInfos, setFileInfos] = useState<FileInfo[]>([]);
    const [archiveName, setArchiveName] = useState<string>('data_export');
    const [processingState, setProcessingState] = useState<ProcessingState>('idle');
    const [statusUpdates, setStatusUpdates] = useState<LogEntry[]>([]);
    const [errorMessage, setErrorMessage] = useState<string>('');
    const [zipUrl, setZipUrl] = useState<string | null>(null);
    const [zipFileName, setZipFileName] = useState<string>('');
    const fileInputRef = useRef<HTMLInputElement>(null);

    const resetState = () => {
        setProcessingState('idle');
        setStatusUpdates([]);
        setErrorMessage('');
        setZipUrl(null);
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
                const { detectedType } = await generateCsvsFromExcel(file, () => {});
                return { index, type: detectedType };
            } catch (error) {
                return { index, type: 'UNKNOWN', error: error instanceof Error ? error.message : "Detection failed" };
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
                const { csvFiles } = await generateCsvsFromExcel(info.file, updateCallback);
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
        } else if (hasProcessedAnyFile && !localErrorMessages) {
            setProcessingState('success');
            setStatusUpdates(prev => [...prev, { message: 'Processing complete. All valid files were processed but contained no data to export.', status: 'success', fileIndex: -1 }]);
        } else {
            setProcessingState('error');
            if (!localErrorMessages) {
                setErrorMessage("No valid files were processed successfully.");
            }
        }

    }, [fileInfos, archiveName]);

    const handleReset = () => {
        setFileInfos([]);
        resetState();
        if (fileInputRef.current) {
            fileInputRef.current.value = '';
        }
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
        if (typeName === 'ITEM_MASTER_V2') return 'Item Master (New Format)';
        if (typeName === 'FACTS') return 'Facts Data';
        if (typeName === 'STORE_ITEMS') return 'Store Items';
        if (typeName === 'STORE') return 'Stores';
        return typeName.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
    };

    const hasValidFiles = fileInfos.some(f => f.type !== 'UNKNOWN' && f.status !== 'error');

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
                        Enhanced for batch processing by a world-class AI engineer.
                    </p>
                </footer>
            </div>
        </div>
    );
};

export default App;
