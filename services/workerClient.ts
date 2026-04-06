import type { StatusUpdate, CsvFile, FileType, CsvGenerationOptions } from '../types.ts';

export const runExcelWorker = (
    file: File,
    updateStatus: (update: StatusUpdate) => void,
    options: CsvGenerationOptions = {},
    selectedColumns: Record<string, boolean> | null = null,
    isDetectionOnly: boolean = false,
    signal?: AbortSignal
): Promise<{ csvFiles: CsvFile[]; detectedType: FileType; headers: string[] }> => {
    return new Promise((resolve, reject) => {
        const worker = new Worker(new URL('./excelWorker.ts', import.meta.url), { type: 'module' });
        const id = Math.random().toString(36).substring(7);

        const cleanup = () => {
            worker.terminate();
            if (signal) {
                signal.removeEventListener('abort', onAbort);
            }
        };

        const onAbort = () => {
            cleanup();
            reject(new Error('Processing cancelled by user'));
        };

        if (signal) {
            if (signal.aborted) {
                return onAbort();
            }
            signal.addEventListener('abort', onAbort);
        }

        worker.onmessage = (e) => {
            const { type, update, result, error, id: msgId } = e.data;
            if (msgId !== id) return;

            if (type === 'status') {
                updateStatus(update);
            } else if (type === 'success') {
                cleanup();
                resolve(result);
            } else if (type === 'error') {
                cleanup();
                reject(new Error(error));
            }
        };

        worker.onerror = (e) => {
            cleanup();
            reject(new Error(`Worker error: ${e.message}`));
        };

        worker.postMessage({
            file,
            options,
            selectedColumns,
            isDetectionOnly,
            id
        });
    });
};
