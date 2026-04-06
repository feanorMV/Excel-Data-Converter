import { generateCsvsFromExcel } from './excelProcessor.ts';
import type { StatusUpdate } from '../types.ts';

self.onmessage = async (e: MessageEvent) => {
    const { file, options, selectedColumns, isDetectionOnly, id } = e.data;
    
    try {
        const updateStatus = (update: StatusUpdate) => {
            self.postMessage({ type: 'status', update, id });
        };
        
        const result = await generateCsvsFromExcel(file, updateStatus, options, selectedColumns, isDetectionOnly);
        self.postMessage({ type: 'success', result, id });
    } catch (error: any) {
        self.postMessage({ type: 'error', error: error.message, id });
    }
};
