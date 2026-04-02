import { AppConfig, CsvGenerationOptions, FileType } from '../types';

const STORAGE_KEY = 'excel_converter_configs';

export function listConfigs(): AppConfig[] {
  try {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  } catch (e) {
    console.error('Failed to read configs from localStorage', e);
    return [];
  }
}

export function saveConfig(config: AppConfig): void {
  try {
    const configs = listConfigs();
    const existingIndex = configs.findIndex(c => c.id === config.id);
    if (existingIndex >= 0) {
      configs[existingIndex] = config;
    } else {
      configs.push(config);
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(configs));
  } catch (e) {
    console.error('Failed to save config to localStorage', e);
    throw new Error('Failed to save profile. Local storage might be unavailable.');
  }
}

export function deleteConfig(id: string): void {
  try {
    const configs = listConfigs();
    const updated = configs.filter(c => c.id !== id);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
  } catch (e) {
    console.error('Failed to delete config from localStorage', e);
  }
}

export function buildConfig(
  name: string,
  state: {
    delimiter: string;
    archiveName: string;
    csvOptions: CsvGenerationOptions;
    selectedColumns: Record<string, boolean>;
    columnMapping: Record<string, string>;
    perTypeRules: Partial<Record<FileType, import('../types').FieldRule[]>>;
  }
): AppConfig {
  const perType: AppConfig['perType'] = {};
  
  for (const [type, rules] of Object.entries(state.perTypeRules)) {
    if (rules && rules.length > 0) {
      perType[type as FileType] = { fieldRules: rules };
    }
  }

  // Filter out non-boolean options from csvOptions
  const globalCsvOptions: Record<string, boolean> = {};
  for (const [key, value] of Object.entries(state.csvOptions)) {
    if (typeof value === 'boolean') {
      globalCsvOptions[key] = value;
    }
  }

  return {
    id: Date.now().toString(),
    name,
    createdAt: new Date().toISOString(),
    delimiter: state.delimiter,
    archiveName: state.archiveName,
    globalCsvOptions,
    globalSelectedColumns: state.selectedColumns,
    globalColumnMapping: state.columnMapping,
    perType
  };
}

export function exportConfigToFile(config: AppConfig): void {
  const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${config.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_profile.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function importConfigFromFile(file: File): Promise<AppConfig> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const config = JSON.parse(content) as AppConfig;
        
        if (!config.name || !config.delimiter) {
          reject(new Error('Invalid profile file: missing required fields (name, delimiter).'));
          return;
        }
        
        // Ensure id is unique if imported
        config.id = Date.now().toString();
        resolve(config);
      } catch (err) {
        reject(new Error('Failed to parse JSON file.'));
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file.'));
    reader.readAsText(file);
  });
}
