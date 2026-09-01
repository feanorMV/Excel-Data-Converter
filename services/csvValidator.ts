import type { Solution, ValidationError, FileValidationResult } from '../types.ts';
import rawSchema from '../config/csv-schema.json';
import { parseCsv } from '../lib/csv.ts';

interface ColumnDef {
  name: string;
  type: 'varchar' | 'int' | 'numeric' | 'date' | 'time';
  required?: boolean;
  solutions?: string[];
  allowed_values?: number[];
  min?: number;
  format?: string;
}

interface SchemaEntry {
  delimiter: string;
  encoding: string;
  unique_fields: string[];
  columns: ColumnDef[];
}

const schema = rawSchema as Record<string, SchemaEntry>;

function patternToRegex(pattern: string): RegExp {
  let r = pattern;
  r = r.replace('<order_number>', '[^_]+');
  r = r.replace('<date>', '\\d{8}');
  r = r.replace(/\.\*/g, '\x00');
  r = r.replace(/\./g, '\\.');
  r = r.replace(/\x00/g, '.*');
  return new RegExp(`^${r}$`, 'i');
}

function detectSchema(filename: string): SchemaEntry | null {
  for (const pattern of Object.keys(schema)) {
    try {
      if (patternToRegex(pattern).test(filename.trim())) {
        return schema[pattern];
      }
    } catch {
      // skip malformed pattern
    }
  }
  return null;
}

function extractSolutions(entry: SchemaEntry): Solution[] {
  const sols = new Set<string>();
  for (const col of entry.columns) {
    (col.solutions ?? []).forEach(s => sols.add(s));
  }
  return [...sols].sort() as Solution[];
}

// Python %Y-%m-%d → regex \d{4}-\d{2}-\d{2}
function dateFormatToRegex(fmt: string): RegExp {
  const r = fmt
    .replace('%Y', '\\d{4}')
    .replace('%m', '\\d{2}')
    .replace('%d', '\\d{2}');
  return new RegExp(`^${r}$`);
}

function validateRow(
  row: Record<string, string>,
  rowNum: number,
  columnDefs: ColumnDef[],
  uniqueFields: string[],
  seenUniques: Set<string>,
  selectedSolutions: Solution[],
): ValidationError[] {
  const errors: ValidationError[] = [];

  if (Object.values(row).every(v => (v ?? '').trim() === '')) {
    errors.push({ row: rowNum, field: '', message: 'Empty row' });
    return errors;
  }

  for (const col of columnDefs) {
    if (col.solutions && !col.solutions.some(s => selectedSolutions.includes(s as Solution))) continue;

    const value = (row[col.name] ?? '').trim();

    if (col.required && value === '') {
      errors.push({ row: rowNum, field: col.name, message: 'Required field is empty' });
      continue;
    }

    if (value === '') continue;

    if (value.includes('\n') || value.includes('\r')) {
      errors.push({ row: rowNum, field: col.name, message: 'Contains line break' });
    }

    let parsedStr: string | null = null;
    let typeError = false;

    if (col.type === 'int') {
      if (/^\d+$/.test(value)) {
        parsedStr = String(parseInt(value, 10));
      } else {
        typeError = true;
        errors.push({ row: rowNum, field: col.name, message: `Not an integer: "${value}"` });
      }
    } else if (col.type === 'numeric') {
      const normalized = value.replace(',', '.');
      if (/^-?\d+(\.\d+)?$/.test(normalized)) {
        const num = parseFloat(normalized);
        parsedStr = String(num);
        if (col.min !== undefined && num < col.min) {
          errors.push({ row: rowNum, field: col.name, message: `Value ${value} is below minimum ${col.min}` });
        }
      } else {
        typeError = true;
        errors.push({ row: rowNum, field: col.name, message: `Not a number: "${value}"` });
      }
    } else if (col.type === 'varchar') {
      parsedStr = value;
    } else if (col.type === 'date') {
      const fmt = col.format ?? '%Y-%m-%d';
      if (!dateFormatToRegex(fmt).test(value)) {
        typeError = true;
        errors.push({ row: rowNum, field: col.name, message: `Invalid date format: "${value}" (expected ${fmt})` });
      } else {
        parsedStr = value;
      }
    } else if (col.type === 'time') {
      const fmt = col.format ?? '%H:%M';
      if (fmt === '%H%M') {
        if (!/^(?:[01]\d|2[0-3])[0-5]\d$/.test(value)) {
          typeError = true;
          errors.push({ row: rowNum, field: col.name, message: `Invalid time format: "${value}" (expected HHmm e.g. 0930)` });
        } else {
          parsedStr = value;
        }
      } else {
        parsedStr = value;
      }
    }

    if (!typeError && col.allowed_values && parsedStr !== null) {
      const allowed = col.allowed_values.map(String);
      if (!allowed.includes(parsedStr)) {
        errors.push({ row: rowNum, field: col.name, message: `Invalid value "${value}", allowed: [${allowed.join(', ')}]` });
      }
    }
  }

  if (uniqueFields.length > 0) {
    const key = uniqueFields.map(f => (row[f] ?? '').trim()).join('\x00');
    if (seenUniques.has(key)) {
      errors.push({ row: rowNum, field: uniqueFields.join('+'), message: `Duplicate unique key (${uniqueFields.join(', ')})` });
    } else {
      seenUniques.add(key);
    }
  }

  return errors;
}

export function validateCsvContent(
  content: string,
  filename: string,
  selectedSolutions: Solution[],
): FileValidationResult {
  const entry = detectSchema(filename);
  if (!entry) {
    return { filename, schemaDetected: false, errorCount: 0, errors: [], availableSolutions: [] };
  }

  const availableSolutions = extractSolutions(entry);
  const effectiveSolutions = selectedSolutions.filter(s => availableSolutions.includes(s));
  // if none of selected solutions match, fall back to all available
  const solutions: Solution[] = effectiveSolutions.length > 0 ? effectiveSolutions : availableSolutions;

  const delimiter = entry.delimiter ?? ',';
  const errors: ValidationError[] = [];

  const rows = parseCsv(content, delimiter);
  const headers = (rows[0] ?? []).map(h => h.trim());

  const expectedHeaders = entry.columns.map(c => c.name);
  const unexpectedHeaders = headers.filter(h => h && !expectedHeaders.includes(h));
  const missingRequiredHeaders = entry.columns
    .filter(c => c.required && !headers.includes(c.name) && (c.solutions ?? []).some(s => solutions.includes(s as Solution)))
    .map(c => c.name);

  if (missingRequiredHeaders.length > 0) {
    errors.push({ row: 0, field: '', message: `Missing required headers: ${missingRequiredHeaders.join(', ')}` });
  }
  if (unexpectedHeaders.length > 0) {
    errors.push({ row: 0, field: '', message: `Unexpected headers: ${unexpectedHeaders.join(', ')}` });
  }

  const seenUniques = new Set<string>();

  for (let i = 1; i < rows.length; i++) {
    const values = rows[i];
    const row: Record<string, string> = {};
    headers.forEach((h, idx) => { row[h] = values[idx] ?? ''; });

    const rowErrors = validateRow(row, i, entry.columns, entry.unique_fields ?? [], seenUniques, solutions);
    errors.push(...rowErrors);
  }

  return {
    filename,
    schemaDetected: true,
    errorCount: errors.length,
    errors,
    availableSolutions,
  };
}
