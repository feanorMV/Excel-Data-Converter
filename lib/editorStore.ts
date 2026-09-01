import type { EntityType } from './entityColumns';
import { CSV_PREFIX_TO_ENTITY } from './entityColumns';
import { detectDelimiter, parseCsv } from './csv';

export interface EditorEntry {
  entityType: EntityType;
  rows: Record<string, string>[];
  filename: string;
}

let pending: EditorEntry[] = [];

export function setPendingEditorData(entries: EditorEntry[]): void {
  pending = entries;
}

export function takePendingEditorData(): EditorEntry[] {
  const data = pending;
  pending = [];
  return data;
}

export function hasPendingEditorData(): boolean {
  return pending.length > 0;
}

export function parseCsvContent(content: string): Record<string, string>[] {
  const trimmed = content.trim();
  if (!trimmed) return [];
  const firstLine = trimmed.split('\n', 1)[0].replace(/\r$/, '');
  const delim = detectDelimiter(firstLine);
  const [headers, ...rows] = parseCsv(trimmed, delim);
  if (!headers) return [];
  return rows.map(cells => Object.fromEntries(headers.map((h, i) => [h, cells[i] ?? ''])));
}

export function csvFilesToEditorEntries(
  files: Array<{ name: string; content: string }>,
): EditorEntry[] {
  const entries: EditorEntry[] = [];
  for (const file of files) {
    const prefix = file.name.replace(/_\d{8}\.csv$/i, '').replace(/\.csv$/i, '');
    const entityType = CSV_PREFIX_TO_ENTITY[prefix];
    if (!entityType) continue;
    entries.push({ entityType, rows: parseCsvContent(file.content), filename: file.name });
  }
  return entries;
}
