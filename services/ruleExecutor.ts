import type { FieldRule, ExcelRow } from '../types';

export function applyFieldRules(
  sourceRow: ExcelRow,
  outputRow: Record<string, any>,
  rules: FieldRule[]
): Record<string, any> {
  const result = { ...outputRow };

  for (const rule of rules) {
    switch (rule.rule) {
      case 'static':
        result[rule.field] = rule.value ?? null;
        break;

      case 'concat':
        result[rule.field] = (rule.sources ?? [])
          .map(s => String(sourceRow[s] ?? '').trim())
          .filter(Boolean)
          .join(rule.separator ?? '_') || null;
        break;

      case 'coalesce':
        result[rule.field] = (rule.sources ?? [])
          .map(s => sourceRow[s])
          .find(v => v !== null && v !== undefined && String(v).trim() !== '')
          ?? null;
        break;

      case 'prefix':
        result[rule.field] = (rule.value ?? '') + String(sourceRow[rule.source ?? ''] ?? '').trim() || null;
        break;

      case 'suffix':
        result[rule.field] = String(sourceRow[rule.source ?? ''] ?? '').trim() + (rule.value ?? '') || null;
        break;
    }
  }

  return result;
}
