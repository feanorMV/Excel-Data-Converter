import React, { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import {
  Plus, Trash2, Download, RefreshCw, Table2,
  CheckCircle, AlertCircle, FolderOpen,
  Copy, ClipboardCheck, Undo2, Redo2, Search, Replace,
  ArrowUp, ArrowDown, Filter, CopyPlus, X, FileDown,
} from 'lucide-react';
import { ENTITY_COLUMNS, ENTITY_LABELS, ITEM_MASTER_BASE_COLS, EDITOR_ENTITY_TYPES } from '../lib/entityColumns';
import type { EntityType } from '../lib/entityColumns';
import { takePendingEditorData } from '../lib/editorStore';
import { parseCsv } from '../lib/csv';

const INITIAL_ROWS = 20;
const ENTITY_OPTIONS: EntityType[] = [...EDITOR_ENTITY_TYPES];

function getTodayDateString(): string {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

function makeEmptyRows(type: EntityType, count: number): Record<string, string>[] {
  return Array.from({ length: count }, () =>
    Object.fromEntries(ENTITY_COLUMNS[type].map(c => [c, '']))
  );
}

function csvEscape(v: string): string {
  if (/[,"\n\r]/.test(v)) return `"${v.replace(/"/g, '""')}"`;
  return v;
}

function buildCsv(type: EntityType, rows: Record<string, string>[]): string {
  const cols = ENTITY_COLUMNS[type];
  const filled = rows.filter(r => cols.some(c => (r[c] ?? '').trim() !== ''));
  return [cols.join(','), ...filled.map(r => cols.map(c => csvEscape(r[c] ?? '')).join(','))].join('\n');
}

function replaceInValue(value: string, find: string, replace: string, caseSensitive: boolean): string {
  const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return value.replace(new RegExp(escaped, caseSensitive ? 'g' : 'gi'), replace);
}

function parseFileToRows(
  data: ArrayBuffer,
  entityType: EntityType,
): { rows: Record<string, string>[]; unmatched: string[]; error?: string } {
  const wb = XLSX.read(data, { type: 'array', raw: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' }) as string[][];
  const knownCols = new Set(ENTITY_COLUMNS[entityType]);

  let bestRowIdx = 0;
  let bestMatchCount = 0;
  for (let i = 0; i < Math.min(rawRows.length, 30); i++) {
    const matches = (rawRows[i] ?? []).filter(cell => knownCols.has(String(cell ?? '').trim())).length;
    if (matches > bestMatchCount) { bestMatchCount = matches; bestRowIdx = i; }
  }

  if (bestMatchCount === 0) {
    return {
      rows: makeEmptyRows(entityType, INITIAL_ROWS),
      unmatched: [],
      error: `No matching columns found. This file uses different column names. For raw Leafio templates, use the Converter tab.`,
    };
  }

  const headers = (rawRows[bestRowIdx] ?? []).map(h => String(h ?? '').trim());
  const dataRows = rawRows.slice(bestRowIdx + 1).filter(r => r.some(c => String(c ?? '').trim() !== ''));
  const unmatched = headers.filter(h => h && !knownCols.has(h));

  const mapped = dataRows.map(srcRow => {
    const row = Object.fromEntries(ENTITY_COLUMNS[entityType].map(c => [c, '']));
    headers.forEach((h, ci) => {
      if (knownCols.has(h)) row[h] = String(srcRow[ci] ?? '');
    });
    return row;
  });

  while (mapped.length < INITIAL_ROWS) mapped.push(Object.fromEntries(ENTITY_COLUMNS[entityType].map(c => [c, ''])));
  return { rows: mapped, unmatched };
}

type Selection = { r1: number; c1: number; r2: number; c2: number };
type SortState = { col: string; dir: 'asc' | 'desc' } | null;

export const ManualEditor: React.FC = () => {
  const [entityType, setEntityType]       = useState<EntityType>('ITEM_MASTER');
  const [rows, setRows]                   = useState<Record<string, string>[]>(() => makeEmptyRows('ITEM_MASTER', INITIAL_ROWS));
  const [showAdditional, setShowAdditional] = useState(false);
  const [importWarning, setImportWarning] = useState<string>('');
  const [loadedEntries, setLoadedEntries] = useState<Map<EntityType, Record<string, string>[]>>(new Map());
  const [selection, setSelection]         = useState<Selection | null>(null);
  const [copyFeedback, setCopyFeedback]   = useState(false);
  const [sortState, setSortState]         = useState<SortState>(null);
  const [showFilledOnly, setShowFilledOnly] = useState(false);
  const [showValidation, setShowValidation] = useState(false);
  const [findOpen, setFindOpen]           = useState(false);
  const [findMode, setFindMode]           = useState<'find' | 'replace'>('find');
  const [findText, setFindText]           = useState('');
  const [replaceText, setReplaceText]     = useState('');
  const [matchCase, setMatchCase]         = useState(false);
  const [findCursor, setFindCursor]       = useState(0);
  const [scrollTop, setScrollTop]         = useState(0);
  const [containerHeight, setContainerHeight] = useState(600);

  const cellRefs     = useRef<Record<string, HTMLInputElement | null>>({});
  const importRef    = useRef<HTMLInputElement | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const pendingFocus = useRef<string | null>(null);
  const findInputRef = useRef<HTMLInputElement | null>(null);
  const selAnchor    = useRef<{ r: number; c: number } | null>(null);
  const isDragging   = useRef(false);
  const historyRef   = useRef<Record<string, string>[][]>([]);
  const historyIdxRef = useRef<number>(-1);
  const rowsRef      = useRef<Record<string, string>[]>(rows);
  const dirtyRef     = useRef(false);

  const isItemMaster = entityType === 'ITEM_MASTER';

  const visibleCols = useMemo(() =>
    isItemMaster && !showAdditional ? ITEM_MASTER_BASE_COLS : ENTITY_COLUMNS[entityType],
    [entityType, showAdditional, isItemMaster],
  );

  // ── Derived: sorted + filtered display rows ────────────────────────────────
  const displayRows = useMemo(() => {
    let result = rows.map((row, realIdx) => ({ row, realIdx }));
    if (showFilledOnly) {
      const allCols = ENTITY_COLUMNS[entityType];
      result = result.filter(({ row }) => allCols.some(c => (row[c] ?? '').trim()));
    }
    if (sortState) {
      result = [...result].sort((a, b) => {
        const va = a.row[sortState.col] ?? '';
        const vb = b.row[sortState.col] ?? '';
        const cmp = va.localeCompare(vb, undefined, { numeric: true, sensitivity: 'base' });
        return sortState.dir === 'asc' ? cmp : -cmp;
      });
    }
    return result;
  }, [rows, entityType, showFilledOnly, sortState]);

  // ── Find matches ───────────────────────────────────────────────────────────
  const findMatches = useMemo(() => {
    if (!findText.trim()) return [];
    const needle = matchCase ? findText : findText.toLowerCase();
    const results: { dispIdx: number; colIdx: number }[] = [];
    displayRows.forEach(({ row }, dispIdx) => {
      visibleCols.forEach((col, colIdx) => {
        const val = matchCase ? (row[col] ?? '') : (row[col] ?? '').toLowerCase();
        if (val.includes(needle)) results.push({ dispIdx, colIdx });
      });
    });
    return results;
  }, [findText, matchCase, displayRows, visibleCols]);

  const findMatchSet = useMemo(() => {
    const s = new Set<string>();
    findMatches.forEach(({ dispIdx, colIdx }) => s.add(`${dispIdx}-${colIdx}`));
    return s;
  }, [findMatches]);

  const currentMatch = findMatches[findCursor] ?? null;

  // ── Normalised selection ───────────────────────────────────────────────────
  const normSel = useMemo((): Selection | null => {
    if (!selection) return null;
    return {
      r1: Math.min(selection.r1, selection.r2), r2: Math.max(selection.r1, selection.r2),
      c1: Math.min(selection.c1, selection.c2), c2: Math.max(selection.c1, selection.c2),
    };
  }, [selection]);

  const isCellSelected = useCallback((dispIdx: number, colIdx: number) => {
    if (!normSel) return false;
    return dispIdx >= normSel.r1 && dispIdx <= normSel.r2 && colIdx >= normSel.c1 && colIdx <= normSel.c2;
  }, [normSel]);

  const filledRowCount = useMemo(() => {
    const cols = ENTITY_COLUMNS[entityType];
    return rows.filter(r => cols.some(c => (r[c] ?? '').trim() !== '')).length;
  }, [rows, entityType]);

  const ROW_H = 29;
  const OVER  = 15;
  const { virtualRows, topPad, botPad } = useMemo(() => {
    const start = Math.max(0, Math.floor(scrollTop / ROW_H) - OVER);
    const end   = Math.min(displayRows.length, Math.ceil((scrollTop + containerHeight) / ROW_H) + OVER);
    return {
      virtualRows: displayRows.slice(start, end).map(({ row, realIdx }, i) => ({ row, realIdx, dispIdx: start + i })),
      topPad: start * ROW_H,
      botPad: Math.max(0, displayRows.length - end) * ROW_H,
    };
  }, [scrollTop, containerHeight, displayRows]);

  // ── History ────────────────────────────────────────────────────────────────
  const recordHistory = useCallback((snapshot: Record<string, string>[]) => {
    historyRef.current = historyRef.current.slice(0, historyIdxRef.current + 1);
    historyRef.current.push(snapshot.map(r => ({ ...r })));
    historyIdxRef.current = historyRef.current.length - 1;
    dirtyRef.current = false;
  }, []);

  const resetHistory = useCallback((initialRows: Record<string, string>[]) => {
    historyRef.current = [initialRows.map(r => ({ ...r }))];
    historyIdxRef.current = 0;
    dirtyRef.current = false;
  }, []);

  const undo = useCallback(() => {
    if (historyIdxRef.current <= 0) return;
    historyIdxRef.current--;
    const prev = historyRef.current[historyIdxRef.current].map(r => ({ ...r }));
    rowsRef.current = prev;
    dirtyRef.current = false;
    setRows(prev);
  }, []);

  const redo = useCallback(() => {
    if (historyIdxRef.current >= historyRef.current.length - 1) return;
    historyIdxRef.current++;
    const next = historyRef.current[historyIdxRef.current].map(r => ({ ...r }));
    rowsRef.current = next;
    dirtyRef.current = false;
    setRows(next);
  }, []);

  const handleCellBlur = useCallback(() => {
    if (!dirtyRef.current) return;
    recordHistory(rowsRef.current);
  }, [recordHistory]);

  // ── Copy feedback ──────────────────────────────────────────────────────────
  const showCopied = useCallback(() => {
    setCopyFeedback(true);
    setTimeout(() => setCopyFeedback(false), 1500);
  }, []);

  // ── Mouse drag for selection ───────────────────────────────────────────────
  useEffect(() => {
    const onUp = () => { isDragging.current = false; };
    window.addEventListener('mouseup', onUp);
    return () => window.removeEventListener('mouseup', onUp);
  }, []);

  // ── Resize observer ────────────────────────────────────────────────────────
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    setContainerHeight(el.clientHeight);
    const ro = new ResizeObserver(() => setContainerHeight(el.clientHeight));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // ── Pending keyboard focus after virtual re-render ─────────────────────────
  useEffect(() => {
    if (!pendingFocus.current) return;
    const key = pendingFocus.current;
    pendingFocus.current = null;
    cellRefs.current[key]?.focus();
  });

  // ── Global keyboard: Ctrl+F / Ctrl+H ──────────────────────────────────────
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.key === 'f') { e.preventDefault(); setFindMode('find'); setFindOpen(true); setTimeout(() => findInputRef.current?.focus(), 50); }
      if (e.key === 'h') { e.preventDefault(); setFindMode('replace'); setFindOpen(true); setTimeout(() => findInputRef.current?.focus(), 50); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  // ── Load CSVs from Converter ───────────────────────────────────────────────
  useEffect(() => {
    const entries = takePendingEditorData();
    if (!entries.length) return;
    const map = new Map<EntityType, Record<string, string>[]>();
    entries.forEach(e => map.set(e.entityType, e.rows));
    setLoadedEntries(map);
    const first = entries[0].entityType;
    setEntityType(first);
    const firstRows = entries[0].rows;
    setRows(firstRows);
    rowsRef.current = firstRows;
    resetHistory(firstRows);
    setShowAdditional(false);
    cellRefs.current = {};
  }, [resetHistory]);

  // ── Handlers ───────────────────────────────────────────────────────────────
  const handleEntityChange = (type: EntityType) => {
    const newRows = loadedEntries.get(type) ?? makeEmptyRows(type, INITIAL_ROWS);
    setEntityType(type);
    setRows(newRows);
    rowsRef.current = newRows;
    resetHistory(newRows);
    setShowAdditional(false);
    setSortState(null);
    setSelection(null);
    setShowFilledOnly(false);
    setShowValidation(false);
    setImportWarning('');
    cellRefs.current = {};
  };

  const handleImportFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportWarning('');
    const reader = new FileReader();
    reader.onload = ev => {
      const data = ev.target?.result as ArrayBuffer;
      const { rows: imported, unmatched, error } = parseFileToRows(data, entityType);
      if (error) { setImportWarning(error); return; }
      setRows(imported);
      rowsRef.current = imported;
      resetHistory(imported);
      cellRefs.current = {};
      setImportWarning(unmatched.length ? `Ignored unknown columns: ${unmatched.join(', ')}` : '');
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const updateCell = useCallback((realIdx: number, col: string, value: string) => {
    setRows(prev => {
      const next = prev.map((r, i) => i === realIdx ? { ...r, [col]: value } : r);
      rowsRef.current = next;
      return next;
    });
    dirtyRef.current = true;
  }, []);

  const handleCellPaste = useCallback((e: React.ClipboardEvent, realIdx: number, startColIdx: number) => {
    e.preventDefault();
    const parsed = parseCsv(e.clipboardData.getData('text/plain'), '\t');
    const allCols = ENTITY_COLUMNS[entityType];

    setRows(prev => {
      const next = prev.map(r => ({ ...r }));
      parsed.forEach((pastedRow, rOff) => {
        const ri = realIdx + rOff;
        while (next.length <= ri) next.push(Object.fromEntries(allCols.map(c => [c, ''])));
        pastedRow.forEach((val, cOff) => {
          const col = visibleCols[startColIdx + cOff];
          if (col) next[ri][col] = val;
        });
      });
      rowsRef.current = next;
      recordHistory(next);
      return next;
    });
  }, [entityType, visibleCols, recordHistory]);

  const handleCellMouseDown = useCallback((e: React.MouseEvent, dispIdx: number, colIdx: number) => {
    if (e.button !== 0) return;
    isDragging.current = true;
    if (e.shiftKey && selAnchor.current) {
      setSelection({ r1: selAnchor.current.r, c1: selAnchor.current.c, r2: dispIdx, c2: colIdx });
    } else {
      selAnchor.current = { r: dispIdx, c: colIdx };
      setSelection({ r1: dispIdx, c1: colIdx, r2: dispIdx, c2: colIdx });
    }
  }, []);

  const handleCellMouseEnter = useCallback((dispIdx: number, colIdx: number) => {
    if (!isDragging.current || !selAnchor.current) return;
    setSelection({ r1: selAnchor.current.r, c1: selAnchor.current.c, r2: dispIdx, c2: colIdx });
  }, []);

  const handleColumnSort = useCallback((col: string) => {
    setSortState(prev => {
      if (!prev || prev.col !== col) return { col, dir: 'asc' };
      if (prev.dir === 'asc') return { col, dir: 'desc' };
      return null;
    });
    setSelection(null);
  }, []);

  const handleFillDown = useCallback(() => {
    if (!normSel) return;
    const { r1, r2, c1, c2 } = normSel;
    if (r1 === r2) return;
    const newRows = rowsRef.current.map(r => ({ ...r }));
    for (let ci = c1; ci <= c2; ci++) {
      const col = visibleCols[ci];
      if (!col) continue;
      const topVal = displayRows[r1]?.row[col] ?? '';
      for (let ri = r1 + 1; ri <= r2; ri++) {
        const realIdx = displayRows[ri]?.realIdx;
        if (realIdx !== undefined) newRows[realIdx][col] = topVal;
      }
    }
    rowsRef.current = newRows;
    setRows(newRows);
    recordHistory(newRows);
  }, [normSel, visibleCols, displayRows, recordHistory]);

  const handleDuplicateRow = useCallback((realIdx: number) => {
    setRows(prev => {
      const next = [...prev];
      next.splice(realIdx + 1, 0, { ...prev[realIdx] });
      rowsRef.current = next;
      recordHistory(next);
      return next;
    });
  }, [recordHistory]);

  const handleInsertRow = useCallback((realIdx: number) => {
    const allCols = ENTITY_COLUMNS[entityType];
    setRows(prev => {
      const next = [...prev];
      next.splice(realIdx, 0, Object.fromEntries(allCols.map(c => [c, ''])));
      rowsRef.current = next;
      recordHistory(next);
      return next;
    });
  }, [entityType, recordHistory]);

  const handleDeleteRow = useCallback((realIdx: number) => {
    setRows(prev => {
      const next = prev.filter((_, i) => i !== realIdx);
      rowsRef.current = next;
      recordHistory(next);
      return next;
    });
  }, [recordHistory]);

  const addRows = (count: number) => {
    const allCols = ENTITY_COLUMNS[entityType];
    setRows(prev => {
      const next = [...prev, ...Array.from({ length: count }, () => Object.fromEntries(allCols.map(c => [c, ''])))];
      rowsRef.current = next;
      return next;
    });
  };

  const handleClear = () => {
    const newRows = makeEmptyRows(entityType, INITIAL_ROWS);
    setRows(newRows);
    rowsRef.current = newRows;
    resetHistory(newRows);
    cellRefs.current = {};
  };

  // ── Copy ───────────────────────────────────────────────────────────────────
  const copySelectionToClipboard = useCallback((): boolean => {
    if (!normSel) return false;
    const { r1, r2, c1, c2 } = normSel;
    if (r1 === r2 && c1 === c2) return false;
    const tsv = Array.from({ length: r2 - r1 + 1 }, (_, ri) => {
      const { row } = displayRows[r1 + ri] ?? { row: {} };
      return Array.from({ length: c2 - c1 + 1 }, (_, ci) =>
        row[visibleCols[c1 + ci]] ?? ''
      ).join('\t');
    }).join('\n');
    navigator.clipboard.writeText(tsv);
    showCopied();
    return true;
  }, [normSel, displayRows, visibleCols, showCopied]);

  const copyAllToClipboard = useCallback(() => {
    const filled = rows.filter(r => visibleCols.some(c => (r[c] ?? '').trim()));
    if (!filled.length) return;
    const tsv = [visibleCols.join('\t'), ...filled.map(r => visibleCols.map(c => r[c] ?? '').join('\t'))].join('\n');
    navigator.clipboard.writeText(tsv);
    showCopied();
  }, [rows, visibleCols, showCopied]);

  const copyColumnToClipboard = useCallback((colIdx: number) => {
    const col = visibleCols[colIdx];
    if (!col) return;
    const values = [col, ...rows.map(r => r[col] ?? '')];
    const lastFilled = values.reduce((acc, v, i) => (v.trim() ? i : acc), 0);
    navigator.clipboard.writeText(values.slice(0, lastFilled + 1).join('\n'));
    setSelection({ r1: 0, c1: colIdx, r2: displayRows.length - 1, c2: colIdx });
    showCopied();
  }, [visibleCols, rows, displayRows.length, showCopied]);

  // ── Find & Replace ─────────────────────────────────────────────────────────
  const scrollToMatch = useCallback((idx: number) => {
    const match = findMatches[idx];
    if (!match || !containerRef.current) return;
    const el = containerRef.current;
    const top = match.dispIdx * ROW_H;
    const bot = top + ROW_H;
    if (top < el.scrollTop) el.scrollTop = top;
    else if (bot > el.scrollTop + el.clientHeight) el.scrollTop = bot - el.clientHeight;
  }, [findMatches]);

  const findNext = useCallback(() => {
    if (!findMatches.length) return;
    const next = (findCursor + 1) % findMatches.length;
    setFindCursor(next);
    scrollToMatch(next);
  }, [findMatches, findCursor, scrollToMatch]);

  const findPrev = useCallback(() => {
    if (!findMatches.length) return;
    const prev = (findCursor - 1 + findMatches.length) % findMatches.length;
    setFindCursor(prev);
    scrollToMatch(prev);
  }, [findMatches, findCursor, scrollToMatch]);

  const handleReplace = useCallback(() => {
    if (!findText || !findMatches.length) return;
    const match = findMatches[findCursor];
    if (!match) return;
    const realIdx = displayRows[match.dispIdx]?.realIdx;
    if (realIdx === undefined) return;
    const col = visibleCols[match.colIdx];
    const newVal = replaceInValue(rows[realIdx][col] ?? '', findText, replaceText, matchCase);
    const newRows = rows.map((r, i) => i === realIdx ? { ...r, [col]: newVal } : { ...r });
    rowsRef.current = newRows;
    setRows(newRows);
    recordHistory(newRows);
    setTimeout(() => findNext(), 0);
  }, [findText, replaceText, matchCase, findMatches, findCursor, displayRows, visibleCols, rows, recordHistory, findNext]);

  const handleReplaceAll = useCallback(() => {
    if (!findText || !findMatches.length) return;
    const newRows = rows.map(r => {
      const updated = { ...r };
      visibleCols.forEach(col => {
        if ((updated[col] ?? '').length) {
          updated[col] = replaceInValue(updated[col], findText, replaceText, matchCase);
        }
      });
      return updated;
    });
    rowsRef.current = newRows;
    setRows(newRows);
    recordHistory(newRows);
  }, [findText, replaceText, matchCase, findMatches.length, rows, visibleCols, recordHistory]);

  const closeFindPanel = useCallback(() => {
    setFindOpen(false);
    setFindText('');
    setReplaceText('');
    setFindCursor(0);
  }, []);

  // ── Key navigation ─────────────────────────────────────────────────────────
  const handleKeyDown = useCallback((e: React.KeyboardEvent, dispIdx: number, realIdx: number, colIdx: number) => {
    const mod = e.ctrlKey || e.metaKey;

    if (mod && e.key === 'z') { e.preventDefault(); undo(); return; }
    if (mod && (e.key === 'y' || (e.shiftKey && e.key === 'z'))) { e.preventDefault(); redo(); return; }
    if (mod && e.key === 'c') { if (copySelectionToClipboard()) e.preventDefault(); return; }
    if (mod && e.key === 'd') { e.preventDefault(); handleFillDown(); return; }
    if (mod && e.key === 'f') { e.preventDefault(); setFindMode('find'); setFindOpen(true); setTimeout(() => findInputRef.current?.focus(), 50); return; }
    if (mod && e.key === 'h') { e.preventDefault(); setFindMode('replace'); setFindOpen(true); setTimeout(() => findInputRef.current?.focus(), 50); return; }

    const lastCol = visibleCols.length - 1;
    const lastRow = displayRows.length - 1;

    const focus = (dispR: number, c: number) => {
      const el = containerRef.current;
      if (el) {
        const top = dispR * ROW_H;
        const bot = top + ROW_H;
        if (top < el.scrollTop) el.scrollTop = top;
        else if (bot > el.scrollTop + el.clientHeight) el.scrollTop = bot - el.clientHeight;
      }
      pendingFocus.current = `${dispR}-${c}`;
    };

    if (e.key === 'Tab') {
      e.preventDefault();
      if (!e.shiftKey) {
        colIdx < lastCol ? focus(dispIdx, colIdx + 1) : dispIdx < lastRow && focus(dispIdx + 1, 0);
      } else {
        colIdx > 0 ? focus(dispIdx, colIdx - 1) : dispIdx > 0 && focus(dispIdx - 1, lastCol);
      }
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (dispIdx < lastRow) {
        focus(dispIdx + 1, colIdx);
      } else if (!showFilledOnly && !sortState) {
        const allCols = ENTITY_COLUMNS[entityType];
        setRows(prev => {
          const next = [...prev, Object.fromEntries(allCols.map(c => [c, '']))];
          rowsRef.current = next;
          return next;
        });
        setTimeout(() => focus(dispIdx + 1, colIdx), 0);
      }
    } else if (e.key === 'ArrowDown' && dispIdx < lastRow) {
      e.preventDefault(); focus(dispIdx + 1, colIdx);
    } else if (e.key === 'ArrowUp' && dispIdx > 0) {
      e.preventDefault(); focus(dispIdx - 1, colIdx);
    } else if (e.key === 'Escape' && findOpen) {
      closeFindPanel();
    }
    void realIdx; // used by updateCell/paste via closure in JSX
  }, [visibleCols.length, displayRows.length, entityType, showFilledOnly, sortState, undo, redo, copySelectionToClipboard, handleFillDown, findOpen, closeFindPanel]);

  // ── Save as file / archive ──────────────────────────────────────────────────
  const handleDownloadCsv = () => {
    const csv = buildCsv(entityType, rows);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    Object.assign(document.createElement('a'), { href: url, download: `${entityType.toLowerCase()}_${getTodayDateString()}.csv` }).click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadZip = async () => {
    const csv = buildCsv(entityType, rows);
    const zip = new JSZip();
    const name = `${entityType.toLowerCase()}_${getTodayDateString()}`;
    zip.file(`${entityType.toLowerCase()}.csv`, csv);
    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    Object.assign(document.createElement('a'), { href: url, download: `${name}.zip` }).click();
    URL.revokeObjectURL(url);
  };

  // ── Cell background helper ─────────────────────────────────────────────────
  const getCellBg = useCallback((dispIdx: number, colIdx: number, row: Record<string, string>, col: string) => {
    const key = `${dispIdx}-${colIdx}`;
    const isCurr = currentMatch?.dispIdx === dispIdx && currentMatch?.colIdx === colIdx;
    const isMatch = !isCurr && findMatchSet.has(key);
    const selected = isCellSelected(dispIdx, colIdx);
    const allCols = ENTITY_COLUMNS[entityType];
    const hasAnyData = allCols.some(c => (row[c] ?? '').trim());
    const invalid = showValidation && !(row[col] ?? '').trim() && hasAnyData;
    if (isCurr)   return 'bg-orange-200';
    if (isMatch)  return 'bg-yellow-100/80';
    if (invalid)  return 'bg-red-100/50';
    if (selected) return 'bg-blue-100/60';
    return '';
  }, [currentMatch, findMatchSet, isCellSelected, entityType, showValidation]);

  // ── Render ─────────────────────────────────────────────────────────────────
  return (
    <div className="space-y-4">
      {/* Toolbar */}
      <div className="bg-white rounded-xl border border-gray-200 p-3 flex flex-wrap items-center gap-2">
        <div className="flex items-center gap-2">
          <Table2 className="w-5 h-5 text-primary shrink-0" />
          <span className="font-semibold text-gray-800 text-sm">Manual Editor</span>
        </div>

        <select
          value={entityType}
          onChange={e => handleEntityChange(e.target.value as EntityType)}
          className="text-sm border border-gray-300 rounded-md px-2 py-1.5 bg-white focus:outline-none focus:ring-2 focus:ring-primary/30"
        >
          {ENTITY_OPTIONS.map(t => <option key={t} value={t}>{ENTITY_LABELS[t]}</option>)}
        </select>

        {/* Undo / Redo */}
        <div className="flex items-center gap-1">
          <button onClick={undo} title="Undo (Ctrl+Z)" className="p-1.5 rounded border border-gray-200 text-gray-500 hover:border-gray-300 transition-colors">
            <Undo2 className="w-3.5 h-3.5" />
          </button>
          <button onClick={redo} title="Redo (Ctrl+Y)" className="p-1.5 rounded border border-gray-200 text-gray-500 hover:border-gray-300 transition-colors">
            <Redo2 className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="w-px h-5 bg-gray-200" />

        {/* Find */}
        <button
          onClick={() => { setFindMode('find'); setFindOpen(v => !v); setTimeout(() => findInputRef.current?.focus(), 50); }}
          title="Find (Ctrl+F)"
          className={`text-xs flex items-center gap-1 px-2 py-1 rounded border transition-colors ${findOpen && findMode === 'find' ? 'bg-primary/10 border-primary/30 text-primary' : 'border-gray-200 text-gray-600 hover:border-gray-300'}`}
        >
          <Search className="w-3 h-3" />Find
        </button>
        <button
          onClick={() => { setFindMode('replace'); setFindOpen(v => !v); setTimeout(() => findInputRef.current?.focus(), 50); }}
          title="Find & Replace (Ctrl+H)"
          className={`text-xs flex items-center gap-1 px-2 py-1 rounded border transition-colors ${findOpen && findMode === 'replace' ? 'bg-primary/10 border-primary/30 text-primary' : 'border-gray-200 text-gray-600 hover:border-gray-300'}`}
        >
          <Replace className="w-3 h-3" />Replace
        </button>

        <div className="w-px h-5 bg-gray-200" />

        {/* Toggles */}
        <button
          onClick={() => setShowFilledOnly(v => !v)}
          title="Show filled rows only"
          className={`text-xs flex items-center gap-1 px-2 py-1 rounded border transition-colors ${showFilledOnly ? 'bg-primary/10 border-primary/30 text-primary' : 'border-gray-200 text-gray-600 hover:border-gray-300'}`}
        >
          <Filter className="w-3 h-3" />{showFilledOnly ? 'All rows' : 'Filled only'}
        </button>
        <button
          onClick={() => setShowValidation(v => !v)}
          title="Highlight empty cells in non-empty rows"
          className={`text-xs flex items-center gap-1 px-2 py-1 rounded border transition-colors ${showValidation ? 'bg-red-50 border-red-300 text-red-600' : 'border-gray-200 text-gray-600 hover:border-gray-300'}`}
        >
          {showValidation ? '✓' : '○'} Validate
        </button>

        <div className="ml-auto flex items-center gap-2 flex-wrap">
          <input ref={importRef} type="file" accept=".xlsx,.csv" className="hidden" onChange={handleImportFile} />
          <button onClick={() => importRef.current?.click()} className="text-xs flex items-center gap-1 px-2 py-1 rounded border border-gray-200 text-gray-600 hover:border-gray-300 transition-colors">
            <FolderOpen className="w-3 h-3" />Import
          </button>

          <span className="text-xs text-gray-400">
            {showFilledOnly ? `${displayRows.length} / ` : ''}{filledRowCount} filled
          </span>

          <button
            onClick={copyAllToClipboard}
            disabled={filledRowCount === 0}
            title="Copy all filled rows as TSV (paste into Excel)"
            className="text-xs flex items-center gap-1 px-2 py-1 rounded border border-gray-200 text-gray-600 hover:border-gray-300 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            {copyFeedback ? <ClipboardCheck className="w-3 h-3 text-green-500" /> : <Copy className="w-3 h-3" />}
            {copyFeedback ? 'Copied!' : 'Copy all'}
          </button>

          {isItemMaster && (
            <button
              onClick={() => setShowAdditional(v => !v)}
              className={`text-xs px-2 py-1 rounded border transition-colors ${showAdditional ? 'bg-primary/10 border-primary/30 text-primary' : 'border-gray-200 text-gray-500 hover:border-gray-300'}`}
            >
              {showAdditional ? 'Hide extra fields' : 'Show extra fields'}
            </button>
          )}
          <button
            onClick={() => addRows(10)}
            className="text-xs flex items-center gap-1 px-2 py-1 rounded border border-gray-200 text-gray-600 hover:border-gray-300 transition-colors"
          >
            <Plus className="w-3 h-3" />+10 rows
          </button>
          <button
            onClick={handleClear}
            className="text-xs flex items-center gap-1 px-2 py-1 rounded border border-gray-200 text-gray-500 hover:border-gray-300 hover:text-red-500 transition-colors"
          >
            <RefreshCw className="w-3 h-3" />Clear
          </button>
        </div>
      </div>

      {/* Find / Replace panel */}
      {findOpen && (
        <div className="bg-white rounded-xl border border-gray-200 p-3 flex flex-wrap items-center gap-2">
          <Search className="w-4 h-4 text-gray-400 shrink-0" />
          <input
            ref={findInputRef}
            value={findText}
            onChange={e => { setFindText(e.target.value); setFindCursor(0); }}
            onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); e.shiftKey ? findPrev() : findNext(); } if (e.key === 'Escape') closeFindPanel(); }}
            placeholder="Find…"
            className="text-sm border border-gray-300 rounded-md px-2 py-1 bg-white focus:outline-none focus:ring-2 focus:ring-primary/30 w-48"
          />
          {findMode === 'replace' && (
            <input
              value={replaceText}
              onChange={e => setReplaceText(e.target.value)}
              placeholder="Replace with…"
              className="text-sm border border-gray-300 rounded-md px-2 py-1 bg-white focus:outline-none focus:ring-2 focus:ring-primary/30 w-48"
            />
          )}
          <label className="flex items-center gap-1 text-xs text-gray-500 cursor-pointer select-none">
            <input type="checkbox" checked={matchCase} onChange={e => setMatchCase(e.target.checked)} className="rounded" />
            Match case
          </label>
          <span className="text-xs text-gray-400">
            {findText ? `${findMatches.length ? findCursor + 1 : 0} / ${findMatches.length}` : ''}
          </span>
          <div className="flex items-center gap-1">
            <button onClick={findPrev} disabled={!findMatches.length} className="p-1 rounded border border-gray-200 disabled:opacity-40 hover:bg-gray-50 transition-colors">
              <ArrowUp className="w-3.5 h-3.5 text-gray-500" />
            </button>
            <button onClick={findNext} disabled={!findMatches.length} className="p-1 rounded border border-gray-200 disabled:opacity-40 hover:bg-gray-50 transition-colors">
              <ArrowDown className="w-3.5 h-3.5 text-gray-500" />
            </button>
          </div>
          {findMode === 'replace' && (<>
            <button onClick={handleReplace} disabled={!findMatches.length} className="text-xs px-2 py-1 rounded border border-gray-300 text-gray-600 hover:bg-gray-50 disabled:opacity-40 transition-colors">Replace</button>
            <button onClick={handleReplaceAll} disabled={!findMatches.length} className="text-xs px-2 py-1 rounded border border-gray-300 text-gray-600 hover:bg-gray-50 disabled:opacity-40 transition-colors">Replace all</button>
          </>)}
          <button onClick={closeFindPanel} className="ml-auto p-1 rounded text-gray-400 hover:text-gray-600 transition-colors">
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* Loaded entries banner */}
      {loadedEntries.size > 0 && (
        <div className="flex items-center gap-2 flex-wrap px-4 py-2.5 bg-blue-50 border border-blue-200 rounded-lg text-blue-700 text-xs">
          <CheckCircle className="w-3.5 h-3.5 shrink-0" />
          <span className="font-medium">Loaded from Converter:</span>
          {[...loadedEntries.entries()].map(([type, r]) => (
            <button
              key={type}
              onClick={() => handleEntityChange(type)}
              className={`px-2 py-0.5 rounded-full border transition-colors ${entityType === type ? 'bg-blue-600 border-blue-600 text-white' : 'bg-white border-blue-300 text-blue-700 hover:bg-blue-100'}`}
            >
              {ENTITY_LABELS[type]} <span className="opacity-70">({r.filter(row => Object.values(row).some(v => String(v).trim())).length})</span>
            </button>
          ))}
        </div>
      )}

      {importWarning && (
        <div className="flex items-center gap-2 px-4 py-2 bg-amber-50 border border-amber-200 rounded-lg text-amber-700 text-xs">
          <AlertCircle className="w-3.5 h-3.5 shrink-0" />{importWarning}
        </div>
      )}

      {/* Grid */}
      <div className="bg-white rounded-xl border border-gray-200 overflow-hidden">
        <div
          ref={containerRef}
          className="overflow-auto max-h-[calc(100vh-340px)]"
          onScroll={e => setScrollTop(e.currentTarget.scrollTop)}
        >
          <table className="text-xs border-collapse w-max min-w-full">
            <thead className="sticky top-0 z-10 bg-gray-100">
              <tr>
                <th className="sticky left-0 z-20 bg-gray-100 border border-gray-200 px-2 py-2 text-gray-400 font-medium w-10 text-center select-none">#</th>
                {visibleCols.map((col, colIdx) => (
                  <th
                    key={col}
                    className="border border-gray-200 px-2 py-1.5 text-gray-600 font-medium whitespace-nowrap min-w-[140px] text-left group/th"
                  >
                    <div className="flex items-center justify-between gap-1">
                      <button
                        onClick={() => handleColumnSort(col)}
                        className="flex items-center gap-1 text-left flex-1 min-w-0 hover:text-primary transition-colors"
                        title={`Sort by "${col}"`}
                      >
                        <span className="truncate">{col}</span>
                        {sortState?.col === col && (
                          sortState.dir === 'asc'
                            ? <ArrowUp className="w-3 h-3 shrink-0 text-primary" />
                            : <ArrowDown className="w-3 h-3 shrink-0 text-primary" />
                        )}
                      </button>
                      <button
                        onClick={e => { e.stopPropagation(); copyColumnToClipboard(colIdx); }}
                        title={`Copy column "${col}"`}
                        className="opacity-0 group-hover/th:opacity-40 hover:!opacity-100 p-0.5 transition-all shrink-0"
                      >
                        <Copy className="w-2.5 h-2.5" />
                      </button>
                    </div>
                  </th>
                ))}
                <th className="border border-gray-200 w-14 bg-gray-100" />
              </tr>
            </thead>
            <tbody>
              {topPad > 0 && <tr style={{ height: topPad }}><td colSpan={visibleCols.length + 2} /></tr>}
              {virtualRows.map(({ row, realIdx, dispIdx }) => (
                <tr key={dispIdx} className="hover:bg-blue-50/30 group">
                  {/* Row number + insert */}
                  <td className="sticky left-0 z-10 bg-white group-hover:bg-blue-50/30 border border-gray-200 w-10 select-none p-0">
                    <div className="relative flex items-center justify-center h-7">
                      <span className="text-gray-400 font-mono text-[10px] group-hover:opacity-0 transition-opacity">
                        {dispIdx + 1}
                      </span>
                      <button
                        onClick={() => handleInsertRow(realIdx)}
                        className="absolute inset-0 flex items-center justify-center opacity-0 group-hover:opacity-100 text-primary hover:bg-blue-100 transition-all"
                        title="Insert row above"
                        tabIndex={-1}
                      >
                        <Plus className="w-3 h-3" />
                      </button>
                    </div>
                  </td>
                  {visibleCols.map((col, colIdx) => {
                    const bg = getCellBg(dispIdx, colIdx, row, col);
                    return (
                      <td
                        key={col}
                        className={`border border-gray-200 p-0 ${bg}`}
                        onMouseDown={e => handleCellMouseDown(e, dispIdx, colIdx)}
                        onMouseEnter={() => handleCellMouseEnter(dispIdx, colIdx)}
                      >
                        <input
                          ref={el => { cellRefs.current[`${dispIdx}-${colIdx}`] = el; }}
                          value={row[col] ?? ''}
                          onChange={e => updateCell(realIdx, col, e.target.value)}
                          onPaste={e => handleCellPaste(e, realIdx, colIdx)}
                          onKeyDown={e => handleKeyDown(e, dispIdx, realIdx, colIdx)}
                          onBlur={handleCellBlur}
                          className={`w-full h-7 px-2 focus:outline-none focus:ring-1 focus:ring-inset focus:ring-primary/40 font-mono ${bg ? 'bg-transparent' : 'bg-transparent focus:bg-blue-50'}`}
                        />
                      </td>
                    );
                  })}
                  {/* Row actions */}
                  <td className="border border-gray-200 w-14 text-center p-0">
                    <div className="flex items-center justify-center gap-0.5 h-7 opacity-0 group-hover:opacity-100 transition-all">
                      <button
                        onClick={() => handleDuplicateRow(realIdx)}
                        className="p-1 text-gray-300 hover:text-blue-400 transition-colors"
                        tabIndex={-1}
                        title="Duplicate row"
                      >
                        <CopyPlus className="w-3 h-3" />
                      </button>
                      <button
                        onClick={() => handleDeleteRow(realIdx)}
                        className="p-1 text-gray-300 hover:text-red-400 transition-colors"
                        tabIndex={-1}
                        title="Delete row"
                      >
                        <Trash2 className="w-3 h-3" />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {botPad > 0 && <tr style={{ height: botPad }}><td colSpan={visibleCols.length + 2} /></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {/* Actions */}
      <div className="bg-white rounded-xl border border-gray-200 p-4 flex items-center gap-3 flex-wrap">
        <button
          onClick={handleDownloadCsv}
          disabled={filledRowCount === 0}
          className="flex items-center gap-2 px-4 py-2 bg-primary text-white text-sm rounded-lg hover:bg-primary/90 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
        >
          <FileDown className="w-4 h-4" />Download CSV
        </button>
        <button
          onClick={handleDownloadZip}
          disabled={filledRowCount === 0}
          className="flex items-center gap-2 px-4 py-2 bg-gray-800 text-white text-sm rounded-lg hover:bg-gray-700 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
        >
          <Download className="w-4 h-4" />Download ZIP
        </button>
      </div>
    </div>
  );
};
