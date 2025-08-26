'use strict';

import { AppState } from '../core/state.js';
import { db } from '../db/indexeddb.js';
import { log } from '../utils/index.js';
import { showToast } from '../ui/toast.js';
import { saveToHistory } from './history-manager.js';

export async function ensureWorkbook() {
  if (AppState.wb) return;
  try {
    // Try restore from IndexedDB
    const savedWb = await db.getWorkbook('current');
    if (savedWb && savedWb.data && savedWb.data.SheetNames) {
      try {
        AppState.wb = savedWb.data;
        AppState.activeSheet = AppState.wb.SheetNames[0] || 'Sheet1';
        log('Restored workbook from IndexedDB');
        return;
      } catch (e) {
        console.warn('Failed to restore workbook:', e);
        showToast('Failed to restore saved workbook, creating new one', 'warning');
      }
    }

    // Create new workbook
    const wb = XLSX.utils.book_new();
    const aoa = [["Name", "Age", "Email"], ["Alice", 30, "alice@example.com"], ["Bob", 28, "bob@example.com"]];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, AppState.activeSheet);
    AppState.wb = wb;

    // Initialize history with the new workbook
    AppState.history = [];
    AppState.historyIndex = -1;
    saveToHistory('Create new workbook', { sheets: wb.SheetNames });

    await persistSnapshot();
    log('Created new workbook');
  } catch (e) {
    console.error('Failed to ensure workbook:', e);
    showToast('Failed to initialize workbook', 'error');
  }
}

export async function persistSnapshot() {
  try {
    await db.saveWorkbook({ id: 'current', data: AppState.wb });
    log('Workbook snapshot saved');
  } catch (e) {
    console.warn('Snapshot failed:', e);
    showToast('Failed to save changes', 'warning', 2000);
  }
}

export function getWorksheet() { return AppState.wb.Sheets[AppState.activeSheet]; }

export function expandRefForCell(ws, addr) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const cell = XLSX.utils.decode_cell(addr);
  range.s.r = Math.min(range.s.r, cell.r);
  range.s.c = Math.min(range.s.c, cell.c);
  range.e.r = Math.max(range.e.r, cell.r);
  range.e.c = Math.max(range.e.c, cell.c);
  ws['!ref'] = XLSX.utils.encode_range(range);
}