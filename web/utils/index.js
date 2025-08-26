'use strict';

import { DEBUG } from '../core/state.js';

export function log(...args) { if (DEBUG) console.log('[DEBUG]', ...args); }

export function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

export function parseCellValue(v) {
  if (v === null || v === undefined) return { t: 'z', v: '' };
  const num = Number(v);
  if (v !== '' && !isNaN(num)) return { t: 'n', v: num };
  if (typeof v === 'boolean') return { t: 'b', v: v };
  return { t: 's', v: String(v) };
}

export function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&')
    .replace(/</g, '<')
    .replace(/>/g, '>')
    .replace(/"/g, '"')
    .replace(/'/g, '&#039;');
}

export function uuid() { return 'id-' + Math.random().toString(36).slice(2) + Date.now().toString(36); }

export function extractFirstJson(text) {
  if (typeof text !== 'string') return null;
  // Code fences
  const fence = text.match(/```json[\s\S]*?```/);
  if (fence) {
    const inner = fence[0].replace(/```json/, '').replace(/```/, '').trim();
    try { return JSON.parse(inner); } catch { }
  }
  // Brute force first {...}
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start >= 0 && end > start) {
    const slice = text.slice(start, end + 1);
    try { return JSON.parse(slice); } catch { }
  }
  return null;
}

export function getSampleDataFromSheet(ws) {
  if (!ws['!ref']) return 'Empty sheet';

  const range = XLSX.utils.decode_range(ws['!ref']);
  const maxSampleRows = 3;
  const maxSampleCols = 5;

  let sample = [];
  for (let r = range.s.r; r <= Math.min(range.s.r + maxSampleRows - 1, range.e.r); r++) {
    let row = [];
    for (let c = range.s.c; c <= Math.min(range.s.c + maxSampleCols - 1, range.e.c); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      row.push(cell ? String(cell.v) : '');
    }
    sample.push(row.join('\t'));
  }

  const truncated = range.e.r > range.s.r + maxSampleRows - 1 || range.e.c > range.s.c + maxSampleCols - 1;
  return sample.join('\n') + (truncated ? '\n...(truncated)' : '');
}
/* global XLSX */
// Spreadsheet helpers
export function expandRefForCell(ws, addr) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const cell = XLSX.utils.decode_cell(addr);
  range.s.r = Math.min(range.s.r, cell.r);
  range.s.c = Math.min(range.s.c, cell.c);
  range.e.r = Math.max(range.e.r, cell.r);
  range.e.c = Math.max(range.e.c, cell.c);
  ws['!ref'] = XLSX.utils.encode_range(range);
}