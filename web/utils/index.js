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

/**
 * Efficient deep copy function that handles spreadsheet objects properly
 * Avoids the performance issues of JSON.parse(JSON.stringify())
 */
export function deepCopy(obj) {
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }

  // Handle Date objects
  if (obj instanceof Date) {
    return new Date(obj.getTime());
  }

  // Handle Arrays
  if (Array.isArray(obj)) {
    return obj.map(item => deepCopy(item));
  }

  // Handle regular objects
  if (typeof obj === 'object' && obj.constructor === Object) {
    const copied = {};
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        copied[key] = deepCopy(obj[key]);
      }
    }
    return copied;
  }

  // Handle special objects (like Map, Set, etc.) - fallback to JSON method
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch (e) {
    console.warn('Deep copy fallback failed, returning shallow copy:', e);
    return { ...obj };
  }
}

/**
 * Fast shallow copy for objects that don't need deep copying
 */
export function shallowCopy(obj) {
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }
  
  if (Array.isArray(obj)) {
    return [...obj];
  }
  
  return { ...obj };
}

/**
 * Convert column letter to number (A=1, B=2, etc.)
 */
export function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Convert column number to letter (1=A, 2=B, etc.)
 */
export function numberToColumn(number) {
  let result = '';
  while (number > 0) {
    number--;
    result = String.fromCharCode(65 + (number % 26)) + result;
    number = Math.floor(number / 26);
  }
  return result;
}

/**
 * Parse cell address (e.g., "A1" -> {row: 1, col: 1, colLetter: "A"})
 */
export function parseCellAddress(address) {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  
  const colLetter = match[1];
  const row = parseInt(match[2]);
  const col = columnToNumber(colLetter);
  
  return { row, col, colLetter };
}

/**
 * Create cell address from row and column (1-indexed)
 */
export function createCellAddress(row, col) {
  return numberToColumn(col) + row;
}

/**
 * Centralized error handling utilities
 */
export class AppError extends Error {
  constructor(message, code = 'GENERIC_ERROR', context = {}) {
    super(message);
    this.name = 'AppError';
    this.code = code;
    this.context = context;
    this.timestamp = new Date().toISOString();
  }
}

export const ERROR_CODES = {
  // File operations
  FILE_NOT_FOUND: 'FILE_NOT_FOUND',
  FILE_READ_ERROR: 'FILE_READ_ERROR',
  FILE_WRITE_ERROR: 'FILE_WRITE_ERROR',
  
  // Formula errors
  FORMULA_PARSE_ERROR: 'FORMULA_PARSE_ERROR',
  CIRCULAR_REFERENCE: 'CIRCULAR_REFERENCE',
  
  // Cell operations
  INVALID_CELL_ADDRESS: 'INVALID_CELL_ADDRESS',
  INVALID_RANGE: 'INVALID_RANGE',
  
  // API errors
  API_KEY_MISSING: 'API_KEY_MISSING',
  API_REQUEST_FAILED: 'API_REQUEST_FAILED',
  
  // Storage errors
  STORAGE_QUOTA_EXCEEDED: 'STORAGE_QUOTA_EXCEEDED',
  STORAGE_ACCESS_DENIED: 'STORAGE_ACCESS_DENIED',
  
  // Validation errors
  INVALID_INPUT: 'INVALID_INPUT',
  
  // Operation errors  
  OPERATION_FAILED: 'OPERATION_FAILED'
};

/**
 * Handle errors consistently across the application
 */
export function handleError(error, context = {}) {
  // Log the error
  console.error('[ERROR]', error.message, {
    code: error.code || 'UNKNOWN',
    context: { ...error.context, ...context },
    stack: error.stack,
    timestamp: error.timestamp || new Date().toISOString()
  });
  
  // Show user-friendly message
  const userMessage = getUserFriendlyMessage(error);
  
  // Import showToast dynamically to avoid circular imports
  import('../ui/toast.js').then(({ showToast }) => {
    showToast(userMessage, 'error');
  }).catch(e => {
    console.error('Failed to show error toast:', e);
  });
  
  return error;
}

/**
 * Convert technical errors to user-friendly messages
 */
function getUserFriendlyMessage(error) {
  if (error instanceof AppError) {
    switch (error.code) {
      case ERROR_CODES.FILE_NOT_FOUND:
        return 'File not found. Please check the file path and try again.';
      case ERROR_CODES.FILE_READ_ERROR:
        return 'Unable to read file. The file may be corrupted or access is denied.';
      case ERROR_CODES.FILE_WRITE_ERROR:
        return 'Unable to save file. Please check permissions and try again.';
      case ERROR_CODES.FORMULA_PARSE_ERROR:
        return 'Invalid formula syntax. Please check your formula and try again.';
      case ERROR_CODES.CIRCULAR_REFERENCE:
        return 'Circular reference detected in formula. Please remove the circular dependency.';
      case ERROR_CODES.INVALID_CELL_ADDRESS:
        return 'Invalid cell address. Please use format like A1, B2, etc.';
      case ERROR_CODES.API_KEY_MISSING:
        return 'API key required. Please configure your API key in settings.';
      case ERROR_CODES.API_REQUEST_FAILED:
        return 'Request failed. Please check your connection and try again.';
      case ERROR_CODES.STORAGE_QUOTA_EXCEEDED:
        return 'Storage limit exceeded. Please free up space or reduce data size.';
      default:
        return error.message;
    }
  }
  
  // Handle standard JavaScript errors
  if (error instanceof TypeError) {
    return 'Invalid operation. Please check your data and try again.';
  }
  
  if (error instanceof ReferenceError) {
    return 'Reference error. A required component may not be loaded.';
  }
  
  // Default fallback
  return error.message || 'An unexpected error occurred. Please try again.';
}

/**
 * Wrap async operations with consistent error handling
 */
export async function withErrorHandling(operation, context = {}) {
  try {
    return await operation();
  } catch (error) {
    throw handleError(error, context);
  }
}

/**
 * Create a standardized validation error
 */
export function createValidationError(message, field = null) {
  return new AppError(message, ERROR_CODES.INVALID_INPUT, { field });
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