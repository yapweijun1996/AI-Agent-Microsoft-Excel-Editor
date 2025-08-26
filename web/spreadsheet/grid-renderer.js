import { AppState } from '../core/state.js';
import { getWorksheet } from './workbook-manager.js';
import { escapeHtml } from '../utils/index.js';
import { bindGridHeaderEvents } from './grid-interactions.js';
import { addResizeHandles, enableAutoResize, getColumnWidth, getRowHeight, initializeGridSizes } from './resizing.js';
/* global XLSX, getFormulaEngine, onCellFocus, updateCell, handleCellKeydown */

function clearPreviousSelection() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  container.querySelectorAll('.ai-selected').forEach(el => {
    el.classList.remove('ai-selected', 'bg-blue-100', 'ring-1', 'ring-blue-300');
  });
}

export function applySelectionHighlight() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  clearPreviousSelection();
  const ws = getWorksheet();
  if (!ws) return;
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

  // Highlight selected rows
  (AppState.selectedRows || []).forEach(rowNumber => {
    if (rowNumber < range.s.r + 1 || rowNumber > range.e.r + 1) return;
    const rowHeader = container.querySelector(`td.row-index[data-row="${rowNumber}"]`);
    if (rowHeader) {
      rowHeader.classList.add('ai-selected', 'bg-blue-100', 'ring-1', 'ring-blue-300');
      const tr = rowHeader.parentElement;
      if (tr) {
        tr.querySelectorAll('td:not(.row-index)').forEach(td => {
          td.classList.add('ai-selected', 'bg-blue-100');
        });
      }
    }
  });

  // Highlight selected columns
  (AppState.selectedCols || []).forEach(colIndex => {
    if (colIndex < range.s.c || colIndex > range.e.c) return;
    const th = container.querySelector(`th.col-header[data-col-index="${colIndex}"]`);
    if (th) th.classList.add('ai-selected', 'bg-blue-100', 'ring-1', 'ring-blue-300');
    container.querySelectorAll(`td[data-col-index="${colIndex}"]`).forEach(td => {
      td.classList.add('ai-selected', 'bg-blue-100');
    });
  });
}

// Enhanced virtual scrolling implementation
let lastScrollTop = 0;
let lastScrollLeft = 0;
let renderTimeout = null;

// Grid dimensions - optimized for performance
const GRID_CONFIG = {
  defaultRowHeight: 32,
  defaultColWidth: 100,
  minRowsBuffer: 5,  // Reduced buffer for better performance
  minColsBuffer: 3,  // Reduced buffer for better performance
  maxRows: 1048576,
  maxCols: 16384
};

export function renderSpreadsheetTable() {
  const container = document.getElementById('spreadsheet');
  const ws = getWorksheet();
  if (!ws) {
    container.innerHTML = '<div class="p-4 text-center text-gray-500">No sheet selected or workbook is empty.</div>';
    return;
  }
  
  // Initialize grid sizes and scroll handler on first render
  initializeGridSizes();
  if (!container.hasAttribute('data-scroll-initialized')) {
    initializeScrollHandler(container);
    container.setAttribute('data-scroll-initialized', 'true');
  }
  
  renderVisibleGrid(container, ws);
}

function initializeScrollHandler(container) {
  const parent = container.parentElement;
  if (!parent) return;
  
  parent.addEventListener('scroll', () => {
    if (renderTimeout) {
      clearTimeout(renderTimeout);
    }
    
    renderTimeout = setTimeout(() => {
      const currentScrollTop = parent.scrollTop;
      const currentScrollLeft = parent.scrollLeft;
      
      // Only re-render if scroll difference is significant
      if (Math.abs(currentScrollTop - lastScrollTop) > GRID_CONFIG.defaultRowHeight * 3 ||
          Math.abs(currentScrollLeft - lastScrollLeft) > GRID_CONFIG.defaultColWidth * 4) {
        lastScrollTop = currentScrollTop;
        lastScrollLeft = currentScrollLeft;
        renderVisibleGrid(container, getWorksheet());
      }
    }, 50); // Reduced frequency for better performance
  }, { passive: true });
}

function renderVisibleGrid(container, ws) {
  const parent = container.parentElement;
  if (!parent) return;
  
  const ref = ws['!ref'] || 'A1:Z50';
  const range = XLSX.utils.decode_range(ref);
  
  // Adaptive limits based on data size to prevent UI crashes
  const totalCells = (range.e.r + 1) * (range.e.c + 1);
  let maxRows, maxCols;
  
  if (totalCells > 10000) {
    // Large dataset - very conservative limits
    maxRows = 100;
    maxCols = 20;
  } else if (totalCells > 5000) {
    // Medium dataset - moderate limits
    maxRows = 150;
    maxCols = 30;
  } else {
    // Small dataset - generous limits
    maxRows = 500;
    maxCols = 100;
  }
  
  const safeRange = {
    s: { r: 0, c: 0 },
    e: { 
      r: Math.min(range.e.r, maxRows),
      c: Math.min(range.e.c, maxCols)
    }
  };
  
  // Log performance warning for large datasets
  if (totalCells > 10000) {
    console.warn(`Large dataset detected (${totalCells} cells). Limiting display to ${maxRows}x${maxCols} for performance.`);
    
    // Show performance warning to user
    const existingWarning = document.getElementById('performance-warning');
    if (!existingWarning) {
      const warning = document.createElement('div');
      warning.id = 'performance-warning';
      warning.className = 'fixed top-16 right-4 bg-yellow-100 border border-yellow-400 text-yellow-700 px-4 py-3 rounded z-50 max-w-md';
      warning.innerHTML = `
        <div class="flex items-center">
          <svg class="w-5 h-5 mr-2" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd"></path>
          </svg>
          <div>
            <strong>Large Dataset Detected</strong>
            <p class="text-sm mt-1">Showing ${maxRows}×${maxCols} cells of ${totalCells.toLocaleString()} total for optimal performance. Use scroll to navigate.</p>
          </div>
          <button onclick="this.parentElement.parentElement.remove()" class="ml-2 text-yellow-600 hover:text-yellow-800">×</button>
        </div>
      `;
      document.body.appendChild(warning);
      
      // Auto-remove warning after 10 seconds
      setTimeout(() => {
        if (warning && warning.parentElement) {
          warning.remove();
        }
      }, 10000);
    }
  }
  
  const scrollTop = parent.scrollTop;
  const containerHeight = parent.clientHeight;
  const containerWidth = parent.clientWidth;
  
  // Calculate visible range efficiently
  const visibleRows = Math.ceil(containerHeight / GRID_CONFIG.defaultRowHeight);
  const visibleCols = Math.ceil(containerWidth / GRID_CONFIG.defaultColWidth);
  
  const firstRow = Math.max(0, Math.floor(scrollTop / GRID_CONFIG.defaultRowHeight) - GRID_CONFIG.minRowsBuffer);
  const lastRow = Math.min(safeRange.e.r, firstRow + visibleRows + (GRID_CONFIG.minRowsBuffer * 2));
  
  const firstCol = Math.max(0, Math.floor(parent.scrollLeft / GRID_CONFIG.defaultColWidth) - GRID_CONFIG.minColsBuffer);
  const lastCol = Math.min(safeRange.e.c, firstCol + visibleCols + (GRID_CONFIG.minColsBuffer * 2));
  
  // Calculate total dimensions
  const totalHeight = (safeRange.e.r + 1) * GRID_CONFIG.defaultRowHeight;
  const totalWidth = (safeRange.e.c + 1) * GRID_CONFIG.defaultColWidth;
  
  let html = `<div class="virtual-scroll-area" style="height: ${totalHeight}px; width: ${totalWidth}px; position: relative;">`;
  
  // Calculate table position
  const tableTop = firstRow * GRID_CONFIG.defaultRowHeight;
  const tableLeft = firstCol * GRID_CONFIG.defaultColWidth;
  
  html += `<table class="ai-grid border-collapse border border-gray-300 bg-white" style="position: absolute; top: ${tableTop}px; left: ${tableLeft}px; transform: translate3d(0, 0, 0);">`;
  
  // Table header for column labels
  html += '<thead class="bg-gray-50"><tr>';
  html += '<th class="w-12 p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500" style="position: sticky; left: 0; z-index: 20;">#</th>';
  
  for (let c = firstCol; c <= lastCol; c++) {
    const colLetter = XLSX.utils.encode_col(c);
    const colWidth = getColumnWidth(c);
    html += `<th class="col-header cursor-pointer select-none p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500" style="min-width: ${colWidth}px; width: ${colWidth}px;" data-col="${colLetter}" data-col-index="${c}">${colLetter}</th>`;
  }
  html += '</tr></thead>';
  
  // Table body with visible rows
  html += '<tbody>';
  for (let r = firstRow; r <= lastRow; r++) {
    const rowHeight = getRowHeight(r);
    html += `<tr class="hover:bg-gray-50" style="height: ${rowHeight}px;">`;
    html += `<td class="row-index cursor-pointer select-none p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500" style="position: sticky; left: 0; z-index: 10; min-width: 48px; height: ${rowHeight}px;" data-row="${r + 1}">${r + 1}</td>`;
    
    for (let c = firstCol; c <= lastCol; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      
      // Safe value calculation with error handling
      let value = '';
      try {
        if (cell) {
          if (cell.f && typeof getFormulaEngine === 'function') {
            try {
              const result = getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet);
              value = (result && typeof result === 'object' && result.error) ? '#ERROR!' : (result || '');
            } catch (formulaError) {
              value = '#FORMULA!';
              console.warn('Formula execution error for', addr, ':', formulaError);
            }
          } else {
            value = cell.v || '';
          }
        }
      } catch (error) {
        value = '#ERROR!';
        console.warn('Cell processing error for', addr, ':', error);
      }
      
      const styles = cell?.s || {};
      const cellStyle = buildCellStyle(styles);
      const hasComment = cell?.c?.t;
      const colWidth = getColumnWidth(c);
      
      html += `
        <td class="border border-gray-300 hover:bg-blue-50 focus-within:bg-blue-50 relative" 
            data-cell="${addr}" data-col-index="${c}" 
            style="min-width: ${colWidth}px; width: ${colWidth}px; height: ${rowHeight}px;">
          ${hasComment ? '<div class="absolute top-0 right-0 w-0 h-0 border-solid border-t-4 border-l-4 border-t-red-500 border-l-transparent"></div>' : ''}
          <input type="text" 
                 value="${escapeHtml(value)}" 
                 style="${cellStyle}" 
                 class="cell-input w-full h-full px-2 py-1 bg-transparent border-none outline-none focus:bg-white focus:shadow-sm focus:ring-1 focus:ring-blue-400 rounded text-sm" 
                 onfocus="onCellFocus('${addr}', this)" 
                 onblur="updateCell('${addr}', this.value)" 
                 onkeydown="handleCellKeydown(event, '${addr}')" />
        </td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table></div>';
  
  container.innerHTML = html;
  bindGridHeaderEvents();
  applySelectionHighlight();
  
  // Add resize handles after rendering
  setTimeout(() => {
    addResizeHandles();
    enableAutoResize();
  }, 10);
}

function buildCellStyle(styles) {
  const parts = [];
  if (styles.bold) parts.push('font-weight: bold');
  if (styles.italic) parts.push('font-style: italic');
  if (styles.underline) parts.push('text-decoration: underline');
  if (styles.color) parts.push(`color: ${styles.color}`);
  if (styles.fill?.fgColor?.rgb) parts.push(`background-color: #${styles.fill.fgColor.rgb}`);
  return parts.join('; ');
}