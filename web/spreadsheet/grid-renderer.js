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

// Grid dimensions
const GRID_CONFIG = {
  defaultRowHeight: 32,
  defaultColWidth: 100,
  minRowsBuffer: 3,
  minColsBuffer: 5,
  maxRows: 1048576, // Excel max rows
  maxCols: 16384    // Excel max cols
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
      if (Math.abs(currentScrollTop - lastScrollTop) > GRID_CONFIG.defaultRowHeight * 2 ||
          Math.abs(currentScrollLeft - lastScrollLeft) > GRID_CONFIG.defaultColWidth * 3) {
        lastScrollTop = currentScrollTop;
        lastScrollLeft = currentScrollLeft;
        renderVisibleGrid(container, getWorksheet());
      }
    }, 16); // ~60fps
  }, { passive: true });
}

function renderVisibleGrid(container, ws) {
  const parent = container.parentElement;
  if (!parent) return;
  
  const ref = ws['!ref'] || 'A1:C20';
  const range = XLSX.utils.decode_range(ref);
  
  // Extend range to support larger sheets
  const extendedRange = {
    s: { r: 0, c: 0 },
    e: { r: Math.max(range.e.r, 100), c: Math.max(range.e.c, 25) }
  };
  
  const scrollTop = parent.scrollTop;
  const containerHeight = parent.clientHeight;
  const containerWidth = parent.clientWidth;
  
  // Calculate visible range with dynamic sizing
  let accumulatedHeight = 0;
  let firstRow = 0;
  for (let r = 0; r <= extendedRange.e.r; r++) {
    const rowHeight = getRowHeight(r);
    if (accumulatedHeight + rowHeight > scrollTop) {
      firstRow = Math.max(0, r - GRID_CONFIG.minRowsBuffer);
      break;
    }
    accumulatedHeight += rowHeight;
  }
  
  let visibleHeight = 0;
  let lastRow = firstRow;
  for (let r = firstRow; r <= extendedRange.e.r; r++) {
    const rowHeight = getRowHeight(r);
    visibleHeight += rowHeight;
    if (visibleHeight > containerHeight + (GRID_CONFIG.minRowsBuffer * GRID_CONFIG.defaultRowHeight)) {
      break;
    }
    lastRow = r;
  }
  
  // Similar calculation for columns
  const firstCol = Math.max(0, Math.floor(parent.scrollLeft / GRID_CONFIG.defaultColWidth) - GRID_CONFIG.minColsBuffer);
  const visibleCols = Math.ceil(containerWidth / GRID_CONFIG.defaultColWidth) + (GRID_CONFIG.minColsBuffer * 2);
  const lastCol = Math.min(extendedRange.e.c, firstCol + visibleCols);
  
  // Calculate total dimensions
  let totalHeight = 0;
  for (let r = 0; r <= extendedRange.e.r; r++) {
    totalHeight += getRowHeight(r);
  }
  
  let totalWidth = 0;
  for (let c = 0; c <= extendedRange.e.c; c++) {
    totalWidth += getColumnWidth(c);
  }
  
  let html = `<div class="virtual-scroll-area" style="height: ${totalHeight}px; width: ${totalWidth}px; position: relative;">`;
  
  // Sticky header row
  html += createStickyHeaders(firstCol, lastCol);
  
  // Calculate table position
  let tableTop = 0;
  for (let r = 0; r < firstRow; r++) {
    tableTop += getRowHeight(r);
  }
  
  let tableLeft = 0;
  for (let c = 0; c < firstCol; c++) {
    tableLeft += getColumnWidth(c);
  }
  
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
      const value = cell ? (cell.f ? getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet) : cell.v) : '';
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

function createStickyHeaders(firstCol, lastCol) {
  let headerHtml = `<div class="sticky-headers" style="position: sticky; top: 0; z-index: 30; background: white; border-bottom: 1px solid #d1d5db;">`;
  headerHtml += '<div class="flex">';
  headerHtml += '<div class="w-12 p-2 border-r border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500 flex-shrink-0">#</div>';
  
  for (let c = firstCol; c <= lastCol; c++) {
    const colLetter = XLSX.utils.encode_col(c);
    const colWidth = getColumnWidth(c);
    headerHtml += `<div class="col-header-sticky cursor-pointer select-none p-2 border-r border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500 flex-shrink-0" style="min-width: ${colWidth}px; width: ${colWidth}px;" data-col="${colLetter}" data-col-index="${c}">${colLetter}</div>`;
  }
  
  headerHtml += '</div></div>';
  return headerHtml;
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