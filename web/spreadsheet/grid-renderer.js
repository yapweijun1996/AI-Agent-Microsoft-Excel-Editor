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

// Grid dimensions - optimized for unlimited scrolling with responsive width
const GRID_CONFIG = {
  defaultRowHeight: 32,
  getOptimalColWidth: () => {
    // Calculate optimal column width based on 75vw container
    const viewportWidth = window.innerWidth;
    const containerWidth = viewportWidth * 0.75; // 75vw
    const headerWidth = 48; // Row header width
    const availableWidth = containerWidth - headerWidth;
    const optimalCols = 20; // Target number of visible columns
    return Math.max(80, Math.floor(availableWidth / optimalCols)); // Min 80px per column
  },
  defaultColWidth: 120, // Fallback if calculation fails
  minRowsBuffer: 8,    // Increased buffer for smoother scrolling
  minColsBuffer: 6,    // Increased buffer for smoother scrolling
  maxRows: 1048576,    // Excel limit
  maxCols: 16384,      // Excel limit
  renderBatchSize: 100, // Maximum cells to render per batch
  scrollThreshold: 1    // Scroll threshold multiplier (more sensitive)
};

// Virtual scrolling state management
let virtualState = {
  renderedCells: new Set(),
  cellCache: new Map(),
  lastRenderTime: 0,
  isRendering: false
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
  
  renderVisibleGridOptimized(container, ws);
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
      const scrollDeltaY = Math.abs(currentScrollTop - lastScrollTop);
      const scrollDeltaX = Math.abs(currentScrollLeft - lastScrollLeft);
      
      const currentOptimalColWidth = GRID_CONFIG.getOptimalColWidth();
      if (scrollDeltaY > GRID_CONFIG.defaultRowHeight * GRID_CONFIG.scrollThreshold ||
          scrollDeltaX > currentOptimalColWidth * GRID_CONFIG.scrollThreshold) {
        
        // Prevent excessive rendering
        if (!virtualState.isRendering) {
          lastScrollTop = currentScrollTop;
          lastScrollLeft = currentScrollLeft;
          renderVisibleGridOptimized(container, getWorksheet());
        }
      }
    }, 16); // Back to 60fps for smooth scrolling
  }, { passive: true });
}

// Optimized rendering function with batching
function renderVisibleGridOptimized(container, ws) {
  if (virtualState.isRendering) return;
  
  virtualState.isRendering = true;
  virtualState.lastRenderTime = performance.now();
  
  // Use requestAnimationFrame for smooth rendering
  requestAnimationFrame(() => {
    try {
      renderVisibleGrid(container, ws);
    } finally {
      virtualState.isRendering = false;
    }
  });
}

function renderVisibleGrid(container, ws) {
  const parent = container.parentElement;
  if (!parent) return;
  
  const ref = ws['!ref'] || 'A1:Z50';
  const range = XLSX.utils.decode_range(ref);
  
  // Full range support with unlimited cells - expand for scrolling
  const fullRange = {
    s: { r: 0, c: 0 },
    e: { 
      r: Math.max(range.e.r, 200),  // Ensure at least 200 rows for scrolling
      c: Math.max(range.e.c, 50)    // Ensure at least 50 columns for scrolling
    }
  };
  
  const totalCells = (fullRange.e.r + 1) * (fullRange.e.c + 1);
  
  const scrollTop = parent.scrollTop;
  const containerHeight = parent.clientHeight;
  const containerWidth = parent.clientWidth;
  
  // Calculate optimal column width for 75vw container
  const optimalColWidth = GRID_CONFIG.getOptimalColWidth();
  
  // Calculate visible range to fill the entire viewport
  const visibleRows = Math.ceil(containerHeight / GRID_CONFIG.defaultRowHeight) + 10; // Extra rows for full coverage
  const visibleCols = Math.ceil(containerWidth / optimalColWidth) + 5;   // Extra cols for full coverage based on optimal width
  
  const firstRow = Math.max(0, Math.floor(scrollTop / GRID_CONFIG.defaultRowHeight) - GRID_CONFIG.minRowsBuffer);
  const lastRow = Math.min(fullRange.e.r, Math.max(firstRow + visibleRows, firstRow + 50)); // At least 50 rows visible
  
  const firstCol = Math.max(0, Math.floor(parent.scrollLeft / optimalColWidth) - GRID_CONFIG.minColsBuffer);
  const lastCol = Math.min(fullRange.e.c, Math.max(firstCol + visibleCols, firstCol + 20)); // At least 20 columns visible
  
  // Calculate total dimensions for full dataset
  const totalHeight = (fullRange.e.r + 1) * GRID_CONFIG.defaultRowHeight;
  const totalWidth = (fullRange.e.c + 1) * optimalColWidth;
  
  // Log dataset info and rendering details
  console.info(`Rendering grid: ${fullRange.e.r + 1} rows × ${fullRange.e.c + 1} columns (${totalCells.toLocaleString()} total cells)`);
  console.info(`Visible range: rows ${firstRow}-${lastRow} (${lastRow - firstRow + 1} rows), columns ${firstCol}-${lastCol} (${lastCol - firstCol + 1} cols)`);
  console.info(`Container size: ${containerWidth}×${containerHeight}px, Virtual size: ${totalWidth}×${totalHeight}px`);
  console.info(`Optimal column width: ${optimalColWidth}px (for 75vw container)`);
  
  let html = `<div class="virtual-scroll-area" style="height: ${totalHeight}px; width: ${totalWidth}px; position: relative; overflow: visible;">`;
  
  // Calculate table position
  const tableTop = firstRow * GRID_CONFIG.defaultRowHeight;
  const tableLeft = firstCol * optimalColWidth;
  
  html += `<table class="ai-grid border-collapse border border-gray-300 bg-white" style="position: absolute; top: ${tableTop}px; left: ${tableLeft}px; transform: translate3d(0, 0, 0);">`;
  
  // Table header for column labels
  html += '<thead class="bg-gray-50"><tr>';
  html += '<th class="w-12 p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500" style="position: sticky; left: 0; z-index: 20;">#</th>';
  
  for (let c = firstCol; c <= lastCol; c++) {
    const colLetter = XLSX.utils.encode_col(c);
    html += `<th class="col-header cursor-pointer select-none p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500" style="min-width: ${optimalColWidth}px; width: ${optimalColWidth}px;" data-col="${colLetter}" data-col-index="${c}">${colLetter}</th>`;
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
      
      // Check cache first for better performance
      let cellData = virtualState.cellCache.get(addr);
      
      if (!cellData) {
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
        
        // Cache the computed cell data
        cellData = {
          value,
          styles: cell?.s || {},
          hasComment: cell?.c?.t,
          originalCell: cell
        };
        
        // Limit cache size to prevent memory issues
        if (virtualState.cellCache.size > 10000) {
          const firstKey = virtualState.cellCache.keys().next().value;
          virtualState.cellCache.delete(firstKey);
        }
        
        virtualState.cellCache.set(addr, cellData);
      }
      
      const cellStyle = buildCellStyle(cellData.styles);
      
      // Track rendered cells
      virtualState.renderedCells.add(addr);
      
      html += `
        <td class="border border-gray-300 hover:bg-blue-50 focus-within:bg-blue-50 relative" 
            data-cell="${addr}" data-col-index="${c}" 
            style="min-width: ${optimalColWidth}px; width: ${optimalColWidth}px; height: ${rowHeight}px;">
          ${cellData.hasComment ? '<div class="absolute top-0 right-0 w-0 h-0 border-solid border-t-4 border-l-4 border-t-red-500 border-l-transparent"></div>' : ''}
          <input type="text" 
                 value="${escapeHtml(cellData.value)}" 
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
  
  // Cleanup old rendered cells from cache
  const currentlyRendered = new Set();
  for (let r = firstRow; r <= lastRow; r++) {
    for (let c = firstCol; c <= lastCol; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      currentlyRendered.add(addr);
    }
  }
  
  // Remove old rendered cells that are no longer visible
  virtualState.renderedCells.forEach(addr => {
    if (!currentlyRendered.has(addr)) {
      virtualState.renderedCells.delete(addr);
    }
  });
  
  // Batch UI enhancements for better performance
  requestAnimationFrame(() => {
    bindGridHeaderEvents();
    applySelectionHighlight();
    
    // Add resize handles after initial render
    setTimeout(() => {
      addResizeHandles();
      enableAutoResize();
    }, 10);
  });
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

// Cache management functions for external use
export function clearCellCache() {
  virtualState.cellCache.clear();
  virtualState.renderedCells.clear();
  console.info('Cell cache cleared');
}

export function invalidateCellCache(cellAddress) {
  if (cellAddress) {
    virtualState.cellCache.delete(cellAddress);
    virtualState.renderedCells.delete(cellAddress);
  } else {
    clearCellCache();
  }
}

export function getCacheStats() {
  return {
    cacheSize: virtualState.cellCache.size,
    renderedCells: virtualState.renderedCells.size,
    isRendering: virtualState.isRendering,
    lastRenderTime: virtualState.lastRenderTime
  };
}