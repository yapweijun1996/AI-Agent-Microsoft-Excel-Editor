import { AppState } from '../core/state.js';
import { renderSpreadsheetTable } from './grid-renderer.js';
import { persistSnapshot } from './workbook-manager.js';
import { showToast } from '../ui/toast.js';

// Resizing state
let resizeState = {
  isResizing: false,
  resizeType: null, // 'row' or 'column'
  resizeIndex: null,
  startPosition: 0,
  originalSize: 0
};

// Default sizes
const DEFAULT_ROW_HEIGHT = 32;
const DEFAULT_COL_WIDTH = 100;
const MIN_ROW_HEIGHT = 20;
const MIN_COL_WIDTH = 50;
const MAX_ROW_HEIGHT = 200;
const MAX_COL_WIDTH = 500;

// Initialize column/row sizes in AppState if not present
export function initializeGridSizes() {
  if (!AppState.columnWidths) {
    AppState.columnWidths = new Map();
  }
  if (!AppState.rowHeights) {
    AppState.rowHeights = new Map();
  }
}

// Get column width
export function getColumnWidth(colIndex) {
  return AppState.columnWidths?.get(colIndex) || DEFAULT_COL_WIDTH;
}

// Get row height
export function getRowHeight(rowIndex) {
  return AppState.rowHeights?.get(rowIndex) || DEFAULT_ROW_HEIGHT;
}

// Set column width
export function setColumnWidth(colIndex, width) {
  initializeGridSizes();
  const clampedWidth = Math.max(MIN_COL_WIDTH, Math.min(MAX_COL_WIDTH, width));
  AppState.columnWidths.set(colIndex, clampedWidth);
  persistSnapshot();
}

// Set row height
export function setRowHeight(rowIndex, height) {
  initializeGridSizes();
  const clampedHeight = Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, height));
  AppState.rowHeights.set(rowIndex, clampedHeight);
  persistSnapshot();
}

// Add resize handles to grid headers
export function addResizeHandles() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  // Add column resize handles
  container.querySelectorAll('th.col-header').forEach(th => {
    addColumnResizeHandle(th);
  });

  // Add row resize handles
  container.querySelectorAll('td.row-index').forEach(td => {
    addRowResizeHandle(td);
  });
}

function addColumnResizeHandle(headerElement) {
  const colIndex = parseInt(headerElement.dataset.colIndex, 10);
  if (!isFinite(colIndex)) return;

  // Create resize handle
  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'col-resize-handle absolute right-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-blue-300 opacity-0 hover:opacity-100 transition-opacity';
  resizeHandle.style.transform = 'translateX(50%)';
  
  // Make header position relative for absolute positioning of handle
  headerElement.style.position = 'relative';
  headerElement.appendChild(resizeHandle);

  // Add resize event listeners
  resizeHandle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    e.stopPropagation();
    startColumnResize(colIndex, e.clientX);
  });

  // Show handle on header hover
  headerElement.addEventListener('mouseenter', () => {
    resizeHandle.style.opacity = '1';
  });

  headerElement.addEventListener('mouseleave', () => {
    if (!resizeState.isResizing) {
      resizeHandle.style.opacity = '0';
    }
  });
}

function addRowResizeHandle(headerElement) {
  const rowIndex = parseInt(headerElement.dataset.row, 10) - 1; // Convert to 0-based
  if (!isFinite(rowIndex)) return;

  // Create resize handle
  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'row-resize-handle absolute left-0 right-0 bottom-0 h-2 cursor-row-resize hover:bg-blue-300 opacity-0 hover:opacity-100 transition-opacity';
  resizeHandle.style.transform = 'translateY(50%)';
  
  // Make header position relative for absolute positioning of handle
  headerElement.style.position = 'relative';
  headerElement.appendChild(resizeHandle);

  // Add resize event listeners
  resizeHandle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    e.stopPropagation();
    startRowResize(rowIndex, e.clientY);
  });

  // Show handle on header hover
  headerElement.addEventListener('mouseenter', () => {
    resizeHandle.style.opacity = '1';
  });

  headerElement.addEventListener('mouseleave', () => {
    if (!resizeState.isResizing) {
      resizeHandle.style.opacity = '0';
    }
  });
}

function startColumnResize(colIndex, startX) {
  resizeState.isResizing = true;
  resizeState.resizeType = 'column';
  resizeState.resizeIndex = colIndex;
  resizeState.startPosition = startX;
  resizeState.originalSize = getColumnWidth(colIndex);

  // Add global mouse move and up listeners
  document.addEventListener('mousemove', handleColumnResize);
  document.addEventListener('mouseup', endResize);
  
  // Add visual feedback
  document.body.style.cursor = 'col-resize';
  document.body.style.userSelect = 'none';
}

function startRowResize(rowIndex, startY) {
  resizeState.isResizing = true;
  resizeState.resizeType = 'row';
  resizeState.resizeIndex = rowIndex;
  resizeState.startPosition = startY;
  resizeState.originalSize = getRowHeight(rowIndex);

  // Add global mouse move and up listeners
  document.addEventListener('mousemove', handleRowResize);
  document.addEventListener('mouseup', endResize);
  
  // Add visual feedback
  document.body.style.cursor = 'row-resize';
  document.body.style.userSelect = 'none';
}

function handleColumnResize(e) {
  if (!resizeState.isResizing || resizeState.resizeType !== 'column') return;

  const deltaX = e.clientX - resizeState.startPosition;
  const newWidth = resizeState.originalSize + deltaX;
  
  setColumnWidth(resizeState.resizeIndex, newWidth);
  
  // Update the specific column width in the current view
  updateColumnWidthInDOM(resizeState.resizeIndex, getColumnWidth(resizeState.resizeIndex));
}

function handleRowResize(e) {
  if (!resizeState.isResizing || resizeState.resizeType !== 'row') return;

  const deltaY = e.clientY - resizeState.startPosition;
  const newHeight = resizeState.originalSize + deltaY;
  
  setRowHeight(resizeState.resizeIndex, newHeight);
  
  // Update the specific row height in the current view
  updateRowHeightInDOM(resizeState.resizeIndex, getRowHeight(resizeState.resizeIndex));
}

function endResize() {
  if (!resizeState.isResizing) return;

  const resizeType = resizeState.resizeType;
  const resizeIndex = resizeState.resizeIndex;
  const newSize = resizeState.resizeType === 'column' 
    ? getColumnWidth(resizeIndex) 
    : getRowHeight(resizeIndex);

  // Clean up
  document.removeEventListener('mousemove', handleColumnResize);
  document.removeEventListener('mousemove', handleRowResize);
  document.removeEventListener('mouseup', endResize);
  
  document.body.style.cursor = '';
  document.body.style.userSelect = '';

  // Reset resize state
  resizeState.isResizing = false;
  resizeState.resizeType = null;
  resizeState.resizeIndex = null;

  // Show toast notification
  const sizeType = resizeType === 'column' ? 'Column' : 'Row';
  const identifier = resizeType === 'column' 
    ? String.fromCharCode(65 + resizeIndex) // Convert to letter
    : (resizeIndex + 1).toString(); // Convert to 1-based row number
  
  showToast(`${sizeType} ${identifier} resized to ${Math.round(newSize)}px`, 'success', 2000);

  // Optionally re-render the entire grid to ensure consistency
  // setTimeout(() => renderSpreadsheetTable(), 100);
}

function updateColumnWidthInDOM(colIndex, newWidth) {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  // Update column header
  const colHeader = container.querySelector(`th.col-header[data-col-index="${colIndex}"]`);
  if (colHeader) {
    colHeader.style.minWidth = `${newWidth}px`;
    colHeader.style.width = `${newWidth}px`;
  }

  // Update all cells in the column
  container.querySelectorAll(`td[data-col-index="${colIndex}"]`).forEach(cell => {
    cell.style.minWidth = `${newWidth}px`;
    cell.style.width = `${newWidth}px`;
  });
}

function updateRowHeightInDOM(rowIndex, newHeight) {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  // Update row header
  const rowHeader = container.querySelector(`td.row-index[data-row="${rowIndex + 1}"]`);
  if (rowHeader) {
    const row = rowHeader.parentElement;
    if (row) {
      row.style.height = `${newHeight}px`;
      row.querySelectorAll('td').forEach(cell => {
        cell.style.height = `${newHeight}px`;
      });
    }
  }
}

// Auto-resize column to fit content
export function autoResizeColumn(colIndex) {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  let maxWidth = MIN_COL_WIDTH;
  
  // Measure header width
  const colHeader = container.querySelector(`th.col-header[data-col-index="${colIndex}"]`);
  if (colHeader) {
    maxWidth = Math.max(maxWidth, colHeader.scrollWidth + 20);
  }

  // Measure cell content widths
  container.querySelectorAll(`td[data-col-index="${colIndex}"] input`).forEach(input => {
    if (input.value) {
      // Create temporary element to measure text width
      const temp = document.createElement('span');
      temp.style.visibility = 'hidden';
      temp.style.position = 'absolute';
      temp.style.font = window.getComputedStyle(input).font;
      temp.textContent = input.value;
      document.body.appendChild(temp);
      
      const textWidth = temp.offsetWidth + 30; // Add padding
      maxWidth = Math.max(maxWidth, textWidth);
      
      document.body.removeChild(temp);
    }
  });

  // Set the new width
  setColumnWidth(colIndex, Math.min(maxWidth, MAX_COL_WIDTH));
  updateColumnWidthInDOM(colIndex, getColumnWidth(colIndex));
  
  const colLetter = String.fromCharCode(65 + colIndex);
  showToast(`Column ${colLetter} auto-resized`, 'success', 1500);
}

// Double-click to auto-resize
export function enableAutoResize() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  // Column auto-resize on double-click
  container.querySelectorAll('.col-resize-handle').forEach(handle => {
    handle.addEventListener('dblclick', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const header = handle.parentElement;
      const colIndex = parseInt(header.dataset.colIndex, 10);
      if (isFinite(colIndex)) {
        autoResizeColumn(colIndex);
      }
    });
  });
}

// Export for use in grid renderer
export { resizeState };