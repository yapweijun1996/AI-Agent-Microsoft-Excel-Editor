import { AppState } from '../core/state.js';
import { getWorksheet } from './workbook-manager.js';
import { escapeHtml } from '../utils/index.js';
import { bindGridHeaderEvents } from './grid-interactions.js';
/* global XLSX, getFormulaEngine, onCellFocus, updateCell, handleCellKeydown, onCellBlur */

// Modern Excel-like Grid Renderer with Clean UI

// Clean, Modern Grid Configuration
const GRID_CONFIG = {
  rowHeight: 28,
  colWidth: 120,
  headerHeight: 32,
  rowHeaderWidth: 60,
  visibleRows: 20,
  visibleCols: 10
};

// Enhanced render state for incremental updates
let renderState = {
  firstRow: 0,
  firstCol: 0,
  isScrolling: false,
  isRenderScheduled: false,
  initialized: false,
  cellCache: new Map(), // Cache for cell DOM elements
  lastUpdateTimestamp: 0
};

export function renderSpreadsheetTable(forceFullRender = false) {
  const now = Date.now();
  if (renderState.isRenderScheduled || (now - renderState.lastUpdateTimestamp < 16 && !forceFullRender)) {
    return;
  }

  renderState.isRenderScheduled = true;
  
  requestAnimationFrame(() => {
    const container = document.getElementById('spreadsheet');
    const ws = getWorksheet();
    
    if (!ws) {
      container.innerHTML = '<div class="flex items-center justify-center h-64 text-gray-500">No worksheet available</div>';
      renderState.initialized = false;
      renderState.isRenderScheduled = false;
      return;
    }
    
    // Use incremental rendering if grid is already initialized
    if (renderState.initialized && !forceFullRender) {
      updateExistingGrid(container, ws);
    } else {
      renderModernGrid(container, ws);
      renderState.initialized = true;
    }
    
    renderState.lastUpdateTimestamp = Date.now();
    renderState.isRenderScheduled = false;
  });
}

/**
 * Efficiently update only changed cells in the existing grid
 */
function updateExistingGrid(container, ws) {
  const ref = ws['!ref'] || 'A1:Z100';
  const range = XLSX.utils.decode_range(ref);
  
  const maxRows = Math.max(range.e.r + 1, GRID_CONFIG.visibleRows);
  const maxCols = Math.max(range.e.c + 1, GRID_CONFIG.visibleCols);
  
  // Find all existing cell inputs
  const cellInputs = container.querySelectorAll('.cell-input');
  
  cellInputs.forEach(input => {
    const cellElement = input.closest('.modern-cell');
    if (!cellElement) return;
    
    const cellAddr = cellElement.dataset.cell;
    if (!cellAddr) return;
    
    const cell = ws[cellAddr];
    let newValue = '';
    let hasFormula = false;
    
    // Calculate the new value
    if (cell) {
      if (cell.f) {
        hasFormula = true;
        try {
          if (typeof getFormulaEngine === 'function') {
            const result = getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet);
            newValue = (result && typeof result === 'object' && result.error) ? '#ERROR!' : (result || '');
          } else {
            newValue = cell.f;
          }
        } catch (error) {
          newValue = '#ERROR!';
        }
      } else {
        newValue = cell.v || '';
      }
    }
    
    // Only update if value changed
    const currentValue = input.value;
    if (currentValue !== String(newValue)) {
      input.value = newValue;
      
      // Update formula class
      if (hasFormula && !cellElement.classList.contains('has-formula')) {
        cellElement.classList.add('has-formula');
      } else if (!hasFormula && cellElement.classList.contains('has-formula')) {
        cellElement.classList.remove('has-formula');
      }
    }
  });
  
  // Check if we need to expand the grid (add more rows/columns)
  expandGridIfNeeded(container, ws, maxRows, maxCols);
}

/**
 * Add additional rows/columns if needed without full re-render
 */
function expandGridIfNeeded(container, ws, newMaxRows, newMaxCols) {
  const gridBody = container.querySelector('.grid-body');
  const gridHeader = container.querySelector('.grid-header');
  
  if (!gridBody || !gridHeader) return;
  
  // Check current dimensions
  const currentRows = gridBody.querySelectorAll('.grid-row').length;
  const currentCols = gridHeader.querySelectorAll('.col-header').length;
  
  // Add new column headers if needed
  if (newMaxCols > currentCols) {
    const headerContainer = gridHeader;
    for (let c = currentCols; c < newMaxCols; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      const headerDiv = document.createElement('div');
      headerDiv.className = 'col-header';
      headerDiv.setAttribute('data-col', colLetter);
      headerDiv.setAttribute('data-col-index', c);
      headerDiv.textContent = colLetter;
      headerContainer.appendChild(headerDiv);
    }
  }
  
  // Add new rows if needed
  if (newMaxRows > currentRows) {
    for (let r = currentRows; r < newMaxRows; r++) {
      const rowDiv = document.createElement('div');
      rowDiv.className = 'grid-row';
      
      // Add row header
      const rowHeader = document.createElement('div');
      rowHeader.className = 'row-header';
      rowHeader.setAttribute('data-row', r + 1);
      rowHeader.textContent = r + 1;
      rowDiv.appendChild(rowHeader);
      
      // Add cells for this row
      for (let c = 0; c < newMaxCols; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        
        let value = '';
        let hasFormula = false;
        
        if (cell) {
          if (cell.f) {
            hasFormula = true;
            try {
              if (typeof getFormulaEngine === 'function') {
                const result = getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet);
                value = (result && typeof result === 'object' && result.error) ? '#ERROR!' : (result || '');
              } else {
                value = cell.f;
              }
            } catch (error) {
              value = '#ERROR!';
            }
          } else {
            value = cell.v || '';
          }
        }
        
        const cellDiv = document.createElement('div');
        cellDiv.className = `modern-cell ${hasFormula ? 'has-formula' : ''}`;
        cellDiv.setAttribute('data-cell', addr);
        cellDiv.setAttribute('data-row', r + 1);
        cellDiv.setAttribute('data-row-index', r);
        cellDiv.setAttribute('data-col', c);
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = value;
        input.className = 'cell-input';
        input.addEventListener('focus', (e) => onCellFocus(addr, e.target));
        input.addEventListener('blur', (e) => onCellBlur(addr, e.target));
        input.addEventListener('keydown', (e) => handleCellKeydown(e, addr));
        
        cellDiv.appendChild(input);
        rowDiv.appendChild(cellDiv);
      }
      
      gridBody.appendChild(rowDiv);
    }
  }
  
  // Also add cells to existing rows if we added columns
  if (newMaxCols > currentCols) {
    const existingRows = gridBody.querySelectorAll('.grid-row');
    existingRows.forEach((rowDiv, rowIndex) => {
      const currentCellsInRow = rowDiv.querySelectorAll('.modern-cell').length;
      
      for (let c = currentCellsInRow; c < newMaxCols; c++) {
        const addr = XLSX.utils.encode_cell({ r: rowIndex, c });
        const cell = ws[addr];
        
        let value = '';
        let hasFormula = false;
        
        if (cell) {
          if (cell.f) {
            hasFormula = true;
            try {
              if (typeof getFormulaEngine === 'function') {
                const result = getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet);
                value = (result && typeof result === 'object' && result.error) ? '#ERROR!' : (result || '');
              } else {
                value = cell.f;
              }
            } catch (error) {
              value = '#ERROR!';
            }
          } else {
            value = cell.v || '';
          }
        }
        
        const cellDiv = document.createElement('div');
        cellDiv.className = `modern-cell ${hasFormula ? 'has-formula' : ''}`;
        cellDiv.setAttribute('data-cell', addr);
        cellDiv.setAttribute('data-row', rowIndex + 1);
        cellDiv.setAttribute('data-row-index', rowIndex);
        cellDiv.setAttribute('data-col', c);
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = value;
        input.className = 'cell-input';
        input.addEventListener('focus', (e) => onCellFocus(addr, e.target));
        input.addEventListener('blur', (e) => onCellBlur(addr, e.target));
        input.addEventListener('keydown', (e) => handleCellKeydown(e, addr));
        
        cellDiv.appendChild(input);
        rowDiv.appendChild(cellDiv);
      }
    });
  }
}

/**
 * Update a single cell efficiently
 */
export function updateSingleCell(cellAddr, newValue) {
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  
  const cellElement = container.querySelector(`[data-cell="${cellAddr}"]`);
  if (!cellElement) return;
  
  const input = cellElement.querySelector('.cell-input');
  if (input && input.value !== String(newValue)) {
    input.value = newValue;
  }
}

function renderModernGrid(container, ws) {
  const ref = ws['!ref'] || 'A1:Z100';
  const range = XLSX.utils.decode_range(ref);
  
  const maxRows = Math.max(range.e.r + 1, GRID_CONFIG.visibleRows);
  const maxCols = Math.max(range.e.c + 1, GRID_CONFIG.visibleCols);
  
  // Create clean modern grid structure
  const html = `
    <div class="modern-spreadsheet">
      <div class="grid-header">
        <div class="corner-cell"></div>
        ${generateColumnHeaders(maxCols)}
      </div>
      <div class="grid-body">
        ${generateGridRows(ws, maxRows, maxCols)}
      </div>
    </div>
  `;
  
  container.innerHTML = html;
  
  // Add modern event handlers
  setTimeout(() => {
    bindGridHeaderEvents();
    addModernInteractions();
  }, 10);
}

function generateColumnHeaders(maxCols) {
  let html = '';
  for (let c = 0; c < maxCols; c++) {
    const colLetter = XLSX.utils.encode_col(c);
    html += `
      <div class="col-header" data-col="${colLetter}" data-col-index="${c}">
        ${colLetter}
      </div>`;
  }
  return html;
}

function generateGridRows(ws, maxRows, maxCols) {
  let html = '';
  
  for (let r = 0; r < maxRows; r++) {
    html += `
      <div class="grid-row">
        <div class="row-header" data-row="${r + 1}" data-row-index="${r}">${r + 1}</div>
        ${generateRowCells(ws, r, maxCols)}
      </div>`;
  }
  
  return html;
}

function generateRowCells(ws, row, maxCols) {
  const fragment = document.createDocumentFragment();
  
  for (let c = 0; c < maxCols; c++) {
    const addr = XLSX.utils.encode_cell({ r: row, c });
    const cell = ws[addr];
    
    let value = '';
    let hasFormula = false;
    
    if (cell) {
      if (cell.f) {
        hasFormula = true;
        try {
          if (typeof getFormulaEngine === 'function') {
            const result = getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet, addr);
            value = (result && typeof result === 'object' && result.error) ? '#ERROR!' : (result || '');
          } else {
            value = cell.f;
          }
        } catch (error) {
          value = '#ERROR!';
        }
      } else {
        value = cell.v || '';
      }
    }
    
    const cellDiv = document.createElement('div');
    cellDiv.className = `modern-cell ${hasFormula ? 'has-formula' : ''}`;
    cellDiv.setAttribute('data-cell', addr);
    cellDiv.setAttribute('data-row', row + 1);
    cellDiv.setAttribute('data-row-index', row);
    cellDiv.setAttribute('data-col', c);
    
    const input = document.createElement('input');
    input.type = 'text';
    input.value = value; // No escaping needed for .value
    input.className = 'cell-input';
    input.addEventListener('focus', (e) => onCellFocus(addr, e.target));
    input.addEventListener('blur', (e) => onCellBlur(addr, e.target));
    input.addEventListener('keydown', (e) => handleCellKeydown(e, addr));
    
    cellDiv.appendChild(input);
    fragment.appendChild(cellDiv);
  }
  
  // This is a trick to return HTML string from a fragment
  const dummyDiv = document.createElement('div');
  dummyDiv.appendChild(fragment);
  return dummyDiv.innerHTML;
}

function addModernInteractions() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  // Delegated event listeners for cells
  container.addEventListener('focusin', (e) => {
    if (e.target.classList.contains('cell-input')) {
      const cellElement = e.target.closest('.modern-cell');
      if (cellElement) {
        onCellFocus(cellElement.dataset.cell, e.target);
      }
    }
  });

  container.addEventListener('focusout', (e) => {
    if (e.target.classList.contains('cell-input')) {
      const cellElement = e.target.closest('.modern-cell');
      if (cellElement) {
        onCellBlur(cellElement.dataset.cell, e.target);
      }
    }
  });

  container.addEventListener('keydown', (e) => {
    if (e.target.classList.contains('cell-input')) {
      const cellElement = e.target.closest('.modern-cell');
      if (cellElement) {
        handleCellKeydown(e, cellElement.dataset.cell);
      }
    }
  });

  // Use pointerenter/pointerleave for more reliable hover
  container.addEventListener('pointerenter', (e) => {
    const cell = e.target.closest('.modern-cell');
    if (cell) {
      cell.classList.add('hovered');
    }
  }, true);
  
  container.addEventListener('pointerleave', (e) => {
    const cell = e.target.closest('.modern-cell');
    if (cell) {
      cell.classList.remove('hovered');
    }
  }, true);
}

// Selection highlighting
function clearPreviousSelection() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  container.querySelectorAll('.selected').forEach(el => {
    el.classList.remove('selected');
  });
}

export function applySelectionHighlight() {
  clearPreviousSelection();
  // Add modern selection highlighting here if needed
}