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
  visibleRows: 30,
  visibleCols: 15
};

// Simple render state
let renderState = {
  firstRow: 0,
  firstCol: 0,
  isScrolling: false
};

export function renderSpreadsheetTable() {
  const container = document.getElementById('spreadsheet');
  const ws = getWorksheet();
  
  if (!ws) {
    container.innerHTML = '<div class="flex items-center justify-center h-64 text-gray-500">No worksheet available</div>';
    return;
  }
  
  renderModernGrid(container, ws);
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
        <div class="row-header" data-row="${r + 1}">${r + 1}</div>
        ${generateRowCells(ws, r, maxCols)}
      </div>`;
  }
  
  return html;
}

function generateRowCells(ws, row, maxCols) {
  let html = '';
  
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
    
    const cellClasses = `modern-cell ${hasFormula ? 'has-formula' : ''}`;
    
    html += `
      <div class="${cellClasses}" data-cell="${addr}" data-row="${row}" data-col="${c}">
        <input type="text" 
               value="${escapeHtml(value)}" 
               class="cell-input"
               onfocus="onCellFocus('${addr}', this)"
               onblur="onCellBlur('${addr}', this)"
               onkeydown="handleCellKeydown(event, '${addr}')" />
      </div>`;
  }
  
  return html;
}

function addModernInteractions() {
  // Add smooth scrolling and interactions
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  
  // Highlight on hover
  container.addEventListener('mouseover', (e) => {
    const cell = e.target.closest('.modern-cell');
    if (cell) {
      cell.classList.add('hovered');
    }
  });
  
  container.addEventListener('mouseout', (e) => {
    const cell = e.target.closest('.modern-cell');
    if (cell) {
      cell.classList.remove('hovered');
    }
  });
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