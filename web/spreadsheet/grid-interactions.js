import { AppState } from '../core/state.js';
import { getWorksheet, persistSnapshot } from './workbook-manager.js';
import { saveToHistory } from './history-manager.js';
import { renderSpreadsheetTable, applySelectionHighlight, updateSingleCell } from './grid-renderer.js';
import { showToast } from '../ui/toast.js';
import { parseCellValue, expandRefForCell } from '../utils/index.js';
import { registerGlobal, createNamespace } from '../core/global-bindings.js';
/* global XLSX */

// Selection state for range selection
let selectionState = {
  isSelecting: false,
  startCell: null,
  endCell: null,
  selectedRange: null
};

// Update the address bar (cell reference display)
function updateAddressBar(addr) {
  const cellRefElement = document.getElementById('cell-reference');
  if (cellRefElement) {
    cellRefElement.textContent = addr;
  }
}

// Helper function to save current cell value if changed
function saveCurrentCellValue(addr, input, ws) {
  if (input && input.value !== undefined) {
    const currentCell = ws[addr];
    const currentValue = currentCell ? (currentCell.f ? '=' + currentCell.f : (currentCell.v || '')) : '';
    if (String(input.value) !== String(currentValue)) {
      updateCell(addr, input.value);
    }
  }
}

// Move to a specific cell and focus it
function moveToCell(row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  AppState.activeCell = { r: row, c: col };
  
  // Update address bar
  updateAddressBar(addr);
  
  // Try to focus the cell input
  const cellInput = document.querySelector(`input[data-cell="${addr}"]`);
  if (cellInput) {
    cellInput.focus();
  }
  
  // Re-render the spreadsheet to show the new selection
  renderSpreadsheetTable();
}

// Cell & Grid Logic

export function updateCell(addr, value) {
  const ws = getWorksheet();
  const oldValue = ws[addr] ? (ws[addr].f || ws[addr].v) : '';

  if (String(oldValue) !== String(value)) {
    saveToHistory(`Edit cell ${addr}`, { addr, oldValue, newValue: value, sheet: AppState.activeSheet });
  }

  if (value.startsWith('=')) {
    ws[addr] = { t: 'f', f: value.substring(1) };
  } else {
    const parsed = parseCellValue(value);
    if (parsed.t === 'z') {
      delete ws[addr];
    } else {
      ws[addr] = { t: parsed.t, v: parsed.v };
    }
  }
  expandRefForCell(ws, addr);
  persistSnapshot();
  
  renderSpreadsheetTable();
}

// Enhanced keyboard navigation
function handleCellKeydown(event, addr) {
  const cell = XLSX.utils.decode_cell(addr);
  const ws = getWorksheet();
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  
  switch (event.key) {
    case 'Enter':
      event.preventDefault();
      event.target.blur();
      // Move to next row
      navigateToCell(cell.r + 1, cell.c, range);
      break;
    
    case 'Tab':
      event.preventDefault();
      // Move to next column (or previous if shift)
      const nextCol = event.shiftKey ? cell.c - 1 : cell.c + 1;
      navigateToCell(cell.r, nextCol, range);
      break;
    
    case 'ArrowUp':
      if (!event.target.selectionStart && !event.target.selectionEnd) {
        event.preventDefault();
        navigateToCell(cell.r - 1, cell.c, range, event.shiftKey);
      }
      break;
    
    case 'ArrowDown':
      if (!event.target.selectionStart && !event.target.selectionEnd) {
        event.preventDefault();
        navigateToCell(cell.r + 1, cell.c, range, event.shiftKey);
      }
      break;
    
    case 'ArrowLeft':
      if (!event.target.selectionStart && !event.target.selectionEnd) {
        event.preventDefault();
        navigateToCell(cell.r, cell.c - 1, range, event.shiftKey);
      }
      break;
    
    case 'ArrowRight':
      if (!event.target.selectionStart && !event.target.selectionEnd) {
        event.preventDefault();
        navigateToCell(cell.r, cell.c + 1, range, event.shiftKey);
      }
      break;
    
    case 'Escape':
      event.preventDefault();
      event.target.blur();
      clearSelection();
      break;
    
    case 'Delete':
    case 'Backspace':
      if (!event.target.value) {
        deleteSelectedCells();
      }
      break;
      
    default:
      // Handle Ctrl/Cmd shortcuts
      if (event.ctrlKey || event.metaKey) {
        handleKeyboardShortcut(event, addr);
      }
      break;
  }
};

function navigateToCell(row, col, range, extendSelection = false) {
  // Bound check
  row = Math.max(0, Math.min(row, range.e.r + 50)); // Allow expanding beyond current range
  col = Math.max(0, Math.min(col, range.e.c + 25));
  
  const newAddr = XLSX.utils.encode_cell({ r: row, c: col });
  const cellElement = document.querySelector(`input[onfocus*="${newAddr}"]`);
  
  if (cellElement) {
    if (extendSelection) {
      // Extend selection range
      extendSelectionRange(row, col);
    } else {
      // Clear selection and focus new cell
      clearSelection();
      cellElement.focus();
      cellElement.select();
    }
  }
}

function handleKeyboardShortcut(event, addr) {
  switch (event.key.toLowerCase()) {
    case 'c':
      event.preventDefault();
      copyCell(addr);
      showToast('Cell copied', 'success', 1000);
      break;
    case 'v':
      event.preventDefault();
      pasteCell(addr);
      showToast('Cell pasted', 'success', 1000);
      break;
    case 'x':
      event.preventDefault();
      cutCell(addr);
      showToast('Cell cut', 'success', 1000);
      break;
    case 'z':
      event.preventDefault();
      // Undo functionality would be implemented here
      break;
    case 'a':
      event.preventDefault();
      selectAllCells();
      break;
  }
}

// Legacy support
function handleCellKeypress(event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    event.target.blur();
  }
};

function onCellFocus(addr, input) {
  try {
    const cell = XLSX.utils.decode_cell(addr);
    AppState.activeCell = cell;

    const refEl = document.getElementById('cell-reference');
    if (refEl) refEl.textContent = addr;

    const formulaBar = document.getElementById('formula-bar');
    if (formulaBar) {
      const ws = getWorksheet();
      const c = ws[addr];
      if (c) {
        if (c.f) {
          formulaBar.value = c.f;
        } else if (c.v !== undefined) {
          formulaBar.value = String(c.v);
        } else {
          formulaBar.value = '';
        }
      } else {
        formulaBar.value = '';
      }
    }
    
    // Update format button states when cell is focused
    updateFormatButtonStates();
  } catch (e) { 
    console.warn('Error in onCellFocus:', e.message);
  }
};

function onCellBlur(addr, input) {
  updateCell(addr, input.value);
};

// Header and context menu interactions
export function bindGridHeaderEvents() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  // Row header click / context menu (modern grid)
  container.querySelectorAll('.row-header').forEach(header => {
    header.addEventListener('click', () => {
      const row = parseInt(header.dataset.row, 10);
      if (!isFinite(row)) return;
      AppState.selectedRows = [row];
      AppState.selectedCols = [];
      AppState.activeCell = { r: row - 1, c: 0 };
      const refEl = document.getElementById('cell-reference');
      if (refEl) refEl.textContent = `A${row}`;
      // Highlight full row via selection range
      const ws = getWorksheet();
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      selectionState.selectedRange = {
        s: { r: row - 1, c: 0 },
        e: { r: row - 1, c: Math.max(range.e.c, 0) }
      };
      applyRangeSelection();
    });

    header.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const row = parseInt(header.dataset.row, 10);
      AppState.selectedRows = [row];
      AppState.selectedCols = [];
      // Also set range for visual consistency
      const ws = getWorksheet();
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      selectionState.selectedRange = {
        s: { r: row - 1, c: 0 },
        e: { r: row - 1, c: Math.max(range.e.c, 0) }
      };
      applyRangeSelection();
      showContextMenu(e.clientX, e.clientY, [
        { label: 'Insert Row Above', action: () => insertRowAtSpecific(row) },
        { label: 'Delete Row', action: () => deleteRowAtSpecific(row) }
      ]);
    });
  });

  // Column header click / context menu (modern grid)
  container.querySelectorAll('.col-header').forEach(header => {
    header.addEventListener('click', () => {
      const colIndex = parseInt(header.dataset.colIndex, 10);
      const colLetter = header.dataset.col;
      if (!isFinite(colIndex)) return;
      AppState.selectedCols = [colIndex];
      AppState.selectedRows = [];
      AppState.activeCell = { r: 0, c: colIndex };
      const refEl = document.getElementById('cell-reference');
      if (refEl) refEl.textContent = `${colLetter}${(AppState.activeCell.r || 0) + 1}`;
      // Highlight full column via selection range
      const ws = getWorksheet();
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      selectionState.selectedRange = {
        s: { r: 0, c: colIndex },
        e: { r: Math.max(range.e.r, 0), c: colIndex }
      };
      applyRangeSelection();
    });

    header.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const colIndex = parseInt(header.dataset.colIndex, 10);
      AppState.selectedCols = [colIndex];
      AppState.selectedRows = [];
      const ws = getWorksheet();
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      selectionState.selectedRange = {
        s: { r: 0, c: colIndex },
        e: { r: Math.max(range.e.r, 0), c: colIndex }
      };
      applyRangeSelection();
      showContextMenu(e.clientX, e.clientY, [
        { label: 'Insert Column Left', action: () => insertColumnAtSpecificIndex(colIndex) },
        { label: 'Delete Column', action: () => deleteColumnAtSpecificIndex(colIndex) }
      ]);
    });
  });

  // Enhanced cell interactions with drag selection (modern grid)
  container.querySelectorAll('.modern-cell[data-cell]').forEach(cellEl => {
    const cellRef = cellEl.dataset.cell;
    const cellCoords = XLSX.utils.decode_cell(cellRef);

    // Mouse down for selection start
    cellEl.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return; // Only left click

      if (e.ctrlKey || e.metaKey) {
        // Multi-select mode
        toggleCellSelection(cellRef);
      } else if (e.shiftKey && selectionState.selectedRange) {
        // Extend selection
        extendSelectionTo(cellCoords);
      } else {
        // Start new selection
        startSelection(cellCoords);
      }
    });

    // Mouse enter for drag selection
    cellEl.addEventListener('mouseenter', (e) => {
      if (selectionState.isSelecting && e.buttons === 1) {
        extendSelectionTo(cellCoords);
      }
    });

    // Mouse up to end selection
    cellEl.addEventListener('mouseup', (e) => {
      if (e.button === 0) {
        endSelection();
      }
    });

    // Context menu
    cellEl.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      showContextMenu(e.clientX, e.clientY, [
        { label: 'Cut', action: () => cutCell(cellRef) },
        { label: 'Copy', action: () => copyCell(cellRef) },
        { label: 'Paste', action: () => pasteCell(cellRef) },
        { label: 'Clear Contents', action: () => clearCell(cellRef) },
        { label: 'Insert Comment', action: () => insertComment(cellRef) }
      ]);
    });
  });
  
  // Global mouse up to handle selection end
  document.addEventListener('mouseup', () => {
    if (selectionState.isSelecting) {
      endSelection();
    }
  });
}

function showContextMenu(x, y, items) {
  // Remove existing
  const existing = document.getElementById('grid-context-menu');
  if (existing) existing.remove();

  const menu = document.createElement('div');
  menu.id = 'grid-context-menu';
  menu.className = 'fixed z-50 bg-white border border-gray-300 rounded shadow-lg text-sm';
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  menu.style.minWidth = '180px';

  items.forEach(item => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'w-full text-left px-3 py-2 hover:bg-gray-100';
    btn.textContent = item.label;
    btn.addEventListener('click', () => {
      hideContextMenu();
      try { item.action(); } catch (e) { console.error(e); }
    });
    menu.appendChild(btn);
  });

  document.body.appendChild(menu);

  const off = (ev) => {
    if (ev && ev.target && menu.contains(ev.target)) return;
    hideContextMenu();
  };
  setTimeout(() => {
    window.addEventListener('click', off, { once: true });
    window.addEventListener('contextmenu', off, { once: true });
    window.addEventListener('scroll', hideContextMenu, { once: true });
    window.addEventListener('resize', hideContextMenu, { once: true });
  }, 0);
}

function hideContextMenu() {
  const m = document.getElementById('grid-context-menu');
  if (m) m.remove();
}

// Specific operations from context menu
async function insertRowAtSpecific(rowNumber) {
  await applyEdits([{ op: 'insertRow', sheet: AppState.activeSheet, row: rowNumber }]);
  showToast(`Inserted row at ${rowNumber}`, 'success');
}

async function deleteRowAtSpecific(rowNumber) {
  await applyEdits([{ op: 'deleteRow', sheet: AppState.activeSheet, row: rowNumber }]);
  showToast(`Deleted row ${rowNumber}`, 'success');
}

async function insertColumnAtSpecificIndex(colIndex) {
  await applyEdits([{ op: 'insertColumn', sheet: AppState.activeSheet, index: colIndex }]);
  const colLetter = XLSX.utils.encode_col(colIndex);
  showToast(`Inserted column ${colLetter}`, 'success');
}

async function deleteColumnAtSpecificIndex(colIndex) {
  await applyEdits([{ op: 'deleteColumn', sheet: AppState.activeSheet, index: colIndex }]);
  const colLetter = XLSX.utils.encode_col(colIndex);
  showToast(`Deleted column ${colLetter}`, 'success');
}

// Operations from ribbon
export async function insertRowAtSelection() {
  const row = AppState.activeCell ? AppState.activeCell.r + 1 : (AppState.selectedRows?.[0] || 1);
  await insertRowAtSpecific(row);
}

export async function insertColumnAtSelectionLeft() {
  const col = AppState.activeCell ? AppState.activeCell.c : (AppState.selectedCols?.[0] || 0);
  await insertColumnAtSpecificIndex(col);
}

export async function deleteSelectedRow() {
  if (AppState.selectedRows && AppState.selectedRows.length > 0) {
    // For now, only delete the first selected row
    await deleteRowAtSpecific(AppState.selectedRows[0]);
  } else if (AppState.activeCell) {
    await deleteRowAtSpecific(AppState.activeCell.r + 1);
  } else {
    showToast('No row selected to delete.', 'warning');
  }
}

export async function deleteSelectedColumn() {
  if (AppState.selectedCols && AppState.selectedCols.length > 0) {
    // For now, only delete the first selected column
    await deleteColumnAtSpecificIndex(AppState.selectedCols[0]);
  } else if (AppState.activeCell) {
    await deleteColumnAtSpecificIndex(AppState.activeCell.c);
  } else {
    showToast('No column selected to delete.', 'warning');
  }
}

// Enhanced selection functions
function startSelection(cellCoords) {
  selectionState.isSelecting = true;
  selectionState.startCell = cellCoords;
  selectionState.endCell = cellCoords;
  updateSelectionRange();
}

function extendSelectionTo(cellCoords) {
  if (!selectionState.startCell) {
    startSelection(cellCoords);
    return;
  }
  selectionState.endCell = cellCoords;
  updateSelectionRange();
}

function extendSelectionRange(row, col) {
  if (!selectionState.selectedRange) {
    startSelection({ r: row, c: col });
    return;
  }
  extendSelectionTo({ r: row, c: col });
}

function endSelection() {
  selectionState.isSelecting = false;
}

function updateSelectionRange() {
  if (!selectionState.startCell || !selectionState.endCell) return;
  
  const startR = Math.min(selectionState.startCell.r, selectionState.endCell.r);
  const endR = Math.max(selectionState.startCell.r, selectionState.endCell.r);
  const startC = Math.min(selectionState.startCell.c, selectionState.endCell.c);
  const endC = Math.max(selectionState.startCell.c, selectionState.endCell.c);
  
  selectionState.selectedRange = {
    s: { r: startR, c: startC },
    e: { r: endR, c: endC }
  };
  
  // Update AppState for compatibility
  AppState.activeCell = selectionState.startCell;
  
  // Apply visual selection
  applyRangeSelection();
}

function applyRangeSelection() {
  // Clear previous selection
  clearPreviousSelection();
  
  if (!selectionState.selectedRange) return;
  
  const container = document.getElementById('spreadsheet');
  if (!container) return;
  
  // Highlight selected range
  for (let r = selectionState.selectedRange.s.r; r <= selectionState.selectedRange.e.r; r++) {
    for (let c = selectionState.selectedRange.s.c; c <= selectionState.selectedRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cellElement = container.querySelector(`.modern-cell[data-cell="${addr}"]`);
      if (cellElement) {
        cellElement.classList.add('selected-range', 'bg-blue-100', 'ring-1', 'ring-blue-300');
        cellElement.setAttribute('aria-selected', 'true');
      }
    }
  }
}

function clearSelection() {
  selectionState.isSelecting = false;
  selectionState.startCell = null;
  selectionState.endCell = null;
  selectionState.selectedRange = null;
  clearPreviousSelection();
}

function clearPreviousSelection() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  container.querySelectorAll('.selected-range, .ai-selected, .multi-selected').forEach(el => {
    el.classList.remove('selected-range', 'ai-selected', 'multi-selected', 'bg-blue-100', 'bg-yellow-100', 'ring-1', 'ring-blue-300', 'ring-yellow-300');
    el.removeAttribute('aria-selected');
  });
}

function toggleCellSelection(cellRef) {
  const container = document.getElementById('spreadsheet');
  const cellElement = container.querySelector(`.modern-cell[data-cell="${cellRef}"]`);
  if (!cellElement) return;

  const nowSelected = !cellElement.classList.contains('multi-selected');
  cellElement.classList.toggle('multi-selected', nowSelected);
  if (nowSelected) {
    cellElement.classList.add('bg-yellow-100', 'ring-1', 'ring-yellow-300');
    cellElement.setAttribute('aria-selected', 'true');
  } else {
    cellElement.classList.remove('bg-yellow-100', 'ring-1', 'ring-yellow-300');
    cellElement.removeAttribute('aria-selected');
  }
}

function selectAllCells() {
  const ws = getWorksheet();
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  
  selectionState.startCell = { r: range.s.r, c: range.s.c };
  selectionState.endCell = { r: range.e.r, c: range.e.c };
  updateSelectionRange();
  
  showToast('All cells selected', 'info', 1500);
}

function deleteSelectedCells() {
  if (!selectionState.selectedRange) return;
  
  const ws = getWorksheet();
  let deletedCount = 0;
  
  for (let r = selectionState.selectedRange.s.r; r <= selectionState.selectedRange.e.r; r++) {
    for (let c = selectionState.selectedRange.s.c; c <= selectionState.selectedRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (ws[addr]) {
        delete ws[addr];
        deletedCount++;
      }
    }
  }
  
  if (deletedCount > 0) {
    persistSnapshot();
    renderSpreadsheetTable();
    showToast(`Cleared ${deletedCount} cells`, 'success');
  }
}

// Enhanced clipboard functions
function cutCell(cellRef) {
  copyCell(cellRef);
  clearCell(cellRef);
}

function copyCell(cellRef) {
  const ws = getWorksheet();
  if (selectionState.selectedRange) {
    // Copy range
    AppState.clipboard = {
      type: 'range',
      range: selectionState.selectedRange,
      data: {}
    };
    
    for (let r = selectionState.selectedRange.s.r; r <= selectionState.selectedRange.e.r; r++) {
      for (let c = selectionState.selectedRange.s.c; c <= selectionState.selectedRange.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (ws[addr]) {
          AppState.clipboard.data[addr] = { ...ws[addr] };
        }
      }
    }
  } else {
    // Copy single cell
    AppState.clipboard = {
      type: 'cell',
      cellRef,
      data: ws[cellRef] ? { ...ws[cellRef] } : null
    };
  }
}

function pasteCell(cellRef) {
  if (!AppState.clipboard) {
    showToast('Nothing to paste', 'warning');
    return;
  }
  
  const ws = getWorksheet();
  const targetCell = XLSX.utils.decode_cell(cellRef);
  
  if (AppState.clipboard.type === 'range') {
    // Paste range
    const sourceRange = AppState.clipboard.range;
    const offsetR = targetCell.r - sourceRange.s.r;
    const offsetC = targetCell.c - sourceRange.s.c;
    
    Object.entries(AppState.clipboard.data).forEach(([sourceAddr, cellData]) => {
      const sourceCell = XLSX.utils.decode_cell(sourceAddr);
      const newAddr = XLSX.utils.encode_cell({
        r: sourceCell.r + offsetR,
        c: sourceCell.c + offsetC
      });
      ws[newAddr] = { ...cellData };
      expandRefForCell(ws, newAddr);
    });
  } else {
    // Paste single cell
    if (AppState.clipboard.data) {
      ws[cellRef] = { ...AppState.clipboard.data };
      expandRefForCell(ws, cellRef);
    }
  }
  
  persistSnapshot();
  renderSpreadsheetTable();
}

function clearCell(cellRef) {
  const ws = getWorksheet();
  delete ws[cellRef];
  persistSnapshot();
  renderSpreadsheetTable();
}

function insertComment(cellRef) {
  const comment = prompt('Enter comment:');
  if (comment) {
    const ws = getWorksheet();
    if (!ws[cellRef]) ws[cellRef] = {};
    ws[cellRef].c = [{ t: comment, a: 'User', T: new Date().toISOString() }];
    persistSnapshot();
    renderSpreadsheetTable();
    showToast('Comment added', 'success');
  }
}

// Formula insertion helper
export function insertFormula(formula) {
  const formulaBar = document.getElementById('formula-bar');
  if (formulaBar) {
    formulaBar.value = formula;
    formulaBar.focus();
    // Place cursor inside parentheses
    const cursorPos = formula.indexOf('()') > -1 ? formula.indexOf('()') + 1 : formula.length;
    formulaBar.setSelectionRange(cursorPos, cursorPos);
  }
}

// Update format button states based on active cell
export function updateFormatButtonStates() {
  if (!AppState.activeCell) return;
  
  const ws = getWorksheet();
  const cellRef = XLSX.utils.encode_cell(AppState.activeCell);
  const cell = ws[cellRef];
  const styles = cell && cell.s ? cell.s : {};
  
  // Update button states
  const boldBtn = document.getElementById('format-bold');
  const italicBtn = document.getElementById('format-italic');
  const underlineBtn = document.getElementById('format-underline');
  const colorBtn = document.getElementById('format-color');
  
  if (boldBtn) {
    boldBtn.classList.toggle('format-btn-active', !!styles.bold);
  }
  if (italicBtn) {
    italicBtn.classList.toggle('format-btn-active', !!styles.italic);
  }
  if (underlineBtn) {
    underlineBtn.classList.toggle('format-btn-active', !!styles.underline);
  }
  if (colorBtn && styles.color) {
    colorBtn.value = styles.color;
  }
}

// Cell formatting helper  
export function applyFormat(formatType, value) {
  const ws = getWorksheet();
  let formattedCount = 0;
  
  // Apply to selected range or active cell
  if (selectionState.selectedRange) {
    for (let r = selectionState.selectedRange.s.r; r <= selectionState.selectedRange.e.r; r++) {
      for (let c = selectionState.selectedRange.s.c; c <= selectionState.selectedRange.e.c; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        if (applyFormatToCell(ws, cellRef, formatType, value)) {
          formattedCount++;
        }
      }
    }
  } else if (AppState.activeCell) {
    const cellRef = XLSX.utils.encode_cell(AppState.activeCell);
    if (applyFormatToCell(ws, cellRef, formatType, value)) {
      formattedCount++;
    }
  } else {
    showToast('Please select a cell to format', 'warning');
    return;
  }
  
  if (formattedCount > 0) {
    persistSnapshot();
    renderSpreadsheetTable();
    updateFormatButtonStates();
    showToast(`Applied ${formatType} formatting to ${formattedCount} cell${formattedCount > 1 ? 's' : ''}`, 'success', 1000);
  }
}

function applyFormatToCell(ws, cellRef, formatType, value) {
  const cell = ws[cellRef] || {};
  
  // Initialize style object if it doesn't exist
  if (!cell.s) {
    cell.s = {};
  }
  
  switch (formatType) {
    case 'bold':
      cell.s.bold = !cell.s.bold;
      break;
    case 'italic':
      cell.s.italic = !cell.s.italic;
      break;
    case 'underline':
      cell.s.underline = !cell.s.underline;
      break;
    case 'color':
      cell.s.color = value;
      break;
    default:
      return false;
  }
  
  ws[cellRef] = cell;
  return true;
}

// Initialize global bindings for HTML event handlers
function initializeGridGlobals() {
  // Create a clean namespace for grid interactions
  createNamespace('GridInteractions', {
    updateCell,
    updateAddressBar,
    moveToCell,
    saveCurrentCellValue,
    handleCellKeydown: function (event, addr) {
      const cell = XLSX.utils.decode_cell(addr);
      const ws = getWorksheet();
      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      
      switch (event.key) {
        case 'ArrowUp':
          event.preventDefault();
          moveToCell(Math.max(0, cell.r - 1), cell.c);
          break;
        case 'ArrowDown':
          event.preventDefault();
          moveToCell(Math.min(range.e.r, cell.r + 1), cell.c);
          break;
        case 'ArrowLeft':
          event.preventDefault();
          moveToCell(cell.r, Math.max(0, cell.c - 1));
          break;
        case 'ArrowRight':
          event.preventDefault();
          moveToCell(cell.r, Math.min(range.e.c, cell.c + 1));
          break;
        case 'Enter':
          event.preventDefault();
          // Save current cell value before moving
          saveCurrentCellValue(addr, event.target, ws);
          moveToCell(Math.min(range.e.r, cell.r + 1), cell.c);
          break;
        case 'Tab':
          event.preventDefault();
          // Save current cell value before moving
          saveCurrentCellValue(addr, event.target, ws);
          if (event.shiftKey) {
            moveToCell(cell.r, Math.max(0, cell.c - 1));
          } else {
            moveToCell(cell.r, Math.min(range.e.c, cell.c + 1));
          }
          break;
        case 'Delete':
          event.preventDefault();
          updateCell(addr, '');
          break;
        case 'F2':
          event.preventDefault();
          event.target.readOnly = false;
          event.target.focus();
          break;
        case 'Escape':
          event.preventDefault();
          event.target.blur();
          break;
      }
    },
    handleCellKeypress: function (event) {
      // Allow typing to immediately start editing
      if (event.target && event.target.readOnly) {
        event.target.readOnly = false;
      }
    },
    onCellFocus: function (addr, input) {
      const cellRef = XLSX.utils.decode_cell(addr);
      AppState.activeCell = cellRef;
      if (input) {
        input.readOnly = false;
        // Show raw value for formulas
        const ws = getWorksheet();
        const cell = ws[addr];
        if (cell && cell.f) {
          input.value = '=' + cell.f;
        }
        
        // Auto-select content for easy editing
        setTimeout(() => {
          if (input && input.select) {
            input.select();
          }
        }, 10);
      }
      updateAddressBar(addr);
    },
    onCellBlur: function (addr, input) {
      if (input && input.value !== undefined) {
        // Get the worksheet to check current value
        const ws = getWorksheet();
        const currentCell = ws[addr];
        const currentValue = currentCell ? (currentCell.f ? '=' + currentCell.f : (currentCell.v || '')) : '';
        
        // Only update if value changed
        if (String(input.value) !== String(currentValue)) {
          updateCell(addr, input.value);
        }
        input.readOnly = true;
      }
    }
  });
  
  // For backward compatibility, also register individual functions
  // Mark these as deprecated to encourage using the namespace
  registerGlobal('updateCell', updateCell, { 
    deprecated: true, 
    description: 'Use GridInteractions.updateCell instead' 
  });
  registerGlobal('handleCellKeydown', window.GridInteractions.handleCellKeydown, { 
    deprecated: true, 
    description: 'Use GridInteractions.handleCellKeydown instead' 
  });
  registerGlobal('handleCellKeypress', window.GridInteractions.handleCellKeypress, { 
    deprecated: true, 
    description: 'Use GridInteractions.handleCellKeypress instead' 
  });
  registerGlobal('onCellFocus', window.GridInteractions.onCellFocus, { 
    deprecated: true, 
    description: 'Use GridInteractions.onCellFocus instead' 
  });
  registerGlobal('onCellBlur', window.GridInteractions.onCellBlur, { 
    deprecated: true, 
    description: 'Use GridInteractions.onCellBlur instead' 
  });
}

// Initialize when module loads
initializeGridGlobals();