import { AppState } from '../core/state.js';
import { getWorksheet, persistSnapshot } from './workbook-manager.js';
import { saveToHistory } from './history-manager.js';
import { renderSpreadsheetTable, applySelectionHighlight } from './grid-renderer.js';
import { showToast } from '../ui/toast.js';
/* global XLSX */

// Cell & Grid Logic
export function expandRefForCell(ws, addr) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const cell = XLSX.utils.decode_cell(addr);
  range.s.r = Math.min(range.s.r, cell.r);
  range.s.c = Math.min(range.s.c, cell.c);
  range.e.r = Math.max(range.e.r, cell.r);
  range.e.c = Math.max(range.e.c, cell.c);
  ws['!ref'] = XLSX.utils.encode_range(range);
}

window.updateCell = function (addr, value) {
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
};

window.handleCellKeypress = function (event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    event.target.blur();
  }
};

window.onCellFocus = function (addr, input) {
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
  } catch (e) { /* no-op */ }
};

// Header and context menu interactions
export function bindGridHeaderEvents() {
  const container = document.getElementById('spreadsheet');
  if (!container) return;

  // Row header click / context menu
  container.querySelectorAll('td.row-index').forEach(td => {
    td.addEventListener('click', () => {
      const row = parseInt(td.dataset.row, 10);
      if (!isFinite(row)) return;
      AppState.selectedRows = [row];
      AppState.selectedCols = [];
      AppState.activeCell = { r: row - 1, c: 0 };
      const refEl = document.getElementById('cell-reference');
      if (refEl) refEl.textContent = `A${row}`;
      applySelectionHighlight();
    });

    td.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const row = parseInt(td.dataset.row, 10);
      AppState.selectedRows = [row];
      AppState.selectedCols = [];
      applySelectionHighlight();
      showContextMenu(e.clientX, e.clientY, [
        { label: 'Insert Row Above', action: () => insertRowAtSpecific(row) },
        { label: 'Delete Row', action: () => deleteRowAtSpecific(row) }
      ]);
    });
  });

  // Column header click / context menu
  container.querySelectorAll('th.col-header').forEach(th => {
    th.addEventListener('click', () => {
      const colIndex = parseInt(th.dataset.colIndex, 10);
      const colLetter = th.dataset.col;
      if (!isFinite(colIndex)) return;
      AppState.selectedCols = [colIndex];
      AppState.selectedRows = [];
      AppState.activeCell = { r: 0, c: colIndex };
      const refEl = document.getElementById('cell-reference');
      if (refEl) refEl.textContent = `${colLetter}${(AppState.activeCell.r || 0) + 1}`;
      applySelectionHighlight();
    });

    th.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const colIndex = parseInt(th.dataset.colIndex, 10);
      AppState.selectedCols = [colIndex];
      AppState.selectedRows = [];
      applySelectionHighlight();
      showContextMenu(e.clientX, e.clientY, [
        { label: 'Insert Column Left', action: () => insertColumnAtSpecificIndex(colIndex) },
        { label: 'Delete Column', action: () => deleteColumnAtSpecificIndex(colIndex) }
      ]);
    });
  });

  // Cell context menu
  container.querySelectorAll('td[data-cell]').forEach(td => {
    td.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const cellRef = td.dataset.cell;
      showContextMenu(e.clientX, e.clientY, [
        { label: 'Cut', action: () => cutCell(cellRef) },
        { label: 'Copy', action: () => copyCell(cellRef) },
        { label: 'Paste', action: () => pasteCell(cellRef) }
      ]);
    });
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

// Cell clipboard functions
function cutCell(cellRef) {
  copyCell(cellRef);
  const ws = getWorksheet();
  delete ws[cellRef];
  renderSpreadsheetTable();
  persistSnapshot();
}

function copyCell(cellRef) {
  const ws = getWorksheet();
  AppState.clipboard = {
    v: ws[cellRef]?.v,
    f: ws[cellRef]?.f,
    t: ws[cellRef]?.t,
    s: ws[cellRef]?.s
  };
}

function pasteCell(cellRef) {
  if (!AppState.clipboard) {
    return;
  }
  const ws = getWorksheet();
  ws[cellRef] = { ...AppState.clipboard };
  renderSpreadsheetTable();
  persistSnapshot();
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

// Cell formatting helper  
export function applyFormat(formatType, value) {
  if (!AppState.activeCell) {
    showToast('Please select a cell to format', 'warning');
    return;
  }

  const ws = getWorksheet();
  const cellRef = XLSX.utils.encode_cell(AppState.activeCell);
  const cell = ws[cellRef] || {};
  
  // Initialize style object if it doesn't exist
  if (!cell.s) {
    cell.s = {};
  }
  
  switch (formatType) {
    case 'bold':
      cell.s.font = cell.s.font || {};
      cell.s.font.bold = !cell.s.font.bold;
      break;
    case 'italic':
      cell.s.font = cell.s.font || {};
      cell.s.font.italic = !cell.s.font.italic;
      break;
    case 'underline':
      cell.s.font = cell.s.font || {};
      cell.s.font.underline = !cell.s.font.underline;
      break;
    case 'color':
      cell.s.font = cell.s.font || {};
      cell.s.font.color = { rgb: value.replace('#', '') };
      break;
  }
  
  ws[cellRef] = cell;
  persistSnapshot();
  renderSpreadsheetTable();
  showToast(`Applied ${formatType} formatting`, 'success', 1000);
}