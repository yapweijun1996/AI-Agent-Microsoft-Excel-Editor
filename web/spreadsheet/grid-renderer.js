import { AppState } from '../core/state.js';
import { getWorksheet } from './workbook-manager.js';
import { escapeHtml } from '../utils/index.js';
import { bindGridHeaderEvents } from './grid-interactions.js';
/* global XLSX, getFormulaEngine */

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

export function renderSpreadsheetTable() {
  const container = document.getElementById('spreadsheet');
  const ws = getWorksheet();
  if (!ws) {
    container.innerHTML = '<div class="p-4 text-center text-gray-500">No sheet selected or workbook is empty.</div>';
    return;
  }
  const ref = ws['!ref'] || 'A1:C20';
  const range = XLSX.utils.decode_range(ref);

  // Virtual scrolling parameters
  const rowHeight = 32;
  const colWidth = 100;
  const visibleRows = Math.ceil(container.clientHeight / rowHeight) + 5; // Buffer rows
  const visibleCols = Math.ceil(container.clientWidth / colWidth) + 10; // Buffer columns

  const firstRow = Math.max(0, Math.floor(container.scrollTop / rowHeight) - 2); // Buffer
  const lastRow = Math.min(range.e.r, firstRow + visibleRows);

  const firstCol = Math.max(range.s.c, Math.floor(container.scrollLeft / colWidth) - 5); // Buffer
  const lastCol = Math.min(range.e.c, firstCol + visibleCols);

  let html = '';
  // Create scrollable area
  html += `<div style="height: ${(range.e.r + 1) * rowHeight}px; width: ${(range.e.c + 1) * colWidth}px; position: relative;">`;
  html += `<table class="ai-grid border-collapse border border-gray-300 bg-white" style="position: absolute; transform: translate(${firstCol * colWidth}px, ${firstRow * rowHeight}px);">`;
  html += '<thead class="bg-gray-50"><tr>';
  html += '<th class="w-12 p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500 sticky left-0 z-10">#</th>';

  // Render visible column headers
  for (let c = firstCol; c <= lastCol; c++) {
    const colLetter = XLSX.utils.encode_col(c);
    html += `<th class="col-header cursor-pointer select-none p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500 min-w-[100px]" data-col="${colLetter}" data-col-index="${c}">${colLetter}</th>`;
  }
  html += '</tr></thead><tbody>';

  // Render visible rows and columns
  for (let r = firstRow; r <= lastRow; r++) {
    html += '<tr class="hover:bg-gray-50">';
    html += `<td class="row-index cursor-pointer select-none p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500 sticky left-0 z-10" data-row="${r + 1}">${r + 1}</td>`;

    for (let c = firstCol; c <= lastCol; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const value = cell ? (cell.f ? getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet) : cell.v) : '';
      const styles = cell && cell.s ? cell.s : {};
      const styleStr = `
        font-weight: ${styles.bold ? 'bold' : 'normal'};
        font-style: ${styles.italic ? 'italic' : 'normal'};
        text-decoration: ${styles.underline ? 'underline' : 'none'};
        background-color: ${styles.fill && styles.fill.fgColor ? `#${styles.fill.fgColor.rgb}` : 'transparent'};
      `;
      const hasComment = cell && cell.c && cell.c.t;
      html += `
        <td class="p-1 border border-gray-300 hover:bg-blue-50 focus-within:bg-blue-50 min-h-[32px] relative" data-cell="${addr}" data-col-index="${c}" style="min-width: 100px;">
          ${hasComment ? '<div class="absolute top-0 right-0 w-0 h-0 border-solid border-t-8 border-l-8 border-t-red-500 border-l-transparent"></div>' : ''}
          <input type="text" value="${escapeHtml(value)}" style="${styleStr}" class="cell-input w-full h-full px-2 py-1 bg-transparent border-none outline-none focus:bg-white focus:shadow-sm focus:ring-1 focus:ring-blue-400 rounded" onfocus="onCellFocus('${addr}', this)" onblur="updateCell('${addr}', this.value)" onkeypress="handleCellKeypress(event)" />
        </td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table></div>';
  container.innerHTML = html;
  // Bind header interactions and re-apply selection highlight after render
  bindGridHeaderEvents();
  applySelectionHighlight();
}