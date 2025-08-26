import { AppState } from '../core/state.js';
import { getFormulaEngine } from '../FormulaEngine.js';
import { escapeHtml } from '../utils/index.js';

/**
 * Gets the display value for a cell, handling formulas and errors.
 * @param {object} cell - The cell object.
 * @returns {string} - The value to display.
 */
function getDisplayValue(cell) {
  if (!cell) {
    return '';
  }
  if (cell.f) {
    const result = getFormulaEngine(AppState.wb, AppState.activeSheet).execute('=' + cell.f, AppState.wb, AppState.activeSheet);
    if (result && typeof result === 'object' && result.error) {
      return '#ERROR!';
    }
    return result;
  }
  return cell.v || '';
}

export function renderSpreadsheetTable() {
  const sheet = AppState.sheets[AppState.activeSheet];
  if (!sheet) {
    console.error("Sheet not found:", AppState.activeSheet);
    return '';
  }
  const { data, columnHeaders, rowHeaders } = sheet;
  const colCount = columnHeaders.length;
  const rowCount = rowHeaders.length;

  let tableHtml = `
    <table class="spreadsheet">
      <thead>
        <tr>
          <th></th>
          ${columnHeaders.map((header, i) => `<th class="col-header" data-col="${i}">${header}<div class="col-resizer"></div></th>`).join('')}
        </tr>
      </thead>
      <tbody>
  `;

  for (let i = 0; i < rowCount; i++) {
    tableHtml += `
      <tr data-row="${i}">
        <th class="row-header" data-row="${i}">${rowHeaders[i]}<div class="row-resizer"></div></th>
    `;
    for (let j = 0; j < colCount; j++) {
      const cellId = `${columnHeaders[j]}${rowHeaders[i]}`;
      const cell = data[cellId];
      const value = getDisplayValue(cell);
      const formula = cell ? cell.f : '';
      tableHtml += `
        <td data-row="${i}" data-col="${j}">
          <input 
            type="text" 
            value="${escapeHtml(value)}" 
            class="cell-input" 
            data-cell-id="${cellId}"
            data-formula="${escapeHtml(formula)}"
          />
        </td>
      `;
    }
    tableHtml += `</tr>`;
  }

  tableHtml += `
      </tbody>
    </table>
  `;

  return tableHtml;
}