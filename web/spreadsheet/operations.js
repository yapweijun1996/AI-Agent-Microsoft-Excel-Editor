/* Spreadsheet Operations Handler */
import { AppState } from '../core/state.js';
import { getWorksheet } from './workbook-manager.js';
import { updateCell } from './grid-interactions.js';
import { renderSpreadsheetTable } from './grid-renderer.js';
import { showToast } from '../ui/toast.js';
import { Modal } from '../ui/modal.js';
import { saveToHistory } from './history-manager.js';

/**
 * Apply spreadsheet edits or show dry run preview
 * @param {Object} result - Executor result with edits array
 */
export async function applyEditsOrDryRun(result) {
  if (!result || !result.edits || !Array.isArray(result.edits)) {
    throw new Error('Invalid result format: missing edits array');
  }

  // If dry run mode is enabled, show preview instead of applying
  if (AppState.dryRun) {
    showDryRunPreview(result);
    return;
  }

  // Apply edits
  const ws = getWorksheet();
  let changeCount = 0;

  try {
    saveToHistory(`Apply ${result.edits.length} AI edits`, { 
      operations: result.edits, 
      sheet: AppState.activeSheet 
    });

    for (const edit of result.edits) {
      await applyEdit(edit, ws);
      changeCount++;
    }

    // Re-render the spreadsheet
    renderSpreadsheetTable();
    
    showToast(`Applied ${changeCount} changes successfully`, 'success');
  } catch (error) {
    showToast(`Failed to apply edit ${changeCount + 1}: ${error.message}`, 'error');
    throw error;
  }
}

/**
 * Apply a single edit operation
 */
async function applyEdit(edit, ws) {
  const { op, sheet, cell, range, value, values, formula, format } = edit;

  // Validate sheet (default to current if not specified)
  if (sheet && sheet !== AppState.activeSheet) {
    showToast(`Warning: Edit targets sheet "${sheet}" but current sheet is "${AppState.activeSheet}"`, 'warning');
  }

  switch (op) {
    case 'setCell':
      if (!cell || value === undefined) {
        throw new Error('setCell operation requires cell and value');
      }
      updateCell(cell, String(value));
      break;

    case 'setFormula':
      if (!cell || !formula) {
        throw new Error('setFormula operation requires cell and formula');
      }
      const formulaValue = formula.startsWith('=') ? formula : `=${formula}`;
      updateCell(cell, formulaValue);
      break;

    case 'setRange':
      if (!range || !values || !Array.isArray(values)) {
        throw new Error('setRange operation requires range and values array');
      }
      await applyRangeValues(range, values, ws);
      break;

    case 'insertRow':
    case 'deleteRow':
    case 'insertColumn':
    case 'deleteColumn':
      showToast(`Operation "${op}" not yet implemented`, 'warning');
      break;

    case 'formatCell':
    case 'formatRange':
      showToast(`Formatting operation "${op}" not yet implemented`, 'warning');
      break;

    default:
      throw new Error(`Unsupported operation: ${op}`);
  }
}

/**
 * Apply values to a range of cells
 */
async function applyRangeValues(range, values, ws) {
  // Parse range (e.g., "A1:C3")
  const [start, end] = range.split(':');
  if (!start || !end) {
    throw new Error(`Invalid range format: ${range}`);
  }

  const startCol = start.match(/[A-Z]+/)[0];
  const startRow = parseInt(start.match(/\d+/)[0]);
  const endCol = end.match(/[A-Z]+/)[0];
  const endRow = parseInt(end.match(/\d+/)[0]);

  const startColIndex = columnToNumber(startCol);
  const endColIndex = columnToNumber(endCol);

  for (let row = 0; row < values.length; row++) {
    if (startRow + row > endRow) break;
    
    for (let col = 0; col < values[row].length; col++) {
      if (startColIndex + col > endColIndex) break;
      
      const cellAddr = numberToColumn(startColIndex + col) + (startRow + row);
      const cellValue = values[row][col];
      
      if (cellValue !== null && cellValue !== undefined) {
        updateCell(cellAddr, String(cellValue));
      }
    }
  }
}

/**
 * Show dry run preview modal
 */
function showDryRunPreview(result) {
  const modal = new Modal();
  
  let previewHtml = `
    <div class="space-y-4">
      <div class="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
        <div class="flex items-center space-x-2">
          <span class="text-yellow-600">🔍</span>
          <strong class="text-yellow-800">Dry Run Mode - Preview Only</strong>
        </div>
        <p class="text-sm text-yellow-700 mt-2">These changes will NOT be applied to your spreadsheet.</p>
      </div>
      
      <div class="space-y-3">
        <h4 class="font-medium text-gray-900">Planned Operations (${result.edits.length}):</h4>
  `;

  result.edits.forEach((edit, index) => {
    const { op, cell, range, value, formula } = edit;
    let description = '';
    
    switch (op) {
      case 'setCell':
        description = `Set cell ${cell} to "${value}"`;
        break;
      case 'setFormula':
        description = `Set formula in ${cell}: ${formula}`;
        break;
      case 'setRange':
        description = `Update range ${range}`;
        break;
      default:
        description = `${op} operation`;
    }

    previewHtml += `
      <div class="flex items-start space-x-3 p-3 bg-gray-50 rounded border">
        <span class="text-sm font-mono text-gray-500">${index + 1}.</span>
        <div>
          <div class="text-sm font-medium text-gray-900">${description}</div>
          <div class="text-xs text-gray-600 mt-1">Operation: ${op}</div>
        </div>
      </div>
    `;
  });

  previewHtml += `
      </div>
      
      <div class="text-sm text-gray-600">
        <strong>Note:</strong> Turn off dry run mode to apply these changes to your spreadsheet.
      </div>
    </div>
  `;

  modal.show({
    title: 'Dry Run Preview',
    content: previewHtml,
    buttons: [
      { text: 'Close', action: 'close', primary: true }
    ],
    size: 'lg'
  });
}

/**
 * Convert column letter to number (A=1, B=2, etc.)
 */
function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Convert column number to letter (1=A, 2=B, etc.)
 */
function numberToColumn(number) {
  let result = '';
  while (number > 0) {
    number--;
    result = String.fromCharCode(65 + (number % 26)) + result;
    number = Math.floor(number / 26);
  }
  return result;
}

// Make function available globally for task-manager.js
window.applyEditsOrDryRun = applyEditsOrDryRun;