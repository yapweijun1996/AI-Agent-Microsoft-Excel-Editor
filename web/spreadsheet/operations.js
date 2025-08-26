/* Spreadsheet Operations Handler */
import { AppState } from '../core/state.js';
import { getWorksheet } from './workbook-manager.js';
import { updateCell } from './grid-interactions.js';
import { renderSpreadsheetTable } from './grid-renderer.js';
import { showToast } from '../ui/toast.js';
import { Modal } from '../ui/modal.js';
import { saveToHistory } from './history-manager.js';
import { columnToNumber, numberToColumn, AppError, ERROR_CODES, handleError } from '../utils/index.js';
import { registerGlobal } from '../core/global-bindings.js';

// XLSX is loaded globally via CDN
const XLSX = window.XLSX;

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
    const wrappedError = error instanceof AppError ? error : 
      new AppError(`Failed to apply edit ${changeCount + 1}: ${error.message}`, ERROR_CODES.OPERATION_FAILED);
    handleError(wrappedError, { operation: 'applyEditsOrDryRun', editCount: changeCount + 1 });
    throw wrappedError;
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
      await insertRow(edit.rowIndex || edit.row || 1, ws);
      break;
      
    case 'deleteRow':
      await deleteRow(edit.rowIndex || edit.row || 1, ws);
      break;
      
    case 'insertColumn':
      await insertColumn(edit.columnIndex || edit.column || 1, ws);
      break;
      
    case 'deleteColumn':
      await deleteColumn(edit.columnIndex || edit.column || 1, ws);
      break;

    case 'formatCell':
      await formatCell(cell, format, ws);
      break;
      
    case 'formatRange':
      await formatRange(range, format, ws);
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
 * Insert a new row at the specified index
 */
async function insertRow(rowIndex, ws) {
  if (!ws || typeof rowIndex !== 'number' || rowIndex < 1) {
    throw new Error('Invalid row index for insertion');
  }

  // Shift all rows down from the insertion point
  const range = ws['!ref'];
  if (range) {
    const { s: start, e: end } = ws['!ref'] ? XLSX.utils.decode_range(range) : { s: { c: 0, r: 0 }, e: { c: 25, r: 1000 } };
    
    // Process from bottom to top to avoid overwriting
    for (let r = end.r; r >= rowIndex - 1; r--) {
      for (let c = start.c; c <= end.c; c++) {
        const currentAddr = XLSX.utils.encode_cell({ c, r });
        const newAddr = XLSX.utils.encode_cell({ c, r: r + 1 });
        
        if (ws[currentAddr]) {
          ws[newAddr] = { ...ws[currentAddr] };
          delete ws[currentAddr];
        }
      }
    }
    
    // Update the range
    ws['!ref'] = XLSX.utils.encode_range({
      s: start,
      e: { ...end, r: end.r + 1 }
    });
  }
  
  showToast(`Row ${rowIndex} inserted successfully`, 'success');
}

/**
 * Delete the row at the specified index
 */
async function deleteRow(rowIndex, ws) {
  if (!ws || typeof rowIndex !== 'number' || rowIndex < 1) {
    throw new Error('Invalid row index for deletion');
  }

  const range = ws['!ref'];
  if (range) {
    const { s: start, e: end } = XLSX.utils.decode_range(range);
    
    if (rowIndex - 1 > end.r) {
      throw new Error('Row index beyond worksheet range');
    }
    
    // Delete the target row
    for (let c = start.c; c <= end.c; c++) {
      const addr = XLSX.utils.encode_cell({ c, r: rowIndex - 1 });
      delete ws[addr];
    }
    
    // Shift all rows up from deletion point
    for (let r = rowIndex; r <= end.r; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const currentAddr = XLSX.utils.encode_cell({ c, r });
        const newAddr = XLSX.utils.encode_cell({ c, r: r - 1 });
        
        if (ws[currentAddr]) {
          ws[newAddr] = { ...ws[currentAddr] };
          delete ws[currentAddr];
        }
      }
    }
    
    // Update the range
    if (end.r > 0) {
      ws['!ref'] = XLSX.utils.encode_range({
        s: start,
        e: { ...end, r: end.r - 1 }
      });
    }
  }
  
  showToast(`Row ${rowIndex} deleted successfully`, 'success');
}

/**
 * Insert a new column at the specified index
 */
async function insertColumn(columnIndex, ws) {
  if (!ws || typeof columnIndex !== 'number' || columnIndex < 1) {
    throw new Error('Invalid column index for insertion');
  }

  const range = ws['!ref'];
  if (range) {
    const { s: start, e: end } = XLSX.utils.decode_range(range);
    
    // Process from right to left to avoid overwriting
    for (let c = end.c; c >= columnIndex - 1; c--) {
      for (let r = start.r; r <= end.r; r++) {
        const currentAddr = XLSX.utils.encode_cell({ c, r });
        const newAddr = XLSX.utils.encode_cell({ c: c + 1, r });
        
        if (ws[currentAddr]) {
          ws[newAddr] = { ...ws[currentAddr] };
          delete ws[currentAddr];
        }
      }
    }
    
    // Update the range
    ws['!ref'] = XLSX.utils.encode_range({
      s: start,
      e: { ...end, c: end.c + 1 }
    });
  }
  
  const columnLetter = numberToColumn(columnIndex);
  showToast(`Column ${columnLetter} inserted successfully`, 'success');
}

/**
 * Delete the column at the specified index
 */
async function deleteColumn(columnIndex, ws) {
  if (!ws || typeof columnIndex !== 'number' || columnIndex < 1) {
    throw new Error('Invalid column index for deletion');
  }

  const range = ws['!ref'];
  if (range) {
    const { s: start, e: end } = XLSX.utils.decode_range(range);
    
    if (columnIndex - 1 > end.c) {
      throw new Error('Column index beyond worksheet range');
    }
    
    // Delete the target column
    for (let r = start.r; r <= end.r; r++) {
      const addr = XLSX.utils.encode_cell({ c: columnIndex - 1, r });
      delete ws[addr];
    }
    
    // Shift all columns left from deletion point
    for (let c = columnIndex; c <= end.c; c++) {
      for (let r = start.r; r <= end.r; r++) {
        const currentAddr = XLSX.utils.encode_cell({ c, r });
        const newAddr = XLSX.utils.encode_cell({ c: c - 1, r });
        
        if (ws[currentAddr]) {
          ws[newAddr] = { ...ws[currentAddr] };
          delete ws[currentAddr];
        }
      }
    }
    
    // Update the range
    if (end.c > 0) {
      ws['!ref'] = XLSX.utils.encode_range({
        s: start,
        e: { ...end, c: end.c - 1 }
      });
    }
  }
  
  const columnLetter = numberToColumn(columnIndex);
  showToast(`Column ${columnLetter} deleted successfully`, 'success');
}

/**
 * Apply formatting to a single cell
 */
async function formatCell(cellAddr, format, ws) {
  if (!cellAddr || !format || !ws) {
    throw new Error('Invalid parameters for cell formatting');
  }

  // Ensure the cell exists
  if (!ws[cellAddr]) {
    ws[cellAddr] = { t: 's', v: '' };
  }

  // Apply formatting
  if (!ws[cellAddr].s) {
    ws[cellAddr].s = {};
  }

  // Apply specific formatting properties
  if (format.bold !== undefined) {
    ws[cellAddr].s.font = ws[cellAddr].s.font || {};
    ws[cellAddr].s.font.bold = format.bold;
  }
  
  if (format.italic !== undefined) {
    ws[cellAddr].s.font = ws[cellAddr].s.font || {};
    ws[cellAddr].s.font.italic = format.italic;
  }
  
  if (format.color) {
    ws[cellAddr].s.font = ws[cellAddr].s.font || {};
    ws[cellAddr].s.font.color = { rgb: format.color.replace('#', '') };
  }
  
  if (format.backgroundColor) {
    ws[cellAddr].s.fill = {
      fgColor: { rgb: format.backgroundColor.replace('#', '') },
      patternType: 'solid'
    };
  }
  
  if (format.numberFormat) {
    ws[cellAddr].s.numFmt = format.numberFormat;
  }
  
  showToast(`Cell ${cellAddr} formatted successfully`, 'success');
}

/**
 * Apply formatting to a range of cells
 */
async function formatRange(range, format, ws) {
  if (!range || !format || !ws) {
    throw new Error('Invalid parameters for range formatting');
  }

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

  // Apply formatting to each cell in the range
  for (let row = startRow; row <= endRow; row++) {
    for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex++) {
      const cellAddr = numberToColumn(colIndex) + row;
      await formatCell(cellAddr, format, ws);
    }
  }
  
  showToast(`Range ${range} formatted successfully`, 'success');
}

// Register for global access with proper deprecation notice
registerGlobal('applyEditsOrDryRun', applyEditsOrDryRun, {
  deprecated: true,
  description: 'Consider using the Operations namespace in future versions'
});