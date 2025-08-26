import { AppState } from '../core/state.js';
import { showToast } from '../ui/toast.js';
import { getWorksheet, persistSnapshot } from '../spreadsheet/workbook-manager.js';
import { renderSheetTabs } from '../spreadsheet/sheet-manager.js';
import { renderSpreadsheetTable } from '../spreadsheet/grid-renderer.js';
/* global XLSX */

export async function importFromFile(file) {
  try {
    if (!file) {
      showToast('No file selected', 'warning');
      return;
    }

    const maxSize = 10 * 1024 * 1024; // 10MB
    if (file.size > maxSize) {
      showToast('File too large (max 10MB)', 'error');
      return;
    }

    showToast('Importing file...', 'info', 2000);

    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array', cellStyles: true });

    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      showToast('Invalid Excel file: no sheets found', 'error');
      return;
    }

    AppState.wb = wb;
    AppState.activeSheet = wb.SheetNames[0] || 'Sheet1';
    await persistSnapshot();
    renderSheetTabs();
    renderSpreadsheetTable();
    showToast(`Imported workbook with ${wb.SheetNames.length} sheet(s)`, 'success');

  } catch (error) {
    console.error('Import failed:', error);
    showToast('Failed to import file: ' + error.message, 'error');
  }
}

export function exportXLSX() {
  try {
    if (!AppState.wb) {
      showToast('No workbook to export', 'warning');
      return;
    }
    XLSX.writeFile(AppState.wb, 'workbook.xlsx', { cellStyles: true });
    showToast('Workbook exported successfully', 'success', 2000);
  } catch (error) {
    console.error('XLSX export failed:', error);
    showToast('Failed to export XLSX: ' + error.message, 'error');
  }
}

export function exportCSV() {
  try {
    if (!AppState.wb) {
      showToast('No workbook to export', 'warning');
      return;
    }

    const ws = getWorksheet();
    if (!ws || !ws['!ref']) {
      showToast('Current sheet is empty', 'warning');
      return;
    }

    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const filename = `${AppState.activeSheet}.csv`;
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 500);
    showToast(`Exported "${filename}" successfully`, 'success', 2000);
  } catch (error) {
    console.error('CSV export failed:', error);
    showToast('Failed to export CSV: ' + error.message, 'error');
  }
}