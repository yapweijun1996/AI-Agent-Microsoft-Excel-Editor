'use strict';

import { AppState } from '../core/state.js';
import { escapeHtml } from '../utils/index.js';
import { showToast } from '../ui/toast.js';
import { saveToHistory } from './history-manager.js';
import { renderSpreadsheetTable } from './grid-renderer.js';
import { persistSnapshot } from './workbook-manager.js';
import { Modal } from '../ui/modal.js';
import { getFormulaEngine } from '../FormulaEngine.js';

export function renderSheetTabs() {
  if (!AppState.wb) return;
  const container = document.getElementById('sheet-tabs');
  if (!container) return;

  let html = '';
  for (const sheetName of AppState.wb.SheetNames) {
    const isActive = sheetName === AppState.activeSheet;
    html += `
      <div class="sheet-tab ${isActive ? 'active' : ''}" data-sheet="${escapeHtml(sheetName)}">
        <span class="sheet-name">${escapeHtml(sheetName)}</span>
        ${AppState.wb.SheetNames.length > 1 ? `<span class="sheet-tab-close" data-sheet-close="${escapeHtml(sheetName)}">&times;</span>` : ''}
      </div>`;
  }
  container.innerHTML = html;

  // Add click handlers
  container.querySelectorAll('.sheet-tab').forEach(tab => {
    tab.addEventListener('click', (e) => {
      if (!e.target.classList.contains('sheet-tab-close')) {
        switchToSheet(tab.dataset.sheet);
      }
    });
  });

  // Add close button handlers
  container.querySelectorAll('.sheet-tab-close').forEach(button => {
    button.addEventListener('click', (e) => {
      e.stopPropagation(); // Prevent tab switch
      deleteSheet(button.dataset.sheetClose);
    });
  });
}

export function switchToSheet(sheetName) {
  if (!AppState.wb || !AppState.wb.Sheets[sheetName]) return;
  AppState.activeSheet = sheetName;
  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
}

export function addNewSheet() {
  if (!AppState.wb) return;

  let newName = 'Sheet1';
  let counter = 1;
  while (AppState.wb.SheetNames.includes(newName)) {
    counter++;
    newName = `Sheet${counter}`;
  }

  saveToHistory(`Add sheet "${newName}"`, { sheetName: newName });

  const ws = XLSX.utils.aoa_to_sheet([['']]);
  XLSX.utils.book_append_sheet(AppState.wb, ws, newName);
  AppState.wbVersion++;
  getFormulaEngine(AppState.wb, AppState.activeSheet).invalidateCache();
  AppState.activeSheet = newName;
  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
  showToast(`Added sheet "${newName}"`, 'success');
}

export async function deleteSheet(sheetName) {
  if (!AppState.wb || AppState.wb.SheetNames.length <= 1) {
    showToast('Cannot delete the last sheet', 'warning');
    return;
  }

  const confirmed = await Modal.confirm(
    `Delete sheet "${sheetName}"? This action cannot be undone.`, 
    { 
      title: 'Delete Sheet',
      confirmText: 'Delete',
      cancelText: 'Cancel',
      dangerousAction: true 
    }
  );
  
  if (!confirmed) return;

  const sheetIndex = AppState.wb.SheetNames.indexOf(sheetName);
  if (sheetIndex === -1) return;

  saveToHistory(`Delete sheet "${sheetName}"`, { sheetName, sheetIndex });

  // Remove from SheetNames and Sheets
  AppState.wb.SheetNames.splice(sheetIndex, 1);
  delete AppState.wb.Sheets[sheetName];
  AppState.wbVersion++;
  getFormulaEngine(AppState.wb, AppState.activeSheet).invalidateCache();

  // Switch to another sheet if this was active
  if (AppState.activeSheet === sheetName) {
    AppState.activeSheet = AppState.wb.SheetNames[Math.max(0, sheetIndex - 1)];
  }

  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
  showToast(`Deleted sheet "${sheetName}"`, 'success');
}