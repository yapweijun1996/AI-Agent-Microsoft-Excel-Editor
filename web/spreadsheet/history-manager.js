'use strict';

import { AppState } from '../core/state.js';
import { log } from '../utils/index.js';
import { showToast } from '../ui/toast.js';
import { renderSheetTabs } from './sheet-manager.js';
import { renderSpreadsheetTable } from './grid-renderer.js';
import { persistSnapshot } from './workbook-manager.js';

export function saveToHistory(action, data) {
  // Don't save if we're in the middle of undo/redo
  if (AppState.isUndoRedoing) return;

  // Remove any history after current index (when user makes new changes after undo)
  if (AppState.historyIndex < AppState.history.length - 1) {
    AppState.history = AppState.history.slice(0, AppState.historyIndex + 1);
  }

  // Create history entry with full workbook state
  const historyEntry = {
    action,
    data,
    workbook: JSON.parse(JSON.stringify(AppState.wb)), // Deep copy
    activeSheet: AppState.activeSheet,
    timestamp: Date.now()
  };

  AppState.history.push(historyEntry);
  AppState.historyIndex = AppState.history.length - 1;

  // Limit history size
  if (AppState.history.length > AppState.maxHistorySize) {
    AppState.history = AppState.history.slice(-AppState.maxHistorySize);
    AppState.historyIndex = AppState.history.length - 1;
  }

  log('Saved to history:', action, `(${AppState.history.length} entries)`);
}

export function canUndo() {
  return AppState.historyIndex > 0;
}

export function canRedo() {
  return AppState.historyIndex < AppState.history.length - 1;
}

export function undo() {
  if (!canUndo()) {
    showToast('Nothing to undo', 'warning', 1000);
    return;
  }

  AppState.isUndoRedoing = true;
  AppState.historyIndex--;

  const historyEntry = AppState.history[AppState.historyIndex];
  AppState.wb = JSON.parse(JSON.stringify(historyEntry.workbook));
  AppState.activeSheet = historyEntry.activeSheet;

  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();

  AppState.isUndoRedoing = false;

  showToast(`Undid: ${historyEntry.action}`, 'success', 1000);
  log('Undo:', historyEntry.action);
}

export function redo() {
  if (!canRedo()) {
    showToast('Nothing to redo', 'warning', 1000);
    return;
  }

  AppState.isUndoRedoing = true;
  AppState.historyIndex++;

  const historyEntry = AppState.history[AppState.historyIndex];
  AppState.wb = JSON.parse(JSON.stringify(historyEntry.workbook));
  AppState.activeSheet = historyEntry.activeSheet;

  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();

  AppState.isUndoRedoing = false;

  showToast(`Redid: ${historyEntry.action}`, 'success', 1000);
  log('Redo:', historyEntry.action);
}