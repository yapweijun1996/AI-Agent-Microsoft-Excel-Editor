import { AppState } from '../core/state.js';
import { showApiKeyModal, showHelpModal, showSortModal, showChartModal, showCommentModal } from './modals.js';
import { onSend } from '../chat/chat-ui.js';
import { exportXLSX, exportCSV, importFromFile } from '../file/import-export.js';
import { addNewSheet, deleteSheet, switchToSheet } from '../spreadsheet/sheet-manager.js';
import { undo, redo } from '../spreadsheet/history-manager.js';
import { insertRowAtSelection, insertColumnAtSelectionLeft, deleteSelectedRow, deleteSelectedColumn, insertFormula, applyFormat } from '../spreadsheet/grid-interactions.js';
import { log } from '../utils/index.js';
import { showToast } from './toast.js';
import { debounce } from '../utils/index.js';
import { renderSpreadsheetTable } from '../spreadsheet/grid-renderer.js';
import { pickProvider, getSelectedModel } from '../services/api-keys.js';
import { executeTasks } from '../tasks/task-manager.js';
/* global XLSX, Chart */

// UI bindings
function initKeyboardShortcuts() {
  document.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.contentEditable === 'true') return;

    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 's':
          e.preventDefault();
          exportXLSX();
          showToast('Workbook saved as XLSX', 'success', 1000);
          break;
        case 'o':
          e.preventDefault();
          document.getElementById('import-xlsx-input').click();
          break;
        case 't':
          e.preventDefault();
          addNewSheet();
          break;
        case 'w':
          e.preventDefault();
          if (AppState.wb && AppState.wb.SheetNames.length > 1) {
            deleteSheet(AppState.activeSheet);
          }
          break;
        case '1':
        case '2':
        case '3':
        case '4':
        case '5':
        case '6':
        case '7':
        case '8':
        case '9':
          e.preventDefault();
          const sheetIndex = parseInt(e.key) - 1;
          if (AppState.wb && AppState.wb.SheetNames[sheetIndex]) {
            switchToSheet(AppState.wb.SheetNames[sheetIndex]);
          }
          break;
        case 'enter':
          e.preventDefault();
          document.getElementById('message-input').focus();
          break;
        case 'z':
          e.preventDefault();
          if (e.shiftKey) {
            redo();
          } else {
            undo();
          }
          break;
        case 'y':
          e.preventDefault();
          redo();
          break;
      }
    }

    if (e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const chatInput = document.getElementById('message-input');
      if (document.activeElement !== chatInput) {
        e.preventDefault();
        const currentIndex = AppState.wb.SheetNames.indexOf(AppState.activeSheet);
        const nextIndex = e.shiftKey ?
          (currentIndex - 1 + AppState.wb.SheetNames.length) % AppState.wb.SheetNames.length :
          (currentIndex + 1) % AppState.wb.SheetNames.length;
        switchToSheet(AppState.wb.SheetNames[nextIndex]);
      }
    }

    if (e.key === 'F2') {
      e.preventDefault();
      document.getElementById('message-input').focus();
    }

    if (e.key === 'Escape') {
      const chatInput = document.getElementById('message-input');
      if (document.activeElement === chatInput) {
        chatInput.blur();
      }
    }
  });
}

function initRibbonTabs() {
  const tabs = document.querySelectorAll('.ribbon-tab');
  const ribbonContent = document.getElementById('ribbon-content');

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      // Remove active class from all tabs
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');

      // Hide all ribbon content sections
      ribbonContent.querySelectorAll('[id$="-ribbon"]').forEach(content => {
        content.classList.add('hidden');
        content.classList.remove('flex');
      });

      // Show the selected ribbon content
      const tabName = tab.dataset.tab;
      const contentToShow = document.getElementById(`${tabName}-ribbon`);
      if (contentToShow) {
        contentToShow.classList.remove('hidden');
        contentToShow.classList.add('flex');
      }
    });
  });

  // Initialize with Home tab active
  const homeTab = document.querySelector('[data-tab="home"]');
  if (homeTab) {
    homeTab.click();
  }
}

export function bindUI() {
  document.getElementById('openai-key-btn')?.addEventListener('click', (e) => {
    e.preventDefault();
    log('OpenAI button clicked');
    showApiKeyModal('OpenAI');
  });

  document.getElementById('gemini-key-btn')?.addEventListener('click', (e) => {
    e.preventDefault();
    log('Gemini button clicked');
    showApiKeyModal('Gemini');
  });

  document.getElementById('help-btn')?.addEventListener('click', showHelpModal);

  document.getElementById('dry-run-toggle')?.addEventListener('change', (e) => { AppState.dryRun = e.target.checked; });
  document.getElementById('auto-execute-toggle')?.addEventListener('change', (e) => { AppState.autoExecute = e.target.checked; });

  document.getElementById('model-select')?.addEventListener('change', (e) => {
    AppState.selectedModel = e.target.value;
    const provider = pickProvider();
    const model = getSelectedModel();
    showToast(`Selected: ${provider === 'mock' ? 'Mock Mode' : `${provider.toUpperCase()} - ${model}`}`, 'info', 2000);
  });

  document.getElementById('send-btn')?.addEventListener('click', onSend);
  document.getElementById('message-input')?.addEventListener('keypress', (e) => { if (e.key === 'Enter') { e.preventDefault(); onSend(); } });

  document.getElementById('export-xlsx')?.addEventListener('click', exportXLSX);
  document.getElementById('export-csv')?.addEventListener('click', exportCSV);

  const importBtn = document.getElementById('import-xlsx');
  const importInput = document.getElementById('import-xlsx-input');
  if (importBtn && importInput) {
    importBtn.addEventListener('click', () => importInput.click());
    importInput.addEventListener('change', () => { if (importInput.files?.[0]) importFromFile(importInput.files[0]); });
  }

  document.getElementById('execute-all-tasks')?.addEventListener('click', () => {
    const pendingTasks = AppState.tasks.filter(t => t.status === 'pending');
    if (pendingTasks.length === 0) {
      showToast('No pending tasks to execute', 'info');
      return;
    }
    executeTasks(pendingTasks.map(t => t.id));
  });

  document.getElementById('add-sheet-btn')?.addEventListener('click', addNewSheet);

  document.getElementById('insert-row-btn')?.addEventListener('click', insertRowAtSelection);
  document.getElementById('insert-col-btn')?.addEventListener('click', insertColumnAtSelectionLeft);
  document.getElementById('delete-row-btn')?.addEventListener('click', deleteSelectedRow);
  document.getElementById('delete-col-btn')?.addEventListener('click', deleteSelectedColumn);

  document.getElementById('toggle-ai-panel')?.addEventListener('click', () => {
    const aiPanel = document.getElementById('ai-panel');
    aiPanel.style.display = aiPanel.style.display === 'none' ? 'flex' : 'none';
  });

  document.getElementById('view-tasks')?.addEventListener('click', () => {
    document.getElementById('task-modal').classList.remove('hidden');
  });

  document.getElementById('close-task-modal')?.addEventListener('click', () => {
    document.getElementById('task-modal').classList.add('hidden');
  });

  document.getElementById('formula-bar')?.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      const cellRef = document.getElementById('cell-reference').textContent;
      updateCell(cellRef, e.target.value);
      e.preventDefault();
    }
  });

  document.getElementById('format-bold')?.addEventListener('click', () => applyFormat('bold'));
  document.getElementById('format-italic')?.addEventListener('click', () => applyFormat('italic'));
  document.getElementById('format-underline')?.addEventListener('click', () => applyFormat('underline'));
  document.getElementById('format-color')?.addEventListener('input', (e) => applyFormat('color', e.target.value));

  document.getElementById('sort-btn')?.addEventListener('click', showSortModal);
  document.getElementById('chart-btn')?.addEventListener('click', showChartModal);
  document.getElementById('comment-btn')?.addEventListener('click', showCommentModal);

  // Page Layout ribbon bindings
  document.getElementById('orientation-btn')?.addEventListener('click', () => showToast('Page orientation feature coming soon', 'info'));
  document.getElementById('margins-btn')?.addEventListener('click', () => showToast('Page margins feature coming soon', 'info'));

  // Formulas ribbon bindings
  document.getElementById('sum-btn')?.addEventListener('click', () => insertFormula('=SUM()'));
  document.getElementById('avg-btn')?.addEventListener('click', () => insertFormula('=AVERAGE()'));
  document.getElementById('count-btn')?.addEventListener('click', () => insertFormula('=COUNT()'));

  document.getElementById('spreadsheet')?.addEventListener('scroll', debounce(renderSpreadsheetTable, 16));

  initKeyboardShortcuts();
  initRibbonTabs();
}