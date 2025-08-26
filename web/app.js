/* AI Excel Editor - App Core */
'use strict';

import { AppState } from './core/state.js';
import { db } from './db/indexeddb.js';
import { showToast } from './ui/toast.js';
import { showWelcomeModal } from './ui/modals.js';
import { restoreApiKeys, updateProviderStatus } from './services/api-keys.js';
import { ensureWorkbook } from './spreadsheet/workbook-manager.js';
import { renderSheetTabs } from './spreadsheet/sheet-manager.js';
import { renderSpreadsheetTable } from './spreadsheet/grid-renderer.js';
import { loadTasks, drawTasks } from './tasks/task-manager.js';
import { bindUI } from './ui/bindings.js';
import './spreadsheet/operations.js'; // Import for applyEditsOrDryRun global function

// Enhanced initialization with loading state
function showLoadingOverlay() {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    overlay.classList.remove('hidden');
  }
}

function hideLoadingOverlay() {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    overlay.classList.add('hidden');
  }
}

function isFirstVisit() {
  return !localStorage.getItem('hasVisited');
}

// Responsive layout handled entirely by CSS media queries
// All desktop/mobile logic removed - unified responsive system

// Main App Initialization
async function main() {
  showLoadingOverlay();
  try {
    await db.init();
    restoreApiKeys();
    await ensureWorkbook();
    await loadTasks();

    renderSheetTabs();
    renderSpreadsheetTable();
    drawTasks();
    bindUI();
    updateProviderStatus();

    document.querySelector('.main-container').classList.add('animate-fade-in-up');
    hideLoadingOverlay();

    const hasKeys = AppState.keys.openai || AppState.keys.gemini;
    if (hasKeys) {
      const provider = AppState.keys.openai ? 'OpenAI' : 'Gemini';
      showToast(`AI Excel Editor ready! Using ${provider} for AI features.`, 'success', 3000);
    } else {
      showToast('AI Excel Editor ready! Set your API keys to enable AI features.', 'success', 3000);
    }

    if (isFirstVisit()) {
      showWelcomeModal();
      localStorage.setItem('hasVisited', 'true');
    }
    
    // CSS media queries handle responsiveness automatically
    // Re-render grid on resize for any dynamic content updates
    let resizeTimeout;
    window.addEventListener('resize', () => {
      clearTimeout(resizeTimeout);
      resizeTimeout = setTimeout(() => {
        renderSpreadsheetTable(true);
      }, 250);
    });
    
    // Handle orientation change
    window.addEventListener('orientationchange', () => {
      setTimeout(() => {
        renderSpreadsheetTable(true);
      }, 300);
    });
  } catch (e) {
    console.error("Initialization failed", e);
    hideLoadingOverlay();
    showToast('Error initializing application: ' + e.message, 'error');
  }
}

document.addEventListener('DOMContentLoaded', main);