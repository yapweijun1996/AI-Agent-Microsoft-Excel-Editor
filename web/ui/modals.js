'use strict';

import { Modal } from './modal.js';
import { log } from '../utils/index.js';
import { saveApiKey } from '../services/api-keys.js';
import { showToast } from './toast.js';
import { AppState } from '../core/state.js';

export function showApiKeyModal(provider) {
  log(`Opening API key modal for ${provider}`);
  const modal = new Modal();
  modal.show({
    title: `Set ${provider} API Key`,
    content: `
      <div class="space-y-4">
        <div class="flex items-start space-x-3 p-3 bg-blue-50 rounded-lg">
          <svg class="w-5 h-5 text-blue-600 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
          </svg>
          <div class="text-sm text-blue-800">
            <p class="font-medium mb-1">API Key Setup</p>
            <p>Enter your ${provider} API key to enable AI-powered spreadsheet automation. The key will be stored securely in your browser.</p>
          </div>
        </div>
        <div class="space-y-2">
          <input type="password" id="api-key-input" placeholder="Enter API key..." class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500" />
          <div class="text-xs text-gray-600">
            <p class="mb-1">Get your API key from:</p>
            <a href="${provider === 'OpenAI' ? 'https://platform.openai.com/api-keys' : 'https://aistudio.google.com/app/apikey'}" target="_blank" class="text-blue-600 hover:text-blue-800 underline">
              ${provider === 'OpenAI' ? 'OpenAI Platform' : 'Google AI Studio'}
            </a>
          </div>
        </div>
        <label class="flex items-center space-x-2 text-xs text-gray-600 bg-yellow-50 p-2 rounded">
          <input id="persist-key" type="checkbox" class="rounded border-gray-300">
          <span>⚠️ Persist to localStorage (less secure, but convenient)</span>
        </label>
      </div>`,
    buttons: [
      { text: 'Cancel', action: 'cancel' },
      {
        text: 'Save Key', action: 'save', primary: true, onClick: () => {
          const key = document.getElementById('api-key-input').value.trim();
          const persist = document.getElementById('persist-key').checked;
          if (key) {
            saveApiKey(provider.toLowerCase(), key, persist);
            showToast(`${provider} API key saved successfully`, 'success');
          } else {
            showToast('Please enter a valid API key', 'warning');
          }
        }
      }
    ],
    size: 'lg'
  });
}

export function showHelpModal() {
  const modal = new Modal();
  modal.show({
    title: 'Keyboard Shortcuts & Help',
    content: `
      <div class="space-y-6 text-sm">
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">File Operations</h4>
          <div class="space-y-1 text-gray-600">
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+S</kbd> Export as XLSX</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+O</kbd> Import XLSX file</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Z</kbd> Undo</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Y</kbd> / <kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Shift+Z</kbd> Redo</div>
          </div>
        </div>
        
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">Sheet Operations</h4>
          <div class="space-y-1 text-gray-600">
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+T</kbd> Add new sheet</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+W</kbd> Delete current sheet</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Tab</kbd> / <kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Shift+Tab</kbd> Switch between sheets</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+1-9</kbd> Switch to sheet by number</div>
          </div>
        </div>
        
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">Chat & AI</h4>
          <div class="space-y-1 text-gray-600">
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">F2</kbd> Focus chat input</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Enter</kbd> Focus chat input</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Enter</kbd> Send message</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Escape</kbd> Clear chat input focus</div>
          </div>
        </div>
        
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">Tips & Example Commands</h4>
          <ul class="space-y-1 text-gray-600 text-xs">
            <li>• Set your OpenAI or Gemini API key to use AI features</li>
            <li>• Enable "Dry Run" to preview AI changes before applying</li>
            <li>• <strong>Excel Operations:</strong> "Add a column after B", "Insert 3 rows at row 5"</li>
            <li>• <strong>Formulas:</strong> "Add SUM formula in C10", "Calculate average in D1"</li>
            <li>• <strong>Data:</strong> "Create header row with Name, Age, Salary", "Sort by column A"</li>
            <li>• <strong>Formatting:</strong> "Format column C as currency", "Make header row bold"</li>
            <li>• AI agents work across multiple sheets in your workbook</li>
          </ul>
        </div>
      </div>`,
    buttons: [{ text: 'Close', action: 'close', primary: true }],
    size: 'lg'
  });
}

export function showSettingsModal() {
  const modal = new Modal();
  modal.show({
    title: 'Settings & Preferences',
    content: `
      <div class="space-y-6">
        <!-- API Configuration -->
        <div>
          <h4 class="font-semibold text-gray-900 mb-3">AI Configuration</h4>
          <div class="space-y-3">
            <div class="flex items-center justify-between p-3 bg-gray-50 rounded-lg">
              <div>
                <div class="font-medium text-gray-900">OpenAI API Key</div>
                <div class="text-sm text-gray-600">Enable GPT-4 powered features</div>
              </div>
              <button onclick="showApiKeyModal('OpenAI')" class="px-3 py-1 bg-blue-500 text-white text-sm rounded hover:bg-blue-600">
                Configure
              </button>
            </div>
            <div class="flex items-center justify-between p-3 bg-gray-50 rounded-lg">
              <div>
                <div class="font-medium text-gray-900">Gemini API Key</div>
                <div class="text-sm text-gray-600">Enable Gemini powered features</div>
              </div>
              <button onclick="showApiKeyModal('Gemini')" class="px-3 py-1 bg-blue-500 text-white text-sm rounded hover:bg-blue-600">
                Configure
              </button>
            </div>
          </div>
        </div>

        <!-- Accessibility -->
        <div>
          <h4 class="font-semibold text-gray-900 mb-3">Accessibility</h4>
          <div class="space-y-3">
            <label class="flex items-center justify-between p-3 bg-gray-50 rounded-lg cursor-pointer">
              <div>
                <div class="font-medium text-gray-900">Reduced Motion</div>
                <div class="text-sm text-gray-600">Disable animations and transitions</div>
              </div>
              <input type="checkbox" id="reduced-motion-toggle" class="rounded border-gray-300" ${AppState.reducedMotion ? 'checked' : ''}>
            </label>
          </div>
        </div>

        <!-- Task Execution -->
        <div>
          <h4 class="font-semibold text-gray-900 mb-3">Task Execution</h4>
          <div class="space-y-3">
            <label class="flex items-center justify-between p-3 bg-gray-50 rounded-lg cursor-pointer">
              <div>
                <div class="font-medium text-gray-900">Auto-execute Tasks</div>
                <div class="text-sm text-gray-600">Automatically run tasks after planning</div>
              </div>
              <input type="checkbox" id="auto-execute-toggle" class="rounded border-gray-300" ${AppState.autoExecute ? 'checked' : ''}>
            </label>
          </div>
        </div>
      </div>`,
    buttons: [
      { text: 'Cancel', action: 'cancel' },
      {
        text: 'Save Settings', action: 'save', primary: true, onClick: () => {
          // Save reduced motion preference
          const reducedMotion = document.getElementById('reduced-motion-toggle').checked;
          AppState.reducedMotion = reducedMotion;
          localStorage.setItem('reducedMotion', reducedMotion);
          
          // Save auto-execute preference
          const autoExecute = document.getElementById('auto-execute-toggle').checked;
          AppState.autoExecute = autoExecute;
          localStorage.setItem('autoExecute', autoExecute);
          
          showToast('Settings saved successfully', 'success');
        }
      }
    ],
    size: 'lg'
  });
}

export function showWelcomeModal() {
  const modal = new Modal();
  modal.show({
    title: 'Welcome to the AI Excel Editor!',
    content: `
            <div class="space-y-4 text-sm">
                <p>This powerful tool combines a familiar spreadsheet interface with advanced AI capabilities to help you automate tasks, analyze data, and streamline your workflows.</p>
                <p><strong>Getting Started:</strong></p>
                <ul class="list-disc list-inside space-y-2">
                    <li><strong>Set Your API Key:</strong> Click on "Set OpenAI Key" or "Set Gemini Key" to connect to your preferred AI provider.</li>
                    <li><strong>Interact with the AI:</strong> Use the chat panel to give commands like "Create a budget for Q3" or "Summarize sales data."</li>
                    <li><strong>Explore the Ribbon:</strong> The ribbon menu provides familiar Excel-like formatting and data manipulation tools.</li>
                </ul>
                <p>For a detailed guide and more examples, click the "Help" button at any time.</p>
            </div>
        `,
    buttons: [{ text: 'Get Started', action: 'close', primary: true }],
    size: 'lg'
  });
}

export function showSortModal() {
  const modal = new Modal();
  modal.show({
    title: 'Sort Data',
    content: `
      <div class="space-y-4">
        <div class="flex items-start space-x-3 p-3 bg-amber-50 rounded-lg">
          <svg class="w-5 h-5 text-amber-600 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v3m0 0v3m0-3h3m-3 0H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z"/>
          </svg>
          <div class="text-sm text-amber-800">
            <p class="font-medium mb-1">Feature Coming Soon</p>
            <p>Sort functionality is currently in development. For now, try using AI commands like:</p>
            <ul class="mt-2 space-y-1 text-xs">
              <li>• "Sort column A alphabetically"</li>
              <li>• "Sort data by column B in descending order"</li>
              <li>• "Sort the table by name column"</li>
            </ul>
          </div>
        </div>
      </div>
    `,
    buttons: [{ text: 'Got it', action: 'close', primary: true }]
  });
}

export function showChartModal() {
  const modal = new Modal();
  modal.show({
    title: 'Create Chart',
    content: `
      <div class="space-y-4">
        <div class="flex items-start space-x-3 p-3 bg-purple-50 rounded-lg">
          <svg class="w-5 h-5 text-purple-600 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/>
          </svg>
          <div class="text-sm text-purple-800">
            <p class="font-medium mb-1">Charts & Visualization</p>
            <p>Chart creation is in development. Try these AI commands instead:</p>
            <ul class="mt-2 space-y-1 text-xs">
              <li>• "Create a bar chart from my data"</li>
              <li>• "Generate a line graph for sales trends"</li>
              <li>• "Make a pie chart showing percentages"</li>
            </ul>
          </div>
        </div>
      </div>
    `,
    buttons: [{ text: 'Try AI Commands', action: 'close', primary: true }]
  });
}

export function showCommentModal() {
  const modal = new Modal();
  modal.show({
    title: 'Add Comment',
    content: '<p>Comment functionality is not yet implemented.</p>',
    buttons: [{ text: 'Close', action: 'close', primary: true }]
  });
}

// Generic showModal function for convenience
export function showModal(title, content, options = {}) {
  const modal = new Modal();
  const modalElement = modal.show({
    title,
    content,
    buttons: options.buttons || [{ text: 'Close', action: 'close', primary: true }],
    size: options.size || 'md',
    closable: options.closable !== false
  });
  return modalElement;
}

// Expose to window for global access
window.showWelcomeModal = showWelcomeModal;
window.showSettingsModal = showSettingsModal;
window.showApiKeyModal = showApiKeyModal;