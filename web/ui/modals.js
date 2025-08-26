'use strict';

import { Modal } from './modal.js';
import { log } from '../utils/index.js';
import { saveApiKey } from '../services/api-keys.js';
import { showToast } from './toast.js';

export function showApiKeyModal(provider) {
  log(`Opening API key modal for ${provider}`);
  const modal = new Modal();
  modal.show({
    title: `Set ${provider} API Key`,
    content: `
      <div class="space-y-4">
        <p class="text-sm text-gray-600">Enter your ${provider} API key. It will be stored in memory; toggle persistence if desired.</p>
        <div class="space-y-2">
          <input type="password" id="api-key-input" placeholder="Enter API key..." class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" />
          <p class="text-xs text-gray-500">Get your API key from: ${provider === 'OpenAI' ? 'https://platform.openai.com/api-keys' : 'https://aistudio.google.com/app/apikey'}</p>
        </div>
        <label class="flex items-center space-x-2 text-xs text-gray-600">
          <input id="persist-key" type="checkbox" class="rounded border-gray-300">
          <span>Persist to localStorage (less secure)</span>
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
    ]
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
    content: '<p>Sort functionality is not yet implemented.</p>',
    buttons: [{ text: 'Close', action: 'close', primary: true }]
  });
}

export function showChartModal() {
  const modal = new Modal();
  modal.show({
    title: 'Create Chart',
    content: '<p>Chart functionality is not yet implemented.</p>',
    buttons: [{ text: 'Close', action: 'close', primary: true }]
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