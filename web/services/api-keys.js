'use strict';

import { AppState, STORAGE_KEYS } from '../core/state.js';
import { log } from '../utils/index.js';

export function saveApiKey(provider, key, persist = false) {
  log(`Saving API key for ${provider}, persist: ${persist}`);

  if (provider === 'openai') {
    AppState.keys.openai = key;
    log('OpenAI key saved to memory');
  }
  if (provider === 'gemini') {
    AppState.keys.gemini = key;
    log('Gemini key saved to memory');
  }

  // Save metadata
  const meta = { openai: !!AppState.keys.openai, gemini: !!AppState.keys.gemini };
  localStorage.setItem(STORAGE_KEYS.keysMeta, JSON.stringify(meta));
  log('API key metadata saved:', meta);

  // Persist the actual key if requested
  if (persist) {
    localStorage.setItem('xlsx_ai_key_' + provider, key);
    log(`${provider} key persisted to localStorage`);
  }

  // Update UI to reflect current provider
  updateProviderStatus();
}

export function restoreApiKeys() {
  const meta = JSON.parse(localStorage.getItem(STORAGE_KEYS.keysMeta) || '{}');
  if (meta.openai) { const k = localStorage.getItem('xlsx_ai_key_openai'); if (k) AppState.keys.openai = k; }
  if (meta.gemini) { const k = localStorage.getItem('xlsx_ai_key_gemini'); if (k) AppState.keys.gemini = k; }
  updateProviderStatus();
}

export function updateProviderStatus() {
  const openaiBtn = document.getElementById('openai-key-btn');
  const geminiBtn = document.getElementById('gemini-key-btn');

  if (openaiBtn) {
    if (AppState.keys.openai) {
      openaiBtn.textContent = '✓ OpenAI Ready';
      openaiBtn.classList.remove('bg-blue-500', 'hover:bg-blue-600');
      openaiBtn.classList.add('bg-green-500', 'hover:bg-green-600');
    } else {
      openaiBtn.textContent = 'Set OpenAI Key';
      openaiBtn.classList.remove('bg-green-500', 'hover:bg-green-600');
      openaiBtn.classList.add('bg-blue-500', 'hover:bg-blue-600');
    }
  }

  if (geminiBtn) {
    if (AppState.keys.gemini) {
      geminiBtn.textContent = '✓ Gemini Ready';
      geminiBtn.classList.remove('bg-green-500', 'hover:bg-green-600');
      geminiBtn.classList.add('bg-green-600', 'hover:bg-green-700');
    } else {
      geminiBtn.textContent = 'Set Gemini Key';
      geminiBtn.classList.remove('bg-green-600', 'hover:bg-green-700');
      geminiBtn.classList.add('bg-green-500', 'hover:bg-green-600');
    }
  }
}

export function pickProvider() {
  // If a specific model is selected, use that provider
  if (AppState.selectedModel !== 'auto') {
    const [provider] = AppState.selectedModel.split(':');
    if (provider === 'openai' && AppState.keys.openai) return 'openai';
    if (provider === 'gemini' && AppState.keys.gemini) return 'gemini';
  }

  // Auto selection - prefer OpenAI if available
  if (AppState.keys.openai) return 'openai';
  if (AppState.keys.gemini) return 'gemini';
  return 'mock';
}

export function getSelectedModel() {
  if (AppState.selectedModel !== 'auto') {
    const [, model] = AppState.selectedModel.split(':');
    // Return the actual API model name (no mapping needed since we're using correct names)
    return model;
  }

  // Default models for auto selection
  const provider = pickProvider();
  if (provider === 'openai') return 'gpt-4o';
  if (provider === 'gemini') return 'gemini-2.5-flash';
  return null;
}