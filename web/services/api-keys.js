'use strict';

import { AppState, STORAGE_KEYS } from '../core/state.js';
import { log } from '../utils/index.js';

// Simple XOR encryption for basic obfuscation (not cryptographically secure)
// For production, consider using Web Crypto API with proper encryption
const OBFUSCATION_KEY = 'xlsx_ai_secure_2024';

function obfuscateKey(key) {
  let result = '';
  for (let i = 0; i < key.length; i++) {
    result += String.fromCharCode(
      key.charCodeAt(i) ^ OBFUSCATION_KEY.charCodeAt(i % OBFUSCATION_KEY.length)
    );
  }
  return btoa(result); // Base64 encode
}

function deobfuscateKey(obfuscated) {
  try {
    const decoded = atob(obfuscated);
    let result = '';
    for (let i = 0; i < decoded.length; i++) {
      result += String.fromCharCode(
        decoded.charCodeAt(i) ^ OBFUSCATION_KEY.charCodeAt(i % OBFUSCATION_KEY.length)
      );
    }
    return result;
  } catch (e) {
    log('Failed to deobfuscate key:', e.message);
    return null;
  }
}

export function saveApiKey(provider, key, persist = false) {
  log(`Saving API key for ${provider}, persist: ${persist}`);

  // Validate API key format
  if (!validateApiKey(provider, key)) {
    throw new Error(`Invalid API key format for ${provider}`);
  }

  // Store in memory
  if (provider === 'openai') {
    AppState.keys.openai = key;
    log('OpenAI key saved to memory');
  }
  if (provider === 'gemini') {
    AppState.keys.gemini = key;
    log('Gemini key saved to memory');
  }

  // Save metadata only (not the actual keys)
  const meta = { 
    openai: !!AppState.keys.openai, 
    gemini: !!AppState.keys.gemini,
    timestamp: Date.now()
  };
  localStorage.setItem(STORAGE_KEYS.keysMeta, JSON.stringify(meta));
  log('API key metadata saved:', meta);

  // Persist with obfuscation if requested (still not recommended for production)
  if (persist) {
    const obfuscated = obfuscateKey(key);
    sessionStorage.setItem(`xlsx_ai_key_${provider}`, obfuscated);
    log(`${provider} key persisted to sessionStorage with obfuscation`);
    
    // Show security warning
    console.warn(`⚠️  API keys are stored in browser storage. For production use, consider:
    1. Server-side proxy to protect API keys
    2. Environment variables
    3. Secure key management service
    4. Short-lived tokens with refresh mechanism`);
  }

  updateProviderStatus();
}

function validateApiKey(provider, key) {
  if (!key || typeof key !== 'string') return false;
  
  // Basic validation patterns
  switch (provider) {
    case 'openai':
      return key.startsWith('sk-') && key.length > 20;
    case 'gemini':
      return key.length > 20 && /^[A-Za-z0-9_-]+$/.test(key);
    default:
      return false;
  }
}

export function restoreApiKeys() {
  try {
    const meta = JSON.parse(localStorage.getItem(STORAGE_KEYS.keysMeta) || '{}');
    
    // Check if keys are too old (security measure)
    const MAX_KEY_AGE = 7 * 24 * 60 * 60 * 1000; // 7 days
    if (meta.timestamp && (Date.now() - meta.timestamp) > MAX_KEY_AGE) {
      log('Stored API keys expired, clearing...');
      clearStoredKeys();
      return;
    }
    
    if (meta.openai) { 
      const k = sessionStorage.getItem('xlsx_ai_key_openai'); 
      if (k) {
        const deobfuscated = deobfuscateKey(k);
        if (deobfuscated && validateApiKey('openai', deobfuscated)) {
          AppState.keys.openai = deobfuscated;
        }
      }
    }
    
    if (meta.gemini) { 
      const k = sessionStorage.getItem('xlsx_ai_key_gemini'); 
      if (k) {
        const deobfuscated = deobfuscateKey(k);
        if (deobfuscated && validateApiKey('gemini', deobfuscated)) {
          AppState.keys.gemini = deobfuscated;
        }
      }
    }
    
    updateProviderStatus();
  } catch (e) {
    log('Error restoring API keys:', e.message);
    clearStoredKeys();
  }
}

export function clearStoredKeys() {
  // Clear from both session and local storage
  sessionStorage.removeItem('xlsx_ai_key_openai');
  sessionStorage.removeItem('xlsx_ai_key_gemini');
  localStorage.removeItem('xlsx_ai_key_openai'); // Clean up old format
  localStorage.removeItem('xlsx_ai_key_gemini'); // Clean up old format
  localStorage.removeItem(STORAGE_KEYS.keysMeta);
  
  // Clear from memory
  AppState.keys.openai = null;
  AppState.keys.gemini = null;
  
  updateProviderStatus();
  log('All API keys cleared from storage and memory');
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