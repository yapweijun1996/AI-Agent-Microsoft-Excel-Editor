'use strict';

/**
 * Global Bindings Manager
 * Provides a clean way to expose necessary functions globally while minimizing namespace pollution
 */

// Centralized registry of global functions
const globalRegistry = new Map();

/**
 * Register a function to be exposed globally
 * @param {string} name - The global name
 * @param {Function} fn - The function to expose
 * @param {Object} options - Options for registration
 */
export function registerGlobal(name, fn, { deprecated = false, description = '' } = {}) {
  if (deprecated) {
    console.warn(`⚠️  Global function '${name}' is deprecated. ${description}`);
  }
  
  globalRegistry.set(name, { fn, deprecated, description });
  window[name] = fn;
}

/**
 * Unregister a global function
 * @param {string} name - The global name to remove
 */
export function unregisterGlobal(name) {
  if (globalRegistry.has(name)) {
    globalRegistry.delete(name);
    delete window[name];
  }
}

/**
 * Get information about registered globals
 */
export function getGlobalInfo() {
  const info = {};
  for (const [name, data] of globalRegistry.entries()) {
    info[name] = {
      deprecated: data.deprecated,
      description: data.description,
      exists: typeof window[name] === 'function'
    };
  }
  return info;
}

/**
 * Clean up all registered globals (for testing or cleanup)
 */
export function cleanupGlobals() {
  for (const name of globalRegistry.keys()) {
    delete window[name];
  }
  globalRegistry.clear();
}

/**
 * Create a namespace object instead of polluting window directly
 * @param {string} namespace - The namespace name (e.g., 'SpreadsheetApp')
 * @param {Object} methods - Object containing methods to expose
 */
export function createNamespace(namespace, methods) {
  if (!window[namespace]) {
    window[namespace] = {};
  }
  
  Object.assign(window[namespace], methods);
  
  // Track the namespace
  globalRegistry.set(namespace, { 
    fn: window[namespace], 
    deprecated: false, 
    description: `Namespace containing: ${Object.keys(methods).join(', ')}`,
    isNamespace: true
  });
}

// Initialize core namespace for essential functions
export function initCoreNamespace() {
  createNamespace('SpreadsheetApp', {
    // Essential functions that need global access
    version: '1.0.0',
    debug: {
      getGlobalInfo,
      cleanupGlobals
    }
  });
}

// Auto-initialize when module is loaded
initCoreNamespace();