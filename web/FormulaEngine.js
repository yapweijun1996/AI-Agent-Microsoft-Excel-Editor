// ===================================================================
// FormulaEngine.js - Advanced Formula Parsing and Execution Engine
// ===================================================================
/* global formulaParser, AppState */

class FormulaEngine {
  constructor(data, activeSheetName = null) {
    this.data = data; // Spreadsheet data
    this.activeSheetName = activeSheetName;
    
    // Formula result cache for performance
    this.formulaCache = new Map();
    this.cacheVersion = 0; // Increment when data changes
    this.maxCacheSize = 1000; // Prevent memory leaks
    this.cacheHitCount = 0;
    this.cacheMissCount = 0;
    
    // Circular reference detection
    this.evaluationStack = new Set(); // Track cells being evaluated
    this.maxStackDepth = 100; // Prevent infinite recursion
    
    // Check if formulaParser is available
    if (typeof formulaParser === 'undefined') {
      console.error('Formula parser library not loaded. Please ensure hot-formula-parser is included.');
      this.parser = null;
      return;
    }
    
    try {
      this.parser = new formulaParser.Parser();
      this.init();
      console.log('FormulaEngine initialized successfully');
    } catch (error) {
      console.error('Error initializing FormulaEngine:', error);
      this.parser = null;
    }
  }

  init() {
    // Register custom functions or event handlers
    this.parser.on('callCellValue', (cellCoord, done) => {
      const value = this.getCellValue(cellCoord);
      done(value);
    });

    this.parser.on('callRangeValue', (startCellCoord, endCellCoord, done) => {
      const values = this.getRangeValues(startCellCoord, endCellCoord);
      done(values);
    });
  }

  execute(formula, data, activeSheetName = null, cellAddress = null) {
    // Normalize: strip leading '=' for parser or fallback evaluation
    // Update engine context before cache lookup
    this.data = data;
    if (activeSheetName) {
      this.activeSheetName = activeSheetName;
    }

    // Normalize: strip leading '=' for parser or fallback evaluation
    const expr = (typeof formula === 'string' && formula.startsWith('=')) ? formula.slice(1) : formula;

    // Create a robust cache key that includes workbook version and cell context
    const wbVersion = (typeof AppState !== 'undefined' && AppState.wbVersion) ? AppState.wbVersion : 0;
    const cacheKey = `${this.cacheVersion}:${activeSheetName || this.activeSheetName}:${cellAddress || ''}:${wbVersion}:${expr}`;
    
    // Check cache first
    if (this.formulaCache.has(cacheKey)) {
      this.cacheHitCount++;
      return this.formulaCache.get(cacheKey);
    }
    
    this.cacheMissCount++;

    // Return evaluated value or plain text if parser is not available
    if (!this.parser) {
      console.log('Parser not available, using fallback for formula:', expr);
      // Safe fallback: evaluate simple arithmetic (digits, + - * / . parentheses, whitespace)
      if (typeof expr === 'string' && /^[0-9+\-*/().\s]+$/.test(expr)) {
        try {
          const val = Function('"use strict";return (' + expr + ')')();
          const result = Number.isFinite(val) ? val : expr;
          console.log('Fallback evaluation result:', result);
          this.setCacheValue(cacheKey, result);
          return result;
        } catch (error) {
          console.log('Fallback evaluation failed:', error);
          this.setCacheValue(cacheKey, expr);
          return expr;
        }
      }
      console.log('Formula does not match simple arithmetic pattern:', expr);
      this.setCacheValue(cacheKey, expr);
      return expr;
    }
    
    try {
      console.log('Parsing formula with parser:', expr);
      const result = this.parser.parse(typeof expr === 'string' ? expr : String(expr));
      console.log('Parser result:', result);
      let finalResult;
      
      if (result.error) {
        console.log('Parser returned error:', result.error);
        finalResult = { error: result.error, details: "Error from parser" };
      } else {
        console.log('Parser successful, result:', result.result);
        finalResult = result.result;
      }
      
      this.setCacheValue(cacheKey, finalResult);
      return finalResult;
    } catch (error) {
      console.log('Parser threw exception:', error);
      const errorResult = { error: "#ERROR!", details: error.message };
      this.setCacheValue(cacheKey, errorResult);
      return errorResult;
    }
  }

  getCellValue(cellCoord) {
    // Use active sheet if specified, otherwise default to first sheet
    const sheetName = this.activeSheetName || this.data.SheetNames[0];
    const sheet = this.data.Sheets[sheetName];
    if (!sheet) return { error: "#REF!", details: `Sheet "${sheetName}" not found.` };

    const cell = sheet[cellCoord.label];
    if (!cell) return null;

    // Handle formula cells recursively
    if (cell.f) {
      const cellKey = `${sheetName}!${cellCoord.label}`;
      
      // Check for circular reference
      if (this.evaluationStack.has(cellKey)) {
        return { error: "#CIRCULAR!", details: `Circular reference detected in ${cellCoord.label}` };
      }
      
      // Check stack depth to prevent infinite recursion
      if (this.evaluationStack.size > this.maxStackDepth) {
        return { error: "#DEPTH!", details: `Maximum evaluation depth exceeded at ${cellCoord.label}` };
      }
      
      // Add to evaluation stack
      this.evaluationStack.add(cellKey);
      
      try {
        const result = this.execute('=' + cell.f, this.data, sheetName, cellCoord.label);
        if (result && result.error) {
          return { error: result.error, details: `Error in cell ${cellCoord.label}: ${result.details}` };
        }
        return result;
      } catch (error) {
        return { error: "#ERROR!", details: `Exception in ${cellCoord.label}: ${error.message}` };
      } finally {
        // Always remove from evaluation stack
        this.evaluationStack.delete(cellKey);
      }
    }
    
    return cell.v;
  }

  /**
   * Set value in cache with size management
   */
  setCacheValue(key, value) {
    // Prevent cache from growing too large
    if (this.formulaCache.size >= this.maxCacheSize) {
      // Remove oldest entries (simple FIFO)
      const keysToDelete = Array.from(this.formulaCache.keys()).slice(0, Math.floor(this.maxCacheSize * 0.3));
      keysToDelete.forEach(oldKey => this.formulaCache.delete(oldKey));
    }
    
    this.formulaCache.set(key, value);
  }

  /**
   * Invalidate cache when spreadsheet data changes
   */
  invalidateCache() {
    this.cacheVersion++;
    this.formulaCache.clear();
    this.evaluationStack.clear(); // Clear circular reference tracking
    
    // Log cache performance if in debug mode
    if (this.cacheHitCount + this.cacheMissCount > 0) {
      const hitRate = (this.cacheHitCount / (this.cacheHitCount + this.cacheMissCount) * 100).toFixed(1);
      console.log(`Formula cache performance - Hits: ${this.cacheHitCount}, Misses: ${this.cacheMissCount}, Hit rate: ${hitRate}%`);
    }
    
    this.cacheHitCount = 0;
    this.cacheMissCount = 0;
  }

  /**
   * Update data and invalidate cache
   */
  updateData(newData, activeSheetName = null) {
    this.data = newData;
    if (activeSheetName) {
      this.activeSheetName = activeSheetName;
    }
    this.invalidateCache();
  }

  getRangeValues(startCellCoord, endCellCoord) {
    const sheetName = this.activeSheetName || this.data.SheetNames[0];
    const sheet = this.data.Sheets[sheetName];
    if (!sheet) return [];
 
    const startCell = XLSX.utils.decode_cell(startCellCoord.label);
    const endCell = XLSX.utils.decode_cell(endCellCoord.label);
    const matrix = [];
 
    for (let r = startCell.r; r <= endCell.r; r++) {
      const row = [];
      for (let c = startCell.c; c <= endCell.c; c++) {
        const addr = XLSX.utils.encode_cell({r, c});
        const cell = sheet[addr];
        let value = 0;
        
        if (cell) {
          if (cell.f) {
            try {
              const result = this.execute('=' + cell.f, this.data, sheetName);
              value = (result && typeof result === 'object' && result.error) ? 0 : (typeof result === 'number' ? result : 0);
            } catch {
              value = 0;
            }
          } else {
            value = typeof cell.v === 'number' ? cell.v : 0;
          }
        }
        
        row.push(value);
      }
      matrix.push(row);
    }
    
    return matrix;
  }
}

// Singleton instance for performance optimization
let globalFormulaEngine = null;

function getFormulaEngine(data, activeSheetName = null) {
  if (!globalFormulaEngine) {
    globalFormulaEngine = new FormulaEngine(data, activeSheetName);
  } else if (globalFormulaEngine.data !== data) {
    // Data has changed, update and invalidate cache
    globalFormulaEngine.updateData(data, activeSheetName);
  } else if (activeSheetName && globalFormulaEngine.activeSheetName !== activeSheetName) {
    globalFormulaEngine.activeSheetName = activeSheetName;
  }
  return globalFormulaEngine;
}

// Export to global scope for use in other scripts
window.getFormulaEngine = getFormulaEngine;

// Immediate test when script loads
console.log('FormulaEngine.js loaded');
console.log('formulaParser available at load time:', typeof formulaParser !== 'undefined');

// Test the formula engine on load
document.addEventListener('DOMContentLoaded', function() {
  console.log('Testing FormulaEngine on DOM load...');
  console.log('formulaParser available after DOM load:', typeof formulaParser !== 'undefined');
  
  if (typeof formulaParser !== 'undefined') {
    try {
      const testEngine = new FormulaEngine({Sheets: {Sheet1: {}}, SheetNames: ['Sheet1']}, 'Sheet1');
      const testResult = testEngine.execute('=1+1', {Sheets: {Sheet1: {}}, SheetNames: ['Sheet1']}, 'Sheet1');
      console.log('FormulaEngine test result for 1+1:', testResult);
    } catch (error) {
      console.error('FormulaEngine test failed:', error);
    }
  } else {
    console.error('formulaParser is not available - trying fallback arithmetic');
    try {
      // Test direct arithmetic evaluation
      const result = Function('"use strict";return (1+1)')();
      console.log('Direct arithmetic evaluation result:', result);
    } catch (error) {
      console.error('Direct arithmetic evaluation failed:', error);
    }
  }
});

// Also add a window load event in case DOM load is too early
window.addEventListener('load', function() {
  console.log('Window loaded, formulaParser available:', typeof formulaParser !== 'undefined');
});