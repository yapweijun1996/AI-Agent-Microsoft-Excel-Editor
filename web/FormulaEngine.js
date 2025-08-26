// ===================================================================
// FormulaEngine.js - Advanced Formula Parsing and Execution Engine
// ===================================================================
/* global formulaParser */

class FormulaEngine {
  constructor(data, activeSheetName = null) {
    this.data = data; // Spreadsheet data
    this.activeSheetName = activeSheetName;
    
    // Check if formulaParser is available
    if (typeof formulaParser === 'undefined') {
      console.error('Formula parser library not loaded. Please ensure hot-formula-parser is included.');
      this.parser = null;
      return;
    }
    
    this.parser = new formulaParser.Parser();
    this.init();
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

  execute(formula, data, activeSheetName = null) {
    // Normalize: strip leading '=' for parser or fallback evaluation
    const expr = (typeof formula === 'string' && formula.startsWith('=')) ? formula.slice(1) : formula;

    // Return evaluated value or plain text if parser is not available
    if (!this.parser) {
      // Safe fallback: evaluate simple arithmetic (digits, + - * / . parentheses, whitespace)
      if (typeof expr === 'string' && /^[0-9+\-*/().\s]+$/.test(expr)) {
        try {
          const val = Function('"use strict";return (' + expr + ')')();
          return Number.isFinite(val) ? val : expr;
        } catch {
          return expr;
        }
      }
      return expr;
    }
    
    this.data = data;
    if (activeSheetName) {
      this.activeSheetName = activeSheetName;
    }
    
    try {
      const result = this.parser.parse(typeof expr === 'string' ? expr : String(expr));
      if (result.error) {
        return { error: result.error, details: "Error from parser" };
      }
      return result.result;
    } catch (error) {
      return { error: "#ERROR!", details: error.message };
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
      try {
        const result = this.execute('=' + cell.f, this.data, sheetName);
        if (result && result.error) {
          return { error: result.error, details: `Error in cell ${cellCoord.label}: ${result.details}` };
        }
        return result;
      } catch (error) {
        return { error: "#ERROR!", details: `Exception in ${cellCoord.label}: ${error.message}` };
      }
    }
    
    return cell.v;
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
  if (!globalFormulaEngine || globalFormulaEngine.data !== data) {
    globalFormulaEngine = new FormulaEngine(data, activeSheetName);
  } else if (activeSheetName && globalFormulaEngine.activeSheetName !== activeSheetName) {
    globalFormulaEngine.activeSheetName = activeSheetName;
  }
  return globalFormulaEngine;
}