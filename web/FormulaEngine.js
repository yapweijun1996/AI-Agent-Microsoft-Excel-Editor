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
    // Return simple value if parser is not available
    if (!this.parser) {
      // Try to evaluate simple expressions or return the formula as text
      if (formula.startsWith('=')) {
        return formula.substring(1); // Return without '=' prefix
      }
      return formula;
    }
    
    this.data = data;
    if (activeSheetName) {
      this.activeSheetName = activeSheetName;
    }
    
    try {
      const result = this.parser.parse(formula);
      return result.error ? { error: result.error } : result.result;
    } catch (error) {
      return { error: error.message };
    }
  }

  getCellValue(cellCoord) {
    // Use active sheet if specified, otherwise default to first sheet
    const sheetName = this.activeSheetName || this.data.SheetNames[0];
    const sheet = this.data.Sheets[sheetName];
    if (!sheet) return undefined;

    const cell = sheet[cellCoord.label];
    if (!cell) return undefined;

    // Handle formula cells recursively
    if (cell.f) {
      try {
        const result = this.execute('=' + cell.f, this.data, sheetName);
        return result.error ? 0 : result;
      } catch (error) {
        return 0;
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
    const values = [];

    for (let r = startCell.r; r <= endCell.r; r++) {
      for (let c = startCell.c; c <= endCell.c; c++) {
        const addr = XLSX.utils.encode_cell({r, c});
        const cell = sheet[addr];
        let value = 0;
        
        if (cell) {
          if (cell.f) {
            try {
              const result = this.execute('=' + cell.f, this.data, sheetName);
              value = result.error ? 0 : (typeof result === 'number' ? result : 0);
            } catch (error) {
              value = 0;
            }
          } else {
            value = typeof cell.v === 'number' ? cell.v : 0;
          }
        }
        
        values.push(value);
      }
    }
    
    return values;
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