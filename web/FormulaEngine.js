/* AI Excel Editor - Formula Engine */
'use strict';

// Using hot-formula-parser
// We will need to add this to our HTML file.
// <script src="https://cdn.jsdelivr.net/npm/hot-formula-parser@4.0.0/dist/formula-parser.min.js"></script>

const FormulaEngine = {
  parser: null,

  init(data) {
    this.parser = new formulaParser.FormulaParser();
    this._registerCustomFunctions();

    // Set up data for the parser
    this.parser.on('callCellValue', (cellCoord, done) => {
      const cellAddress = `${cellCoord.label}`;
      const sheet = data.Sheets[data.activeSheet];
      if (sheet[cellAddress]) {
        done(sheet[cellAddress].v);
      } else {
        done('');
      }
    });

    this.parser.on('callRangeValue', (startCellCoord, endCellCoord, done) => {
        const fragment = [];
        for (let row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
            const rowData = [];
            for (let col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
                const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
                const sheet = data.Sheets[data.activeSheet];
                if (sheet[cellAddress]) {
                    rowData.push(sheet[cellAddress].v);
                } else {
                    rowData.push('');
                }
            }
            fragment.push(rowData);
        }
        done(fragment);
    });
  },

  execute(formula, data) {
    if (!formula || !formula.startsWith('=')) {
      return formula;
    }

    this.init(data);

    try {
      const result = this.parser.parse(formula.substring(1));
      if (result.error) {
        return result.error;
      }
      return result.result;
    } catch (error) {
      console.error("Formula evaluation error:", error);
      return "#ERROR!";
    }
  },

  // TODO: Implement the rest of the 400+ Excel functions.
  // This is a sample of how to add custom functions.
  _registerCustomFunctions() {
    this.parser.on('VLOOKUP', (params, done) => {
        // VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
        const [lookupValue, tableArray, colIndexNum, rangeLookup] = params;
        // This is a simplified implementation. A full implementation would be more complex.
        for (let i = 0; i < tableArray.length; i++) {
            if (tableArray[i][0] == lookupValue) {
                done(tableArray[i][colIndexNum - 1]);
                return;
            }
        }
        done('#N/A');
    });

    this.parser.on('INDEX', (params, done) => {
        // INDEX(array, row_num, [column_num])
        const [array, rowNum, columnNum] = params;
        if (array[rowNum - 1] && array[rowNum - 1][columnNum - 1]) {
            done(array[rowNum - 1][columnNum - 1]);
        } else {
            done('#REF!');
        }
    });

    this.parser.on('MATCH', (params, done) => {
        // MATCH(lookup_value, lookup_array, [match_type])
        const [lookupValue, lookupArray, matchType] = params;
        for (let i = 0; i < lookupArray.length; i++) {
            if (lookupArray[i][0] == lookupValue) {
                done(i + 1);
                return;
            }
        }
        done('#N/A');
    });
  }
};