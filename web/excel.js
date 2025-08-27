class Excel {
    constructor() {
        this.hot = null; // Handsontable instance
        this.workbook = {
            'Sheet1': Handsontable.helper.createEmptySpreadsheetData(100, 26),
            'Sheet2': Handsontable.helper.createEmptySpreadsheetData(100, 26),
            'Sheet3': Handsontable.helper.createEmptySpreadsheetData(100, 26),
        };
        this.activeSheet = 'Sheet1';
        this.init();
    }

    init() {
        console.log('🚀 Excel initializing...');
        const container = document.getElementById('grid-container');

        // Load sample data into Sheet1
        this.workbook['Sheet1'] = [
            ['Product', 'Quantity', 'Price', 'Total'],
            ['Laptop', 2, 999.99, '=B2*C2'],
            ['Mouse', 5, 29.99, '=B3*C3'],
            ['Keyboard', 3, 79.99, '=B4*C4'],
            [],
            ['TOTAL:', null, null, '=SUM(D2:D4)']
        ];

        this.hot = new Handsontable(container, {
            data: this.workbook[this.activeSheet],
            rowHeaders: true,
            colHeaders: true,
            height: 'auto',
            licenseKey: 'non-commercial-and-evaluation',
            formulas: {
                engine: HyperFormula
            },
            afterSelection: (r, c, r2, c2) => {
                this.updateFormulaBar(r, c);
            }
        });

        this.setupEvents();
        console.log('✅ Excel loaded');
    }

    updateFormulaBar(row, col) {
        const cellRef = this.hot.getCell(row, col);
        const cellValue = this.hot.getDataAtCell(row, col);
        const formula = this.hot.getCellMeta(row, col).formulaValue;

        document.getElementById('cell-ref').textContent = this.hot.getColHeader(col) + (row + 1);
        document.getElementById('formula-input').value = formula || cellValue || '';
    }

    setupEvents() {
        // Toolbar buttons
        document.getElementById('new-btn').addEventListener('click', () => this.newWorkbook());
        document.getElementById('save-btn').addEventListener('click', () => this.save());
        document.getElementById('load-btn').addEventListener('click', () => this.load());
        document.getElementById('export-btn').addEventListener('click', () => this.export());

        // Formula bar editing
        const formulaInput = document.getElementById('formula-input');
        formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                const selected = this.hot.getSelected();
                if (selected) {
                    const [startRow, startCol] = selected[0];
                    this.hot.setDataAtCell(startRow, startCol, formulaInput.value);
                }
            }
        });

        // Sheet tabs
        document.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                this.switchSheet(tab.dataset.sheet);
            });
        });
    }

    switchSheet(sheetName) {
        // Save current sheet data
        this.workbook[this.activeSheet] = this.hot.getData();

        this.activeSheet = sheetName;
        this.hot.loadData(this.workbook[this.activeSheet]);

        // Update active tab UI
        document.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.sheet === sheetName);
        });
    }

    newWorkbook() {
        if (confirm('Create new workbook? This will clear all data.')) {
            this.workbook = {
                'Sheet1': Handsontable.helper.createEmptySpreadsheetData(100, 26),
                'Sheet2': Handsontable.helper.createEmptySpreadsheetData(100, 26),
                'Sheet3': Handsontable.helper.createEmptySpreadsheetData(100, 26),
            };
            this.switchSheet('Sheet1');
        }
    }

    save() {
        // Save current sheet data before saving workbook
        this.workbook[this.activeSheet] = this.hot.getData();
        const json = JSON.stringify(this.workbook, null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `excel-workbook-${new Date().toISOString().split('T')[0]}.json`;
        a.click();
        URL.revokeObjectURL(url);
    }

    load() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.json';

        input.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    this.workbook = JSON.parse(e.target.result);
                    this.switchSheet(Object.keys(this.workbook)[0]); // Switch to the first sheet
                } catch (err) {
                    alert('Invalid file format.');
                }
            };
            reader.readAsText(file);
        });
        input.click();
    }

    export() {
        const exportPlugin = this.hot.getPlugin('exportFile');
        exportPlugin.downloadFile('csv', { filename: `${this.activeSheet}-export` });
    }
}

// START EXCEL
let excel = null;
document.addEventListener('DOMContentLoaded', () => {
    excel = new Excel();
});