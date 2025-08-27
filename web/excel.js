// BULLETPROOF EXCEL
// IMPOSSIBLE TO BREAK

class Excel {
    constructor() {
        this.workbook = {
            Sheet1: {},
            Sheet2: {},
            Sheet3: {}
        };
        this.activeSheet = 'Sheet1';
        this.currentCell = 'A1';
        this.rows = 200;         // More rows
        this.cols = 26;          // A-Z  
        this.isLoaded = false;
        this.selection = {
            start: null,
            end: null,
            type: 'cell' // 'cell', 'row', 'col'
        };
        
        // Store event handlers for cleanup
        this.eventHandlers = [];
        
        this.init();
    }
    
    init() {
        console.log('🚀 Excel initializing...');
        
        // Generate the grid
        this.generateGrid();
        
        // Set up events
        this.setupEvents();
        
        // Load sample data
        this.loadSampleData();
        
        // Hide loading
        document.getElementById('loading').style.display = 'none';
        
        this.updateCellCount();
        this.isLoaded = true;
        console.log('✅ Excel loaded');
    }
    
    generateGrid() {
        try {
            const container = document.getElementById('grid-container');
            if (!container) {
                console.error('Grid container not found');
                return;
            }
            
            let html = '<table class="excel-table">';
        
        // HEADER ROW
        html += '<tr>';
        html += '<th></th>'; // Corner
        for (let c = 0; c < this.cols; c++) {
            const colLetter = this.getColumnLetter(c);
            html += `<th data-col="${c}" title="Column ${colLetter}">${colLetter}</th>`;
        }
        html += '</tr>';
        
        // DATA ROWS
        for (let r = 0; r < this.rows; r++) {
            html += '<tr>';
            html += `<td data-row="${r}" title="Row ${r + 1}">${r + 1}</td>`; // Row header
            
            for (let c = 0; c < this.cols; c++) {
                const addr = this.getCellAddress(r, c);
                const value = this.getCurrentSheetData()[addr] || '';
                html += `<td data-addr="${addr}">`;
                html += `<input type="text" class="cell-input" data-addr="${addr}" value="${this.escapeHtml(value)}" spellcheck="false">`;
                html += '</td>';
            }
            html += '</tr>';
        }
        
        html += '</table>';
        
        // ATOMIC REPLACEMENT
        const loading = document.getElementById('loading');
        container.innerHTML = html;
        container.appendChild(loading);
        
            console.log(`📊 Generated ${this.rows}×${this.cols} grid (${this.rows * this.cols} cells)`);
        } catch (error) {
            console.error('Error generating grid:', error);
            this.updateStatus('Error generating grid');
        }
    }
    
    setupEvents() {
        // Cell focus/blur
        document.addEventListener('focusin', (e) => {
            if (e.target.classList.contains('cell-input')) {
                this.handleCellFocus(e.target);
            }
        });
        
        document.addEventListener('focusout', (e) => {
            if (e.target.classList.contains('cell-input')) {
                this.handleCellBlur(e.target);
            }
        });
        
        // Keyboard navigation
        document.addEventListener('keydown', (e) => {
            if (e.target.classList.contains('cell-input')) {
                this.handleCellKeydown(e);
            }
        });
        
        // Formula bar
        const formulaInput = document.getElementById('formula-input');
        formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                this.handleFormulaEnter();
            }
        });
        
        formulaInput.addEventListener('input', () => {
            // Real-time formula preview
            const currentInput = document.querySelector(`[data-addr="${this.currentCell}"]`);
            if (currentInput) {
                currentInput.value = formulaInput.value;
            }
        });
        
        // Sheet tabs
        document.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                this.switchSheet(tab.dataset.sheet);
            });
        });
        
        // Toolbar - use event delegation for reliability and dynamic elements
        const toolbar = document.querySelector('.toolbar');
        if (toolbar) {
            const toolbarClickHandler = (e) => {
                const btn = e.target.closest('button');
                if (!btn) return;
                switch (btn.id) {
                    case 'new-btn':
                        this.newWorkbook();
                        break;
                    case 'save-btn':
                        this.save();
                        break;
                    case 'load-btn':
                        this.load();
                        break;
                    case 'export-btn':
                        this.export();
                        break;
                }
            };
            toolbar.addEventListener('click', toolbarClickHandler);
            // store handler for cleanup to avoid leaks and allow re-initialization
            this.eventHandlers.push({ element: toolbar, event: 'click', handler: toolbarClickHandler });
        }
        
        // Column/Row selection
        document.addEventListener('click', (e) => {
            if (e.target.dataset.col !== undefined) {
                this.selectColumn(parseInt(e.target.dataset.col));
            } else if (e.target.dataset.row !== undefined) {
                this.selectRow(parseInt(e.target.dataset.row));
            }
        });
        
        // Context menu
        document.addEventListener('contextmenu', (e) => {
            if (e.target.classList.contains('cell-input')) {
                e.preventDefault();
                this.showContextMenu(e.pageX, e.pageY, e.target.dataset.addr);
            }
        });
        
        // Click outside to hide context menu
        document.addEventListener('click', () => {
            this.hideContextMenu();
        });
        
        console.log('⚡ Events set up');
    }
    
    handleCellFocus(input) {
        const addr = input.dataset.addr;
        this.currentCell = addr;
        
        // Clear previous selection
        this.clearSelection();
        
        // Update UI
        document.getElementById('cell-ref').textContent = addr;
        document.getElementById('formula-input').value = this.getCurrentSheetData()[addr] || '';
        
        // Highlight cell
        input.parentElement.classList.add('cell-selected');
        
        // Show cell value/formula in input
        const value = this.getCurrentSheetData()[addr] || '';
        if (value.startsWith('=')) {
            input.value = value; // Show formula
        }
        
        this.updateStatus(`Selected cell ${addr}`);
    }
    
    handleCellBlur(input) {
        const addr = input.dataset.addr;
        const value = input.value.trim();
        
        // Save value to current sheet
        if (value) {
            this.getCurrentSheetData()[addr] = value;
        } else {
            delete this.getCurrentSheetData()[addr];
        }
        
        // Process formulas (basic)
        if (value.startsWith('=')) {
            const result = this.evaluateFormula(value);
            if (result !== null) {
                input.setAttribute('title', `Formula: ${value}\nResult: ${result}`);
            }
        }
        
        input.parentElement.classList.remove('cell-selected');
        this.updateCellCount();
    }
    
    handleCellKeydown(e) {
        const input = e.target;
        const addr = input.dataset.addr;
        const pos = this.parseAddress(addr);
        
        if (!pos) return;
        
        let newAddr = null;
        
        switch (e.key) {
            case 'ArrowUp':
                if (pos.row > 0) {
                    newAddr = this.getCellAddress(pos.row - 1, pos.col);
                }
                break;
            case 'ArrowDown':
                if (pos.row < this.rows - 1) {
                    newAddr = this.getCellAddress(pos.row + 1, pos.col);
                }
                break;
            case 'ArrowLeft':
                if (pos.col > 0 && input.selectionStart === 0) {
                    newAddr = this.getCellAddress(pos.row, pos.col - 1);
                }
                break;
            case 'ArrowRight':
                if (pos.col < this.cols - 1 && input.selectionStart === input.value.length) {
                    newAddr = this.getCellAddress(pos.row, pos.col + 1);
                }
                break;
            case 'Enter':
                if (pos.row < this.rows - 1) {
                    newAddr = this.getCellAddress(pos.row + 1, pos.col);
                }
                break;
            case 'Tab':
                e.preventDefault();
                if (e.shiftKey) {
                    if (pos.col > 0) {
                        newAddr = this.getCellAddress(pos.row, pos.col - 1);
                    }
                } else {
                    if (pos.col < this.cols - 1) {
                        newAddr = this.getCellAddress(pos.row, pos.col + 1);
                    }
                }
                break;
            case 'Delete':
                input.value = '';
                delete this.getCurrentSheetData()[addr];
                this.updateCellCount();
                break;
        }
        
        if (newAddr && (e.key === 'ArrowUp' || e.key === 'ArrowDown' || 
                       e.key === 'Enter' || e.key === 'Tab' ||
                       (e.key === 'ArrowLeft' && input.selectionStart === 0) ||
                       (e.key === 'ArrowRight' && input.selectionStart === input.value.length))) {
            e.preventDefault();
            const nextInput = document.querySelector(`input[data-addr="${newAddr}"]`);
            if (nextInput) {
                nextInput.focus();
                nextInput.select();
            }
        }
    }
    
    handleFormulaEnter() {
        const formulaInput = document.getElementById('formula-input');
        const value = formulaInput.value;
        const addr = this.currentCell;
        
        if (addr) {
            // Update data
            if (value) {
                this.getCurrentSheetData()[addr] = value;
            } else {
                delete this.getCurrentSheetData()[addr];
            }
            
            // Update cell input
            const cellInput = document.querySelector(`[data-addr="${addr}"]`);
            if (cellInput) {
                cellInput.value = value;
                
                // Process formula
                if (value.startsWith('=')) {
                    const result = this.evaluateFormula(value);
                    if (result !== null) {
                        cellInput.setAttribute('title', `Formula: ${value}\nResult: ${result}`);
                    }
                }
            }
            
            this.updateCellCount();
            this.updateStatus(`Formula entered in ${addr}`);
        }
    }
    
    switchSheet(sheetName) {
        // Save current sheet state
        const currentInputs = document.querySelectorAll('.cell-input');
        currentInputs.forEach(input => {
            const addr = input.dataset.addr;
            const value = input.value.trim();
            if (value) {
                this.getCurrentSheetData()[addr] = value;
            } else {
                delete this.getCurrentSheetData()[addr];
            }
        });
        
        // Switch active sheet
        this.activeSheet = sheetName;
        
        // Create sheet if it doesn't exist
        if (!this.workbook[this.activeSheet]) {
            this.workbook[this.activeSheet] = {};
        }
        
        // Update UI
        document.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.sheet === sheetName);
        });
        
        // Load new sheet data
        currentInputs.forEach(input => {
            const addr = input.dataset.addr;
            const value = this.getCurrentSheetData()[addr] || '';
            input.value = value;
        });
        
        this.updateCellCount();
        this.updateStatus(`Switched to ${sheetName}`);
    }
    
    selectColumn(colIndex) {
        this.clearSelection();
        this.selection = { type: 'col', col: colIndex };
        
        // Highlight column
        document.querySelectorAll(`[data-addr]`).forEach(input => {
            const pos = this.parseAddress(input.dataset.addr);
            if (pos && pos.col === colIndex) {
                input.parentElement.classList.add('col-selected');
            }
        });
        
        // Highlight header
        document.querySelector(`[data-col="${colIndex}"]`).style.background = '#fbbf24';
        
        const colLetter = this.getColumnLetter(colIndex);
        this.updateStatus(`Selected column ${colLetter}`);
    }
    
    selectRow(rowIndex) {
        this.clearSelection();
        this.selection = { type: 'row', row: rowIndex };
        
        // Highlight row
        document.querySelectorAll(`[data-addr]`).forEach(input => {
            const pos = this.parseAddress(input.dataset.addr);
            if (pos && pos.row === rowIndex) {
                input.parentElement.classList.add('row-selected');
            }
        });
        
        // Highlight header
        document.querySelector(`[data-row="${rowIndex}"]`).style.background = '#fbbf24';
        
        this.updateStatus(`Selected row ${rowIndex + 1}`);
    }
    
    clearSelection() {
        document.querySelectorAll('.cell-selected, .row-selected, .col-selected').forEach(el => {
            el.classList.remove('cell-selected', 'row-selected', 'col-selected');
        });
        
        document.querySelectorAll('[data-col], [data-row]').forEach(header => {
            header.style.background = '';
        });
        
        this.selection = { type: 'cell' };
    }
    
    showContextMenu(x, y, addr) {
        this.hideContextMenu();
        
        const menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.innerHTML = `
            <div class="context-menu-item" data-action="copy">Copy</div>
            <div class="context-menu-item" data-action="paste">Paste</div>
            <div class="context-menu-item" data-action="clear">Clear</div>
            <div class="context-menu-item" data-action="format">Format Cell</div>
            <div class="context-menu-item" data-action="insert-row-above">Insert Row Above</div>
            <div class="context-menu-item" data-action="insert-col-left">Insert Column Left</div>
        `;
        
        menu.style.left = x + 'px';
        menu.style.top = y + 'px';
        
        menu.addEventListener('click', (e) => {
            const action = e.target.dataset.action;
            if (action) {
                this.handleContextAction(action, addr);
            }
            this.hideContextMenu();
        });
        
        document.body.appendChild(menu);
        menu.id = 'context-menu';
    }
    
    hideContextMenu() {
        const menu = document.getElementById('context-menu');
        if (menu) {
            menu.remove();
        }
    }
    
    handleContextAction(action, addr) {
        switch (action) {
            case 'copy':
                this.copyCell(addr);
                break;
            case 'paste':
                this.pasteCell(addr);
                break;
            case 'clear':
                this.clearCell(addr);
                break;
            case 'format':
                this.formatCell(addr);
                break;
            case 'insert-row-above':
                const { row } = this.parseAddress(addr);
                this.addRow(row);
                break;
            case 'insert-col-left':
                const { col } = this.parseAddress(addr);
                this.addColumn(col);
                break;
        }
    }
    
    copyCell(addr) {
        const value = this.getCurrentSheetData()[addr] || '';
        if (navigator.clipboard) {
            navigator.clipboard.writeText(value);
        }
        this.updateStatus(`Copied cell ${addr}`);
    }
    
    async pasteCell(addr) {
        if (!navigator.clipboard) {
            this.updateStatus('Clipboard not supported');
            return;
        }
        
        try {
            const text = await navigator.clipboard.readText();
            if (!text) return;
            
            this.getCurrentSheetData()[addr] = text;
            const input = document.querySelector(`[data-addr="${addr}"]`);
            if (input) {
                input.value = text;
            }
            this.updateCellCount();
            this.updateStatus(`Pasted to cell ${addr}`);
        } catch (err) {
            console.error('Paste error:', err);
            this.updateStatus('Paste failed: ' + (err.message || 'Unknown error'));
        }
    }
    
    clearCell(addr) {
        delete this.getCurrentSheetData()[addr];
        const input = document.querySelector(`[data-addr="${addr}"]`);
        if (input) {
            input.value = '';
            input.removeAttribute('title');
        }
        this.updateCellCount();
        this.updateStatus(`Cleared cell ${addr}`);
    }
    
    formatCell(addr) {
        // Simple formatting - could be enhanced
        const color = prompt('Enter text color (hex, e.g., #ff0000):');
        if (color) {
            const input = document.querySelector(`[data-addr="${addr}"]`);
            if (input) {
                input.style.color = color;
                this.updateStatus(`Formatted cell ${addr}`);
            }
        }
    }

    addRow(atIndex) {
        // Increment total rows
        this.rows++;
        
        // Re-generate grid for simplicity, or implement more complex DOM manipulation
        // Re-generating is simpler but less performant for very large grids
        this.generateGrid(); 
        
        // Adjust data in workbook if necessary (e.g., shift existing data down)
        // This is a complex part: if you insert a row, all cells below need their row index incremented.
        // For example, A10 becomes A11, B10 becomes B11, etc.
        this.shiftData('row', atIndex, 1); // Shift data down by 1 row
        
        this.updateStatus(`Added row at index ${atIndex + 1}`);
    }

    addColumn(atIndex) {
        // Increment total columns
        this.cols++;
        
        // Re-generate grid
        this.generateGrid();
        
        // Adjust data in workbook (e.g., shift existing data right)
        // For example, C1 becomes D1, C2 becomes D2, etc.
        this.shiftData('col', atIndex, 1); // Shift data right by 1 column
        
        this.updateStatus(`Added column at index ${this.getColumnLetter(atIndex)}`);
    }

    // Helper to shift data when rows/columns are added/deleted
    shiftData(type, index, offset) {
        const currentSheet = this.getCurrentSheetData();
        const newSheet = {};

        // Get all existing addresses and sort them to ensure correct shifting order
        const addresses = Object.keys(currentSheet).sort((a, b) => {
            const posA = this.parseAddress(a);
            const posB = this.parseAddress(b);
            if (posA.row !== posB.row) return posA.row - posB.row;
            return posA.col - posB.col;
        });

        addresses.forEach(addr => {
            const pos = this.parseAddress(addr);
            let newRow = pos.row;
            let newCol = pos.col;

            if (type === 'row' && pos.row >= index) {
                newRow += offset;
            } else if (type === 'col' && pos.col >= index) {
                newCol += offset;
            }

            const newAddr = this.getCellAddress(newRow, newCol);
            newSheet[newAddr] = currentSheet[addr];
        });

        this.workbook[this.activeSheet] = newSheet;
        // Re-evaluate all formulas after shifting data
        this.recalculateAllCells(); 
    }
    
    evaluateFormula(formula) {
        // Safe formula evaluation without eval()
        try {
            const expr = formula.substring(1); // Remove =
            
            // Simple SUM function
            if (expr.toUpperCase().startsWith('SUM(')) {
                const range = expr.match(/SUM\(([A-Z0-9:]+)\)/i);
                if (range) {
                    return this.calculateSum(range[1]);
                }
            }
            
            // Simple arithmetic - safe evaluation
            if (/^[\d\+\-\*\/\(\)\s\.]+$/.test(expr)) {
                return this.safeArithmeticEval(expr);
            }
            
            return null;
        } catch (e) {
            return null;
        }
    }

    safeArithmeticEval(expr) {
        // Safe arithmetic evaluation without eval()
        try {
            // Remove whitespace
            expr = expr.replace(/\s/g, '');
            
            // Simple expression parser for basic arithmetic
            
            // For very simple expressions like "2*3" or "10+5"
            if (/^\d+[\+\-\*\/]\d+$/.test(expr)) {
                const operator = expr.match(/[\+\-\*\/]/)[0];
                const parts = expr.split(operator);
                const a = parseFloat(parts[0]);
                const b = parseFloat(parts[1]);
                
                switch (operator) {
                    case '+': return a + b;
                    case '-': return a - b;
                    case '*': return a * b;
                    case '/': return b !== 0 ? a / b : '#DIV/0!';
                    default: return null;
                }
            }
            
            // For more complex expressions, return null (requires proper parser)
            return null;
        } catch (e) {
            return null;
        }
    }
    
    calculateSum(range) {
        // Simple SUM implementation
        const parts = range.split(':');
        if (parts.length !== 2) return 0;
        
        const start = this.parseAddress(parts[0]);
        const end = this.parseAddress(parts[1]);
        if (!start || !end) return 0;
        
        let sum = 0;
        for (let r = start.row; r <= end.row; r++) {
            for (let c = start.col; c <= end.col; c++) {
                const addr = this.getCellAddress(r, c);
                const value = this.getCurrentSheetData()[addr];
                if (value && !isNaN(value)) {
                    sum += parseFloat(value);
                }
            }
        }
        return sum;
    }

    recalculateAllCells() {
        const currentSheet = this.getCurrentSheetData();
        Object.keys(currentSheet).forEach(addr => {
            const value = currentSheet[addr];
            if (value.startsWith('=')) {
                const result = this.evaluateFormula(value);
                const input = document.querySelector(`[data-addr="${addr}"]`);
                if (input) {
                    input.setAttribute('title', `Formula: ${value}\nResult: ${result}`);
                }
            }
        });
    }
    
    loadSampleData() {
        this.workbook.Sheet1 = {
            'A1': 'Product',
            'B1': 'Quantity',
            'C1': 'Price',
            'D1': 'Total',
            'A2': 'Laptop',
            'B2': '2',
            'C2': '999.99',
            'D2': '=B2*C2',
            'A3': 'Mouse',
            'B3': '5',
            'C3': '29.99',
            'D3': '=B3*C3',
            'A4': 'Keyboard',
            'B4': '3',
            'C4': '79.99',
            'D4': '=B4*C4',
            'A6': 'TOTAL:',
            'D6': '=SUM(D2:D4)'
        };
        
        // Update inputs with sample data
        Object.entries(this.getCurrentSheetData()).forEach(([addr, value]) => {
            const input = document.querySelector(`[data-addr="${addr}"]`);
            if (input) {
                input.value = value;
                if (value.startsWith('=')) {
                    const result = this.evaluateFormula(value);
                    if (result !== null) {
                        input.setAttribute('title', `Formula: ${value}\nResult: ${result}`);
                    }
                }
            }
        });
        
        console.log('📋 Sample data loaded');
    }
    
    cleanup() {
        // Remove all event listeners to prevent memory leaks
        this.eventHandlers.forEach(({ element, event, handler }) => {
            element.removeEventListener(event, handler);
        });
        this.eventHandlers = [];
    }

    newWorkbook() {
        if (confirm('Create new workbook? This will clear all data.')) {
            // Cleanup existing event handlers
            this.cleanup();
            
            this.workbook = {
                Sheet1: {},
                Sheet2: {},
                Sheet3: {}
            };
            this.activeSheet = 'Sheet1';
            
            // Clear all inputs
            document.querySelectorAll('.cell-input').forEach(input => {
                input.value = '';
                input.removeAttribute('title');
                input.style.color = '';
            });
            
            // Re-setup events after cleanup
            this.setupEvents();
            
            this.updateCellCount();
            this.updateStatus('New workbook created');
        }
    }
    
    save() {
        const data = {
            workbook: this.workbook,
            activeSheet: this.activeSheet,
            timestamp: new Date().toISOString()
        };
        
        const json = JSON.stringify(data, null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `excel-workbook-${new Date().toISOString().split('T')[0]}.json`;
        a.click();
        
        URL.revokeObjectURL(url);
        this.updateStatus('Workbook saved');
    }
    
    load() {
const input = document.createElement('input');
        input.type = 'file';
        // Allow JSON, XLSX, and CSV files
        input.accept = '.json,.xlsx,.csv'; 

        input.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    if (file.name.endsWith('.json')) {
                        const data = JSON.parse(e.target.result);
                        this.workbook = data.workbook || data;
                        this.activeSheet = data.activeSheet || 'Sheet1';
                    } else {
                        if (typeof XLSX === 'undefined') {
                            alert('Error: XLSX library not loaded. Cannot read Excel files.');
                            return;
                        }
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const sheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[sheetName];
                        const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        this.workbook = {
                            [sheetName]: this.convertSheetDataToInternalFormat(sheetData)
                        };
                        this.activeSheet = sheetName;
                    }

                    this.switchSheet(this.activeSheet);
                    this.updateStatus('Workbook loaded');

                } catch (err) {
                    console.error('Error loading file:', err);
                    alert('Invalid file format or error processing file.');
                }
            };

            // Read file based on type
            if (file.name.endsWith('.json')) {
                reader.readAsText(file);
            } else {
                reader.readAsArrayBuffer(file); // Read as ArrayBuffer for XLSX/CSV
            }
        });

        document.body.appendChild(input);
        input.click();
        document.body.removeChild(input);
    }

    
    export() {
        if (typeof XLSX === 'undefined') {
            this.updateStatus('Error: XLSX library not loaded');
            return;
        }
        
        const wb = XLSX.utils.book_new();
        
        // Export all sheets
        Object.entries(this.workbook).forEach(([sheetName, data]) => {
            const ws = {};
            
            // Add data to worksheet
            Object.entries(data).forEach(([addr, value]) => {
                if (value.startsWith && value.startsWith('=')) {
                    // Formula
                    ws[addr] = { f: value.substring(1), t: 'f' };
                } else if (!isNaN(value) && value !== '') {
                    // Number
                    ws[addr] = { v: parseFloat(value), t: 'n' };
                } else {
                    // String
                    ws[addr] = { v: value, t: 's' };
                }
            });
            
            // Set range
            if (Object.keys(data).length > 0) {
                const addresses = Object.keys(data);
                let minR = 999, maxR = 0, minC = 999, maxC = 0;
                
                addresses.forEach(addr => {
                    const decoded = XLSX.utils.decode_cell(addr);
                    minR = Math.min(minR, decoded.r);
                    maxR = Math.max(maxR, decoded.r);
                    minC = Math.min(minC, decoded.c);
                    maxC = Math.max(maxC, decoded.c);
                });
                
                ws['!ref'] = XLSX.utils.encode_range({
                    s: { r: minR, c: minC },
                    e: { r: maxR, c: maxC }
                });
            }
            
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });
        
        XLSX.writeFile(wb, `excel-export-${new Date().toISOString().split('T')[0]}.xlsx`);
        this.updateStatus('Excel file exported');
    }
    
    updateCellCount() {
        const count = Object.keys(this.getCurrentSheetData()).length;
        document.getElementById('cell-count').textContent = count;
    }
    
    updateStatus(message) {
        document.getElementById('status-left').textContent = message;
        setTimeout(() => {
            if (document.getElementById('status-left').textContent === message) {
                document.getElementById('status-left').textContent = 'Ready';
            }
        }, 3000);
    }

    convertSheetDataToInternalFormat(sheetData) {
        const internalSheet = {};
        sheetData.forEach((row, rIdx) => {
            row.forEach((cellValue, cIdx) => {
                const addr = this.getCellAddress(rIdx, cIdx);
                if (cellValue !== undefined && cellValue !== null && cellValue !== '') {
                    internalSheet[addr] = String(cellValue); // Store as string
                }
            });
        });
        return internalSheet;
    }
    
    getCurrentSheetData() {
        return this.workbook[this.activeSheet];
    }
    
    getColumnLetter(index) {
        let letter = '';
        let num = index + 1; // Convert to 1-based
        
        while (num > 0) {
            num--;
            letter = String.fromCharCode(num % 26 + 65) + letter;
            num = Math.floor(num / 26);
        }
        return letter;
    }
    
    getCellAddress(row, col) {
        return this.getColumnLetter(col) + (row + 1);
    }
    
    parseAddress(addr) {
        const match = addr.match(/^([A-Z]+)(\d+)$/);
        if (match) {
            const colStr = match[1];
            let col = 0;
            for (let i = 0; i < colStr.length; i++) {
                col = col * 26 + (colStr.charCodeAt(i) - 64);
            }
            return {
                col: col - 1,
                row: parseInt(match[2]) - 1
            };
        }
        return null;
    }
    
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
}

// START EXCEL
let excel = null;

document.addEventListener('DOMContentLoaded', () => {
    excel = new Excel();
});

// PERFORMANCE MONITORING
let perfStats = {
    renderTime: 0,
    cellCount: 0,
    memoryUsage: 0
};

function updatePerfStats() {
    if (performance.memory) {
        perfStats.memoryUsage = Math.round(performance.memory.usedJSHeapSize / 1048576);
    }
    
    // Update status with perf info
    if (excel && excel.isLoaded) {
        const status = document.getElementById('status-right');
        status.innerHTML = `Excel v2.0 | <span id="cell-count">${Object.keys(excel.getCurrentSheetData()).length}</span> cells | ${perfStats.memoryUsage}MB`;
    }
}

setInterval(updatePerfStats, 5000);

// ERROR HANDLING
window.addEventListener('error', (e) => {
    console.error('💥 Error:', e.error);
    document.getElementById('status-left').textContent = `Error: ${e.message}`;
    document.getElementById('status-left').style.color = '#ef4444';
    
    setTimeout(() => {
        document.getElementById('status-left').style.color = '';
        document.getElementById('status-left').textContent = 'Ready';
    }, 5000);
});

// KEYBOARD SHORTCUTS
document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && !e.target.classList.contains('cell-input') && !e.target.classList.contains('formula-input')) {
        switch (e.key.toLowerCase()) {
            case 's':
                e.preventDefault();
                if (excel) excel.save();
                break;
            case 'o':
                e.preventDefault();
                if (excel) excel.load();
                break;
            case 'n':
                e.preventDefault();
                if (excel) excel.newWorkbook();
                break;
            case 'e':
                e.preventDefault();
                if (excel) excel.export();
                break;
        }
    }
});

console.log('📦 Excel script loaded');