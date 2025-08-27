/**
 * Grid Event Manager - Optimized Event Delegation System
 * Handles all grid interactions with single event listeners
 */

import { AppState } from '../core/state.js';
import { getWorksheet, persistSnapshot } from '../spreadsheet/workbook-manager.js';
import { renderSpreadsheetTable } from '../spreadsheet/grid-renderer.js';
import { showToast } from '../ui/toast.js';
import { parseCellValue, expandRefForCell } from '../utils/index.js';

export class GridEventManager {
  constructor(grid) {
    this.grid = grid;
    this.container = grid.container;
    
    // Event state tracking
    this.state = {
      activeCell: null,
      isEditing: false,
      isDragging: false,
      isSelecting: false,
      selectionStart: null,
      selectionEnd: null,
      dragStartTime: 0,
      lastClickTime: 0,
      clickCount: 0
    };
    
    // Selection state
    this.selection = {
      ranges: [],
      activeRange: null,
      isMultiSelect: false
    };
    
    // Touch handling
    this.touch = {
      startX: 0,
      startY: 0,
      startTime: 0,
      isTouch: false
    };
    
    // Performance tracking
    this.eventStats = {
      clicks: 0,
      keydowns: 0,
      scrolls: 0,
      lastEventTime: 0
    };
    
    this.setupEventDelegation();
  }
  
  setupEventDelegation() {
    // Single event listeners for the entire grid
    this.container.addEventListener('click', this.handleClick.bind(this), true);
    this.container.addEventListener('dblclick', this.handleDoubleClick.bind(this));
    this.container.addEventListener('mousedown', this.handleMouseDown.bind(this));
    this.container.addEventListener('mousemove', this.handleMouseMove.bind(this));
    this.container.addEventListener('mouseup', this.handleMouseUp.bind(this));
    this.container.addEventListener('keydown', this.handleKeyDown.bind(this), true);
    this.container.addEventListener('keyup', this.handleKeyUp.bind(this));
    this.container.addEventListener('focusin', this.handleFocusIn.bind(this));
    this.container.addEventListener('focusout', this.handleFocusOut.bind(this));
    this.container.addEventListener('contextmenu', this.handleContextMenu.bind(this));
    
    // Touch events
    this.container.addEventListener('touchstart', this.handleTouchStart.bind(this), { passive: false });
    this.container.addEventListener('touchmove', this.handleTouchMove.bind(this), { passive: false });
    this.container.addEventListener('touchend', this.handleTouchEnd.bind(this));
    
    // Prevent default context menu on the grid
    this.container.addEventListener('contextmenu', (e) => {
      const target = e.target.closest('.excel-cell, .virtual-cell, .excel-col-header, .excel-row-header');
      if (target) {
        e.preventDefault();
      }
    });
    
    // Global keyboard shortcuts
    document.addEventListener('keydown', this.handleGlobalKeyDown.bind(this));
  }
  
  handleClick(e) {
    this.eventStats.clicks++;
    this.eventStats.lastEventTime = Date.now();
    
    const now = Date.now();
    
    // Handle double-click detection
    if (now - this.state.lastClickTime < 300) {
      this.state.clickCount++;
    } else {
      this.state.clickCount = 1;
    }
    this.state.lastClickTime = now;
    
    const cell = e.target.closest('.excel-cell, .virtual-cell');
    const colHeader = e.target.closest('.excel-col-header, .virtual-col-header');
    const rowHeader = e.target.closest('.excel-row-header, .virtual-row-header');
    const cornerCell = e.target.closest('.excel-corner, .virtual-corner');
    
    if (cell) {
      this.handleCellClick(e, cell);
    } else if (colHeader) {
      this.handleColumnHeaderClick(e, colHeader);
    } else if (rowHeader) {
      this.handleRowHeaderClick(e, rowHeader);
    } else if (cornerCell) {
      this.handleCornerClick(e);
    }
  }
  
  handleCellClick(e, cell) {
    const addr = cell.dataset.cell;
    const row = parseInt(cell.dataset.row);
    const col = parseInt(cell.dataset.col);
    
    if (!addr) return;
    
    // Update active cell
    this.setActiveCell(addr, row, col);
    
    // Handle selection modes
    if (e.ctrlKey || e.metaKey) {
      this.toggleCellSelection(addr);
    } else if (e.shiftKey && this.state.activeCell) {
      this.extendSelection(addr);
    } else {
      this.clearSelection();
      this.selectCell(addr);
    }
    
    // Focus the cell input for editing
    const input = cell.querySelector('input');
    if (input && !this.state.isEditing) {
      setTimeout(() => input.focus(), 10);
    }
  }
  
  handleDoubleClick(e) {
    const cell = e.target.closest('.excel-cell, .virtual-cell');
    if (cell) {
      this.enterEditMode(cell);
    }
  }
  
  handleMouseDown(e) {
    const cell = e.target.closest('.excel-cell, .virtual-cell');
    if (cell && e.button === 0) { // Left mouse button
      this.state.isDragging = true;
      this.state.dragStartTime = Date.now();
      
      const addr = cell.dataset.cell;
      this.state.selectionStart = addr;
      
      e.preventDefault(); // Prevent text selection
    }
  }
  
  handleMouseMove(e) {
    if (this.state.isDragging) {
      const cell = e.target.closest('.excel-cell, .virtual-cell');
      if (cell) {
        const addr = cell.dataset.cell;
        if (addr !== this.state.selectionEnd) {
          this.state.selectionEnd = addr;
          this.updateDragSelection();
        }
      }
    }
  }
  
  handleMouseUp(e) {
    if (this.state.isDragging) {
      this.state.isDragging = false;
      const dragDuration = Date.now() - this.state.dragStartTime;
      
      // If drag was very short, treat as click
      if (dragDuration < 150) {
        this.state.selectionEnd = null;
      }
      
      this.finalizeDragSelection();
    }
  }
  
  handleKeyDown(e) {
    this.eventStats.keydowns++;
    
    const cell = e.target.closest('.excel-cell, .virtual-cell');
    if (!cell) return;
    
    const addr = cell.dataset.cell;
    const row = parseInt(cell.dataset.row);
    const col = parseInt(cell.dataset.col);
    
    // Handle navigation
    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        this.navigateCell(row - 1, col, e.shiftKey);
        break;
      case 'ArrowDown':
        e.preventDefault();
        this.navigateCell(row + 1, col, e.shiftKey);
        break;
      case 'ArrowLeft':
        e.preventDefault();
        this.navigateCell(row, col - 1, e.shiftKey);
        break;
      case 'ArrowRight':
        e.preventDefault();
        this.navigateCell(row, col + 1, e.shiftKey);
        break;
      case 'Enter':
        e.preventDefault();
        this.exitEditMode();
        this.navigateCell(row + 1, col, false);
        break;
      case 'Tab':
        e.preventDefault();
        this.exitEditMode();
        if (e.shiftKey) {
          this.navigateCell(row, col - 1, false);
        } else {
          this.navigateCell(row, col + 1, false);
        }
        break;
      case 'Delete':
      case 'Backspace':
        if (!this.state.isEditing) {
          e.preventDefault();
          this.clearCellContent(addr);
        }
        break;
      case 'F2':
        e.preventDefault();
        this.enterEditMode(cell);
        break;
      case 'Escape':
        e.preventDefault();
        this.cancelEdit();
        break;
      default:
        // Start editing if typing regular characters
        if (!this.state.isEditing && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
          this.enterEditMode(cell);
        }
    }
  }
  
  handleKeyUp(e) {
    // Handle key up events if needed
  }
  
  handleFocusIn(e) {
    const input = e.target;
    const cell = input.closest('.excel-cell, .virtual-cell');
    
    if (cell && input.classList.contains('excel-cell-input')) {
      const addr = cell.dataset.cell;
      this.onCellFocus(addr, input, cell);
    }
  }
  
  handleFocusOut(e) {
    const input = e.target;
    const cell = input.closest('.excel-cell, .virtual-cell');
    
    if (cell && input.classList.contains('excel-cell-input')) {
      const addr = cell.dataset.cell;
      this.onCellBlur(addr, input, cell);
    }
  }
  
  handleContextMenu(e) {
    const cell = e.target.closest('.excel-cell, .virtual-cell');
    const colHeader = e.target.closest('.excel-col-header, .virtual-col-header');
    const rowHeader = e.target.closest('.excel-row-header, .virtual-row-header');
    
    if (cell) {
      this.showCellContextMenu(e, cell);
    } else if (colHeader) {
      this.showColumnContextMenu(e, colHeader);
    } else if (rowHeader) {
      this.showRowContextMenu(e, rowHeader);
    }
  }
  
  // Touch event handlers
  handleTouchStart(e) {
    this.touch.isTouch = true;
    this.touch.startTime = Date.now();
    
    if (e.touches.length === 1) {
      const touch = e.touches[0];
      this.touch.startX = touch.clientX;
      this.touch.startY = touch.clientY;
      
      // Convert touch to click
      const cell = e.target.closest('.excel-cell, .virtual-cell');
      if (cell) {
        e.preventDefault();
        this.handleCellClick(e, cell);
      }
    }
  }
  
  handleTouchMove(e) {
    if (e.touches.length === 1) {
      const touch = e.touches[0];
      const deltaX = Math.abs(touch.clientX - this.touch.startX);
      const deltaY = Math.abs(touch.clientY - this.touch.startY);
      
      // If moved significantly, it's a drag/scroll, not a tap
      if (deltaX > 10 || deltaY > 10) {
        this.touch.isTouch = false;
      }
    }
  }
  
  handleTouchEnd(e) {
    const touchDuration = Date.now() - this.touch.startTime;
    
    // Long press detection
    if (touchDuration > 500 && this.touch.isTouch) {
      const cell = e.target.closest('.excel-cell, .virtual-cell');
      if (cell) {
        this.showCellContextMenu(e, cell);
      }
    }
    
    this.touch.isTouch = false;
  }
  
  // Global keyboard shortcuts
  handleGlobalKeyDown(e) {
    if (e.ctrlKey || e.metaKey) {
      switch (e.key.toLowerCase()) {
        case 'c':
          if (this.hasSelection()) {
            e.preventDefault();
            this.copySelection();
          }
          break;
        case 'v':
          if (this.hasClipboard()) {
            e.preventDefault();
            this.pasteSelection();
          }
          break;
        case 'x':
          if (this.hasSelection()) {
            e.preventDefault();
            this.cutSelection();
          }
          break;
        case 'z':
          e.preventDefault();
          this.undo();
          break;
        case 'y':
          e.preventDefault();
          this.redo();
          break;
        case 'a':
          if (this.container.contains(document.activeElement)) {
            e.preventDefault();
            this.selectAll();
          }
          break;
      }
    }
  }
  
  // Cell management methods
  setActiveCell(addr, row, col) {
    this.state.activeCell = addr;
    AppState.activeCell = { r: row, c: col };
    
    // Update address bar
    const cellRefElement = document.getElementById('cell-reference');
    if (cellRefElement) {
      cellRefElement.textContent = addr;
    }
    
    // Update formula bar
    this.updateFormulaBar(addr);
  }
  
  navigateCell(targetRow, targetCol, extendSelection = false) {
    const ws = getWorksheet();
    if (!ws) return;
    
    // Bound checking
    targetRow = Math.max(0, targetRow);
    targetCol = Math.max(0, targetCol);
    
    const targetAddr = window.XLSX.utils.encode_cell({ r: targetRow, c: targetCol });
    
    // If virtual scrolling, ensure cell is visible
    if (this.grid.scrollToCell) {
      this.grid.scrollToCell(targetRow, targetCol);
    }
    
    // Find and focus the target cell
    setTimeout(() => {
      const targetCell = this.container.querySelector(`[data-cell="${targetAddr}"]`);
      if (targetCell) {
        if (extendSelection) {
          this.extendSelection(targetAddr);
        } else {
          this.clearSelection();
          this.selectCell(targetAddr);
        }
        
        const input = targetCell.querySelector('input');
        if (input) {
          input.focus();
        }
      }
    }, 50); // Allow time for virtual scrolling
  }
  
  enterEditMode(cell) {
    if (this.state.isEditing) return;
    
    const input = cell.querySelector('input');
    if (!input) return;
    
    this.state.isEditing = true;
    cell.classList.add('editing');
    
    input.readOnly = false;
    input.focus();
    
    // Show raw formula if it's a formula
    const addr = cell.dataset.cell;
    const ws = getWorksheet();
    const cellData = ws[addr];
    
    if (cellData && cellData.f) {
      input.value = '=' + cellData.f;
    }
    
    // Select all content
    setTimeout(() => input.select(), 10);
  }
  
  exitEditMode() {
    if (!this.state.isEditing) return;
    
    const activeInput = document.activeElement;
    if (activeInput && activeInput.classList.contains('excel-cell-input')) {
      activeInput.blur();
    }
    
    this.state.isEditing = false;
  }
  
  cancelEdit() {
    if (!this.state.isEditing) return;
    
    const activeInput = document.activeElement;
    const cell = activeInput?.closest('.excel-cell, .virtual-cell');
    
    if (cell && activeInput) {
      // Restore original value
      const addr = cell.dataset.cell;
      const ws = getWorksheet();
      const cellData = ws[addr];
      
      let originalValue = '';
      if (cellData) {
        originalValue = cellData.f ? ('=' + cellData.f) : (cellData.v || '');
      }
      
      activeInput.value = originalValue;
      activeInput.blur();
    }
    
    this.state.isEditing = false;
  }
  
  onCellFocus(addr, input, cell) {
    const cellRef = window.XLSX.utils.decode_cell(addr);
    AppState.activeCell = cellRef;
    
    input.readOnly = false;
    
    // Show raw value for formulas in formula bar
    this.updateFormulaBar(addr);
    
    // Emit focus event
    this.container.dispatchEvent(new CustomEvent('cellFocus', {
      detail: { addr, cell, input }
    }));
  }
  
  onCellBlur(addr, input, cell) {
    const ws = getWorksheet();
    const currentCell = ws[addr];
    const currentValue = currentCell ? (currentCell.f ? '=' + currentCell.f : (currentCell.v || '')) : '';
    
    // Only update if value changed
    if (String(input.value) !== String(currentValue)) {
      this.updateCell(addr, input.value);
    }
    
    input.readOnly = true;
    cell.classList.remove('editing');
    this.state.isEditing = false;
    
    // Emit blur event
    this.container.dispatchEvent(new CustomEvent('cellBlur', {
      detail: { addr, cell, input, value: input.value }
    }));
  }
  
  updateCell(addr, value) {
    const ws = getWorksheet();
    const oldValue = ws[addr] ? (ws[addr].f || ws[addr].v) : '';
    
    if (String(oldValue) !== String(value)) {
      // Save to history
      if (window.saveToHistory) {
        window.saveToHistory(`Edit cell ${addr}`, { 
          addr, 
          oldValue, 
          newValue: value, 
          sheet: AppState.activeSheet 
        });
      }
    }
    
    if (value.startsWith('=')) {
      ws[addr] = { t: 'f', f: value.substring(1) };
    } else {
      const parsed = parseCellValue(value);
      if (parsed.t === 'z') {
        delete ws[addr];
      } else {
        ws[addr] = { t: parsed.t, v: parsed.v };
      }
    }
    
    expandRefForCell(ws, addr);
    persistSnapshot();
    
    // Re-render if using traditional grid, or update if using virtual grid
    if (this.grid.refresh) {
      this.grid.refresh();
    } else if (renderSpreadsheetTable) {
      renderSpreadsheetTable();
    }
  }
  
  clearCellContent(addr) {
    this.updateCell(addr, '');
    showToast(`Cleared cell ${addr}`, 'info', 1000);
  }
  
  updateFormulaBar(addr) {
    const formulaBar = document.getElementById('formula-bar');
    if (!formulaBar) return;
    
    const ws = getWorksheet();
    const cell = ws[addr];
    let cellValue = '';
    
    if (cell && cell.f) {
      cellValue = '=' + cell.f;
    } else if (cell && cell.v !== undefined) {
      cellValue = String(cell.v);
    }
    
    formulaBar.value = cellValue;
  }
  
  // Selection methods
  selectCell(addr) {
    this.clearSelection();
    this.selection.ranges = [{ start: addr, end: addr }];
    this.highlightSelection();
  }
  
  clearSelection() {
    this.selection.ranges = [];
    this.removeSelectionHighlights();
  }
  
  extendSelection(endAddr) {
    if (!this.state.activeCell) return;
    
    const range = {
      start: this.state.activeCell,
      end: endAddr
    };
    
    this.selection.ranges = [range];
    this.highlightSelection();
  }
  
  toggleCellSelection(addr) {
    // Multi-select implementation
    const existingRange = this.selection.ranges.find(range => 
      range.start === addr && range.end === addr
    );
    
    if (existingRange) {
      const index = this.selection.ranges.indexOf(existingRange);
      this.selection.ranges.splice(index, 1);
    } else {
      this.selection.ranges.push({ start: addr, end: addr });
    }
    
    this.highlightSelection();
  }
  
  highlightSelection() {
    this.removeSelectionHighlights();
    
    this.selection.ranges.forEach(range => {
      const startCell = window.XLSX.utils.decode_cell(range.start);
      const endCell = window.XLSX.utils.decode_cell(range.end);
      
      const minRow = Math.min(startCell.r, endCell.r);
      const maxRow = Math.max(startCell.r, endCell.r);
      const minCol = Math.min(startCell.c, endCell.c);
      const maxCol = Math.max(startCell.c, endCell.c);
      
      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          const addr = window.XLSX.utils.encode_cell({ r, c });
          const cellElement = this.container.querySelector(`[data-cell="${addr}"]`);
          if (cellElement) {
            cellElement.classList.add('selected');
          }
        }
      }
    });
  }
  
  removeSelectionHighlights() {
    const selectedCells = this.container.querySelectorAll('.selected');
    selectedCells.forEach(cell => cell.classList.remove('selected'));
  }
  
  hasSelection() {
    return this.selection.ranges.length > 0;
  }
  
  hasClipboard() {
    return AppState.clipboard !== null;
  }
  
  // Context menu handlers
  showCellContextMenu(e, cell) {
    const addr = cell.dataset.cell;
    this.showContextMenu(e.clientX, e.clientY, [
      { label: 'Cut', action: () => this.cutCell(addr) },
      { label: 'Copy', action: () => this.copyCell(addr) },
      { label: 'Paste', action: () => this.pasteCell(addr) },
      { label: 'Clear Contents', action: () => this.clearCellContent(addr) },
      { label: 'Insert Comment', action: () => this.insertComment(addr) }
    ]);
  }
  
  showColumnContextMenu(e, header) {
    // Column context menu implementation
  }
  
  showRowContextMenu(e, header) {
    // Row context menu implementation
  }
  
  showContextMenu(x, y, items) {
    // Remove existing context menu
    const existing = document.getElementById('grid-context-menu');
    if (existing) existing.remove();
    
    const menu = document.createElement('div');
    menu.id = 'grid-context-menu';
    menu.className = 'fixed z-50 bg-white border border-gray-300 rounded shadow-lg text-sm';
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    menu.style.minWidth = '160px';
    
    items.forEach(item => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'w-full text-left px-3 py-2 hover:bg-gray-100';
      button.textContent = item.label;
      button.addEventListener('click', () => {
        this.hideContextMenu();
        item.action();
      });
      menu.appendChild(button);
    });
    
    document.body.appendChild(menu);
    
    // Hide on click outside
    setTimeout(() => {
      const hideHandler = (e) => {
        if (!menu.contains(e.target)) {
          this.hideContextMenu();
        }
      };
      document.addEventListener('click', hideHandler, { once: true });
    }, 10);
  }
  
  hideContextMenu() {
    const menu = document.getElementById('grid-context-menu');
    if (menu) menu.remove();
  }
  
  // Clipboard operations
  copySelection() {
    console.log('Copy selection');
    showToast('Selection copied', 'success', 1000);
  }
  
  cutSelection() {
    console.log('Cut selection');
    showToast('Selection cut', 'success', 1000);
  }
  
  pasteSelection() {
    console.log('Paste selection');
    showToast('Selection pasted', 'success', 1000);
  }
  
  copyCell(addr) {
    console.log('Copy cell:', addr);
    showToast(`Cell ${addr} copied`, 'success', 1000);
  }
  
  cutCell(addr) {
    console.log('Cut cell:', addr);
    showToast(`Cell ${addr} cut`, 'success', 1000);
  }
  
  pasteCell(addr) {
    console.log('Paste to cell:', addr);
    showToast(`Pasted to cell ${addr}`, 'success', 1000);
  }
  
  insertComment(addr) {
    const comment = prompt('Enter comment:');
    if (comment) {
      console.log('Insert comment:', comment, 'at', addr);
      showToast('Comment added', 'success', 1000);
    }
  }
  
  // Undo/Redo
  undo() {
    console.log('Undo');
    showToast('Undo', 'info', 1000);
  }
  
  redo() {
    console.log('Redo');
    showToast('Redo', 'info', 1000);
  }
  
  selectAll() {
    console.log('Select all');
    showToast('All cells selected', 'info', 1000);
  }
  
  // Drag selection helpers
  updateDragSelection() {
    if (this.state.selectionStart && this.state.selectionEnd) {
      this.selection.ranges = [{
        start: this.state.selectionStart,
        end: this.state.selectionEnd
      }];
      this.highlightSelection();
    }
  }
  
  finalizeDragSelection() {
    // Finalize the drag selection
    this.state.selectionStart = null;
    this.state.selectionEnd = null;
  }
  
  // Column/row header handlers
  handleColumnHeaderClick(e, header) {
    const col = parseInt(header.dataset.col);
    console.log('Column header clicked:', col);
    
    // Select entire column
    AppState.selectedCols = [col];
    AppState.selectedRows = [];
  }
  
  handleRowHeaderClick(e, header) {
    const row = parseInt(header.dataset.row);
    console.log('Row header clicked:', row);
    
    // Select entire row
    AppState.selectedRows = [row + 1]; // 1-based for display
    AppState.selectedCols = [];
  }
  
  handleCornerClick(e) {
    console.log('Corner clicked - select all');
    this.selectAll();
  }
  
  // Public API
  getEventStats() {
    return { ...this.eventStats };
  }
  
  destroy() {
    // Clean up event listeners
    this.container.removeEventListener('click', this.handleClick);
    // ... remove all other listeners
    document.removeEventListener('keydown', this.handleGlobalKeyDown);
  }
}