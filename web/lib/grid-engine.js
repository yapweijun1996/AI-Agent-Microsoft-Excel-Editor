/**
 * Excel Grid Engine - Core Architecture
 * Modern, high-performance Excel-like grid implementation
 */

import { AppState } from '../core/state.js';
import { getWorksheet } from '../spreadsheet/workbook-manager.js';

export class ExcelGridEngine {
  constructor(container, options = {}) {
    this.container = container;
    this.options = {
      rowHeight: 20,
      colWidth: 64,
      headerHeight: 20,
      bufferSize: 10,
      enableVirtualScroll: true,
      maxVisibleRows: 100,
      maxVisibleCols: 50,
      ...options
    };
    
    this.plugins = new Map();
    this.hooks = new Map();
    this.viewport = {
      startRow: 0,
      endRow: 0,
      startCol: 0,
      endCol: 0,
      scrollTop: 0,
      scrollLeft: 0,
      width: 0,
      height: 0
    };
    
    this.cellPool = new CellPool(1000); // Reuse DOM elements
    this.renderState = {
      isRendering: false,
      needsFullRender: true,
      dirtyRegions: new Set()
    };
    
    this.eventManager = null;
    this.performanceMonitor = null;
    
    this.init();
  }
  
  init() {
    this.setupContainer();
    this.calculateViewport();
    this.bindScrollEvents();
  }
  
  setupContainer() {
    this.container.className = 'excel-grid-container';
    this.container.innerHTML = `
      <div class="excel-grid-viewport">
        <div class="excel-grid-canvas" style="position: relative;">
          <div class="excel-header-row"></div>
          <div class="excel-grid-body"></div>
        </div>
      </div>
    `;
    
    this.viewport.width = this.container.clientWidth;
    this.viewport.height = this.container.clientHeight;
    
    // Cache DOM elements
    this.elements = {
      viewport: this.container.querySelector('.excel-grid-viewport'),
      canvas: this.container.querySelector('.excel-grid-canvas'),
      headerRow: this.container.querySelector('.excel-header-row'),
      body: this.container.querySelector('.excel-grid-body')
    };
  }
  
  calculateViewport() {
    const ws = getWorksheet();
    if (!ws || !ws['!ref']) return;
    
    const range = window.XLSX.utils.decode_range(ws['!ref']);
    const visibleRows = Math.min(
      Math.ceil(this.viewport.height / this.options.rowHeight) + this.options.bufferSize,
      this.options.maxVisibleRows
    );
    const visibleCols = Math.min(
      Math.ceil(this.viewport.width / this.options.colWidth) + this.options.bufferSize,
      this.options.maxVisibleCols
    );
    
    this.viewport.startRow = Math.floor(this.viewport.scrollTop / this.options.rowHeight);
    this.viewport.endRow = Math.min(this.viewport.startRow + visibleRows, range.e.r + 1);
    this.viewport.startCol = Math.floor(this.viewport.scrollLeft / this.options.colWidth);
    this.viewport.endCol = Math.min(this.viewport.startCol + visibleCols, range.e.c + 1);
  }
  
  bindScrollEvents() {
    let scrollTimeout;
    
    this.elements.viewport.addEventListener('scroll', (e) => {
      this.viewport.scrollTop = e.target.scrollTop;
      this.viewport.scrollLeft = e.target.scrollLeft;
      
      // Throttle scroll updates
      if (scrollTimeout) clearTimeout(scrollTimeout);
      scrollTimeout = setTimeout(() => {
        this.handleScroll();
      }, 16); // 60fps
    });
  }
  
  handleScroll() {
    const oldViewport = { ...this.viewport };
    this.calculateViewport();
    
    // Check if viewport changed significantly
    if (this.viewportChanged(oldViewport)) {
      this.renderVisibleCells();
    }
  }
  
  viewportChanged(oldViewport) {
    return (
      Math.abs(this.viewport.startRow - oldViewport.startRow) > 5 ||
      Math.abs(this.viewport.startCol - oldViewport.startCol) > 5
    );
  }
  
  render(forceFullRender = false) {
    if (this.renderState.isRendering) return;
    
    this.renderState.isRendering = true;
    
    requestAnimationFrame(() => {
      const startTime = performance.now();
      
      try {
        if (forceFullRender || this.renderState.needsFullRender) {
          this.renderFullGrid();
        } else {
          this.renderVisibleCells();
        }
        
        this.renderState.needsFullRender = false;
      } catch (error) {
        console.error('Grid render error:', error);
      } finally {
        const renderTime = performance.now() - startTime;
        this.renderState.isRendering = false;
        
        // Performance tracking
        if (this.performanceMonitor) {
          this.performanceMonitor.trackRender(renderTime);
        }
      }
    });
  }
  
  renderFullGrid() {
    this.renderHeaders();
    this.renderVisibleCells();
    this.updateScrollArea();
  }
  
  renderHeaders() {
    const headerFragment = document.createDocumentFragment();
    
    // Corner cell
    const corner = document.createElement('div');
    corner.className = 'excel-corner-cell';
    corner.style.cssText = `
      position: absolute;
      left: 0;
      top: 0;
      width: 42px;
      height: ${this.options.headerHeight}px;
      background: linear-gradient(to bottom, #f6f8fa 0%, #e9ecef 100%);
      border: 1px solid #c6cbd1;
      z-index: 10;
    `;
    headerFragment.appendChild(corner);
    
    // Column headers
    for (let c = this.viewport.startCol; c < this.viewport.endCol; c++) {
      const colLetter = window.XLSX.utils.encode_col(c);
      const header = this.cellPool.getHeader() || this.createColumnHeader();
      
      header.textContent = colLetter;
      header.className = 'excel-col-header';
      header.dataset.col = c;
      header.style.cssText = `
        position: absolute;
        left: ${42 + (c - this.viewport.startCol) * this.options.colWidth}px;
        top: 0;
        width: ${this.options.colWidth}px;
        height: ${this.options.headerHeight}px;
        background: linear-gradient(to bottom, #f6f8fa 0%, #e9ecef 100%);
        border: 1px solid #c6cbd1;
        text-align: center;
        line-height: ${this.options.headerHeight}px;
        font-size: 11px;
        font-family: 'Calibri', sans-serif;
        z-index: 9;
      `;
      
      headerFragment.appendChild(header);
    }
    
    this.elements.headerRow.innerHTML = '';
    this.elements.headerRow.appendChild(headerFragment);
  }
  
  renderVisibleCells() {
    const ws = getWorksheet();
    if (!ws) return;
    
    const bodyFragment = document.createDocumentFragment();
    
    for (let r = this.viewport.startRow; r < this.viewport.endRow; r++) {
      // Row header
      const rowHeader = this.cellPool.getRowHeader() || this.createRowHeader();
      rowHeader.textContent = r + 1;
      rowHeader.className = 'excel-row-header';
      rowHeader.dataset.row = r;
      rowHeader.style.cssText = `
        position: absolute;
        left: 0;
        top: ${(r - this.viewport.startRow) * this.options.rowHeight + this.options.headerHeight}px;
        width: 42px;
        height: ${this.options.rowHeight}px;
        background: linear-gradient(to bottom, #f6f8fa 0%, #e9ecef 100%);
        border: 1px solid #c6cbd1;
        text-align: center;
        line-height: ${this.options.rowHeight}px;
        font-size: 11px;
        font-family: 'Calibri', sans-serif;
        z-index: 8;
      `;
      bodyFragment.appendChild(rowHeader);
      
      // Cells in row
      for (let c = this.viewport.startCol; c < this.viewport.endCol; c++) {
        const addr = window.XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        
        const cellElement = this.cellPool.get() || this.createCell();
        this.updateCellContent(cellElement, addr, cell, r, c);
        
        cellElement.style.cssText = `
          position: absolute;
          left: ${42 + (c - this.viewport.startCol) * this.options.colWidth}px;
          top: ${(r - this.viewport.startRow) * this.options.rowHeight + this.options.headerHeight}px;
          width: ${this.options.colWidth}px;
          height: ${this.options.rowHeight}px;
        `;
        
        bodyFragment.appendChild(cellElement);
      }
    }
    
    this.elements.body.innerHTML = '';
    this.elements.body.appendChild(bodyFragment);
  }
  
  createCell() {
    const cell = document.createElement('div');
    cell.className = 'excel-cell';
    
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'excel-cell-input';
    
    cell.appendChild(input);
    return cell;
  }
  
  createColumnHeader() {
    return document.createElement('div');
  }
  
  createRowHeader() {
    return document.createElement('div');
  }
  
  updateCellContent(cellElement, addr, cellData, row, col) {
    const input = cellElement.querySelector('.excel-cell-input');
    
    cellElement.dataset.cell = addr;
    cellElement.dataset.row = row;
    cellElement.dataset.col = col;
    
    let value = '';
    let hasFormula = false;
    
    if (cellData) {
      if (cellData.f) {
        hasFormula = true;
        try {
          // Calculate formula result
          if (typeof getFormulaEngine === 'function') {
            const result = getFormulaEngine(AppState.wb, AppState.activeSheet)
              .execute('=' + cellData.f, AppState.wb, AppState.activeSheet, addr);
            value = (result && typeof result === 'object' && result.error) ? '#ERROR!' : (result || '');
          } else {
            value = '=' + cellData.f;
          }
        } catch (error) {
          value = '#ERROR!';
        }
      } else {
        value = cellData.v || '';
      }
    }
    
    input.value = String(value);
    cellElement.classList.toggle('has-formula', hasFormula);
  }
  
  updateScrollArea() {
    const ws = getWorksheet();
    if (!ws || !ws['!ref']) return;
    
    const range = window.XLSX.utils.decode_range(ws['!ref']);
    const totalHeight = (range.e.r + 1) * this.options.rowHeight + this.options.headerHeight;
    const totalWidth = (range.e.c + 1) * this.options.colWidth + 42; // 42px for row headers
    
    this.elements.canvas.style.height = totalHeight + 'px';
    this.elements.canvas.style.width = totalWidth + 'px';
  }
  
  // Plugin system
  use(plugin) {
    if (typeof plugin.install === 'function') {
      plugin.install(this);
      this.plugins.set(plugin.name, plugin);
    }
  }
  
  // Hook system for extensibility
  addHook(name, callback) {
    if (!this.hooks.has(name)) {
      this.hooks.set(name, []);
    }
    this.hooks.get(name).push(callback);
  }
  
  triggerHook(name, ...args) {
    const callbacks = this.hooks.get(name);
    if (callbacks) {
      callbacks.forEach(callback => callback(...args));
    }
  }
  
  // Public API
  refresh() {
    this.render(true);
  }
  
  scrollToCell(row, col) {
    const scrollTop = row * this.options.rowHeight;
    const scrollLeft = col * this.options.colWidth;
    
    this.elements.viewport.scrollTop = scrollTop;
    this.elements.viewport.scrollLeft = scrollLeft;
  }
  
  getVisibleRange() {
    return {
      startRow: this.viewport.startRow,
      endRow: this.viewport.endRow,
      startCol: this.viewport.startCol,
      endCol: this.viewport.endCol
    };
  }
}

// Cell Pool for DOM element reuse
class CellPool {
  constructor(initialSize = 500) {
    this.cells = [];
    this.headers = [];
    this.rowHeaders = [];
    this.maxSize = initialSize * 2;
    
    // Pre-create initial pool
    for (let i = 0; i < initialSize; i++) {
      this.cells.push(this.createCell());
    }
  }
  
  createCell() {
    const cell = document.createElement('div');
    cell.className = 'excel-cell';
    
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'excel-cell-input';
    input.style.cssText = `
      width: 100%;
      height: 100%;
      border: 1px solid #d0d7de;
      background: transparent;
      font-family: 'Calibri', sans-serif;
      font-size: 11px;
      padding: 1px 2px;
      outline: none;
      vertical-align: bottom;
    `;
    
    cell.appendChild(input);
    return cell;
  }
  
  get() {
    return this.cells.pop() || this.createCell();
  }
  
  release(cell) {
    if (this.cells.length < this.maxSize) {
      // Clean up cell before returning to pool
      const input = cell.querySelector('input');
      if (input) {
        input.value = '';
        input.readOnly = false;
      }
      
      cell.className = 'excel-cell';
      delete cell.dataset.cell;
      delete cell.dataset.row;
      delete cell.dataset.col;
      
      this.cells.push(cell);
    }
  }
  
  getHeader() {
    return this.headers.pop();
  }
  
  releaseHeader(header) {
    if (this.headers.length < 100) {
      this.headers.push(header);
    }
  }
  
  getRowHeader() {
    return this.rowHeaders.pop();
  }
  
  releaseRowHeader(header) {
    if (this.rowHeaders.length < 100) {
      this.rowHeaders.push(header);
    }
  }
}