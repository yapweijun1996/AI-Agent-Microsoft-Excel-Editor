/**
 * Virtual Scrolling Grid Engine
 * High-performance virtualized grid for large datasets
 */

import { AppState } from '../core/state.js';
import { getWorksheet } from '../spreadsheet/workbook-manager.js';

export class VirtualGrid {
  constructor(container, options = {}) {
    this.container = container;
    this.options = {
      rowHeight: 20,
      colWidth: 64,
      headerHeight: 20,
      bufferRows: 10,
      bufferCols: 10,
      overscan: 5, // Extra rows/cols to render outside viewport
      recycleThreshold: 100, // Recycle DOM elements when pool gets large
      smoothScrolling: true,
      ...options
    };
    
    // Viewport state
    this.viewport = {
      width: 0,
      height: 0,
      scrollTop: 0,
      scrollLeft: 0,
      startRow: 0,
      endRow: 0,
      startCol: 0,
      endCol: 0,
      visibleRows: 0,
      visibleCols: 0
    };
    
    // Virtual state
    this.virtualState = {
      totalRows: 1000,
      totalCols: 100,
      virtualHeight: 0,
      virtualWidth: 0
    };
    
    // DOM element pools for recycling
    this.elementPools = {
      cells: [],
      rowHeaders: [],
      colHeaders: []
    };
    
    // Rendered elements tracking
    this.renderedElements = {
      cells: new Map(), // key: "row,col" -> element
      rowHeaders: new Map(), // key: row -> element  
      colHeaders: new Map() // key: col -> element
    };
    
    // Performance tracking
    this.performance = {
      lastRenderTime: 0,
      renderCount: 0,
      scrollEvents: 0
    };
    
    this.init();
  }
  
  init() {
    this.setupVirtualContainer();
    this.calculateVirtualSize();
    this.bindEvents();
    this.render();
  }
  
  setupVirtualContainer() {
    this.container.className = 'virtual-grid-container';
    this.container.innerHTML = `
      <div class="virtual-scrollable" style="
        position: relative;
        overflow: auto;
        width: 100%;
        height: 100%;
      ">
        <div class="virtual-spacer" style="
          position: absolute;
          top: 0;
          left: 0;
          pointer-events: none;
        "></div>
        <div class="virtual-content" style="
          position: absolute;
          top: 0;
          left: 0;
          will-change: transform;
        ">
          <div class="virtual-headers"></div>
          <div class="virtual-body"></div>
        </div>
      </div>
    `;
    
    // Cache DOM references
    this.elements = {
      scrollable: this.container.querySelector('.virtual-scrollable'),
      spacer: this.container.querySelector('.virtual-spacer'),
      content: this.container.querySelector('.virtual-content'),
      headers: this.container.querySelector('.virtual-headers'),
      body: this.container.querySelector('.virtual-body')
    };
    
    // Set initial dimensions
    this.viewport.width = this.container.clientWidth;
    this.viewport.height = this.container.clientHeight;
    
    this.calculateVisibleRange();
  }
  
  calculateVirtualSize() {
    const ws = getWorksheet();
    if (ws && ws['!ref']) {
      const range = window.XLSX.utils.decode_range(ws['!ref']);
      this.virtualState.totalRows = Math.max(range.e.r + 1, 100);
      this.virtualState.totalCols = Math.max(range.e.c + 1, 26);
    }
    
    this.virtualState.virtualHeight = this.virtualState.totalRows * this.options.rowHeight + this.options.headerHeight;
    this.virtualState.virtualWidth = this.virtualState.totalCols * this.options.colWidth + 42; // 42px for row headers
    
    // Update spacer to create scrollable area
    this.elements.spacer.style.height = this.virtualState.virtualHeight + 'px';
    this.elements.spacer.style.width = this.virtualState.virtualWidth + 'px';
  }
  
  calculateVisibleRange() {
    // Calculate how many rows/cols can fit in viewport
    this.viewport.visibleRows = Math.ceil((this.viewport.height - this.options.headerHeight) / this.options.rowHeight);
    this.viewport.visibleCols = Math.ceil((this.viewport.width - 42) / this.options.colWidth);
    
    // Calculate which rows/cols should be rendered (with overscan)
    const startRow = Math.floor(this.viewport.scrollTop / this.options.rowHeight);
    const startCol = Math.floor(this.viewport.scrollLeft / this.options.colWidth);
    
    this.viewport.startRow = Math.max(0, startRow - this.options.overscan);
    this.viewport.endRow = Math.min(
      this.virtualState.totalRows,
      startRow + this.viewport.visibleRows + this.options.overscan * 2
    );
    
    this.viewport.startCol = Math.max(0, startCol - this.options.overscan);
    this.viewport.endCol = Math.min(
      this.virtualState.totalCols,
      startCol + this.viewport.visibleCols + this.options.overscan * 2
    );
  }
  
  bindEvents() {
    let rafId;
    let isScrolling = false;
    
    this.elements.scrollable.addEventListener('scroll', (e) => {
      this.performance.scrollEvents++;
      
      if (rafId) cancelAnimationFrame(rafId);
      
      if (!isScrolling) {
        isScrolling = true;
        this.container.classList.add('scrolling');
      }
      
      rafId = requestAnimationFrame(() => {
        this.handleScroll(e);
        
        // Debounce scroll end
        setTimeout(() => {
          if (isScrolling) {
            isScrolling = false;
            this.container.classList.remove('scrolling');
          }
        }, 150);
      });
    });
    
    // Handle resize
    const resizeObserver = new ResizeObserver(entries => {
      for (const entry of entries) {
        this.viewport.width = entry.contentRect.width;
        this.viewport.height = entry.contentRect.height;
        this.handleResize();
      }
    });
    
    resizeObserver.observe(this.container);
  }
  
  handleScroll(e) {
    const newScrollTop = e.target.scrollTop;
    const newScrollLeft = e.target.scrollLeft;
    
    const scrollChanged = (
      Math.abs(this.viewport.scrollTop - newScrollTop) > this.options.rowHeight / 2 ||
      Math.abs(this.viewport.scrollLeft - newScrollLeft) > this.options.colWidth / 2
    );
    
    this.viewport.scrollTop = newScrollTop;
    this.viewport.scrollLeft = newScrollLeft;
    
    if (scrollChanged) {
      const oldRange = { ...this.viewport };
      this.calculateVisibleRange();
      
      // Only re-render if visible range changed significantly
      if (this.rangeChanged(oldRange)) {
        this.render();
      }
    }
    
    // Update content position for smooth scrolling
    if (this.options.smoothScrolling) {
      const translateY = this.viewport.startRow * this.options.rowHeight;
      const translateX = this.viewport.startCol * this.options.colWidth;
      this.elements.content.style.transform = `translate(${translateX}px, ${translateY}px)`;
    }
  }
  
  handleResize() {
    this.calculateVisibleRange();
    this.render();
  }
  
  rangeChanged(oldRange) {
    return (
      oldRange.startRow !== this.viewport.startRow ||
      oldRange.endRow !== this.viewport.endRow ||
      oldRange.startCol !== this.viewport.startCol ||
      oldRange.endCol !== this.viewport.endCol
    );
  }
  
  render() {
    const startTime = performance.now();
    
    try {
      this.renderHeaders();
      this.renderCells();
      this.recycleUnusedElements();
      
      this.performance.renderCount++;
      this.performance.lastRenderTime = performance.now() - startTime;
      
      // Emit render event
      this.container.dispatchEvent(new CustomEvent('virtualRender', {
        detail: {
          renderTime: this.performance.lastRenderTime,
          visibleRange: this.getVisibleRange()
        }
      }));
      
    } catch (error) {
      console.error('Virtual grid render error:', error);
    }
  }
  
  renderHeaders() {
    // Clear existing headers
    this.elements.headers.innerHTML = '';
    
    const fragment = document.createDocumentFragment();
    
    // Corner header
    const corner = this.createElement('div', 'virtual-corner', {
      position: 'absolute',
      left: '0px',
      top: '0px',
      width: '42px',
      height: this.options.headerHeight + 'px',
      background: 'linear-gradient(to bottom, #f6f8fa 0%, #e9ecef 100%)',
      border: '1px solid #c6cbd1',
      zIndex: '10'
    });
    fragment.appendChild(corner);
    
    // Column headers
    for (let col = this.viewport.startCol; col < this.viewport.endCol; col++) {
      const colLetter = window.XLSX.utils.encode_col(col);
      const header = this.getOrCreateColHeader(col);
      
      header.textContent = colLetter;
      header.style.cssText = `
        position: absolute;
        left: ${42 + (col - this.viewport.startCol) * this.options.colWidth}px;
        top: 0px;
        width: ${this.options.colWidth}px;
        height: ${this.options.headerHeight}px;
        background: linear-gradient(to bottom, #f6f8fa 0%, #e9ecef 100%);
        border: 1px solid #c6cbd1;
        text-align: center;
        line-height: ${this.options.headerHeight}px;
        font: 11px Calibri, sans-serif;
        z-index: 9;
      `;
      
      fragment.appendChild(header);
    }
    
    this.elements.headers.appendChild(fragment);
  }
  
  renderCells() {
    const ws = getWorksheet();
    if (!ws) return;
    
    // Clear existing body
    this.elements.body.innerHTML = '';
    const fragment = document.createDocumentFragment();
    
    // Render visible rows
    for (let row = this.viewport.startRow; row < this.viewport.endRow; row++) {
      // Row header
      const rowHeader = this.getOrCreateRowHeader(row);
      rowHeader.textContent = (row + 1).toString();
      rowHeader.style.cssText = `
        position: absolute;
        left: 0px;
        top: ${(row - this.viewport.startRow) * this.options.rowHeight + this.options.headerHeight}px;
        width: 42px;
        height: ${this.options.rowHeight}px;
        background: linear-gradient(to bottom, #f6f8fa 0%, #e9ecef 100%);
        border: 1px solid #c6cbd1;
        text-align: center;
        line-height: ${this.options.rowHeight}px;
        font: 11px Calibri, sans-serif;
        z-index: 8;
      `;
      fragment.appendChild(rowHeader);
      
      // Cells in row
      for (let col = this.viewport.startCol; col < this.viewport.endCol; col++) {
        const addr = window.XLSX.utils.encode_cell({ r: row, c: col });
        const cellData = ws[addr];
        
        const cell = this.getOrCreateCell(row, col);
        this.updateCellContent(cell, addr, cellData, row, col);
        
        cell.style.cssText = `
          position: absolute;
          left: ${42 + (col - this.viewport.startCol) * this.options.colWidth}px;
          top: ${(row - this.viewport.startRow) * this.options.rowHeight + this.options.headerHeight}px;
          width: ${this.options.colWidth}px;
          height: ${this.options.rowHeight}px;
        `;
        
        fragment.appendChild(cell);
      }
    }
    
    this.elements.body.appendChild(fragment);
  }
  
  getOrCreateCell(row, col) {
    const key = `${row},${col}`;
    let cell = this.renderedElements.cells.get(key);
    
    if (!cell) {
      cell = this.elementPools.cells.pop() || this.createCellElement();
      this.renderedElements.cells.set(key, cell);
    }
    
    return cell;
  }
  
  getOrCreateRowHeader(row) {
    let header = this.renderedElements.rowHeaders.get(row);
    
    if (!header) {
      header = this.elementPools.rowHeaders.pop() || this.createRowHeaderElement();
      this.renderedElements.rowHeaders.set(row, header);
    }
    
    return header;
  }
  
  getOrCreateColHeader(col) {
    let header = this.renderedElements.colHeaders.get(col);
    
    if (!header) {
      header = this.elementPools.colHeaders.pop() || this.createColHeaderElement();
      this.renderedElements.colHeaders.set(col, header);
    }
    
    return header;
  }
  
  createCellElement() {
    const cell = document.createElement('div');
    cell.className = 'virtual-cell';
    
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'virtual-cell-input';
    input.style.cssText = `
      width: 100%;
      height: 100%;
      border: 1px solid #d0d7de;
      background: white;
      font: 11px Calibri, sans-serif;
      padding: 1px 2px;
      outline: none;
      box-sizing: border-box;
    `;
    
    cell.appendChild(input);
    return cell;
  }
  
  createRowHeaderElement() {
    const header = document.createElement('div');
    header.className = 'virtual-row-header';
    return header;
  }
  
  createColHeaderElement() {
    const header = document.createElement('div');
    header.className = 'virtual-col-header';
    return header;
  }
  
  updateCellContent(cellElement, addr, cellData, row, col) {
    const input = cellElement.querySelector('.virtual-cell-input');
    
    cellElement.dataset.cell = addr;
    cellElement.dataset.row = row;
    cellElement.dataset.col = col;
    
    let value = '';
    let hasFormula = false;
    
    if (cellData) {
      if (cellData.f) {
        hasFormula = true;
        try {
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
  
  recycleUnusedElements() {
    // Recycle cells that are no longer visible
    const currentCells = new Set();
    for (let row = this.viewport.startRow; row < this.viewport.endRow; row++) {
      for (let col = this.viewport.startCol; col < this.viewport.endCol; col++) {
        currentCells.add(`${row},${col}`);
      }
    }
    
    for (const [key, cell] of this.renderedElements.cells) {
      if (!currentCells.has(key)) {
        this.recycleCell(key, cell);
      }
    }
    
    // Recycle headers
    const currentRows = new Set();
    for (let row = this.viewport.startRow; row < this.viewport.endRow; row++) {
      currentRows.add(row);
    }
    
    for (const [row, header] of this.renderedElements.rowHeaders) {
      if (!currentRows.has(row)) {
        this.recycleRowHeader(row, header);
      }
    }
    
    const currentCols = new Set();
    for (let col = this.viewport.startCol; col < this.viewport.endCol; col++) {
      currentCols.add(col);
    }
    
    for (const [col, header] of this.renderedElements.colHeaders) {
      if (!currentCols.has(col)) {
        this.recycleColHeader(col, header);
      }
    }
  }
  
  recycleCell(key, cell) {
    this.renderedElements.cells.delete(key);
    
    // Clean up cell
    const input = cell.querySelector('input');
    if (input) {
      input.value = '';
      input.readOnly = false;
    }
    
    delete cell.dataset.cell;
    delete cell.dataset.row;
    delete cell.dataset.col;
    cell.className = 'virtual-cell';
    
    if (this.elementPools.cells.length < this.options.recycleThreshold) {
      this.elementPools.cells.push(cell);
    }
  }
  
  recycleRowHeader(row, header) {
    this.renderedElements.rowHeaders.delete(row);
    
    header.textContent = '';
    if (this.elementPools.rowHeaders.length < 50) {
      this.elementPools.rowHeaders.push(header);
    }
  }
  
  recycleColHeader(col, header) {
    this.renderedElements.colHeaders.delete(col);
    
    header.textContent = '';
    if (this.elementPools.colHeaders.length < 50) {
      this.elementPools.colHeaders.push(header);
    }
  }
  
  createElement(tag, className, styles) {
    const element = document.createElement(tag);
    if (className) element.className = className;
    if (styles) {
      Object.assign(element.style, styles);
    }
    return element;
  }
  
  // Public API
  scrollToCell(row, col) {
    const scrollTop = row * this.options.rowHeight;
    const scrollLeft = col * this.options.colWidth;
    
    this.elements.scrollable.scrollTo({
      top: scrollTop,
      left: scrollLeft,
      behavior: 'smooth'
    });
  }
  
  getVisibleRange() {
    return {
      startRow: this.viewport.startRow,
      endRow: this.viewport.endRow,
      startCol: this.viewport.startCol,
      endCol: this.viewport.endCol
    };
  }
  
  refresh() {
    this.calculateVirtualSize();
    this.calculateVisibleRange();
    this.render();
  }
  
  getPerformanceStats() {
    return {
      ...this.performance,
      poolSizes: {
        cells: this.elementPools.cells.length,
        rowHeaders: this.elementPools.rowHeaders.length,
        colHeaders: this.elementPools.colHeaders.length
      },
      renderedElements: {
        cells: this.renderedElements.cells.size,
        rowHeaders: this.renderedElements.rowHeaders.size,
        colHeaders: this.renderedElements.colHeaders.size
      }
    };
  }
}