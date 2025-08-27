/**
 * Enhanced Excel Grid - Main Integration Module
 * Entry point for the new high-performance Excel grid system
 */

import { ExcelGridEngine } from './grid-engine.js';
import { VirtualGrid } from './virtual-grid.js';
import { GridEventManager } from './event-manager.js';
import { GridPluginManager, AdvancedFormulaPlugin, CSVPlugin, PerformancePlugin } from './plugin-system.js';
import { GridPerformanceMonitor } from './performance-monitor.js';

/**
 * Enhanced Excel Grid Factory
 * Creates a complete Excel-like grid system with all features
 */
export class EnhancedExcelGrid {
  constructor(container, options = {}) {
    this.container = container;
    this.options = {
      // Performance options
      enableVirtualScrolling: true,
      enablePerformanceMonitoring: true,
      
      // Grid options  
      rowHeight: 20,
      colWidth: 64,
      headerHeight: 20,
      bufferRows: 10,
      bufferCols: 10,
      
      // Plugin options
      enableAdvancedFormulas: true,
      enableCSVSupport: true,
      preloadExternalLibs: true,
      
      // Event options
      enableKeyboardShortcuts: true,
      enableContextMenu: true,
      enableMultiSelect: true,
      
      // UI options
      excelTheme: true,
      enableResizing: true,
      enableFreezePane: false,
      
      ...options
    };
    
    this.components = {
      grid: null,
      eventManager: null,
      pluginManager: null,
      performanceMonitor: null
    };
    
    this.isInitialized = false;
    this.isDestroyed = false;
  }
  
  /**
   * Initialize the complete grid system
   */
  async init() {
    if (this.isInitialized) {
      console.warn('Enhanced Excel Grid already initialized');
      return this;
    }
    
    try {
      console.log('🚀 Initializing Enhanced Excel Grid System...');
      
      // Step 1: Initialize core grid
      await this.initializeGrid();
      
      // Step 2: Initialize event system
      this.initializeEventSystem();
      
      // Step 3: Initialize plugin system
      await this.initializePluginSystem();
      
      // Step 4: Initialize performance monitoring
      this.initializePerformanceMonitoring();
      
      // Step 5: Apply Excel theme
      this.applyExcelTheme();
      
      this.isInitialized = true;
      
      console.log('✅ Enhanced Excel Grid System initialized successfully');
      
      // Emit initialization event
      this.container.dispatchEvent(new CustomEvent('excelGridReady', {
        detail: {
          grid: this.components.grid,
          features: this.getEnabledFeatures()
        }
      }));
      
      return this;
      
    } catch (error) {
      console.error('❌ Failed to initialize Enhanced Excel Grid:', error);
      throw error;
    }
  }
  
  async initializeGrid() {
    if (this.options.enableVirtualScrolling) {
      console.log('📊 Initializing Virtual Grid...');
      this.components.grid = new VirtualGrid(this.container, {
        rowHeight: this.options.rowHeight,
        colWidth: this.options.colWidth,
        headerHeight: this.options.headerHeight,
        bufferRows: this.options.bufferRows,
        bufferCols: this.options.bufferCols,
        overscan: 5,
        smoothScrolling: true,
        recycleThreshold: 100
      });
    } else {
      console.log('📊 Initializing Standard Grid...');
      this.components.grid = new ExcelGridEngine(this.container, {
        rowHeight: this.options.rowHeight,
        colWidth: this.options.colWidth,
        headerHeight: this.options.headerHeight,
        bufferSize: this.options.bufferRows
      });
    }
  }
  
  initializeEventSystem() {
    console.log('🎯 Initializing Event System...');
    this.components.eventManager = new GridEventManager(this.components.grid);
  }
  
  async initializePluginSystem() {
    console.log('🔌 Initializing Plugin System...');
    this.components.pluginManager = new GridPluginManager(this.components.grid);
    
    // Register core plugins
    if (this.options.enableAdvancedFormulas) {
      this.components.pluginManager.register(AdvancedFormulaPlugin);
    }
    
    if (this.options.enableCSVSupport) {
      this.components.pluginManager.register(CSVPlugin);
    }
    
    this.components.pluginManager.register(PerformancePlugin);
    
    // Preload external libraries
    if (this.options.preloadExternalLibs) {
      try {
        await this.components.pluginManager.preloadCore();
        console.log('📚 External libraries preloaded successfully');
      } catch (error) {
        console.warn('⚠️ Some external libraries failed to preload:', error);
      }
    }
  }
  
  initializePerformanceMonitoring() {
    if (this.options.enablePerformanceMonitoring) {
      console.log('📈 Initializing Performance Monitoring...');
      this.components.performanceMonitor = new GridPerformanceMonitor(this.components.grid, {
        enableVisualIndicators: this.options.debug || false,
        enableConsoleReports: this.options.debug || false,
        trackRender: true,
        trackScroll: true,
        trackMemory: true,
        trackInput: true
      });
    }
  }
  
  applyExcelTheme() {
    if (this.options.excelTheme) {
      console.log('🎨 Applying Excel Theme...');
      this.container.classList.add('excel-grid-container');
      
      // Ensure Excel theme CSS is loaded
      if (!document.querySelector('link[href*="excel-theme.css"]')) {
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = './styles/excel-theme.css';
        document.head.appendChild(link);
      }
    }
  }
  
  /**
   * Get information about enabled features
   */
  getEnabledFeatures() {
    return {
      virtualScrolling: this.options.enableVirtualScrolling,
      performanceMonitoring: this.options.enablePerformanceMonitoring,
      advancedFormulas: this.options.enableAdvancedFormulas,
      csvSupport: this.options.enableCSVSupport,
      keyboardShortcuts: this.options.enableKeyboardShortcuts,
      contextMenu: this.options.enableContextMenu,
      multiSelect: this.options.enableMultiSelect,
      excelTheme: this.options.excelTheme,
      resizing: this.options.enableResizing,
      freezePane: this.options.enableFreezePane
    };
  }
  
  /**
   * Add a custom plugin
   */
  addPlugin(plugin) {
    if (!this.components.pluginManager) {
      throw new Error('Plugin system not initialized');
    }
    
    return this.components.pluginManager.register(plugin);
  }
  
  /**
   * Load an external library
   */
  async loadExternalLibrary(libName, options = {}) {
    if (!this.components.pluginManager) {
      throw new Error('Plugin system not initialized');
    }
    
    return this.components.pluginManager.loadExternal(libName, options);
  }
  
  /**
   * Get performance metrics
   */
  getPerformanceMetrics() {
    if (!this.components.performanceMonitor) {
      return null;
    }
    
    return this.components.performanceMonitor.getMetrics();
  }
  
  /**
   * Start performance profiling
   */
  startProfiling(duration = 10000) {
    if (!this.components.performanceMonitor) {
      console.warn('Performance monitoring not enabled');
      return null;
    }
    
    return this.components.performanceMonitor.startProfiling(duration);
  }
  
  /**
   * Export performance report
   */
  exportPerformanceReport() {
    if (!this.components.performanceMonitor) {
      console.warn('Performance monitoring not enabled');
      return null;
    }
    
    return this.components.performanceMonitor.exportReport();
  }
  
  /**
   * Refresh the grid
   */
  refresh() {
    if (this.components.grid && this.components.grid.refresh) {
      this.components.grid.refresh();
    }
  }
  
  /**
   * Scroll to a specific cell
   */
  scrollToCell(row, col) {
    if (this.components.grid && this.components.grid.scrollToCell) {
      this.components.grid.scrollToCell(row, col);
    }
  }
  
  /**
   * Get the visible range
   */
  getVisibleRange() {
    if (this.components.grid && this.components.grid.getVisibleRange) {
      return this.components.grid.getVisibleRange();
    }
    return null;
  }
  
  /**
   * Import CSV data
   */
  async importCSV(csvText) {
    if (!this.components.pluginManager) {
      throw new Error('Plugin system not initialized');
    }
    
    const csvPlugin = this.components.pluginManager.get('csvSupport');
    if (!csvPlugin) {
      throw new Error('CSV plugin not loaded');
    }
    
    return this.components.grid.importCSV(csvText);
  }
  
  /**
   * Export to CSV
   */
  async exportCSV(range = null) {
    if (!this.components.pluginManager) {
      throw new Error('Plugin system not initialized');
    }
    
    const csvPlugin = this.components.pluginManager.get('csvSupport');
    if (!csvPlugin) {
      throw new Error('CSV plugin not loaded');
    }
    
    return this.components.grid.exportCSV(range);
  }
  
  /**
   * Resize the grid container
   */
  resize() {
    if (this.components.grid) {\n      // Trigger resize handling\n      const resizeEvent = new Event('resize');\n      window.dispatchEvent(resizeEvent);\n      \n      if (this.components.grid.refresh) {\n        this.components.grid.refresh();\n      }\n    }\n  }\n  \n  /**\n   * Enable/disable specific features\n   */\n  toggleFeature(featureName, enabled) {\n    switch (featureName) {\n      case 'performanceMonitoring':\n        if (enabled && !this.components.performanceMonitor) {\n          this.initializePerformanceMonitoring();\n        } else if (!enabled && this.components.performanceMonitor) {\n          this.components.performanceMonitor.destroy();\n          this.components.performanceMonitor = null;\n        }\n        break;\n        \n      case 'excelTheme':\n        if (enabled) {\n          this.applyExcelTheme();\n        } else {\n          this.container.classList.remove('excel-grid-container');\n        }\n        break;\n        \n      default:\n        console.warn(`Unknown feature: ${featureName}`);\n    }\n    \n    this.options[`enable${featureName.charAt(0).toUpperCase() + featureName.slice(1)}`] = enabled;\n  }\n  \n  /**\n   * Get grid statistics\n   */\n  getStats() {\n    const stats = {\n      isInitialized: this.isInitialized,\n      isDestroyed: this.isDestroyed,\n      enabledFeatures: this.getEnabledFeatures(),\n      components: {\n        grid: !!this.components.grid,\n        eventManager: !!this.components.eventManager,\n        pluginManager: !!this.components.pluginManager,\n        performanceMonitor: !!this.components.performanceMonitor\n      }\n    };\n    \n    if (this.components.grid && this.components.grid.getPerformanceStats) {\n      stats.gridStats = this.components.grid.getPerformanceStats();\n    }\n    \n    if (this.components.eventManager && this.components.eventManager.getEventStats) {\n      stats.eventStats = this.components.eventManager.getEventStats();\n    }\n    \n    if (this.components.performanceMonitor) {\n      stats.performanceMetrics = this.components.performanceMonitor.getMetrics();\n    }\n    \n    return stats;\n  }\n  \n  /**\n   * Destroy the grid system and clean up resources\n   */\n  destroy() {\n    if (this.isDestroyed) {\n      console.warn('Enhanced Excel Grid already destroyed');\n      return;\n    }\n    \n    console.log('🧹 Destroying Enhanced Excel Grid System...');\n    \n    // Destroy components in reverse order\n    if (this.components.performanceMonitor) {\n      this.components.performanceMonitor.destroy();\n    }\n    \n    if (this.components.pluginManager) {\n      // Unregister all plugins\n      const plugins = this.components.pluginManager.list();\n      plugins.forEach(pluginName => {\n        this.components.pluginManager.unregister(pluginName);\n      });\n    }\n    \n    if (this.components.eventManager) {\n      this.components.eventManager.destroy();\n    }\n    \n    if (this.components.grid && this.components.grid.destroy) {\n      this.components.grid.destroy();\n    }\n    \n    // Clear references\n    this.components = {\n      grid: null,\n      eventManager: null,\n      pluginManager: null,\n      performanceMonitor: null\n    };\n    \n    this.isDestroyed = true;\n    this.isInitialized = false;\n    \n    // Emit destruction event\n    this.container.dispatchEvent(new CustomEvent('excelGridDestroyed'));\n    \n    console.log('✅ Enhanced Excel Grid System destroyed');\n  }\n}\n\n/**\n * Factory function to create an Enhanced Excel Grid\n */\nexport function createEnhancedExcelGrid(container, options = {}) {\n  return new EnhancedExcelGrid(container, options);\n}\n\n/**\n * Auto-initialize if container with specific ID exists\n */\nexport function autoInitialize() {\n  const container = document.getElementById('spreadsheet');\n  if (container && !container._enhancedExcelGrid) {\n    const grid = new EnhancedExcelGrid(container);\n    container._enhancedExcelGrid = grid;\n    \n    // Initialize when DOM is ready\n    if (document.readyState === 'loading') {\n      document.addEventListener('DOMContentLoaded', () => grid.init());\n    } else {\n      grid.init();\n    }\n    \n    return grid;\n  }\n  return null;\n}\n\n// Export all components for advanced usage\nexport {\n  ExcelGridEngine,\n  VirtualGrid,\n  GridEventManager,\n  GridPluginManager,\n  GridPerformanceMonitor,\n  AdvancedFormulaPlugin,\n  CSVPlugin,\n  PerformancePlugin\n};