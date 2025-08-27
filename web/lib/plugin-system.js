/**
 * External JS Library Integration System
 * Supports dynamic loading and integration of external libraries
 */

export class GridPluginManager {
  constructor(grid) {
    this.grid = grid;
    this.plugins = new Map();
    this.externalLibs = new Map();
    this.loadingPromises = new Map();
    
    // CDN URLs for external libraries
    this.cdnUrls = {
      hyperformula: 'https://cdn.jsdelivr.net/npm/hyperformula@2.6.2/dist/hyperformula.min.js',
      luckysheet: 'https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/js/plugin.min.js',
      agGrid: 'https://cdn.jsdelivr.net/npm/ag-grid-community@31.1.1/dist/ag-grid-community.min.js',
      chartjs: 'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js',
      papaparse: 'https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js',
      rxjs: 'https://cdn.jsdelivr.net/npm/rxjs@7.8.1/dist/bundles/rxjs.umd.min.js',
      lodash: 'https://cdn.jsdelivr.net/npm/lodash@4.17.21/lodash.min.js',
      moment: 'https://cdn.jsdelivr.net/npm/moment@2.30.1/moment.min.js'
    };
    
    this.localFallbacks = {
      hyperformula: '/web/external/hyperformula.min.js',
      luckysheet: '/web/external/luckysheet.min.js',
      agGrid: '/web/external/ag-grid.min.js'
    };
  }
  
  /**
   * Load external library dynamically
   */
  async loadExternal(libName, options = {}) {
    // Return cached promise if already loading
    if (this.loadingPromises.has(libName)) {
      return this.loadingPromises.get(libName);
    }
    
    // Return cached library if already loaded
    if (this.externalLibs.has(libName)) {
      return this.externalLibs.get(libName);
    }
    
    const loadPromise = this._loadLibrary(libName, options);
    this.loadingPromises.set(libName, loadPromise);
    
    try {
      const lib = await loadPromise;
      this.externalLibs.set(libName, lib);
      this.loadingPromises.delete(libName);
      
      console.log(`✅ Loaded external library: ${libName}`);
      return lib;
    } catch (error) {
      this.loadingPromises.delete(libName);
      throw error;
    }
  }
  
  async _loadLibrary(libName, options) {
    const cdnUrl = this.cdnUrls[libName];
    const fallbackUrl = this.localFallbacks[libName];
    
    if (!cdnUrl) {
      throw new Error(`Unknown library: ${libName}`);
    }
    
    // Try CDN first
    try {
      return await this._loadScript(cdnUrl, libName, options);
    } catch (cdnError) {
      console.warn(`CDN failed for ${libName}, trying fallback:`, cdnError);
      
      // Try local fallback
      if (fallbackUrl) {
        try {
          return await this._loadScript(fallbackUrl, libName, options);
        } catch (fallbackError) {
          throw new Error(`Failed to load ${libName} from both CDN and fallback: ${fallbackError.message}`);
        }
      } else {
        throw cdnError;
      }
    }
  }
  
  _loadScript(url, libName, options) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = url;
      script.async = true;
      
      const timeout = setTimeout(() => {
        reject(new Error(`Timeout loading ${libName} from ${url}`));
      }, options.timeout || 10000);
      
      script.onload = () => {
        clearTimeout(timeout);
        
        // Get library from global scope
        const globalName = options.globalName || this._getGlobalName(libName);
        const lib = window[globalName];
        
        if (!lib) {
          reject(new Error(`Library ${libName} not found in global scope as ${globalName}`));
          return;
        }
        
        resolve(lib);
      };
      
      script.onerror = () => {
        clearTimeout(timeout);
        reject(new Error(`Failed to load script: ${url}`));
      };
      
      document.head.appendChild(script);
    });
  }
  
  _getGlobalName(libName) {
    const globalNames = {
      hyperformula: 'HyperFormula',
      luckysheet: 'luckysheet',
      agGrid: 'agGrid',
      chartjs: 'Chart',
      papaparse: 'Papa',
      rxjs: 'rxjs',
      lodash: '_',
      moment: 'moment'
    };
    
    return globalNames[libName] || libName;
  }
  
  /**
   * Create HyperFormula integration
   */
  async createFormulaEngine() {
    const HyperFormula = await this.loadExternal('hyperformula');
    
    const engine = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3',
      useColumnIndex: true,
      useArrayArithmetic: true,
      smartRounding: true,
      useRegularExpressions: true
    });
    
    return {
      engine,
      calculate: (formula, context = {}) => {
        try {
          // Add formula to engine
          const sheetId = engine.addSheet('temp');
          const cellId = engine.addSheet ? engine.setCellContents({ sheet: sheetId, row: 0, col: 0 }, formula) : null;
          
          const result = engine.getCellValue({ sheet: sheetId, row: 0, col: 0 });
          engine.removeSheet(sheetId);
          
          return result;
        } catch (error) {
          return { error: error.message };
        }
      },
      validateFormula: (formula) => {
        try {
          return engine.isFormula(formula);
        } catch {
          return false;
        }
      }
    };
  }
  
  /**
   * Create Chart.js integration
   */
  async createChartEngine() {
    const Chart = await this.loadExternal('chartjs');
    
    return {
      Chart,
      createChart: (canvas, config) => {
        return new Chart(canvas.getContext('2d'), config);
      },
      getChartTypes: () => ['line', 'bar', 'pie', 'doughnut', 'scatter', 'area']
    };
  }
  
  /**
   * Create CSV parser integration
   */
  async createCSVParser() {
    const Papa = await this.loadExternal('papaparse');
    
    return {
      parse: (csvText, options = {}) => {
        return Papa.parse(csvText, {
          header: true,
          skipEmptyLines: true,
          dynamicTyping: true,
          ...options
        });
      },
      unparse: (data, options = {}) => {
        return Papa.unparse(data, options);
      }
    };
  }
  
  /**
   * Create reactive programming utilities
   */
  async createRxEngine() {
    const rxjs = await this.loadExternal('rxjs');
    
    return {
      rxjs,
      createCellStream: (cellAddress) => {
        return new rxjs.Subject();
      },
      combineFormulas: (...streams) => {
        return rxjs.combineLatest(streams);
      }
    };
  }
  
  /**
   * Plugin registration system
   */
  register(plugin) {
    if (!plugin.name || typeof plugin.install !== 'function') {
      throw new Error('Plugin must have name and install function');
    }
    
    if (this.plugins.has(plugin.name)) {
      console.warn(`Plugin ${plugin.name} already registered, overriding...`);
    }
    
    this.plugins.set(plugin.name, plugin);
    
    // Install plugin
    try {
      plugin.install(this.grid, this);
      console.log(`✅ Plugin registered: ${plugin.name}`);
    } catch (error) {
      console.error(`❌ Plugin installation failed: ${plugin.name}`, error);
      this.plugins.delete(plugin.name);
      throw error;
    }
  }
  
  /**
   * Get registered plugin
   */
  get(pluginName) {
    return this.plugins.get(pluginName);
  }
  
  /**
   * Unregister plugin
   */
  unregister(pluginName) {
    const plugin = this.plugins.get(pluginName);
    if (plugin && typeof plugin.uninstall === 'function') {
      plugin.uninstall(this.grid, this);
    }
    return this.plugins.delete(pluginName);
  }
  
  /**
   * List all registered plugins
   */
  list() {
    return Array.from(this.plugins.keys());
  }
  
  /**
   * Check if library is loaded
   */
  isLoaded(libName) {
    return this.externalLibs.has(libName);
  }
  
  /**
   * Preload commonly used libraries
   */
  async preloadCore() {
    const coreLibs = ['hyperformula', 'papaparse', 'lodash'];
    
    const results = await Promise.allSettled(
      coreLibs.map(lib => this.loadExternal(lib))
    );
    
    results.forEach((result, index) => {
      if (result.status === 'rejected') {
        console.warn(`Failed to preload ${coreLibs[index]}:`, result.reason);
      }
    });
    
    return results;
  }
}

/**
 * Built-in plugins
 */

// Advanced Formula Plugin
export const AdvancedFormulaPlugin = {
  name: 'advancedFormula',
  async install(grid, pluginManager) {
    const formulaEngine = await pluginManager.createFormulaEngine();
    
    grid.formulaEngine = formulaEngine;
    
    // Override formula calculation
    grid.addHook('beforeCellUpdate', (address, value) => {
      if (value.startsWith('=')) {
        try {
          const result = formulaEngine.calculate(value);
          if (result && result.error) {
            return { error: result.error };
          }
          return { calculated: result };
        } catch (error) {
          return { error: error.message };
        }
      }
    });
  }
};

// CSV Import/Export Plugin
export const CSVPlugin = {
  name: 'csvSupport',
  async install(grid, pluginManager) {
    const csvParser = await pluginManager.createCSVParser();
    
    grid.importCSV = (csvText) => {
      const result = csvParser.parse(csvText);
      if (result.errors.length > 0) {
        throw new Error('CSV parse errors: ' + result.errors.map(e => e.message).join(', '));
      }
      
      // Convert to grid data format
      return result.data;
    };
    
    grid.exportCSV = (range) => {
      const data = grid.getRangeData(range);
      return csvParser.unparse(data);
    };
  }
};

// Performance Monitoring Plugin
export const PerformancePlugin = {
  name: 'performance',
  install(grid, pluginManager) {
    const metrics = {
      renderTimes: [],
      scrollEvents: 0,
      cellUpdates: 0
    };
    
    grid.addHook('afterRender', (renderTime) => {
      metrics.renderTimes.push(renderTime);
      if (metrics.renderTimes.length > 100) {
        metrics.renderTimes.shift();
      }
    });
    
    grid.addHook('onScroll', () => {
      metrics.scrollEvents++;
    });
    
    grid.addHook('afterCellUpdate', () => {
      metrics.cellUpdates++;
    });
    
    grid.getPerformanceMetrics = () => ({
      averageRenderTime: metrics.renderTimes.reduce((a, b) => a + b, 0) / metrics.renderTimes.length,
      scrollEvents: metrics.scrollEvents,
      cellUpdates: metrics.cellUpdates
    });
  }
};