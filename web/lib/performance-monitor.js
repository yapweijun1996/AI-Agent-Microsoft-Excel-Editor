/**
 * Grid Performance Monitor
 * Tracks and reports performance metrics for the Excel grid
 */

export class GridPerformanceMonitor {
  constructor(grid, options = {}) {
    this.grid = grid;
    this.options = {
      trackRender: true,
      trackScroll: true,
      trackMemory: true,
      trackInput: true,
      sampleSize: 100,
      alertThreshold: 16, // 60fps = 16ms per frame
      enableConsoleReports: false,
      enableVisualIndicators: true,
      ...options
    };
    
    // Metrics storage
    this.metrics = {
      render: {
        times: [],
        totalRenders: 0,
        slowRenders: 0,
        averageTime: 0,
        lastRenderTime: 0
      },
      scroll: {
        events: 0,
        smoothEvents: 0,
        laggedEvents: 0,
        averageHandleTime: 0,
        times: []
      },
      input: {
        keystrokes: 0,
        cellUpdates: 0,
        formulaCalculations: 0,
        updateTimes: [],
        averageUpdateTime: 0
      },
      memory: {
        domNodes: 0,
        poolSizes: {},
        memoryUsage: 0,
        lastMemoryCheck: Date.now()
      },
      viewport: {
        visibleCells: 0,
        totalCells: 0,
        recycledCells: 0,
        cacheHitRate: 0
      }
    };
    
    // Performance observers
    this.observers = new Map();
    
    // FPS tracking
    this.fps = {
      frames: 0,
      lastTime: Date.now(),
      current: 60,
      history: []
    };
    
    // Warning thresholds
    this.thresholds = {
      slowRender: 33, // 30fps
      verySlowRender: 66, // 15fps
      highMemory: 50 * 1024 * 1024, // 50MB
      lowFPS: 30
    };
    
    this.init();
  }
  
  init() {
    this.setupPerformanceObservers();
    this.startFPSTracking();
    this.bindGridEvents();
    
    if (this.options.enableVisualIndicators) {
      this.createPerformanceIndicator();
    }
  }
  
  setupPerformanceObservers() {
    // Performance Observer for measuring render times
    if ('PerformanceObserver' in window) {
      const renderObserver = new PerformanceObserver((list) => {
        const entries = list.getEntries();
        entries.forEach(entry => {
          if (entry.name.includes('grid-render')) {
            this.recordRenderTime(entry.duration);
          }
        });
      });
      
      renderObserver.observe({ entryTypes: ['measure'] });
      this.observers.set('render', renderObserver);
    }
    
    // Memory usage tracking
    if ('memory' in performance) {
      this.trackMemoryPeriodically();
    }
  }
  
  startFPSTracking() {
    const trackFPS = () => {
      const now = Date.now();
      this.fps.frames++;
      
      if (now - this.fps.lastTime >= 1000) {
        this.fps.current = Math.round((this.fps.frames * 1000) / (now - this.fps.lastTime));
        this.fps.history.push(this.fps.current);
        
        if (this.fps.history.length > 60) {
          this.fps.history.shift();
        }
        
        this.fps.frames = 0;
        this.fps.lastTime = now;
        
        this.updatePerformanceIndicator();
      }
      
      requestAnimationFrame(trackFPS);
    };
    
    requestAnimationFrame(trackFPS);
  }
  
  bindGridEvents() {
    if (!this.grid.container) return;
    
    // Track render events
    this.grid.container.addEventListener('virtualRender', (e) => {
      this.recordRenderTime(e.detail.renderTime);
      this.trackViewport(e.detail.visibleRange);
    });
    
    // Track scroll events
    const scrollable = this.grid.container.querySelector('.virtual-scrollable, .excel-grid-viewport');
    if (scrollable) {
      let scrollStart = 0;
      
      scrollable.addEventListener('scroll', () => {
        if (scrollStart === 0) {
          scrollStart = performance.now();
        }
        
        this.metrics.scroll.events++;
        
        // Track scroll performance
        requestAnimationFrame(() => {
          if (scrollStart > 0) {
            const scrollTime = performance.now() - scrollStart;
            this.recordScrollTime(scrollTime);
            scrollStart = 0;
          }
        });
      });
    }
    
    // Track input events
    this.grid.container.addEventListener('cellBlur', (e) => {
      this.metrics.input.cellUpdates++;
      
      if (e.detail.value?.startsWith('=')) {
        this.metrics.input.formulaCalculations++;
      }
    });
    
    this.grid.container.addEventListener('keydown', () => {
      this.metrics.input.keystrokes++;
    });
  }
  
  recordRenderTime(duration) {
    if (!this.options.trackRender) return;
    
    this.metrics.render.times.push(duration);
    this.metrics.render.totalRenders++;
    this.metrics.render.lastRenderTime = duration;
    
    if (duration > this.options.alertThreshold) {
      this.metrics.render.slowRenders++;
      
      if (this.options.enableConsoleReports) {
        console.warn(`🐌 Slow render: ${duration.toFixed(2)}ms`);
      }
    }
    
    // Keep only recent samples
    if (this.metrics.render.times.length > this.options.sampleSize) {
      this.metrics.render.times.shift();
    }
    
    // Update average
    this.metrics.render.averageTime = this.metrics.render.times.reduce((a, b) => a + b, 0) / this.metrics.render.times.length;
    
    // Mark render in performance timeline
    if ('performance' in window && 'mark' in performance) {
      performance.mark(`render-${Date.now()}`);
    }
  }
  
  recordScrollTime(duration) {
    if (!this.options.trackScroll) return;
    
    this.metrics.scroll.times.push(duration);
    
    if (duration > this.options.alertThreshold) {
      this.metrics.scroll.laggedEvents++;
    } else {
      this.metrics.scroll.smoothEvents++;
    }
    
    // Keep only recent samples
    if (this.metrics.scroll.times.length > this.options.sampleSize) {
      this.metrics.scroll.times.shift();
    }
    
    // Update average
    this.metrics.scroll.averageHandleTime = this.metrics.scroll.times.reduce((a, b) => a + b, 0) / this.metrics.scroll.times.length;
  }
  
  trackViewport(visibleRange) {
    if (!visibleRange) return;
    
    const visibleCells = (visibleRange.endRow - visibleRange.startRow) * (visibleRange.endCol - visibleRange.startCol);
    this.metrics.viewport.visibleCells = visibleCells;
    
    // Track grid stats if available
    if (this.grid.getPerformanceStats) {
      const gridStats = this.grid.getPerformanceStats();
      this.metrics.memory.poolSizes = gridStats.poolSizes || {};
      this.metrics.viewport.recycledCells = Object.values(gridStats.poolSizes || {}).reduce((a, b) => a + b, 0);
    }
  }
  
  trackMemoryPeriodically() {
    const checkMemory = () => {
      if ('memory' in performance) {
        this.metrics.memory.memoryUsage = performance.memory.usedJSHeapSize;
        
        if (this.metrics.memory.memoryUsage > this.thresholds.highMemory) {
          if (this.options.enableConsoleReports) {
            console.warn(`🧠 High memory usage: ${(this.metrics.memory.memoryUsage / 1024 / 1024).toFixed(1)}MB`);
          }
        }
      }
      
      // Count DOM nodes
      this.metrics.memory.domNodes = this.grid.container.querySelectorAll('*').length;
      
      setTimeout(checkMemory, 5000); // Check every 5 seconds
    };
    
    checkMemory();
  }
  
  createPerformanceIndicator() {
    const indicator = document.createElement('div');
    indicator.id = 'grid-performance-indicator';
    indicator.style.cssText = `
      position: fixed;
      top: 10px;
      right: 10px;
      background: rgba(0, 0, 0, 0.8);
      color: white;
      padding: 8px 12px;
      border-radius: 4px;
      font-family: monospace;
      font-size: 12px;
      z-index: 1000;
      min-width: 120px;
      backdrop-filter: blur(4px);
      transition: opacity 0.3s ease;
    `;
    
    document.body.appendChild(indicator);
    this.performanceIndicator = indicator;
    
    // Hide/show on hover
    let hideTimeout;
    indicator.addEventListener('mouseenter', () => {
      clearTimeout(hideTimeout);
      indicator.style.opacity = '1';
    });
    
    indicator.addEventListener('mouseleave', () => {
      hideTimeout = setTimeout(() => {
        indicator.style.opacity = '0.3';
      }, 2000);
    });
    
    // Initial state
    indicator.style.opacity = '0.3';
    
    this.updatePerformanceIndicator();
  }
  
  updatePerformanceIndicator() {
    if (!this.performanceIndicator) return;
    
    const fps = this.fps.current;
    const renderTime = this.metrics.render.averageTime || 0;
    const memoryMB = this.metrics.memory.memoryUsage ? (this.metrics.memory.memoryUsage / 1024 / 1024).toFixed(1) : 'N/A';
    const visibleCells = this.metrics.viewport.visibleCells;
    
    // Color coding for FPS
    let fpsColor = '#00ff00'; // Green
    if (fps < this.thresholds.lowFPS) fpsColor = '#ff9900'; // Orange
    if (fps < 20) fpsColor = '#ff0000'; // Red
    
    // Color coding for render time
    let renderColor = '#00ff00'; // Green
    if (renderTime > this.options.alertThreshold) renderColor = '#ff9900'; // Orange
    if (renderTime > this.thresholds.slowRender) renderColor = '#ff0000'; // Red
    
    this.performanceIndicator.innerHTML = `
      <div style="margin-bottom: 4px;">
        <span style="color: ${fpsColor};">FPS: ${fps}</span>
      </div>
      <div style="margin-bottom: 4px;">
        <span style="color: ${renderColor};">Render: ${renderTime.toFixed(1)}ms</span>
      </div>
      <div style="margin-bottom: 4px;">
        Memory: ${memoryMB}MB
      </div>
      <div>
        Cells: ${visibleCells}
      </div>
    `;
  }
  
  // Public API
  getMetrics() {
    return {
      ...this.metrics,
      fps: this.fps.current,
      fpsHistory: [...this.fps.history],
      summary: this.generateSummary()
    };
  }
  
  generateSummary() {
    const renderPerf = this.metrics.render.averageTime < this.options.alertThreshold ? 'Good' : 'Poor';
    const fpsPerf = this.fps.current >= this.thresholds.lowFPS ? 'Good' : 'Poor';
    const memoryPerf = this.metrics.memory.memoryUsage < this.thresholds.highMemory ? 'Good' : 'High';
    
    return {
      overall: renderPerf === 'Good' && fpsPerf === 'Good' && memoryPerf === 'Good' ? 'Excellent' : 
              renderPerf === 'Good' && fpsPerf === 'Good' ? 'Good' : 'Needs Improvement',
      render: renderPerf,
      fps: fpsPerf,
      memory: memoryPerf,
      recommendations: this.generateRecommendations()
    };
  }
  
  generateRecommendations() {
    const recommendations = [];
    
    if (this.metrics.render.averageTime > this.thresholds.slowRender) {
      recommendations.push('Consider reducing viewport size or enabling virtual scrolling');
    }
    
    if (this.fps.current < this.thresholds.lowFPS) {
      recommendations.push('Reduce animation complexity or disable transitions');
    }
    
    if (this.metrics.memory.memoryUsage > this.thresholds.highMemory) {
      recommendations.push('Enable cell recycling or reduce cache size');
    }
    
    if (this.metrics.scroll.laggedEvents > this.metrics.scroll.smoothEvents) {
      recommendations.push('Enable scroll throttling or reduce overscan buffer');
    }
    
    return recommendations;
  }
  
  startProfiling(duration = 10000) {
    console.log('🔍 Starting grid performance profiling...');
    
    const startTime = Date.now();
    const initialMetrics = { ...this.metrics };
    
    // Mark start of profiling
    if ('performance' in window && 'mark' in performance) {
      performance.mark('profiling-start');
    }
    
    setTimeout(() => {
      this.stopProfiling(startTime, initialMetrics);
    }, duration);
    
    return startTime;
  }
  
  stopProfiling(startTime, initialMetrics) {
    const endTime = Date.now();
    const duration = endTime - startTime;
    
    // Mark end of profiling
    if ('performance' in window && 'mark' in performance) {
      performance.mark('profiling-end');
      performance.measure('profiling-session', 'profiling-start', 'profiling-end');
    }
    
    const report = this.generateProfilingReport(duration, initialMetrics);
    console.log('📊 Grid Performance Report:', report);
    
    return report;
  }
  
  generateProfilingReport(duration, initialMetrics) {
    return {
      session: {
        duration: duration,
        timestamp: new Date().toISOString()
      },
      performance: {
        averageFPS: this.fps.history.slice(-10).reduce((a, b) => a + b, 0) / Math.min(10, this.fps.history.length),
        averageRenderTime: this.metrics.render.averageTime,
        totalRenders: this.metrics.render.totalRenders - initialMetrics.render.totalRenders,
        slowRenders: this.metrics.render.slowRenders - initialMetrics.render.slowRenders,
        scrollEvents: this.metrics.scroll.events - initialMetrics.scroll.events,
        inputEvents: this.metrics.input.keystrokes - initialMetrics.input.keystrokes
      },
      memory: {
        currentUsage: this.metrics.memory.memoryUsage,
        domNodes: this.metrics.memory.domNodes,
        poolSizes: this.metrics.memory.poolSizes
      },
      recommendations: this.generateRecommendations()
    };
  }
  
  reset() {
    // Reset all metrics
    this.metrics.render.times = [];
    this.metrics.render.totalRenders = 0;
    this.metrics.render.slowRenders = 0;
    this.metrics.scroll.events = 0;
    this.metrics.scroll.smoothEvents = 0;
    this.metrics.scroll.laggedEvents = 0;
    this.metrics.input.keystrokes = 0;
    this.metrics.input.cellUpdates = 0;
    this.metrics.input.formulaCalculations = 0;
    
    this.fps.history = [];
    
    console.log('🔄 Performance metrics reset');
  }
  
  exportReport() {
    const report = {
      timestamp: new Date().toISOString(),
      metrics: this.getMetrics(),
      browser: {
        userAgent: navigator.userAgent,
        memory: 'memory' in performance ? performance.memory : null,
        timing: performance.timing ? {
          domContentLoaded: performance.timing.domContentLoadedEventEnd - performance.timing.navigationStart,
          loadComplete: performance.timing.loadEventEnd - performance.timing.navigationStart
        } : null
      },
      grid: {
        type: this.grid.constructor.name,
        options: this.grid.options || {},
        viewport: this.metrics.viewport
      }
    };
    
    const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = `grid-performance-${Date.now()}.json`;
    a.click();
    
    URL.revokeObjectURL(url);
    
    return report;
  }
  
  destroy() {
    // Clean up observers
    this.observers.forEach(observer => observer.disconnect());
    this.observers.clear();
    
    // Remove performance indicator
    if (this.performanceIndicator) {
      this.performanceIndicator.remove();
    }
    
    console.log('🧹 Performance monitor destroyed');
  }
}