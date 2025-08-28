# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Running the Application
```bash
# No build step required - open directly in browser
python3 -m http.server 8000
# Then navigate to http://localhost:8000/index.html
```

### Testing
```bash
# Use the built-in test runner in the application
# Click "Advanced > Run Tests" in the UI, or programmatically call the test function
```

## Architecture Overview

This is a **vanilla JavaScript spreadsheet application** that runs entirely in the browser with no build step or dependencies. The architecture is single-page application with modular functionality.

### Core Architecture
- **Single HTML file** (`index.html`) - Main UI structure with toolbar, grid, and status areas
- **Single JavaScript module** (`app.js`) - All application logic in one file with clear functional separation
- **Single CSS file** (`style.css`) - Dark theme styling with CSS custom properties

### Key Components (within app.js)
- **Sheet Management**: Multi-sheet workbook support with tab interface
- **Grid System**: Dynamic table rendering with resizable columns/rows 
- **Formula Engine**: Safe formula parser supporting SUM, MIN, MAX, AVERAGE functions and cell references
- **File I/O**: CSV and XLSX import/export with lazy-loaded libraries
- **Undo/Redo**: Full state management with snapshot system
- **Cell Formatting**: Bold, italic, background colors stored per cell

### Data Structure
```javascript
// Each sheet contains:
{
  name: string,           // Sheet name
  rows: number,           // Row count  
  cols: number,           // Column count
  data: Cell[][],         // 2D array of cell objects
  colWidths: number[],    // Column width overrides
  rowHeights: number[]    // Row height overrides
}

// Each cell contains:
{
  value: string,          // Raw cell value (formula or literal)
  bold: boolean,          // Formatting flags
  italic: boolean,
  bgColor: string         // Hex color
}
```

### External Libraries (Lazy-loaded)
- **XLSX.js**: Loaded from CDN only when needed for Excel import/export
- **JSZip**: Loaded when exporting multiple CSV sheets as ZIP
- Libraries are loaded with fallback CDNs and graceful degradation

### Formula System
- **Safe evaluation**: No `eval()` usage, custom tokenizer and parser
- **Circular reference detection**: Prevents infinite loops
- **Error handling**: `#VALUE!` and `#CIRC!` errors with visual indicators
- **Reference types**: Supports absolute (`$A$1`), mixed (`$A1`, `A$1`), and relative (`A1`) references

## Key Implementation Details

### State Management
- No external state management - uses closure-scoped variables
- Undo/redo system creates deep copies of entire sheet state
- Active sheet switching preserves state per sheet

### Performance Optimizations
- **Throttled recalculation**: 16ms delay to avoid excessive updates during typing
- **Caret preservation**: Maintains cursor position during grid refreshes
- **Auto-fit columns**: Canvas-based text measurement for optimal widths
- **Limited autofit sampling**: Only measures first 50 rows for performance

### Testing System
- Built-in self-test suite accessible via UI
- Tests formula evaluation, CSV parsing, and library loading
- Includes edge cases like circular references and error propagation

## Development Guidelines

### Code Organization
The single `app.js` file is organized into logical sections:
1. DOM element references and sheet initialization
2. Sheet management functions
3. Rendering functions (header, body, tabs)
4. Formula evaluation engine
5. Event handlers and user interactions
6. File I/O operations
7. Testing infrastructure

### Adding New Features
- Formula functions: Add to `fnMap` object and implement handler
- File formats: Extend file input event handler with new MIME types
- UI controls: Add to `index.html` and wire up event listeners in `app.js`

### Error Handling
- All errors are logged to both console and debug panel
- Formula errors are displayed inline with red styling
- File operation errors show in status bar with graceful fallbacks

## Migration Path

The `development.md` file contains a comprehensive plan for migrating this vanilla JS prototype to a modern React/TypeScript stack with proper tooling and build process. The current implementation serves as a working reference for the feature set and user interactions.