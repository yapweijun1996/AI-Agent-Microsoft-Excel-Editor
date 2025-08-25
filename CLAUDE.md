# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a client-side Excel editor with multi-agent AI automation that runs entirely in the browser. It uses vanilla JavaScript, HTML5, and CSS3 with Tailwind CSS styling. The application features a multi-agent AI system (Planner, Executor, Validator) that can automate spreadsheet tasks using OpenAI GPT-4 and Google Gemini APIs.

## Development Commands

Since this is a pure frontend application, there is no build process required:

```bash
# To run locally - open directly in browser
open web/index.html

# For development with a local server (recommended)
python -m http.server 8000  # Then visit http://localhost:8000/web/
# or
npx serve web/

# Optional: Install Tailwind CSS for style development
npm install
npx tailwindcss -i ./web/styles.css -o ./web/styles.css --watch
```

## Architecture

### Core Components
- **AppState**: Central state management in `app.js:14-30`
- **Multi-Agent System**: 
  - Planner Agent: Breaks down complex requests
  - Executor Agent: Performs spreadsheet operations  
  - Validator Agent: Ensures data integrity
- **Storage Layer**: IndexedDB wrapper at `app.js:35-80` with localStorage fallback
- **Formula Engine**: `FormulaEngine.js` handles Excel formula parsing and calculation
- **Spreadsheet Engine**: SheetJS (XLSX) for import/export functionality

### File Structure
```
web/
├── index.html          # Main HTML with modal system and UI components
├── styles.css          # Tailwind CSS with custom animations
├── app.js              # Core application logic (1400+ lines)
└── FormulaEngine.js    # Excel formula parsing and calculation
```

### Key Systems
- **Modal System**: Reusable modal dialogs for settings, imports, exports
- **Toast Notifications**: User feedback system with animations
- **Task Manager**: AI task coordination and visual status tracking  
- **History Manager**: Undo/redo with 50-level history at `app.js:26-29`
- **Sheet Renderer**: Dynamic spreadsheet display with live editing
- **Keyboard Shortcuts**: Professional Excel-like shortcuts

### Multi-Agent Workflow
1. User inputs natural language command
2. Planner Agent breaks down into tasks
3. Executor Agent performs operations
4. Validator Agent checks integrity
5. Results displayed with task status updates

## Important Implementation Details

### State Management
- **AppState** object manages workbook, active sheet, selected cells, tasks, and history
- IndexedDB for persistent storage with localStorage fallback
- All state changes go through history system for undo/redo

### AI Integration
- API keys stored locally in browser (security consideration)
- Support for multiple AI providers (OpenAI GPT-4o, Gemini Pro/Flash)
- Model selection via `AppState.selectedModel`
- Dry run mode available for previewing AI changes

### Spreadsheet Operations
- Uses SheetJS for Excel compatibility
- Multi-sheet support with tab management
- Live cell editing with instant updates
- Import/export for .xlsx and .csv files

## Common Tasks

### Adding New AI Agents
1. Create agent function in `app.js` following existing pattern
2. Add to agent selection system
3. Implement task coordination logic
4. Update validation pipeline

### Modifying Spreadsheet Features
1. Check existing patterns in sheet rendering functions
2. Update AppState structure if needed
3. Add keyboard shortcuts in event handlers
4. Ensure history system captures changes

### Styling Changes
- Use existing Tailwind classes in `styles.css`
- Follow component-based modal system
- Maintain responsive design patterns
- Use CSS custom properties for theming

## Security Considerations

- API keys are stored in browser localStorage
- All AI calls made directly from client browser
- No server-side components - pure frontend architecture
- For production, consider server proxy to protect API keys

## Dependencies

- **SheetJS**: Excel file processing
- **Tailwind CSS**: Styling framework  
- **hot-formula-parser**: Excel formula calculation
- **IndexedDB**: Browser storage API
- External CDN resources loaded in HTML

## Browser Compatibility

- Chrome 60+, Firefox 55+, Safari 11+, Edge 79+
- Requires ES6+ support and IndexedDB
- Mobile responsive design included