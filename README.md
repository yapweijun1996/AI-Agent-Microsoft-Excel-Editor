# AI-Agent-Microsoft-Excel-Editor

This project is a web-based spreadsheet application with AI-powered features. It allows users to work with spreadsheets, leveraging AI agents for various tasks.

## Project Structure

- `index.html`: The main entry point of the application.
- `package.json`: Defines project dependencies and scripts.
- `server.log`: Log file for server-side operations.
- `web/`: Contains all the front-end source code.
  - `app.js`: The main application logic.
  - `FormulaEngine.js`: Handles formula parsing and calculations.
  - `styles.css`: Defines the application's styling.
  - `chat/`: Chat-related functionalities.
    - `chat-ui.js`: Manages the chat user interface.
  - `core/`: Core application state and global bindings.
    - `global-bindings.js`: Manages global event listeners.
    - `state.js`: Holds the application's state.
  - `db/`: Database-related functionalities.
    - `indexeddb.js`: Manages the IndexedDB storage.
  - `file/`: File import and export functionalities.
    - `import-export.js`: Handles file import and export operations.
  - `services/`: Services for interacting with external APIs.
    - `ai-agents.js`: Manages AI agent interactions.
    - `api-keys.js`: Manages API key settings.
  - `spreadsheet/`: Spreadsheet-related functionalities.
    - `grid-interactions.js`: Manages user interactions with the grid.
    - `grid-renderer.js`: Renders the spreadsheet grid.
    - `history-manager.js`: Manages undo/redo functionality.
    - `operations.js`: Defines spreadsheet operations.
    - `resizing.js`: Handles column and row resizing.
    - `sheet-manager.js`: Manages individual sheets.
    - `workbook-manager.js`: Manages the entire workbook.
  - `tasks/`: Task management functionalities.
    - `task-manager.js`: Manages tasks and their execution.
  - `ui/`: User interface components.
    - `bindings.js`: Manages UI event bindings.
    - `modal.js`: A reusable modal component.
    - `modals.js`: Manages all modals in the application.
    - `resizable-panel.js`: A resizable panel component.
    - `toast.js`: A toast notification component.
  - `utils/`: Utility functions.
    - `index.js`: A collection of utility functions.

## Getting Started

To run the application, open `index.html` in your web browser.