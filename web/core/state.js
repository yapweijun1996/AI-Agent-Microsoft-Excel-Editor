'use strict';

export const DEBUG = new URLSearchParams(location.search).get('debug') === 'true' || location.hostname === 'localhost';

export const STORAGE_KEYS = {
  tasks: 'xlsx_ai_tasks_v1',
  keysMeta: 'xlsx_ai_keys_meta',
  wb: 'xlsx_ai_wb_b64',
  panelLayout: 'panelLayout'
};

export const AppState = {
  wb: null,
  activeSheet: 'Sheet1',
  activeCell: { r: 0, c: 0 }, // Excel-like active cell (0-based r,c)
  selectedRows: [], // 1-based row numbers selected via row header
  selectedCols: [], // 0-based column indices selected via column header
  tasks: [],
  messages: [],
  keys: { openai: null, gemini: null },
  dryRun: false,
  selectedModel: 'auto', // auto, openai:gpt-4o, gemini:gemini-2.5-flash, etc.
  autoExecute: true,
  history: [],
  historyIndex: -1,
  maxHistorySize: 50,
  clipboard: null
};