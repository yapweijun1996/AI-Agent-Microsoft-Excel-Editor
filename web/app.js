/* AI Excel Editor - App Core */
'use strict';

// Debug flag via ?debug=true
const DEBUG = new URLSearchParams(location.search).get('debug') === 'true' || location.hostname === 'localhost';

const STORAGE_KEYS = {
  tasks: 'xlsx_ai_tasks_v1',
  keysMeta: 'xlsx_ai_keys_meta',
  wb: 'xlsx_ai_wb_b64',
  panelLayout: 'panelLayout'
};

const AppState = {
  wb: null,
  activeSheet: 'Sheet1',
  tasks: [],
  messages: [],
  keys: { openai: null, gemini: null },
  dryRun: false,
  selectedModel: 'auto', // auto, openai:gpt-4o, gemini:gemini-2.5-flash, etc.
  history: [],
  historyIndex: -1,
  maxHistorySize: 50
};

function log(...args){ if(DEBUG) console.log('[DEBUG]', ...args); }

// IndexedDB wrapper
const db = {
  name: 'ExcelAIDB',
  version: 1,
  db: null,
  
  async init() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.name, this.version);
      
      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        this.db = request.result;
        resolve();
      };
      
      request.onupgradeneeded = (event) => {
        const db = event.target.result;
        
        // Create workbooks store
        if (!db.objectStoreNames.contains('workbooks')) {
          db.createObjectStore('workbooks', { keyPath: 'id' });
        }
        
        // Create tasks store
        if (!db.objectStoreNames.contains('tasks')) {
          const taskStore = db.createObjectStore('tasks', { keyPath: 'id' });
          taskStore.createIndex('workbookId', 'workbookId', { unique: false });
        }
      };
    });
  },
  
  async saveWorkbook(workbook) {
    const tx = this.db.transaction(['workbooks'], 'readwrite');
    const store = tx.objectStore('workbooks');
    return store.put(workbook);
  },
  
  async getWorkbook(id) {
    const tx = this.db.transaction(['workbooks'], 'readonly');
    const store = tx.objectStore('workbooks');
    return store.get(id);
  },
  
  async saveTask(task) {
    const tx = this.db.transaction(['tasks'], 'readwrite');
    const store = tx.objectStore('tasks');
    return store.put(task);
  },
  
  async getTasksByWorkbook(workbookId) {
    const tx = this.db.transaction(['tasks'], 'readonly');
    const store = tx.objectStore('tasks');
    const index = store.index('workbookId');
    return new Promise((resolve, reject) => {
      const request = index.getAll(workbookId);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }
};

// Toast
class Toast {
  constructor(){ this.container = document.getElementById('toast-container'); }
  show(message, type='info', duration=5000){
    const id = 't'+Date.now();
    const typeStyles = {success:'bg-green-500 text-white', error:'bg-red-500 text-white', warning:'bg-yellow-500 text-white', info:'bg-blue-500 text-white'};
    const html = `
      <div id="${id}" class="toast flex items-center p-4 rounded-lg shadow-lg ${typeStyles[type]} transform translate-x-full transition-transform duration-300 ease-in-out">
        <div class="flex-1"><p class="text-sm font-medium">${message}</p></div>
        <button onclick="this.parentElement.remove()" class="ml-3 text-white hover:text-gray-200 focus:outline-none">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
        </button>
      </div>`;
    this.container.insertAdjacentHTML('beforeend', html);
    const el = document.getElementById(id);
    setTimeout(()=>el.classList.remove('translate-x-full'), 10);
    if(duration>0){
      setTimeout(()=>{ if(el){ el.classList.add('translate-x-full'); setTimeout(()=>el.remove(),300);} }, duration);
    }
    return el;
  }
}
const toast = new Toast();
function showToast(msg,type='info',dur=5000){ return toast.show(msg,type,dur); }

// Modal
class Modal {
  constructor(){ this.container = document.getElementById('modal-container'); this.currentModal = null; }
  show({title, content, buttons=[], size='md', closable=true}){
    const sizeClasses={sm:'max-w-sm',md:'max-w-md',lg:'max-w-lg',xl:'max-w-xl',full:'max-w-full'};
    const html=`
    <div class="fixed inset-0 z-50 overflow-y-auto" id="modal-overlay">
      <div class="flex items-center justify-center min-h-screen px-4 pt-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 transition-opacity bg-gray-500 bg-opacity-75" id="modal-backdrop"></div>
        <div class="inline-block w-full ${sizeClasses[size]} p-6 my-8 overflow-hidden text-left align-middle transition-all transform bg-white shadow-xl rounded-lg">
          <div class="flex items-center justify-between mb-4">
            <h3 class="text-lg font-medium text-gray-900">${title}</h3>
            ${closable ? `<button id="modal-close" class="text-gray-400 hover:text-gray-600 focus:outline-none">
              <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
            </button>`:''}
          </div>
          <div class="mb-6">${content}</div>
          <div class="flex justify-end space-x-3">
            ${buttons.map(btn=>`
              <button data-action="${btn.action}" class="px-4 py-2 text-sm font-medium rounded-md focus:outline-none focus:ring-2 focus:ring-offset-2 ${btn.primary ? 'bg-blue-500 hover:bg-blue-600 text-white focus:ring-blue-500':'bg-gray-300 hover:bg-gray-400 text-gray-700 focus:ring-gray-500'}">${btn.text}</button>
            `).join('')}
          </div>
        </div>
      </div>
    </div>`;
    this.container.innerHTML = html;
    this.currentModal = document.getElementById('modal-overlay');
    if(closable){
      document.getElementById('modal-close').addEventListener('click',()=>this.close());
      document.getElementById('modal-backdrop').addEventListener('click',()=>this.close());
    }
    buttons.forEach(btn=>{
      const el = this.container.querySelector(`[data-action="${btn.action}"]`);
      if(el && btn.onClick){ 
        el.addEventListener('click',e=>{ 
          e.preventDefault();
          btn.onClick(e); 
          if(btn.closeOnClick!==false) this.close(); 
        }); 
      }
    });
    return this.currentModal;
  }
  close(){ if(this.currentModal){ this.currentModal.remove(); this.currentModal=null; } }
}

function showApiKeyModal(provider){
  log(`Opening API key modal for ${provider}`);
  const modal = new Modal();
  modal.show({
    title: `Set ${provider} API Key`,
    content: `
      <div class="space-y-4">
        <p class="text-sm text-gray-600">Enter your ${provider} API key. It will be stored in memory; toggle persistence if desired.</p>
        <div class="space-y-2">
          <input type="password" id="api-key-input" placeholder="Enter API key..." class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500" />
          <p class="text-xs text-gray-500">Get your API key from: ${provider === 'OpenAI' ? 'https://platform.openai.com/api-keys' : 'https://aistudio.google.com/app/apikey'}</p>
        </div>
        <label class="flex items-center space-x-2 text-xs text-gray-600">
          <input id="persist-key" type="checkbox" class="rounded border-gray-300">
          <span>Persist to localStorage (less secure)</span>
        </label>
      </div>`,
    buttons:[
      {text:'Cancel', action:'cancel'},
      {text:'Save Key', action:'save', primary:true, onClick:()=>{
        const key = document.getElementById('api-key-input').value.trim();
        const persist = document.getElementById('persist-key').checked;
        if(key){ 
          saveApiKey(provider.toLowerCase(), key, persist); 
          showToast(`${provider} API key saved successfully`, 'success'); 
        } else {
          showToast('Please enter a valid API key', 'warning');
        }
      }}
    ]
  });
}

function showHelpModal(){
  const modal = new Modal();
  modal.show({
    title: 'Keyboard Shortcuts & Help',
    content: `
      <div class="space-y-6 text-sm">
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">File Operations</h4>
          <div class="space-y-1 text-gray-600">
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+S</kbd> Export as XLSX</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+O</kbd> Import XLSX file</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Z</kbd> Undo</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Y</kbd> / <kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Shift+Z</kbd> Redo</div>
          </div>
        </div>
        
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">Sheet Operations</h4>
          <div class="space-y-1 text-gray-600">
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+T</kbd> Add new sheet</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+W</kbd> Delete current sheet</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Tab</kbd> / <kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Shift+Tab</kbd> Switch between sheets</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+1-9</kbd> Switch to sheet by number</div>
          </div>
        </div>
        
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">Chat & AI</h4>
          <div class="space-y-1 text-gray-600">
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">F2</kbd> Focus chat input</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Ctrl+Enter</kbd> Focus chat input</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Enter</kbd> Send message</div>
            <div><kbd class="px-2 py-1 bg-gray-100 rounded text-xs">Escape</kbd> Clear chat input focus</div>
          </div>
        </div>
        
        <div>
          <h4 class="font-semibold text-gray-900 mb-2">Tips & Example Commands</h4>
          <ul class="space-y-1 text-gray-600 text-xs">
            <li>• Set your OpenAI or Gemini API key to use AI features</li>
            <li>• Enable "Dry Run" to preview AI changes before applying</li>
            <li>• <strong>Excel Operations:</strong> "Add a column after B", "Insert 3 rows at row 5"</li>
            <li>• <strong>Formulas:</strong> "Add SUM formula in C10", "Calculate average in D1"</li>
            <li>• <strong>Data:</strong> "Create header row with Name, Age, Salary", "Sort by column A"</li>
            <li>• <strong>Formatting:</strong> "Format column C as currency", "Make header row bold"</li>
            <li>• AI agents work across multiple sheets in your workbook</li>
          </ul>
        </div>
      </div>`,
    buttons:[{text:'Close', action:'close', primary:true}],
    size: 'lg'
  });
}

function saveApiKey(provider, key, persist=false){
  log(`Saving API key for ${provider}, persist: ${persist}`);
  
  if(provider==='openai') {
    AppState.keys.openai = key;
    log('OpenAI key saved to memory');
  }
  if(provider==='gemini') {
    AppState.keys.gemini = key;
    log('Gemini key saved to memory');
  }
  
  // Save metadata
  const meta = {openai: !!AppState.keys.openai, gemini: !!AppState.keys.gemini};
  localStorage.setItem(STORAGE_KEYS.keysMeta, JSON.stringify(meta));
  log('API key metadata saved:', meta);
  
  // Persist the actual key if requested
  if(persist){ 
    localStorage.setItem('xlsx_ai_key_'+provider, key); 
    log(`${provider} key persisted to localStorage`);
  }
  
  // Update UI to reflect current provider
  updateProviderStatus();
}
function restoreApiKeys(){
  const meta = JSON.parse(localStorage.getItem(STORAGE_KEYS.keysMeta)||'{}');
  if(meta.openai){ const k = localStorage.getItem('xlsx_ai_key_openai'); if(k) AppState.keys.openai = k; }
  if(meta.gemini){ const k = localStorage.getItem('xlsx_ai_key_gemini'); if(k) AppState.keys.gemini = k; }
  updateProviderStatus();
}

function updateProviderStatus(){
  const openaiBtn = document.getElementById('openai-key-btn');
  const geminiBtn = document.getElementById('gemini-key-btn');
  
  if(openaiBtn) {
    if(AppState.keys.openai) {
      openaiBtn.textContent = '✓ OpenAI Ready';
      openaiBtn.classList.remove('bg-blue-500', 'hover:bg-blue-600');
      openaiBtn.classList.add('bg-green-500', 'hover:bg-green-600');
    } else {
      openaiBtn.textContent = 'Set OpenAI Key';
      openaiBtn.classList.remove('bg-green-500', 'hover:bg-green-600');
      openaiBtn.classList.add('bg-blue-500', 'hover:bg-blue-600');
    }
  }
  
  if(geminiBtn) {
    if(AppState.keys.gemini) {
      geminiBtn.textContent = '✓ Gemini Ready';
      geminiBtn.classList.remove('bg-green-500', 'hover:bg-green-600');
      geminiBtn.classList.add('bg-green-600', 'hover:bg-green-700');
    } else {
      geminiBtn.textContent = 'Set Gemini Key';
      geminiBtn.classList.remove('bg-green-600', 'hover:bg-green-700');
      geminiBtn.classList.add('bg-green-500', 'hover:bg-green-600');
    }
  }
}

// Panel layout persistence
function savePanelLayout(){
  const left = document.getElementById('excel-panel');
  const right = document.getElementById('chat-panel');
  localStorage.setItem(STORAGE_KEYS.panelLayout, JSON.stringify({leftWidth:left.style.flex, rightWidth:right.style.flex}));
}
function restorePanelLayout(){
  const saved = localStorage.getItem(STORAGE_KEYS.panelLayout);
  if(saved){ const layout = JSON.parse(saved); const l = document.getElementById('excel-panel'); const r = document.getElementById('chat-panel'); l.style.flex = layout.leftWidth||''; r.style.flex = layout.rightWidth||''; }
}

// Draggable divider
function initDraggableDivider(){
  // Skip divider initialization if elements don't exist (new Excel-like layout)
  const divider = document.getElementById('divider');
  const leftPanel = document.getElementById('excel-panel');
  const rightPanel = document.getElementById('chat-panel');
  
  if(!divider || !leftPanel || !rightPanel) {
    log('Divider elements not found, skipping draggable divider initialization');
    return;
  }
  
  const container = document.querySelector('.main-container');
  let isDragging=false;
  divider.addEventListener('mousedown',()=>{ isDragging=true; document.body.style.cursor='col-resize'; document.body.style.userSelect='none'; });
  document.addEventListener('mousemove',(e)=>{
    if(!isDragging) return;
    const rect = container.getBoundingClientRect();
    const newLeft = e.clientX - rect.left;
    const minLeft=300, minRight=250;
    if(newLeft>=minLeft && (rect.width - newLeft)>=minRight){ leftPanel.style.flex = `0 0 ${newLeft}px`; rightPanel.style.flex='1'; }
  });
  document.addEventListener('mouseup',()=>{ if(isDragging){ isDragging=false; document.body.style.cursor=''; document.body.style.userSelect=''; savePanelLayout(); } });
}

// Workbook helpers
async function ensureWorkbook(){
  if(AppState.wb) return;
  try {
    // Try restore from IndexedDB
    const savedWb = await db.getWorkbook('current');
    if(savedWb && savedWb.data && savedWb.data.SheetNames){
      try{ 
        AppState.wb = savedWb.data; 
        AppState.activeSheet = AppState.wb.SheetNames[0] || 'Sheet1'; 
        log('Restored workbook from IndexedDB'); 
        return; 
      }catch(e){ 
        console.warn('Failed to restore workbook:', e);
        showToast('Failed to restore saved workbook, creating new one', 'warning');
      }
    }
    
    // Create new workbook
    const wb = XLSX.utils.book_new();
    const aoa = [["Name","Age","Email"],["Alice",30,"alice@example.com"],["Bob",28,"bob@example.com"]];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, AppState.activeSheet);
    AppState.wb = wb;
    
    // Initialize history with the new workbook
    AppState.history = [];
    AppState.historyIndex = -1;
    saveToHistory('Create new workbook', { sheets: wb.SheetNames });
    
    await persistSnapshot();
    log('Created new workbook');
  } catch(e) {
    console.error('Failed to ensure workbook:', e);
    showToast('Failed to initialize workbook', 'error');
  }
}

async function persistSnapshot(){
  try{ 
    await db.saveWorkbook({ id: 'current', data: AppState.wb }); 
    log('Workbook snapshot saved');
  }catch(e){ 
    console.warn('Snapshot failed:', e);
    showToast('Failed to save changes', 'warning', 2000);
  }
}

function getWorksheet(){ return AppState.wb.Sheets[AppState.activeSheet]; }

// History management for undo/redo
function saveToHistory(action, data) {
  // Don't save if we're in the middle of undo/redo
  if(AppState.isUndoRedoing) return;
  
  // Remove any history after current index (when user makes new changes after undo)
  if(AppState.historyIndex < AppState.history.length - 1) {
    AppState.history = AppState.history.slice(0, AppState.historyIndex + 1);
  }
  
  // Create history entry with full workbook state
  const historyEntry = {
    action,
    data,
    workbook: JSON.parse(JSON.stringify(AppState.wb)), // Deep copy
    activeSheet: AppState.activeSheet,
    timestamp: Date.now()
  };
  
  AppState.history.push(historyEntry);
  AppState.historyIndex = AppState.history.length - 1;
  
  // Limit history size
  if(AppState.history.length > AppState.maxHistorySize) {
    AppState.history = AppState.history.slice(-AppState.maxHistorySize);
    AppState.historyIndex = AppState.history.length - 1;
  }
  
  log('Saved to history:', action, `(${AppState.history.length} entries)`);
}

function canUndo() {
  return AppState.historyIndex > 0;
}

function canRedo() {
  return AppState.historyIndex < AppState.history.length - 1;
}

function undo() {
  if(!canUndo()) {
    showToast('Nothing to undo', 'warning', 1000);
    return;
  }
  
  AppState.isUndoRedoing = true;
  AppState.historyIndex--;
  
  const historyEntry = AppState.history[AppState.historyIndex];
  AppState.wb = JSON.parse(JSON.stringify(historyEntry.workbook));
  AppState.activeSheet = historyEntry.activeSheet;
  
  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
  
  AppState.isUndoRedoing = false;
  
  showToast(`Undid: ${historyEntry.action}`, 'success', 1000);
  log('Undo:', historyEntry.action);
}

function redo() {
  if(!canRedo()) {
    showToast('Nothing to redo', 'warning', 1000);
    return;
  }
  
  AppState.isUndoRedoing = true;
  AppState.historyIndex++;
  
  const historyEntry = AppState.history[AppState.historyIndex];
  AppState.wb = JSON.parse(JSON.stringify(historyEntry.workbook));
  AppState.activeSheet = historyEntry.activeSheet;
  
  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
  
  AppState.isUndoRedoing = false;
  
  showToast(`Redid: ${historyEntry.action}`, 'success', 1000);
  log('Redo:', historyEntry.action);
}

// Sheet management
function renderSheetTabs(){
  if(!AppState.wb) return;
  const container = document.getElementById('sheet-tabs');
  if(!container) return;
  
  let html = '';
  for(const sheetName of AppState.wb.SheetNames){
    const isActive = sheetName === AppState.activeSheet;
    html += `
      <div class="sheet-tab ${isActive ? 'active' : ''}" data-sheet="${escapeHtml(sheetName)}">
        <span class="sheet-name">${escapeHtml(sheetName)}</span>
        ${AppState.wb.SheetNames.length > 1 ? `<span class="sheet-tab-close" onclick="deleteSheet('${escapeHtml(sheetName)}')">&times;</span>` : ''}
      </div>`;
  }
  container.innerHTML = html;
  
  // Add click handlers
  container.querySelectorAll('.sheet-tab').forEach(tab => {
    tab.addEventListener('click', (e) => {
      if(!e.target.classList.contains('sheet-tab-close')){
        switchToSheet(tab.dataset.sheet);
      }
    });
  });
}

function switchToSheet(sheetName){
  if(!AppState.wb || !AppState.wb.Sheets[sheetName]) return;
  AppState.activeSheet = sheetName;
  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
}

function addNewSheet(){
  if(!AppState.wb) return;
  
  let newName = 'Sheet1';
  let counter = 1;
  while(AppState.wb.SheetNames.includes(newName)){
    counter++;
    newName = `Sheet${counter}`;
  }
  
  saveToHistory(`Add sheet "${newName}"`, { sheetName: newName });
  
  const ws = XLSX.utils.aoa_to_sheet([['']]);
  XLSX.utils.book_append_sheet(AppState.wb, ws, newName);
  AppState.activeSheet = newName;
  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
  showToast(`Added sheet "${newName}"`, 'success');
}

window.deleteSheet = function(sheetName){
  if(!AppState.wb || AppState.wb.SheetNames.length <= 1) {
    showToast('Cannot delete the last sheet', 'warning');
    return;
  }
  
  if(!confirm(`Delete sheet "${sheetName}"?`)) return;
  
  const sheetIndex = AppState.wb.SheetNames.indexOf(sheetName);
  if(sheetIndex === -1) return;
  
  saveToHistory(`Delete sheet "${sheetName}"`, { sheetName, sheetIndex });
  
  // Remove from SheetNames and Sheets
  AppState.wb.SheetNames.splice(sheetIndex, 1);
  delete AppState.wb.Sheets[sheetName];
  
  // Switch to another sheet if this was active
  if(AppState.activeSheet === sheetName){
    AppState.activeSheet = AppState.wb.SheetNames[Math.max(0, sheetIndex - 1)];
  }
  
  renderSheetTabs();
  renderSpreadsheetTable();
  persistSnapshot();
  showToast(`Deleted sheet "${sheetName}"`, 'success');
};

// Spreadsheet render
function renderSpreadsheetTable(){
  const ws = getWorksheet();
  const ref = ws['!ref'] || 'A1:C20';
  const range = XLSX.utils.decode_range(ref);
  let html = '';
  html += '<table class="ai-grid min-w-full border-collapse border border-gray-300 bg-white">';
  html += '<thead class="bg-gray-50"><tr>';
  html += '<th class="w-12 p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500">#</th>';
  for(let c=range.s.c; c<=range.e.c; c++){
    const colLetter = XLSX.utils.encode_col(c);
    html += `<th class="p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500 min-w-[100px]">${colLetter}</th>`;
  }
  html += '</tr></thead><tbody>';
  for(let r=range.s.r; r<=range.e.r; r++){
    html += '<tr class="hover:bg-gray-50">';
    html += `<td class=\"p-2 border border-gray-300 bg-gray-100 text-center text-xs font-medium text-gray-500\">${r+1}</td>`;
    for(let c=range.s.c; c<=range.e.c; c++){
      const addr = XLSX.utils.encode_cell({r, c});
      const cell = ws[addr];
      const value = cell ? cell.v : '';
      html += `
        <td class=\"p-1 border border-gray-300 hover:bg-blue-50 focus-within:bg-blue-50 min-h-[32px]\" data-cell=\"${addr}\">
          <input type=\"text\" value=\"${escapeHtml(value)}\" class=\"cell-input w-full h-full px-2 py-1 bg-transparent border-none outline-none focus:bg-white focus:shadow-sm focus:ring-1 focus:ring-blue-400 rounded\" onblur=\"updateCell('${addr}', this.value)\" onkeypress=\"handleCellKeypress(event)\" />
        </td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table>';
  document.getElementById('spreadsheet').innerHTML = html;
}

function parseCellValue(v){
  if(v===null || v===undefined) return {t:'z', v:''};
  const num = Number(v);
  if(v!=='' && !isNaN(num)) return {t:'n', v:num};
  if(typeof v === 'boolean') return {t:'b', v:v};
  return {t:'s', v:String(v)};
}

function escapeHtml(str){
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function expandRefForCell(ws, addr){
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const cell = XLSX.utils.decode_cell(addr);
  range.s.r = Math.min(range.s.r, cell.r);
  range.s.c = Math.min(range.s.c, cell.c);
  range.e.r = Math.max(range.e.r, cell.r);
  range.e.c = Math.max(range.e.c, cell.c);
  ws['!ref'] = XLSX.utils.encode_range(range);
}

window.updateCell = function(addr, value){
  const ws = getWorksheet();
  const oldValue = ws[addr] ? ws[addr].v : '';
  
  // Only save to history if the value actually changed
  if(String(oldValue) !== String(value)) {
    saveToHistory(`Edit cell ${addr}`, { addr, oldValue, newValue: value, sheet: AppState.activeSheet });
  }
  
  const parsed = parseCellValue(value);
  if (parsed.t === 'z') {
    delete ws[addr];
  } else {
    ws[addr] = { t: parsed.t, v: parsed.v };
    expandRefForCell(ws, addr);
  }
  persistSnapshot();
  renderSpreadsheetTable();
};

window.handleCellKeypress = function(event){
  if(event.key==='Enter'){ event.preventDefault(); event.target.blur(); }
};

// Direct spreadsheet manipulation functions
function insertRowAtEnd() {
  const ws = getWorksheet();
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const newRowIndex = range.e.r + 1;
  
  saveToHistory(`Insert row at ${newRowIndex + 1}`, { row: newRowIndex + 1 });
  
  // Expand range to include new row
  range.e.r = newRowIndex;
  ws['!ref'] = XLSX.utils.encode_range(range);
  
  persistSnapshot();
  renderSpreadsheetTable();
  showToast(`Added row ${newRowIndex + 1}`, 'success');
}

function insertColumnAtEnd() {
  const ws = getWorksheet();
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const newColIndex = range.e.c + 1;
  const newColLetter = XLSX.utils.encode_col(newColIndex);
  
  saveToHistory(`Insert column ${newColLetter}`, { column: newColLetter });
  
  // Expand range to include new column
  range.e.c = newColIndex;
  ws['!ref'] = XLSX.utils.encode_range(range);
  
  persistSnapshot();
  renderSpreadsheetTable();
  showToast(`Added column ${newColLetter}`, 'success');
}

function deleteLastRow() {
  const ws = getWorksheet();
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  
  if(range.e.r <= range.s.r) {
    showToast('Cannot delete the only row', 'warning');
    return;
  }
  
  const deleteRowIndex = range.e.r;
  
  saveToHistory(`Delete row ${deleteRowIndex + 1}`, { row: deleteRowIndex + 1 });
  
  // Delete cells in the last row
  for(let c = range.s.c; c <= range.e.c; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r: deleteRowIndex, c });
    delete ws[cellAddr];
  }
  
  // Shrink range
  range.e.r = Math.max(range.s.r, range.e.r - 1);
  ws['!ref'] = XLSX.utils.encode_range(range);
  
  persistSnapshot();
  renderSpreadsheetTable();
  showToast(`Deleted row ${deleteRowIndex + 1}`, 'success');
}

function deleteLastColumn() {
  const ws = getWorksheet();
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  
  if(range.e.c <= range.s.c) {
    showToast('Cannot delete the only column', 'warning');
    return;
  }
  
  const deleteColIndex = range.e.c;
  const deleteColLetter = XLSX.utils.encode_col(deleteColIndex);
  
  saveToHistory(`Delete column ${deleteColLetter}`, { column: deleteColLetter });
  
  // Delete cells in the last column
  for(let r = range.s.r; r <= range.e.r; r++) {
    const cellAddr = XLSX.utils.encode_cell({ r, c: deleteColIndex });
    delete ws[cellAddr];
  }
  
  // Shrink range
  range.e.c = Math.max(range.s.c, range.e.c - 1);
  ws['!ref'] = XLSX.utils.encode_range(range);
  
  persistSnapshot();
  renderSpreadsheetTable();
  showToast(`Deleted column ${deleteColLetter}`, 'success');
}

// Tasks
async function loadTasks(){
  try{ AppState.tasks = await db.getTasksByWorkbook('current') || []; }catch{ AppState.tasks=[]; }
}
async function saveTasks(){
  for(const task of AppState.tasks){
    await db.saveTask({ ...task, workbookId: 'current' });
  }
}

function renderTask(task){
  const statusColors = {
    pending:'bg-gray-100 text-gray-800', 
    in_progress:'bg-blue-100 text-blue-800', 
    done:'bg-green-100 text-green-800', 
    failed:'bg-red-100 text-red-800', 
    blocked:'bg-yellow-100 text-yellow-800'
  };
  
  const statusIcons = {
    pending: '⏳',
    in_progress: '🔄', 
    done: '✅',
    failed: '❌',
    blocked: '🚫'
  };
  
  const canExecute = task.status === 'pending' || task.status === 'failed' || task.status === 'blocked';
  const showRetry = task.status === 'failed' || task.status === 'blocked';
  
  return `
    <div class=\"task-item flex items-start justify-between p-3 bg-white rounded-lg border border-gray-200 hover:border-gray-300 transition-colors\" data-task-id=\"${task.id}\">
      <div class=\"flex-1 min-w-0\">
        <div class=\"flex items-center space-x-2 mb-1\">
          <h4 class=\"text-sm font-medium text-gray-900 truncate\">${escapeHtml(task.title)}</h4>
          <span class=\"text-xs\">${statusIcons[task.status] || statusIcons.pending}</span>
        </div>
        ${task.description ? `<p class=\"text-xs text-gray-500 mb-2 line-clamp-2\">${escapeHtml(task.description)}</p>` : ''}
        <div class=\"flex items-center justify-between\">
          <span class=\"inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${statusColors[task.status]||statusColors.pending}\">${(task.status||'pending').replace('_',' ')}</span>
          ${task.context?.sheet ? `<span class=\"text-xs text-gray-400\">📊 ${escapeHtml(task.context.sheet)}</span>` : ''}
        </div>
        ${task.result && task.status === 'blocked' ? `<div class=\"mt-2 p-2 bg-yellow-50 rounded text-xs text-yellow-800\">${escapeHtml(typeof task.result === 'object' ? task.result.errors?.join(', ') || 'Task blocked' : task.result)}</div>` : ''}
        ${task.result && task.status === 'failed' ? `<div class=\"mt-2 p-2 bg-red-50 rounded text-xs text-red-800\">${escapeHtml(typeof task.result === 'string' ? task.result : 'Task failed')}</div>` : ''}
        ${task.createdAt ? `<div class=\"text-xs text-gray-400 mt-1\">${new Date(task.createdAt).toLocaleString()}</div>` : ''}
      </div>
      <div class=\"flex items-center space-x-1 ml-3 flex-shrink-0\">
        ${canExecute ? `<button onclick=\"executeTask('${task.id}')\" class=\"p-1 text-blue-600 hover:text-blue-800 transition-colors\" title=\"${showRetry ? 'Retry' : 'Execute'}\">
          ${showRetry ? 
            '<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>' :
            '<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M14.828 14.828a4 4 0 01-5.656 0M9 10h1m4 0h1m-6 4h1m4 0h1m-6-8h8a2 2 0 012 2v8a2 2 0 01-2 2H8a2 2 0 01-2-2V6a2 2 0 012-2z"></path></svg>'
          }</button>`:''}
        ${task.status === 'done' ? `<button onclick=\"viewTaskResult('${task.id}')\" class=\"p-1 text-green-600 hover:text-green-800 transition-colors\" title=\"View Result\">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>
          </button>`:''}
        <button onclick=\"deleteTask('${task.id}')\" class=\"p-1 text-red-600 hover:text-red-800 transition-colors\" title=\"Delete\">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
        </button>
      </div>
    </div>`;
}

function drawTasks(){
  const list = document.getElementById('task-list');
  const summary = document.getElementById('task-summary');
  
  // Update task summary in AI panel
  if(summary) {
    const pending = AppState.tasks.filter(t => t.status === 'pending').length;
    const inProgress = AppState.tasks.filter(t => t.status === 'in_progress').length;
    const completed = AppState.tasks.filter(t => t.status === 'done').length;
    
    if(AppState.tasks.length === 0) {
      summary.textContent = 'No active tasks';
    } else {
      summary.textContent = `${pending} pending, ${inProgress} running, ${completed} done`;
    }
  }
  
  if(AppState.tasks.length === 0) {
    list.innerHTML = '<div class="text-center text-gray-500 text-sm py-4">No tasks yet. Chat with AI to create tasks!</div>';
    return;
  }
  
  // Group tasks by status
  const tasksByStatus = {
    in_progress: AppState.tasks.filter(t => t.status === 'in_progress'),
    pending: AppState.tasks.filter(t => t.status === 'pending'),
    blocked: AppState.tasks.filter(t => t.status === 'blocked'), 
    failed: AppState.tasks.filter(t => t.status === 'failed'),
    done: AppState.tasks.filter(t => t.status === 'done')
  };
  
  let html = '';
  
  // Show active tasks first (in progress, pending, blocked, failed)
  const activeTasks = [...tasksByStatus.in_progress, ...tasksByStatus.pending, ...tasksByStatus.blocked, ...tasksByStatus.failed];
  if(activeTasks.length > 0) {
    html += '<div class="space-y-2">';
    html += activeTasks.map(renderTask).join('');
    html += '</div>';
  }
  
  // Show completed tasks in a collapsible section
  if(tasksByStatus.done.length > 0) {
    html += `
      <div class="mt-4 pt-4 border-t border-gray-200">
        <button onclick="toggleCompletedTasks()" class="flex items-center space-x-2 text-sm text-gray-600 hover:text-gray-800 mb-2">
          <svg id="completed-toggle-icon" class="w-4 h-4 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"/>
          </svg>
          <span>Completed (${tasksByStatus.done.length})</span>
        </button>
        <div id="completed-tasks" class="hidden space-y-2">
          ${tasksByStatus.done.map(renderTask).join('')}
        </div>
      </div>`;
  }
  
  list.innerHTML = html;
}

window.toggleCompletedTasks = function() {
  const completedTasks = document.getElementById('completed-tasks');
  const toggleIcon = document.getElementById('completed-toggle-icon');
  
  if(completedTasks.classList.contains('hidden')) {
    completedTasks.classList.remove('hidden');
    toggleIcon.style.transform = 'rotate(90deg)';
  } else {
    completedTasks.classList.add('hidden');
    toggleIcon.style.transform = 'rotate(0deg)';
  }
};

window.viewTaskResult = function(id) {
  const task = AppState.tasks.find(t => t.id === id);
  if(!task || !task.result) return;
  
  const modal = new Modal();
  const result = typeof task.result === 'object' ? JSON.stringify(task.result, null, 2) : String(task.result);
  
  modal.show({
    title: `Task Result: ${task.title}`,
    content: `
      <div class="space-y-3">
        <div class="text-sm text-gray-600">
          <strong>Status:</strong> ${task.status} <br>
          <strong>Sheet:</strong> ${task.context?.sheet || 'Unknown'} <br>
          <strong>Completed:</strong> ${new Date(task.createdAt).toLocaleString()}
        </div>
        <div class="bg-gray-50 p-3 rounded-lg">
          <pre class="text-sm text-gray-800 whitespace-pre-wrap">${escapeHtml(result)}</pre>
        </div>
      </div>`,
    buttons: [{text: 'Close', action: 'close', primary: true}],
    size: 'lg'
  });
};

window.deleteTask = function(id){
  AppState.tasks = AppState.tasks.filter(t=>t.id!==id);
  saveTasks();
  drawTasks();
};

// Orchestrator Agent - coordinates multi-agent workflows
async function runOrchestrator(tasks) {
  const provider = pickProvider();
  
  if(provider === 'mock') {
    return {
      executionPlan: tasks.map((t, i) => ({taskId: t.id, order: i + 1, dependencies: []})),
      estimatedTime: tasks.length * 2000,
      riskAssessment: 'low',
      recommendations: ['Execute tasks sequentially']
    };
  }
  
  // Get comprehensive context
  const ws = getWorksheet();
  const sheetContext = ws['!ref'] ? `Sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
  const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';
  
  const system = `You are the Orchestrator Agent - the master coordinator responsible for optimizing multi-agent workflows and ensuring successful task execution.

ROLE: Analyze task dependencies, optimize execution order, assess risks, and coordinate between Planner, Executor, and Validator agents.

CAPABILITIES:
- Dependency analysis and resolution
- Risk assessment and mitigation planning
- Resource optimization and conflict detection
- Parallel execution planning where safe
- Error recovery and rollback strategies
- Performance optimization

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Sheet structure: ${sheetContext}
- Sample data preview: ${sampleData}
- Available sheets: [${AppState.wb.SheetNames.join(', ')}]

ORCHESTRATION STRATEGY:
1. Analyze task dependencies and conflicts
2. Identify opportunities for parallel execution
3. Assess risks and plan mitigation strategies
4. Optimize execution order for efficiency
5. Plan error recovery and rollback procedures

REQUIRED OUTPUT FORMAT:
{
  "executionPlan": [
    {"taskId": "task1", "order": 1, "dependencies": [], "canParallel": false, "estimatedDuration": 1000},
    {"taskId": "task2", "order": 2, "dependencies": ["task1"], "canParallel": false, "estimatedDuration": 2000}
  ],
  "parallelGroups": [
    {"group": 1, "tasks": ["task3", "task4"], "description": "Independent formatting tasks"}
  ],
  "riskAssessment": "medium",
  "risks": ["Potential data overwrite in range A1:C3", "Large dataset may impact performance"],
  "mitigations": ["Create backup before execution", "Implement incremental saves"],
  "estimatedTime": 5000,
  "rollbackComplexity": "medium",
  "recommendations": ["Execute in dry-run mode first", "Monitor memory usage during large operations"],
  "validationStrategy": "incremental",
  "optimizations": ["Batch similar operations", "Use range operations instead of individual cells"]
}

ORCHESTRATION PRINCIPLES:
- Minimize data conflicts and race conditions
- Maximize safe parallelization opportunities
- Provide comprehensive error recovery plans
- Optimize for both speed and data integrity
- Include detailed risk assessment and mitigation`;

  const tasksSummary = tasks.map(t => ({
    id: t.id,
    title: t.title,
    description: t.description,
    dependencies: t.dependencies || [],
    priority: t.priority || 3,
    context: t.context || {}
  }));
  
  const user = `Orchestrate execution of ${tasks.length} tasks:\n${JSON.stringify(tasksSummary, null, 2)}`;
  const messages = [{role:'system', content:system}, {role:'user', content:user}];
  
  try {
    let data;
    const selectedModel = getSelectedModel();
    if(provider === 'openai') {
      data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel);
    } else {
      data = await fetchGemini(AppState.keys.gemini, messages, selectedModel);
    }
    
    let text = '';
    if(provider === 'openai') {
      text = data.choices?.[0]?.message?.content || '';
    } else {
      text = data.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '';
    }
    
    let result = null;
    try {
      result = JSON.parse(text);
    } catch {
      result = extractFirstJson(text);
    }
    
    return result || {
      executionPlan: tasks.map((t, i) => ({taskId: t.id, order: i + 1, dependencies: t.dependencies || []})),
      riskAssessment: 'unknown',
      recommendations: ['Execute with caution - orchestrator analysis failed']
    };
    
  } catch(error) {
    console.error('Orchestrator failed:', error);
    return {
      executionPlan: tasks.map((t, i) => ({taskId: t.id, order: i + 1, dependencies: t.dependencies || []})),
      riskAssessment: 'high',
      recommendations: ['Manual review recommended - orchestrator unavailable']
    };
  }
}

// Enhanced task execution with orchestration
window.executeTask = async function(id) {
  const task = AppState.tasks.find(t => t.id === id);
  if(!task) return;
  
  // Check dependencies
  const uncompletedDeps = task.dependencies?.filter(depId => {
    const depTask = AppState.tasks.find(t => t.id === depId);
    return !depTask || depTask.status !== 'done';
  }) || [];
  
  if(uncompletedDeps.length > 0) {
    showToast(`Cannot execute: waiting for dependencies (${uncompletedDeps.join(', ')})`, 'warning');
    return;
  }
  
  task.status = 'in_progress';
  task.startTime = Date.now();
  saveTasks(); drawTasks();
  
  try {
    const result = await runExecutor(task);
    if(!result) throw new Error('No executor result');
    
    const validation = await runValidator(result, task);
    if(!validation.valid) {
      task.status = 'blocked';
      task.result = validation;
      task.retryCount = (task.retryCount || 0) + 1;
      saveTasks(); drawTasks();
      
      if(task.retryCount < task.maxRetries) {
        showToast(`Task blocked - ${task.maxRetries - task.retryCount} retries remaining`, 'warning');
      } else {
        showToast('Task failed after maximum retries', 'error');
        task.status = 'failed';
      }
      return;
    }
    
    await applyEditsOrDryRun(result);
    task.status = 'done';
    task.result = result;
    task.completedAt = Date.now();
    task.duration = task.completedAt - task.startTime;
    
    saveTasks(); drawTasks();
    showToast(`Task completed: ${task.title}`, 'success');
    
    // Check if this completion enables other tasks
    const enabledTasks = AppState.tasks.filter(t => 
      t.status === 'pending' && 
      t.dependencies?.includes(id) &&
      t.dependencies.every(depId => {
        const depTask = AppState.tasks.find(dt => dt.id === depId);
        return depTask?.status === 'done';
      })
    );
    
    if(enabledTasks.length > 0) {
      showToast(`${enabledTasks.length} task(s) now ready to execute`, 'info');
    }
    
  } catch(e) {
    console.error('Task execution failed:', e);
    task.status = 'failed';
    task.result = String(e);
    task.retryCount = (task.retryCount || 0) + 1;
    saveTasks(); drawTasks();
    showToast(`Task failed: ${task.title}`, 'error');
  }
};

// Execute multiple tasks with orchestration
window.executeTasks = async function(taskIds) {
  const tasks = taskIds.map(id => AppState.tasks.find(t => t.id === id)).filter(Boolean);
  if(tasks.length === 0) return;
  
  showToast(`Orchestrating execution of ${tasks.length} tasks...`, 'info');
  
  try {
    const orchestration = await runOrchestrator(tasks);
    log('Orchestration plan:', orchestration);
    
    // Execute according to orchestration plan
    if(orchestration.executionPlan) {
      const sortedTasks = orchestration.executionPlan
        .sort((a, b) => a.order - b.order)
        .map(plan => tasks.find(t => t.id === plan.taskId))
        .filter(Boolean);
      
      for(const task of sortedTasks) {
        await executeTask(task.id);
        // Small delay between tasks
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }
    
    showToast('Task orchestration completed', 'success');
    
  } catch(error) {
    console.error('Task orchestration failed:', error);
    showToast('Orchestration failed, executing tasks sequentially', 'warning');
    
    // Fallback to sequential execution
    for(const task of tasks) {
      await executeTask(task.id);
    }
  }
};

// Chat
function renderChatMessage(msg){
  const isUser = msg.role==='user';
  const isTyping = msg.isTyping || false;
  
  // Simple markdown-like formatting for AI messages
  let content = escapeHtml(msg.content);
  if(!isUser) {
    // Bold text: **text** -> <strong>text</strong>
    content = content.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    // Line breaks
    content = content.replace(/\n/g, '<br>');
  }
  
  return `
    <div class=\"flex ${isUser ? 'justify-end' : 'justify-start'} ${isTyping ? 'animate-pulse' : ''}\">
      <div class=\"max-w-xs lg:max-w-md px-4 py-2 rounded-lg ${isUser ? 'bg-blue-500 text-white' : (isTyping ? 'bg-yellow-100 text-yellow-800' : 'bg-gray-200 text-gray-900')}\">
        ${isUser ? '' : `<div class="text-xs font-medium ${isTyping ? 'text-yellow-600' : 'text-gray-500'} mb-1">${isTyping ? '🤖 AI Agents' : 'AI Assistant'}</div>`}
        <div class=\"text-sm\">${content}</div>
        <div class=\"text-xs ${isUser ? 'text-blue-100' : (isTyping ? 'text-yellow-600' : 'text-gray-500')} mt-1\">${new Date(msg.timestamp).toLocaleTimeString()}</div>
      </div>
    </div>`;
}

function drawChat(){
  const el = document.getElementById('chat-messages');
  el.innerHTML = AppState.messages.map(renderChatMessage).join('');
  el.scrollTop = el.scrollHeight;
}

async function onSend(){
  const input = document.getElementById('message-input');
  const text = input.value.trim();
  if(!text) return;
  
  // Add user message
  const userMsg = {role:'user', content:text, timestamp:Date.now()};
  AppState.messages.push(userMsg);
  drawChat();
  input.value='';
  
  // Show typing indicator
  const typingMsg = {role:'assistant', content:'🤔 AI agents are planning your request...', timestamp:Date.now(), isTyping:true};
  AppState.messages.push(typingMsg);
  drawChat();
  
  try {
    // Planner: create tasks
    const tasks = await runPlanner(text);
    
    // Remove typing indicator
    AppState.messages = AppState.messages.filter(m => !m.isTyping);
    
    let responseContent = '';
    
    if(tasks && tasks.length){
      AppState.tasks.push(...tasks);
      saveTasks(); drawTasks();
      
      responseContent = `✅ I've analyzed your request and created ${tasks.length} task(s):\n\n`;
      tasks.forEach((task, i) => {
        responseContent += `${i+1}. **${task.title}**: ${task.description}\n`;
      });
      responseContent += `\n🎯 Click the execute button on each task to run them, or use "Execute All" for orchestrated execution.`;
      
      showToast(`Planned ${tasks.length} task(s)`, 'success');
    }else{
      responseContent = `⚠️ I couldn't break down your request into specific tasks. Let me try a direct approach...`;
      
      // fallback single-step
      const task = {
        id:'task-'+Date.now(), 
        title:text, 
        description:text, 
        status:'pending', 
        createdAt:new Date().toISOString()
      };
      AppState.tasks.push(task); 
      saveTasks(); 
      drawTasks();
      
      responseContent += `\n\n✅ Created a single task: "${task.title}"\nClick execute to run it.`;
      showToast('Created fallback task','warning');
    }
    
    // Add AI response message
    const aiMsg = {role:'assistant', content:responseContent, timestamp:Date.now()};
    AppState.messages.push(aiMsg);
    drawChat();
    
  } catch(error) {
    // Remove typing indicator
    AppState.messages = AppState.messages.filter(m => !m.isTyping);
    
    const errorMsg = {
      role:'assistant', 
      content:`❌ I encountered an error processing your request: ${error.message}\n\nPlease check your API keys and try again.`, 
      timestamp:Date.now()
    };
    AppState.messages.push(errorMsg);
    drawChat();
    
    console.error('Chat error:', error);
    showToast('Chat processing failed', 'error');
  }
}

// Agent connectors
async function fetchOpenAI(apiKey, messages, model='gpt-4o'){
  const res = await fetch('https://api.openai.com/v1/chat/completions',{
    method:'POST',
    headers:{'Content-Type':'application/json','Authorization':`Bearer ${apiKey}`},
    body: JSON.stringify({model, messages})
  });
  return res.json();
}

async function fetchGemini(apiKey, messages, model='gemini-2.5-flash'){
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  const res = await fetch(url,{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({
      contents: messages.map(m=>({role: m.role==='assistant'?'model':'user', parts:[{text: m.content}]}))
    })
  });
  return res.json();
}

function pickProvider(){
  // If a specific model is selected, use that provider
  if(AppState.selectedModel !== 'auto') {
    const [provider] = AppState.selectedModel.split(':');
    if(provider === 'openai' && AppState.keys.openai) return 'openai';
    if(provider === 'gemini' && AppState.keys.gemini) return 'gemini';
  }
  
  // Auto selection - prefer OpenAI if available
  if(AppState.keys.openai) return 'openai';
  if(AppState.keys.gemini) return 'gemini';
  return 'mock';
}

function getSelectedModel(){
  if(AppState.selectedModel !== 'auto') {
    const [, model] = AppState.selectedModel.split(':');
    // Return the actual API model name (no mapping needed since we're using correct names)
    return model;
  }
  
  // Default models for auto selection
  const provider = pickProvider();
  if(provider === 'openai') return 'gpt-4o';
  if(provider === 'gemini') return 'gemini-2.5-flash';
  return null;
}

function extractFirstJson(text){
  if(typeof text !== 'string') return null;
  // Code fences
  const fence = text.match(/```json[\s\S]*?```/);
  if(fence){
    const inner = fence[0].replace(/```json/,'').replace(/```/,'').trim();
    try{ return JSON.parse(inner); }catch{}
  }
  // Brute force first {...}
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if(start>=0 && end>start){
    const slice = text.slice(start, end+1);
    try{ return JSON.parse(slice); }catch{}
  }
  return null;
}

function uuid(){ return 'id-'+Math.random().toString(36).slice(2)+Date.now().toString(36); }

function getSampleDataFromSheet(ws) {
  if (!ws['!ref']) return 'Empty sheet';
  
  const range = XLSX.utils.decode_range(ws['!ref']);
  const maxSampleRows = 3;
  const maxSampleCols = 5;
  
  let sample = [];
  for (let r = range.s.r; r <= Math.min(range.s.r + maxSampleRows - 1, range.e.r); r++) {
    let row = [];
    for (let c = range.s.c; c <= Math.min(range.s.c + maxSampleCols - 1, range.e.c); c++) {
      const addr = XLSX.utils.encode_cell({r, c});
      const cell = ws[addr];
      row.push(cell ? String(cell.v) : '');
    }
    sample.push(row.join('\t'));
  }
  
  const truncated = range.e.r > range.s.r + maxSampleRows - 1 || range.e.c > range.s.c + maxSampleCols - 1;
  return sample.join('\n') + (truncated ? '\n...(truncated)' : '');
}

async function runPlanner(userText){
  const provider = pickProvider();
  const tasks=[];
  
  try {
    if(provider==='mock'){
      tasks.push({id:uuid(), title:'Insert header row', description:'Add Name, Age, Email', status:'pending', context:{range:'A1:C1', sheet:AppState.activeSheet}, createdAt:new Date().toISOString()});
      return tasks;
    }
    
    // Get current sheet context for better planning
    const ws = getWorksheet();
    const sheetContext = ws['!ref'] ? `Current sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
    const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';
    
    const system = `You are the Planner Agent - an expert at analyzing user requests and breaking them down into precise, executable tasks for spreadsheet automation.

ROLE: Decompose complex spreadsheet operations into logical, sequential tasks that can be executed by specialized agents.

CAPABILITIES:
- Analyze natural language requests for spreadsheet operations
- Understand data patterns, relationships, and structure
- Plan multi-step workflows with dependencies
- Consider data validation and error handling needs
- Optimize task sequencing for efficiency

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Sheet structure: ${sheetContext}
- Sample data preview:
${sampleData}
- Available sheets: [${AppState.wb.SheetNames.join(', ')}]
- Total sheets: ${AppState.wb.SheetNames.length}

TASK BREAKDOWN STRATEGY:
1. Analyze the user request for complexity and dependencies
2. Identify required data operations (create, read, update, delete)
3. Consider data validation and formatting requirements
4. Plan for potential errors or edge cases
5. Sequence tasks logically with clear dependencies

OUTPUT FORMAT: Return a JSON array of task objects. Each task must include:
- "id": unique identifier
- "title": brief descriptive title
- "description": detailed operation description
- "priority": number 1-5 (1=highest)
- "dependencies": array of task IDs that must complete first
- "context": {"range": "A1:C10", "sheet": "SheetName", "operation": "type"}
- "validation": expected outcome or validation criteria

EXAMPLES:
For "Add totals row with formulas":
[
  {"id":"task1", "title":"Detect data range", "description":"Find the extent of existing data", "priority":1, "dependencies":[], "context":{"range":"detect", "sheet":"${AppState.activeSheet}", "operation":"analyze"}, "validation":"Data range identified"},
  {"id":"task2", "title":"Insert totals row", "description":"Add row below data for totals", "priority":2, "dependencies":["task1"], "context":{"range":"below_data", "sheet":"${AppState.activeSheet}", "operation":"insertRow"}, "validation":"Row inserted successfully"},
  {"id":"task3", "title":"Add SUM formulas", "description":"Create SUM formulas for numeric columns", "priority":3, "dependencies":["task2"], "context":{"range":"totals_row", "sheet":"${AppState.activeSheet}", "operation":"setFormula"}, "validation":"Formulas calculate correctly"}
]

IMPORTANT: 
- Always consider data integrity and user intent
- Plan for edge cases (empty data, invalid formats, etc.)
- Keep tasks atomic and focused
- Ensure proper sequencing with dependencies
- Include validation criteria for each task`;

    const messages=[{role:'system', content:system},{role:'user', content:userText}];
    let data;
    
    try {
      const selectedModel = getSelectedModel();
      if(provider==='openai'){ 
        data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel); 
      } else { 
        data = await fetchGemini(AppState.keys.gemini, messages, selectedModel); 
      }
    } catch(apiError) {
      console.error('API call failed:', apiError);
      showToast(`${provider} API call failed. Check your API key and internet connection.`, 'error');
      return [];
    }
    
    let text='';
    try{
      if(provider==='openai'){ 
        text = data.choices?.[0]?.message?.content || ''; 
        if(!text && data.error) {
          throw new Error(data.error.message || 'OpenAI API error');
        }
      } else { 
        text = data.candidates?.[0]?.content?.parts?.map(p=>p.text).join('') || ''; 
        if(!text && data.error) {
          throw new Error(data.error.message || 'Gemini API error');
        }
      }
    } catch(parseError) {
      console.error('Failed to parse API response:', parseError);
      showToast('Failed to parse AI response', 'error');
      return [];
    }
    
    if(!text) {
      showToast('AI returned empty response', 'warning');
      return [];
    }
    
    let arr=null;
    try{ 
      arr = JSON.parse(text); 
    } catch { 
      arr = extractFirstJson(text); 
    }
    
    if(Array.isArray(arr)){
      return arr.map(t=>({
        id: t.id || uuid(),
        title: t.title || (t.description || 'Task'),
        description: t.description || '',
        status: 'pending',
        priority: t.priority || 3,
        dependencies: t.dependencies || [],
        context: { ...t.context, sheet: t.context?.sheet || AppState.activeSheet },
        validation: t.validation || null,
        createdAt: new Date().toISOString(),
        estimatedDuration: t.estimatedDuration || null,
        retryCount: 0,
        maxRetries: 3
      }));
    } else {
      showToast('AI response was not in expected format', 'warning');
      return [];
    }
  } catch(error) {
    console.error('Planner failed:', error);
    showToast('Planning failed: ' + error.message, 'error');
    return [];
  }
}

async function runExecutor(task){
  const provider = pickProvider();
  if(provider==='mock'){
    return {
      edits:[
        {op:'setCell', sheet:AppState.activeSheet, cell:'A1', value:'Total'},
        {op:'setRange', sheet:AppState.activeSheet, range:'A2:C3', values:[['a',1,2],['b',3,4]]}
      ],
      export:null,
      message:`Mock applied 2 edits for ${task.title}`
    };
  }
  
  // Get current sheet context
  const ws = getWorksheet();
  const sheetContext = ws['!ref'] ? `Sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
  const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';
  
  const system = `You are the Executor Agent - a specialist in translating planned tasks into precise spreadsheet operations with intelligent analysis and error handling.

ROLE: Execute planned tasks by analyzing current spreadsheet state and generating optimal operation sequences.

CAPABILITIES:
- Analyze spreadsheet data patterns and structure
- Generate precise SheetJS-compatible operations
- Handle complex data transformations and calculations
- Implement intelligent error handling and rollback strategies
- Optimize operations for performance and data integrity

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Sheet structure: ${sheetContext}
- Sample data preview:
${sampleData}
- Available sheets: [${AppState.wb.SheetNames.join(', ')}]

EXECUTION STRATEGY:
1. Analyze current data structure and patterns
2. Determine optimal operation sequence
3. Consider data types and formatting requirements
4. Plan for edge cases and error conditions
5. Generate atomic, reversible operations

OPERATION SCHEMA (REQUIRED OUTPUT FORMAT):
{
  "success": true,
  "analysis": "Brief analysis of current state and planned changes",
  "edits": [
    {"op":"setCell","sheet":"SheetName","cell":"A1","value":"Total","dataType":"string"},
    {"op":"setRange","sheet":"SheetName","range":"A2:C3","values":[["a",1,2],["b",3,4]],"preserveTypes":true},
    {"op":"setFormula","sheet":"SheetName","cell":"D1","formula":"=SUM(A:A)"},
    {"op":"insertRow","sheet":"SheetName","row":2,"count":1},
    {"op":"deleteRow","sheet":"SheetName","row":2,"count":1},
    {"op":"insertColumn","sheet":"SheetName","col":"B","count":1},
    {"op":"deleteColumn","sheet":"SheetName","col":"B","count":1},
    {"op":"formatCell","sheet":"SheetName","cell":"A1","format":"0.00"},
    {"op":"formatRange","sheet":"SheetName","range":"A1:C3","format":"General"}
  ],
  "validation": {
    "expectedChanges": ["description of expected changes"],
    "rollbackPlan": ["steps to undo if needed"],
    "dataIntegrityChecks": ["validation points to verify"]
  },
  "warnings": ["any potential issues or considerations"],
  "message": "Detailed description of what was accomplished"
}

INTELLIGENT FEATURES:
- Auto-detect data types (numbers, dates, text, formulas)
- Preserve existing formatting where appropriate
- Handle formula dependencies and references
- Optimize range operations for efficiency
- Provide detailed rollback plans for safety

IMPORTANT: 
- Always analyze before executing
- Preserve data integrity and user intent
- Generate atomic, reversible operations
- Include comprehensive validation plans
- Handle edge cases gracefully`;

  const user = `Task: ${task.title}\nDescription: ${task.description||''}\nContext: ${JSON.stringify(task.context||{})}`;
  const messages=[{role:'system', content:system},{role:'user', content:user}];
  let data;
  const selectedModel = getSelectedModel();
  if(provider==='openai'){ data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel); }
  else { data = await fetchGemini(AppState.keys.gemini, messages, selectedModel); }
  let text='';
  try{
    if(provider==='openai'){ text = data.choices?.[0]?.message?.content || ''; }
    else{ text = data.candidates?.[0]?.content?.parts?.map(p=>p.text).join('') || ''; }
  }catch{ text=''; }
  let obj=null;
  try{ obj = JSON.parse(text); }catch{ obj = extractFirstJson(text); }
  log('Executor raw', text);
  return obj;
}

async function runValidator(executorObj, task) {
  const provider = pickProvider();
  
  // Basic schema validation first
  const basicResult = { valid: true, errors: [], warnings: [] };
  if(!executorObj || !Array.isArray(executorObj.edits)) {
    basicResult.valid = false;
    basicResult.errors.push('Missing edits array');
    return basicResult;
  }
  
  const supportedOps = ['setCell', 'setRange', 'setFormula', 'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'formatCell', 'formatRange'];
  for(const e of executorObj.edits) {
    if(!e.op) {
      basicResult.valid = false;
      basicResult.errors.push('Edit missing operation type');
      break;
    }
    if(supportedOps.indexOf(e.op) === -1) {
      basicResult.valid = false;
      basicResult.errors.push(`Unsupported operation: ${e.op}`);
      break;
    }
  }
  
  if(!basicResult.valid) return basicResult;
  
  // Advanced AI-powered validation
  if(provider === 'mock') {
    return {
      valid: true,
      confidence: 0.8,
      analysis: 'Mock validation - basic schema checks passed',
      risks: [],
      recommendations: [],
      dataIntegrityScore: 0.9
    };
  }
  
  try {
    const ws = getWorksheet();
    const sheetContext = ws['!ref'] ? `Sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
    const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';
    
    const system = `You are the Validator Agent - an expert in data integrity, conflict detection, and intelligent validation of spreadsheet operations.

ROLE: Analyze planned operations for potential conflicts, data integrity issues, and optimization opportunities while ensuring user intent is preserved.

CAPABILITIES:
- Deep data integrity analysis and conflict detection
- Formula dependency and reference validation  
- Performance impact assessment for large operations
- Data type consistency and format validation
- User intent preservation and goal alignment
- Risk assessment with confidence scoring

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Sheet structure: ${sheetContext}
- Sample data preview: ${sampleData}
- Available sheets: [${AppState.wb.SheetNames.join(', ')}]

VALIDATION STRATEGY:
1. Analyze data integrity and potential conflicts
2. Validate formula references and dependencies
3. Assess performance impact and optimization opportunities
4. Check data type consistency and formatting
5. Verify alignment with user intent and task goals
6. Identify potential risks and provide recommendations

REQUIRED OUTPUT FORMAT:
{
  "valid": true,
  "confidence": 0.95,
  "analysis": "Detailed analysis of the planned operations and their impact",
  "dataIntegrityScore": 0.9,
  "risks": [
    {"level": "medium", "description": "Potential data overwrite", "mitigation": "Create backup"},
    {"level": "low", "description": "Performance impact on large dataset", "mitigation": "Use batch operations"}
  ],
  "conflicts": [
    {"type": "formula_reference", "description": "Formula may reference moved cells", "severity": "high"}
  ],
  "optimizations": [
    "Batch similar cell operations for better performance",
    "Use range operations instead of individual cell updates"
  ],
  "recommendations": [
    "Execute in dry-run mode first",
    "Consider creating a backup before major structural changes"
  ],
  "userIntentAlignment": 0.95,
  "expectedOutcome": "Operations will successfully add totals row with proper formulas",
  "rollbackComplexity": "low",
  "warnings": ["Large dataset may impact browser performance"]
}

VALIDATION CRITERIA:
- Data integrity and consistency preservation
- Formula reference validity and dependency management
- Performance impact on current dataset size
- Alignment with original user request and task goals
- Potential for data loss or corruption
- Reversibility and rollback complexity

INTELLIGENCE FEATURES:
- Context-aware conflict detection
- Performance impact prediction
- User intent analysis and preservation
- Advanced risk assessment with mitigation strategies
- Optimization recommendations for efficiency`;

    const operations = {
      task: {
        id: task?.id,
        title: task?.title,
        description: task?.description,
        context: task?.context
      },
      executorResult: executorObj
    };
    
    const user = `Validate these planned operations:\n${JSON.stringify(operations, null, 2)}`;
    const messages = [{role:'system', content:system}, {role:'user', content:user}];
    
    let data;
    const selectedModel = getSelectedModel();
    if(provider === 'openai') {
      data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel);
    } else {
      data = await fetchGemini(AppState.keys.gemini, messages, selectedModel);
    }
    
    let text = '';
    if(provider === 'openai') {
      text = data.choices?.[0]?.message?.content || '';
    } else {
      text = data.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '';
    }
    
    let result = null;
    try {
      result = JSON.parse(text);
    } catch {
      result = extractFirstJson(text);
    }
    
    if(result && typeof result.valid === 'boolean') {
      return result;
    }
    
    // Fallback to basic validation
    return {
      valid: true,
      confidence: 0.7,
      analysis: 'AI validation failed, using basic schema validation',
      warnings: ['Advanced validation unavailable']
    };
    
  } catch(error) {
    console.error('Validator failed:', error);
    return {
      valid: true, // Don't block on validator failure
      confidence: 0.5,
      analysis: `Validation error: ${error.message}`,
      warnings: ['Validator agent unavailable - proceeding with basic validation only']
    };
  }
}

async function applyEditsOrDryRun(result){
  if(AppState.dryRun){
    const modal = new Modal();
    const content = `<pre class=\"text-xs bg-gray-50 p-3 rounded border border-gray-200 overflow-auto max-h-64\">${JSON.stringify(result, null, 2)}</pre>`;
    modal.show({
      title:'Dry Run: Review edits',
      content, buttons:[
        {text:'Cancel', action:'cancel'},
        {text:'Apply', action:'apply', primary:true, onClick:()=>applyEdits(result.edits)}
      ]
    });
  }else{
    await applyEdits(result.edits);
  }
}

async function applyEdits(edits){
  // Save to history before applying edits
  saveToHistory(`Apply AI edits (${edits.length} operations)`, { 
    edits: edits.map(e => ({...e})), 
    sheet: AppState.activeSheet 
  });
  
  const ws = getWorksheet();
  for(const e of edits){
    switch(e.op){
      case 'setCell':{
        const parsed = parseCellValue(e.value);
        ws[e.cell] = { t: parsed.t==='z'?'s':parsed.t, v: parsed.v };
        expandRefForCell(ws, e.cell);
        break;
      }
      case 'setRange':{
        if(!e.values || !Array.isArray(e.values)){ showToast('setRange missing values','warning'); break; }
        const start = (e.range||'A1').split(':')[0];
        XLSX.utils.sheet_add_aoa(ws, e.values, { origin: start });
        const r = XLSX.utils.decode_range(e.range||`${start}:${start}`);
        expandRefForCell(ws, XLSX.utils.encode_cell({r:r.e.r, c:r.e.c}));
        break;
      }
      case 'setFormula':{
        const { cell, formula } = e;
        if(!cell || !formula) {
          showToast('Invalid setFormula operation', 'warning');
          break;
        }
        ws[cell] = { t: 'str', f: formula.startsWith('=') ? formula : '=' + formula };
        expandRefForCell(ws, cell);
        break;
      }
      case 'insertRow': {
        const rowNumber = (typeof e.row === 'number' ? e.row : (typeof e.index === 'number' ? e.index : null));
        if (rowNumber === null || rowNumber < 1) {
          showToast('Invalid row/index for insertRow', 'warning');
          break;
        }
        const rowIndex = rowNumber - 1;
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        for (let R = range.e.r; R >= rowIndex; --R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const from = XLSX.utils.encode_cell({ r: R, c: C });
            const to = XLSX.utils.encode_cell({ r: R + 1, c: C });
            ws[to] = ws[from];
            delete ws[from];
          }
        }
        range.e.r++;
        ws['!ref'] = XLSX.utils.encode_range(range);
        break;
      }
      case 'deleteRow': {
        const rowNumber = (typeof e.row === 'number' ? e.row : (typeof e.index === 'number' ? e.index : null));
        if (rowNumber === null || rowNumber < 1) {
          showToast('Invalid row/index for deleteRow', 'warning');
          break;
        }
        const rowIndex = rowNumber - 1;
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        for (let R = rowIndex; R < range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const from = XLSX.utils.encode_cell({ r: R + 1, c: C });
            const to = XLSX.utils.encode_cell({ r: R, c: C });
            ws[to] = ws[from];
            delete ws[from];
          }
        }
        range.e.r--;
        ws['!ref'] = XLSX.utils.encode_range(range);
        break;
      }
      case 'insertColumn': {
        let colIndex;
        if (typeof e.col === 'string' && /^[A-Z]+$/.test(e.col)) {
          colIndex = XLSX.utils.decode_col(e.col);
        } else if (typeof e.index === 'number' && e.index >= 0) {
          colIndex = e.index;
        } else {
          showToast('Invalid column/index for insertColumn', 'warning');
          break;
        }
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        for (let C = range.e.c; C >= colIndex; --C) {
          for (let R = range.s.r; R <= range.e.r; ++R) {
            const from = XLSX.utils.encode_cell({ r: R, c: C });
            const to = XLSX.utils.encode_cell({ r: R, c: C + 1 });
            ws[to] = ws[from];
            delete ws[from];
          }
        }
        range.e.c++;
        ws['!ref'] = XLSX.utils.encode_range(range);
        break;
      }
      case 'deleteColumn': {
        let colIndex;
        if (typeof e.col === 'string' && /^[A-Z]+$/.test(e.col)) {
          colIndex = XLSX.utils.decode_col(e.col);
        } else if (typeof e.index === 'number' && e.index >= 0) {
          colIndex = e.index;
        } else {
          showToast('Invalid column/index for deleteColumn', 'warning');
          break;
        }
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        for (let C = colIndex; C < range.e.c; ++C) {
          for (let R = range.s.r; R <= range.e.r; ++R) {
            const from = XLSX.utils.encode_cell({ r: R, c: C + 1 });
            const to = XLSX.utils.encode_cell({ r: R, c: C });
            ws[to] = ws[from];
            delete ws[from];
          }
        }
        range.e.c--;
        ws['!ref'] = XLSX.utils.encode_range(range);
        break;
      }
      case 'formatCell': {
        const { cell, format } = e;
        if (!cell || !format) {
          showToast('Invalid formatCell operation', 'warning');
          break;
        }
        if (!ws[cell]) ws[cell] = { t: 'z', v: '' };
        ws[cell].z = format; // z is the number format string
        break;
      }
      case 'formatRange': {
        const { range, format } = e;
        if (!range || !format) {
          showToast('Invalid formatRange operation', 'warning');
          break;
        }
        const r = XLSX.utils.decode_range(range);
        for (let row = r.s.r; row <= r.e.r; row++) {
          for (let col = r.s.c; col <= r.e.c; col++) {
            const addr = XLSX.utils.encode_cell({ r: row, c: col });
            if (!ws[addr]) ws[addr] = { t: 'z', v: '' };
            ws[addr].z = format;
          }
        }
        break;
      }
      default:
        showToast(`Unknown op ${e.op}`, 'error');
    }
  }
  persistSnapshot();
  renderSpreadsheetTable();
}

// Import/Export
async function importFromFile(file){
  try {
    if(!file) {
      showToast('No file selected', 'warning');
      return;
    }
    
    const maxSize = 10 * 1024 * 1024; // 10MB
    if(file.size > maxSize) {
      showToast('File too large (max 10MB)', 'error');
      return;
    }
    
    showToast('Importing file...', 'info', 2000);
    
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {type:'array'});
    
    if(!wb.SheetNames || wb.SheetNames.length === 0) {
      showToast('Invalid Excel file: no sheets found', 'error');
      return;
    }
    
    AppState.wb = wb;
    AppState.activeSheet = wb.SheetNames[0] || 'Sheet1';
    await persistSnapshot();
    renderSheetTabs();
    renderSpreadsheetTable();
    showToast(`Imported workbook with ${wb.SheetNames.length} sheet(s)`, 'success');
    
  } catch(error) {
    console.error('Import failed:', error);
    showToast('Failed to import file: ' + error.message, 'error');
  }
}

function exportXLSX(){ 
  try {
    if(!AppState.wb) {
      showToast('No workbook to export', 'warning');
      return;
    }
    XLSX.writeFile(AppState.wb, 'workbook.xlsx');
    showToast('Workbook exported successfully', 'success', 2000);
  } catch(error) {
    console.error('XLSX export failed:', error);
    showToast('Failed to export XLSX: ' + error.message, 'error');
  }
}

function exportCSV(){
  try {
    if(!AppState.wb) {
      showToast('No workbook to export', 'warning');
      return;
    }
    
    const ws = getWorksheet();
    if(!ws || !ws['!ref']) {
      showToast('Current sheet is empty', 'warning');
      return;
    }
    
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const filename = `${AppState.activeSheet}.csv`;
    const a = document.createElement('a'); 
    a.href = url; 
    a.download = filename; 
    a.click(); 
    setTimeout(()=>URL.revokeObjectURL(url), 500);
    showToast(`Exported "${filename}" successfully`, 'success', 2000);
  } catch(error) {
    console.error('CSV export failed:', error);
    showToast('Failed to export CSV: ' + error.message, 'error');
  }
}

// UI bindings
// Keyboard shortcuts
function initKeyboardShortcuts(){
  document.addEventListener('keydown', (e) => {
    // Skip if user is typing in an input field
    if(e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.contentEditable === 'true') return;
    
    // Ctrl/Cmd + key combinations
    if(e.ctrlKey || e.metaKey) {
      switch(e.key.toLowerCase()) {
        case 's':
          e.preventDefault();
          exportXLSX();
          showToast('Workbook saved as XLSX', 'success', 1000);
          break;
        case 'o':
          e.preventDefault();
          document.getElementById('import-xlsx-input').click();
          break;
        case 't':
          e.preventDefault();
          addNewSheet();
          break;
        case 'w':
          e.preventDefault();
          if(AppState.wb && AppState.wb.SheetNames.length > 1) {
            deleteSheet(AppState.activeSheet);
          }
          break;
        case '1':
        case '2':
        case '3':
        case '4':
        case '5':
        case '6':
        case '7':
        case '8':
        case '9':
          e.preventDefault();
          const sheetIndex = parseInt(e.key) - 1;
          if(AppState.wb && AppState.wb.SheetNames[sheetIndex]) {
            switchToSheet(AppState.wb.SheetNames[sheetIndex]);
          }
          break;
        case 'enter':
          e.preventDefault();
          document.getElementById('message-input').focus();
          break;
        case 'z':
          e.preventDefault();
          if(e.shiftKey) {
            redo();
          } else {
            undo();
          }
          break;
        case 'y':
          e.preventDefault();
          redo();
          break;
      }
    }
    
    // Tab to switch between sheets
    if(e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const chatInput = document.getElementById('message-input');
      if(document.activeElement !== chatInput) {
        e.preventDefault();
        const currentIndex = AppState.wb.SheetNames.indexOf(AppState.activeSheet);
        const nextIndex = e.shiftKey ? 
          (currentIndex - 1 + AppState.wb.SheetNames.length) % AppState.wb.SheetNames.length :
          (currentIndex + 1) % AppState.wb.SheetNames.length;
        switchToSheet(AppState.wb.SheetNames[nextIndex]);
      }
    }
    
    // F2 to focus chat input
    if(e.key === 'F2') {
      e.preventDefault();
      document.getElementById('message-input').focus();
    }
    
    // Escape to clear chat input
    if(e.key === 'Escape') {
      const chatInput = document.getElementById('message-input');
      if(document.activeElement === chatInput) {
        chatInput.blur();
      }
    }
  });
}

function bindUI(){
  const openaiBtn = document.getElementById('openai-key-btn');
  if(openaiBtn) {
    openaiBtn.addEventListener('click', (e)=>{
      e.preventDefault();
      log('OpenAI button clicked');
      showApiKeyModal('OpenAI');
    });
  }
  
  const geminiBtn = document.getElementById('gemini-key-btn');
  if(geminiBtn) {
    geminiBtn.addEventListener('click', (e)=>{
      e.preventDefault();
      log('Gemini button clicked');
      showApiKeyModal('Gemini');
    });
  }
  
  const helpBtn = document.getElementById('help-btn');
  if(helpBtn) {
    helpBtn.addEventListener('click', showHelpModal);
  }
  
  const dryRunToggle = document.getElementById('dry-run-toggle');
  if(dryRunToggle) {
    dryRunToggle.addEventListener('change', (e)=>{ AppState.dryRun = e.target.checked; });
  }
  
  const modelSelect = document.getElementById('model-select');
  if(modelSelect) {
    modelSelect.addEventListener('change', (e)=>{ 
      AppState.selectedModel = e.target.value; 
      const provider = pickProvider();
      const model = getSelectedModel();
      showToast(`Selected: ${provider === 'mock' ? 'Mock Mode' : `${provider.toUpperCase()} - ${model}`}`, 'info', 2000);
    });
  }
  const sendBtn = document.getElementById('send-btn');
  if(sendBtn) {
    sendBtn.addEventListener('click', onSend);
  }
  
  const messageInput = document.getElementById('message-input');
  if(messageInput) {
    messageInput.addEventListener('keypress', (e)=>{ if(e.key==='Enter'){ e.preventDefault(); onSend(); } });
  }
  
  const exportXlsx = document.getElementById('export-xlsx');
  if(exportXlsx) {
    exportXlsx.addEventListener('click', exportXLSX);
  }
  
  const exportCsv = document.getElementById('export-csv');
  if(exportCsv) {
    exportCsv.addEventListener('click', exportCSV);
  }
  
  const importBtn = document.getElementById('import-xlsx');
  const importInput = document.getElementById('import-xlsx-input');
  if(importBtn && importInput) {
    importBtn.addEventListener('click', ()=>importInput.click());
    importInput.addEventListener('change', ()=>{ if(importInput.files?.[0]) importFromFile(importInput.files[0]); });
  }
  
  const addMock = document.getElementById('add-mock-task');
  if(addMock){ 
    addMock.addEventListener('click', ()=>{
      const t = {id:uuid(), title:'Mock: Add totals row', description:'Sum column B into C1', status:'pending', createdAt:new Date().toISOString()};
      AppState.tasks.push(t); saveTasks(); drawTasks();
    }); 
  }
  
  const executeAll = document.getElementById('execute-all-tasks');
  if(executeAll) { 
    executeAll.addEventListener('click', ()=>{
      const pendingTasks = AppState.tasks.filter(t => t.status === 'pending');
      if(pendingTasks.length === 0) {
        showToast('No pending tasks to execute', 'info');
        return;
      }
      executeTasks(pendingTasks.map(t => t.id));
    }); 
  }
  
  const addSheetBtn = document.getElementById('add-sheet-btn');
  if(addSheetBtn) {
    addSheetBtn.addEventListener('click', addNewSheet);
  }
  
  // Bind spreadsheet control buttons
  const insertRowBtn = document.getElementById('insert-row-btn');
  if(insertRowBtn) {
    insertRowBtn.addEventListener('click', insertRowAtEnd);
  }
  
  const insertColBtn = document.getElementById('insert-col-btn');
  if(insertColBtn) {
    insertColBtn.addEventListener('click', insertColumnAtEnd);
  }
  
  const deleteRowBtn = document.getElementById('delete-row-btn');
  if(deleteRowBtn) {
    deleteRowBtn.addEventListener('click', deleteLastRow);
  }
  
  const deleteColBtn = document.getElementById('delete-col-btn');
  if(deleteColBtn) {
    deleteColBtn.addEventListener('click', deleteLastColumn);
  }
  
  // Bind new Excel-like UI elements
  const toggleAiPanel = document.getElementById('toggle-ai-panel');
  if(toggleAiPanel) {
    toggleAiPanel.addEventListener('click', ()=>{
      const aiPanel = document.getElementById('ai-panel');
      aiPanel.style.display = aiPanel.style.display === 'none' ? 'flex' : 'none';
    });
  }
  
  const viewTasks = document.getElementById('view-tasks');
  if(viewTasks) {
    viewTasks.addEventListener('click', ()=>{
      document.getElementById('task-modal').classList.remove('hidden');
    });
  }
  
  const closeTaskModal = document.getElementById('close-task-modal');
  if(closeTaskModal) {
    closeTaskModal.addEventListener('click', ()=>{
      document.getElementById('task-modal').classList.add('hidden');
    });
  }
  
  // Formula bar functionality
  const formulaBar = document.getElementById('formula-bar');
  if(formulaBar) {
    formulaBar.addEventListener('keypress', (e)=>{
      if(e.key === 'Enter') {
        const cellRef = document.getElementById('cell-reference').textContent;
        updateCell(cellRef, formulaBar.value);
        e.preventDefault();
      }
    });
  }
  
  // Initialize keyboard shortcuts
  initKeyboardShortcuts();
}

// Enhanced initialization with loading state
function showLoadingOverlay() {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    overlay.classList.remove('hidden');
  }
}

function hideLoadingOverlay() {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    overlay.classList.add('hidden');
  }
}

// Init
document.addEventListener('DOMContentLoaded', async ()=>{
  showLoadingOverlay();
  
  try{
    await db.init();
    restoreApiKeys();
    await ensureWorkbook();
    await loadTasks();
    renderSheetTabs();
    renderSpreadsheetTable();
    drawTasks();
    bindUI();
    initDraggableDivider();
    restorePanelLayout();
    updateProviderStatus(); // Update button states based on available keys
    
    // Add fade-in animation to main content
    document.querySelector('.main-container').classList.add('animate-fade-in-up');
    
    hideLoadingOverlay();
    
    // Show appropriate welcome message
    const hasKeys = AppState.keys.openai || AppState.keys.gemini;
    if(hasKeys) {
      const provider = AppState.keys.openai ? 'OpenAI' : 'Gemini';
      showToast(`AI Excel Editor ready! Using ${provider} for AI features.`, 'success', 3000);
    } else {
      showToast('AI Excel Editor ready! Set your API keys to enable AI features.', 'success', 3000);
    }
  }catch(e){
    console.error("Initialization failed", e);
    hideLoadingOverlay();
    showToast('Error initializing application: ' + e.message, 'error');
  }
});