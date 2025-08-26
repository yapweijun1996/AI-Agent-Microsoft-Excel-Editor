'use strict';

export const db = {
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