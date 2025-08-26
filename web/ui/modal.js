'use strict';

export class Modal {
  constructor() { this.container = document.getElementById('modal-container'); this.currentModal = null; }
  show({ title, content, buttons = [], size = 'md', closable = true }) {
    const sizeClasses = { sm: 'max-w-sm', md: 'max-w-md', lg: 'max-w-lg', xl: 'max-w-xl', full: 'max-w-full' };
    const html = `
    <div class="fixed inset-0 z-50 overflow-y-auto" id="modal-overlay">
      <div class="flex items-center justify-center min-h-screen px-4 pt-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 transition-opacity bg-gray-500 bg-opacity-75" id="modal-backdrop"></div>
        <div class="inline-block w-full ${sizeClasses[size]} p-6 my-8 overflow-hidden text-left align-middle transition-all transform bg-white shadow-xl rounded-lg">
          <div class="flex items-center justify-between mb-4">
            <h3 class="text-lg font-medium text-gray-900">${title}</h3>
            ${closable ? `<button id="modal-close" class="text-gray-400 hover:text-gray-600 focus:outline-none">
              <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
            </button>`: ''}
          </div>
          <div class="mb-6">${content}</div>
          <div class="flex justify-end space-x-3">
            ${buttons.map(btn => `
              <button data-action="${btn.action}" class="px-4 py-2 text-sm font-medium rounded-md focus:outline-none focus:ring-2 focus:ring-offset-2 ${btn.primary ? 'bg-blue-500 hover:bg-blue-600 text-white focus:ring-blue-500' : 'bg-gray-300 hover:bg-gray-400 text-gray-700 focus:ring-gray-500'}">${btn.text}</button>
            `).join('')}
          </div>
        </div>
      </div>
    </div>`;
    this.container.innerHTML = html;
    this.currentModal = document.getElementById('modal-overlay');
    if (closable) {
      document.getElementById('modal-close').addEventListener('click', () => this.close());
      document.getElementById('modal-backdrop').addEventListener('click', () => this.close());
    }
    buttons.forEach(btn => {
      const el = this.container.querySelector(`[data-action="${btn.action}"]`);
      if (el && btn.onClick) {
        el.addEventListener('click', e => {
          e.preventDefault();
          btn.onClick(e);
          if (btn.closeOnClick !== false) this.close();
        });
      }
    });
    return this.currentModal;
  }
  close() { if (this.currentModal) { this.currentModal.remove(); this.currentModal = null; } }
}