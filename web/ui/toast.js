'use strict';

class Toast {
    constructor() {
        this.container = document.getElementById('toast-container');
    }

    show(message, type = 'info', duration = 5000) {
        const id = 't' + Date.now();
        const toast = document.createElement('div');
        toast.id = id;
        toast.className = `toast toast-${type}`;
        toast.setAttribute('role', 'alert');
        toast.innerHTML = `
            <div class="toast-icon"></div>
            <div class="toast-content">
                <p class="toast-message">${message}</p>
            </div>
            <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
        `;

        this.container.appendChild(toast);

        setTimeout(() => {
            toast.classList.add('show');
        }, 100);

        if (duration > 0) {
            setTimeout(() => {
                toast.classList.remove('show');
                setTimeout(() => toast.remove(), 500);
            }, duration);
        }

        return toast;
    }
}

const toast = new Toast();
export function showToast(msg, type = 'info', dur = 5000) { return toast.show(msg, type, dur); }