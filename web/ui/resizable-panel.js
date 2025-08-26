document.addEventListener('DOMContentLoaded', () => {
    const resizer = document.getElementById('divider');
    const panel = document.getElementById('ai-panel');

    let isResizing = false;

    resizer.addEventListener('mousedown', (e) => {
        isResizing = true;
        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', () => {
            isResizing = false;
            document.removeEventListener('mousemove', handleMouseMove);
            // Optional: Save the new width to localStorage to persist it
            // localStorage.setItem('aiPanelWidth', panel.style.width);
        });
    });

    function handleMouseMove(e) {
        if (!isResizing) return;
        const newWidth = document.body.clientWidth - e.clientX;
        if (newWidth > 200 && newWidth < 800) { // Min and max width constraints
            panel.style.width = `${newWidth}px`;
        }
    }

    // Optional: Restore the width from localStorage on page load
    // const savedWidth = localStorage.getItem('aiPanelWidth');
    // if (savedWidth) {
    //     panel.style.width = savedWidth;
    // }
});