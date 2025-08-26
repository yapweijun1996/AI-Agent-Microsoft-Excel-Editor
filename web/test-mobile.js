// Test mobile responsiveness after render
function testMobileResponsiveness() {
  console.log('=== Mobile Responsiveness Test ===');
  console.log('Window width:', window.innerWidth);
  console.log('Is Mobile:', window.innerWidth <= 768);
  console.log('Is Small Mobile:', window.innerWidth <= 480);
  
  const body = document.body;
  console.log('Body classes:', body.className);
  
  const spreadsheet = document.querySelector('.modern-spreadsheet');
  if (spreadsheet) {
    console.log('Spreadsheet classes:', spreadsheet.className);
  }
  
  const cells = document.querySelectorAll('.modern-cell');
  if (cells.length > 0) {
    const firstCell = cells[0];
    const computedStyle = window.getComputedStyle(firstCell);
    console.log('First cell width:', computedStyle.width);
    console.log('First cell height:', computedStyle.height);
  }
  
  const colHeaders = document.querySelectorAll('.col-header');
  if (colHeaders.length > 0) {
    const firstHeader = colHeaders[0];
    const computedStyle = window.getComputedStyle(firstHeader);
    console.log('First header width:', computedStyle.width);
    console.log('First header height:', computedStyle.height);
  }
}

// Run test when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    setTimeout(testMobileResponsiveness, 1000);
  });
} else {
  setTimeout(testMobileResponsiveness, 1000);
}

// Also run on window resize
window.addEventListener('resize', () => {
  setTimeout(testMobileResponsiveness, 500);
});