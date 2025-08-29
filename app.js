document.addEventListener('DOMContentLoaded', () => {
    const gridEl = document.getElementById('grid');
    const thead = gridEl.querySelector('thead');
    const tbody = gridEl.querySelector('tbody');
      const calcState = document.getElementById('calcState');
      const fileInfo = document.getElementById('fileInfo');
      const xlsxState = document.getElementById('xlsxState');
      const sheetTabs = document.getElementById('sheetTabs');
      const saveXLSXBtn = document.getElementById('saveXLSX');
      const debugBox = document.getElementById('debugBox');
      const debugOut = document.getElementById('debugOut');
      const formulaBar = document.getElementById('formulaBar');
      const boldBtn = document.getElementById('boldBtn');
      const italicBtn = document.getElementById('italicBtn');
      const fillColorInput = document.getElementById('fillColor');
      const undoBtn = document.getElementById('undoBtn');
      const redoBtn = document.getElementById('redoBtn');

    // Sheet management
      function createSheet(name, r=30, c=12){
        return {name, rows:r, cols:c, data:createEmpty(r,c), colWidths:Array(c).fill(null), rowHeights:Array(r).fill(null)};
      }
      const sheets = [createSheet('Sheet1')];
      let activeSheetIndex = 0;
      let rows = sheets[0].rows, cols = sheets[0].cols, data = sheets[0].data;
      let colWidths = sheets[0].colWidths, rowHeights = sheets[0].rowHeights;
      let copyOrigin = null; // track source cell for copy/paste
      let activeCell = {r:0, c:0};
      let lastHeader = {r:-1, c:-1};
      const undoStack = [];
      const redoStack = [];

    // Error map: key "r,c" -> message
    const errMap = new Map();

    // XLSX lib reference (avoid global window access)
    let XLSXRef = null; // set after ensureXLSX()

    // Sheet tab helpers
    function renderTabs(){
      sheetTabs.innerHTML='';
      sheets.forEach((s,i)=>{
        const tab=document.createElement('div');
        tab.className='sheetTab'+(i===activeSheetIndex?' active':'');
        tab.textContent=s.name;
        tab.dataset.idx=i;
        const close=document.createElement('span');
        close.textContent='×';
        close.className='close';
        tab.appendChild(close);
        sheetTabs.appendChild(tab);
      });
      const add=document.createElement('div');
      add.className='sheetTab add';
      add.textContent='+';
      sheetTabs.appendChild(add);
    }
    function saveActiveState(){
      const s=sheets[activeSheetIndex];
      s.rows=rows; s.cols=cols; s.data=data; s.colWidths=colWidths; s.rowHeights=rowHeights;
    }
    function loadSheet(idx){
      const s=sheets[idx];
      rows=s.rows; cols=s.cols; data=s.data; colWidths=s.colWidths; rowHeights=s.rowHeights;
      activeSheetIndex=idx;
      renderHeader(); renderBody(); recalc();
    }
    function addSheet(){
      saveActiveState();
      const name=`Sheet${sheets.length+1}`;
      sheets.push(createSheet(name));
      loadSheet(sheets.length-1);
      renderTabs();
    }
    function switchSheet(idx){
      if(idx===activeSheetIndex) return;
      saveActiveState();
      loadSheet(idx);
      renderTabs();
    }
    function deleteSheet(idx){
      if(sheets.length===1) return;
      sheets.splice(idx,1);
      if(activeSheetIndex>=sheets.length) activeSheetIndex=sheets.length-1;
      loadSheet(activeSheetIndex);
      renderTabs();
    }
    sheetTabs.addEventListener('click', e=>{
      const t=e.target;
      if(t.classList.contains('close')){
        deleteSheet(+t.parentElement.dataset.idx);
      }else if(t.classList.contains('add')){
        addSheet();
      }else{
        const tab=t.closest('.sheetTab');
        if(tab) switchSheet(+tab.dataset.idx);
      }
    });
    sheetTabs.addEventListener('dblclick', e=>{
      const tab=e.target.closest('.sheetTab');
      if(!tab || tab.classList.contains('add')) return;
      const idx=+tab.dataset.idx;
      const name=prompt('Rename sheet', sheets[idx].name);
      if(name){ sheets[idx].name=name; renderTabs(); }
    });

    // ===== Debug helpers =====
    function log(...args){
      const line = `[${new Date().toISOString()}] ` + args.map(a=>{
        try{ return typeof a==='string'? a : JSON.stringify(a); }catch{ return String(a); }
      }).join(' ');
      console.log(...args);
      debugOut.textContent += line + "\n";
      debugOut.parentElement.scrollTop = debugOut.parentElement.scrollHeight;
    }
    document.getElementById('toggleDebug').onclick = ()=>{
      debugBox.style.display = debugBox.style.display==='none' ? 'block' : 'none';
    };
    document.getElementById('copyDebug').onclick = async ()=>{
      const report = collectReport();
      try {
        await navigator.clipboard.writeText(report);
        fileInfo.textContent = 'Debug report copied to clipboard';
      } catch (e) {
        fileInfo.textContent = 'Failed to copy debug report';
        log('Failed to copy debug report', e);
      }
    };
    function collectReport(){
      // basic environment + last 200 lines of debug
      const env = {
        ua: navigator.userAgent,
        xlsxLoaded: !!(XLSXRef||window.XLSX),
        rows, cols,
        errCount: errMap.size
      };
      const lines = debugOut.textContent.split(/\n/);
      const tail = lines.slice(Math.max(0, lines.length-200)).join('\n');
      return `ENV: ${JSON.stringify(env, null, 2)}\n\nLOG:\n${tail}`;
    }

    // ===== Robust XLSX loader (prevents ReferenceError) =====
    async function ensureXLSX(){
      if (XLSXRef) { xlsxState.textContent = 'XLSX: loaded'; return XLSXRef; }
      xlsxState.textContent = 'XLSX: loading…';
      const sources = [
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
        'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
        'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
      ];
      for (const src of sources){
        try{
          await loadScript(src, 12000);
          if (window.XLSX){
            XLSXRef = window.XLSX;
            xlsxState.textContent = 'XLSX: loaded';
            log('XLSX loaded from', src);
            return XLSXRef;
          }
        } catch(e){
          log('XLSX load failed from', src, e.message);
        }
      }
      xlsxState.textContent = 'XLSX: unavailable (CSV only)';
      log('XLSX unavailable after trying all CDNs');
      return null;
    }
    function loadScript(src, timeout=10000){
      return new Promise((resolve,reject)=>{
        const s = document.createElement('script');
        s.src = src; s.async = true; s.onload = () => resolve(); s.onerror = () => reject(new Error('load failed'));
        document.head.appendChild(s);
        const to = setTimeout(()=>{ reject(new Error('timeout')); }, timeout);
        s.addEventListener('load', ()=> clearTimeout(to));
        s.addEventListener('error', ()=> clearTimeout(to));
      });
    }

    // Optional ZIP library for multi-sheet CSV export
    let JSZipRef = null;
    async function ensureJSZip(){
      if (JSZipRef) return JSZipRef;
      const sources = [
        'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
        'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
        'https://unpkg.com/jszip@3.10.1/dist/jszip.min.js'
      ];
      for (const src of sources){
        try{
          await loadScript(src,12000);
          if (window.JSZip){ JSZipRef = window.JSZip; log('JSZip loaded from', src); return JSZipRef; }
        }catch(e){ log('JSZip load failed from', src, e.message); }
      }
      log('JSZip unavailable'); return null;
    }

    // Utilities
    function createCell(){ return { value:'', bold:false, italic:false, bgColor:'' }; }
    function createEmpty(r,c){ return Array.from({length:r},()=>Array.from({length:c},()=>createCell())); }
    function colLabel(n){ // 0->A, 25->Z, 26->AA
      let s=''; n = n>>>0; do{ s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n/26) - 1; } while(n>=0); return s;
    }
    function parseA1(ref){
      const m = /^(\$?)([A-Z]+)(\$?)(\d+)$/.exec(ref.toUpperCase());
      if(!m) return null;
      const absCol = !!m[1];
      const absRow = !!m[3];
      const c = m[2].split('').reduce((a,ch)=>a*26 + (ch.charCodeAt(0)-64),0)-1;
      const r = parseInt(m[4],10)-1;
      return {r, c, absRow, absCol};
    }

    function shiftFormulaRefs(expr, dr, dc){
      return expr.replace(/\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?/g, ref=>{
        if(ref.includes(':')){
          const [a,b] = ref.split(':');
          return shiftSingle(a)+':'+shiftSingle(b);
        }
        return shiftSingle(ref);
      });
      function shiftSingle(rf){
        const p = parseA1(rf);
        if(!p) return rf;
        const row = p.absRow ? p.r : p.r + dr;
        const col = p.absCol ? p.c : p.c + dc;
        return (p.absCol?'$':'') + colLabel(col) + (p.absRow?'$':'') + (row+1);
      }
    }

    function snapshot(){
      return { data: data.map(r=>r.map(cell=>({...cell}))), rows, cols };
    }
    function pushUndo(){
      undoStack.push(snapshot());
      redoStack.length = 0;
    }
    function restore(state){
      data = state.data.map(r=>r.map(cell=>({...cell})));
      rows = state.rows;
      cols = state.cols;
      renderHeader();
      renderBody();
      recalc();
    }
    function undo(){
      if(!undoStack.length) return;
      redoStack.push(snapshot());
      const prev = undoStack.pop();
      restore(prev);
    }
    function redo(){
      if(!redoStack.length) return;
      undoStack.push(snapshot());
      const next = redoStack.pop();
      restore(next);
    }

    // Rendering
    function applyCellStyles(el, cell){
      el.style.fontWeight = cell.bold ? 'bold' : '';
      el.style.fontStyle = cell.italic ? 'italic' : '';
      el.style.backgroundColor = cell.bgColor || '';
    }

    function syncSizeArrays(){
      while(colWidths.length < cols) colWidths.push(null);
      if(colWidths.length > cols) colWidths.length = cols;
      while(rowHeights.length < rows) rowHeights.push(null);
      if(rowHeights.length > rows) rowHeights.length = rows;
    }

    function applyColWidth(c, w){
      const header = thead.querySelectorAll('th')[c+1];
      if(header){
        header.style.width = w + 'px';
        header.style.maxWidth = w + 'px';
      }
      const cells = tbody.querySelectorAll(`td:nth-child(${c+2})`);
      cells.forEach(td=>{ td.style.width = w + 'px'; td.style.maxWidth = w + 'px'; });
    }

    function applyRowHeight(r, h){
      const tr = tbody.querySelectorAll('tr')[r];
      if(tr){
        tr.style.height = h + 'px';
        const th = tr.querySelector('th');
        if(th) th.style.height = h + 'px';
        tr.querySelectorAll('td').forEach(td=>{
          td.style.height = h + 'px';
          const div = td.firstElementChild;
          if(div) div.style.height = h + 'px';
        });
      }
    }

    let resizingCol = null, startX = 0, startWidth = 0;
    function startColResize(e){
      e.preventDefault();
      resizingCol = parseInt(e.target.dataset.c,10);
      startX = e.clientX;
      const th = thead.querySelectorAll('th')[resizingCol+1];
      startWidth = th.offsetWidth;
      document.addEventListener('mousemove', onColResize);
      document.addEventListener('mouseup', stopColResize);
    }
    function onColResize(e){
      if(resizingCol===null) return;
      const delta = e.clientX - startX;
      const newW = Math.max(40, startWidth + delta);
      colWidths[resizingCol] = newW;
      applyColWidth(resizingCol, newW);
    }
    function stopColResize(){
      document.removeEventListener('mousemove', onColResize);
      document.removeEventListener('mouseup', stopColResize);
      resizingCol = null;
    }

    let resizingRow = null, startY = 0, startHeight = 0;
    function startRowResize(e){
      e.preventDefault();
      resizingRow = parseInt(e.target.dataset.r,10);
      startY = e.clientY;
      const tr = tbody.querySelectorAll('tr')[resizingRow];
      startHeight = tr.offsetHeight;
      document.addEventListener('mousemove', onRowResize);
      document.addEventListener('mouseup', stopRowResize);
    }
    function onRowResize(e){
      if(resizingRow===null) return;
      const delta = e.clientY - startY;
      const newH = Math.max(24, startHeight + delta);
      rowHeights[resizingRow] = newH;
      applyRowHeight(resizingRow, newH);
    }
    function stopRowResize(){
      document.removeEventListener('mousemove', onRowResize);
      document.removeEventListener('mouseup', stopRowResize);
      resizingRow = null;
    }
    function renderHeader(){
      syncSizeArrays();
      const tr = document.createElement('tr');
      tr.appendChild(document.createElement('th')); // corner
      for(let c=0;c<cols;c++){
        const th = document.createElement('th');
        th.textContent = colLabel(c);
        if(colWidths[c]!=null){
          th.style.width = colWidths[c] + 'px';
          th.style.maxWidth = colWidths[c] + 'px';
        }
        const handle = document.createElement('div');
        handle.className = 'col-resizer';
        handle.dataset.c = c;
        handle.addEventListener('mousedown', startColResize);
        th.appendChild(handle);
        tr.appendChild(th);
      }
      thead.innerHTML='';
      thead.appendChild(tr);
    }
    function renderBody(){
      syncSizeArrays();
      tbody.innerHTML='';
      for(let r=0;r<rows;r++){
        const tr = document.createElement('tr');
        if(rowHeights[r]!=null) tr.style.height = rowHeights[r] + 'px';
        const rowTh = document.createElement('th');
        rowTh.textContent = r+1;
        if(rowHeights[r]!=null) rowTh.style.height = rowHeights[r] + 'px';
        const rHandle = document.createElement('div');
        rHandle.className = 'row-resizer';
        rHandle.dataset.r = r;
        rHandle.addEventListener('mousedown', startRowResize);
        rowTh.appendChild(rHandle);
        tr.appendChild(rowTh);
        for(let c=0;c<cols;c++){
          const td = document.createElement('td');
          if(colWidths[c]!=null){
            td.style.width = colWidths[c] + 'px';
            td.style.maxWidth = colWidths[c] + 'px';
          }
          if(rowHeights[r]!=null) td.style.height = rowHeights[r] + 'px';
          const div = document.createElement('div');
            div.className='cell';
            div.contentEditable = true;
            div.dataset.r = r;
            div.dataset.c = c;
            div.textContent = displayValue(r,c);
            applyCellStyles(div, data[r][c]);
            if(rowHeights[r]!=null) div.style.height = rowHeights[r] + 'px';
            div.addEventListener('input', onEdit);
            div.addEventListener('blur', onBlurNormalize);
            div.addEventListener('focus', onCellFocus);
            td.appendChild(div);
            tr.appendChild(td);
          }
          tbody.appendChild(tr);
        }
        applyErrorDecorations();
        // initial autofit for the newly rendered table
        autofitColumns();
        setActiveCell(activeCell.r, activeCell.c);
      }

      // Caret helpers for flicker-free refresh
      function setActiveCell(r,c){
        r = Math.max(0, Math.min(rows-1, r));
        c = Math.max(0, Math.min(cols-1, c));
        activeCell = {r,c};
        const cell = data[r][c];
        formulaBar.value = rawValue(r,c);
        formulaBar.dataset.r = r;
        formulaBar.dataset.c = c;
        if(fillColorInput) fillColorInput.value = cell.bgColor || '#ffffff';

        // Highlight row/column headers for the active cell
        if (typeof lastHeader !== 'undefined') {
          if (lastHeader.r >= 0) {
            const prevRowTh = tbody.querySelectorAll('tr')[lastHeader.r]?.querySelector('th');
            if (prevRowTh) prevRowTh.classList.remove('active');
          }
          if (lastHeader.c >= 0) {
            const prevColTh = thead.querySelectorAll('th')[lastHeader.c+1];
            if (prevColTh) prevColTh.classList.remove('active');
          }
        }
        const rowTh = tbody.querySelectorAll('tr')[r]?.querySelector('th');
        if (rowTh) rowTh.classList.add('active');
        const colTh = thead.querySelectorAll('th')[c+1];
        if (colTh) colTh.classList.add('active');
        lastHeader = { r, c };
      }
      function getCaret() {
        const el = document.activeElement;
        if (!el || !el.classList || !el.classList.contains('cell')) return null;
        const sel = window.getSelection();
        if (!sel || sel.rangeCount === 0){
          setActiveCell(+el.dataset.r, +el.dataset.c);
          return null;
        }
        const range = sel.getRangeAt(0);
        const r = +el.dataset.r, c = +el.dataset.c;
        setActiveCell(r,c);
        return { el, r, c, start: range.startOffset, end: range.endOffset };
      }
    function setCaret(snap) {
      if (!snap) return;
      const { r, c, start, end } = snap;
      const el = tbody.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
      if (!el) return;
      const node = el.firstChild || el;
      const range = document.createRange();
      const len = (node.textContent || '').length;
      range.setStart(node, Math.min(start, len));
      range.setEnd(node, Math.min(end, len));
      const sel = window.getSelection();
      sel.removeAllRanges(); sel.addRange(range);
    }

      function refreshAllDisplay(){
        const snap = getCaret();
        const active = document.activeElement;
        for(const el of tbody.querySelectorAll('.cell')){
          const r = +el.dataset.r, c = +el.dataset.c;
          if (el === active || (active === formulaBar && activeCell.r === r && activeCell.c === c)) continue; // keep user input while editing
          const newText = displayValue(r,c);
          if (el.textContent !== newText) el.textContent = newText;
          applyCellStyles(el, data[r][c]);
        }
        applyErrorDecorations();
        setCaret(snap);
      }

    // Formula evaluation (simple & safe-ish)
    function rawValue(r,c){ return data[r]?.[c]?.value ?? ''; }
    const VALUE_ERROR = {error:'#VALUE!'};
    const CIRC_ERROR = {error:'#CIRC!'};
    function isErr(v){ return v && typeof v === 'object' && 'error' in v; }
    function numeric(v){
      if(isErr(v)) return v;
      if(v === '') return VALUE_ERROR;
      const n = Number(v);
      return isFinite(n) ? n : VALUE_ERROR;
    }
    function flatten(args){
      const out=[];
      for(const a of args){
        if(Array.isArray(a)) out.push(...a);
        else out.push(numeric(a));
      }
      return out;
    }
    function sumFn(args){
      const v=flatten(args);
      for(const x of v) if(isErr(x)) return x;
      return v.reduce((a,b)=>a+b,0);
    }
    function minFn(args){
      const v=flatten(args);
      for(const x of v) if(isErr(x)) return x;
      return v.length?Math.min(...v):0;
    }
    function maxFn(args){
      const v=flatten(args);
      for(const x of v) if(isErr(x)) return x;
      return v.length?Math.max(...v):0;
    }
    function avgFn(args){
      const v=flatten(args);
      for(const x of v) if(isErr(x)) return x;
      const s=sumFn(v);
      return isErr(s)?s:(v.length?s/v.length:0);
    }
    const fnMap = {SUM: sumFn, MIN: minFn, MAX: maxFn, AVERAGE: avgFn};

    function isBlankFormula(s){ return typeof s==='string' && /^=\s*$/.test(s); }

    function tokenize(str){
      const tokens=[]; let i=0;
      const len=str.length;
      const isDigit=ch=>/\d/.test(ch);
      const isAlpha=ch=>/[A-Za-z]/.test(ch);
      while(i<len){
        const ch=str[i];
        if(ch===" "||ch==="\t"||ch==="\n"||ch==="\r"){ i++; continue; }
        if("+-*/(),".includes(ch)){ tokens.push({type:ch}); i++; continue; }
        if(isDigit(ch)){
          let s=i; while(isDigit(str[i])) i++; if(str[i]==='.') { i++; while(isDigit(str[i])) i++; }
          tokens.push({type:'num', value:parseFloat(str.slice(s,i))});
          continue;
        }
        if(isAlpha(ch) || ch==='$'){
          const m = /^\$?[A-Za-z]+\$?\d+/.exec(str.slice(i));
          if(m){
            const cell1=parseA1(m[0]); i+=m[0].length;
            if(str[i]===':'){
              const m2 = /^\$?[A-Za-z]+\$?\d+/.exec(str.slice(i+1));
              if(!m2) throw new Error('Invalid range');
              const cell2=parseA1(m2[0]);
              i += 1 + m2[0].length;
              tokens.push({type:'range', start:cell1, end:cell2});
            }else{
              tokens.push({type:'cell', pos:cell1});
            }
            continue;
          }
          let s=i; while(isAlpha(str[i])) i++; const letters=str.slice(s,i).toUpperCase();
          tokens.push({type:'id', value:letters});
          continue;
        }
        throw new Error('Unexpected character '+ch);
      }
      return tokens;
    }

    function evaluateFormula(expr, visited){
      const tokens=tokenize(expr.trim()); let i=0;
      function peek(){ return tokens[i]; }
      function consume(type){ const t=tokens[i]; if(!t||t.type!==type) throw new Error('Expected '+type); i++; return t; }
      function parseExpression(){
        let val=parseTerm();
        while(peek() && (peek().type==='+'||peek().type==='-')){
          if(isErr(val)) return val;
          const op=consume(peek().type).type; const rhs=parseTerm();
          if(isErr(rhs)) return rhs;
          if(Array.isArray(val)||Array.isArray(rhs)) throw new Error('Invalid range in expression');
          val = op==='+'? val+rhs : val-rhs;
        }
        return val;
      }
      function parseTerm(){
        let val=parseFactor();
        while(peek() && (peek().type==='*'||peek().type==='/')){
          if(isErr(val)) return val;
          const op=consume(peek().type).type; const rhs=parseFactor();
          if(isErr(rhs)) return rhs;
          if(Array.isArray(val)||Array.isArray(rhs)) throw new Error('Invalid range in expression');
          val = op==='*'? val*rhs : val/rhs;
        }
        return val;
      }
      function parseFactor(){
        const t=peek();
        if(t && t.type==='+'){ consume('+'); return parseFactor(); }
        if(t && t.type==='-'){ consume('-'); const v=parseFactor(); if(isErr(v)) return v; return Array.isArray(v)?(()=>{throw new Error('Invalid range in expression');})(): -v; }
        return parsePrimary();
      }
      function parsePrimary(){
        const t=peek(); if(!t) throw new Error('Unexpected end');
        if(t.type==='num'){ consume('num'); return t.value; }
        if(t.type==='cell'){ consume('cell'); return numeric(valueAt(t.pos.r,t.pos.c, visited)); }
        if(t.type==='range'){
          consume('range'); const out=[];
          const r1=Math.min(t.start.r,t.end.r), r2=Math.max(t.start.r,t.end.r);
          const c1=Math.min(t.start.c,t.end.c), c2=Math.max(t.start.c,t.end.c);
          for(let r=r1;r<=r2;r++) for(let c=c1;c<=c2;c++) out.push(numeric(valueAt(r,c, visited)));
          return out;
        }
        if(t.type==='id') return parseFunctionCall();
        if(t.type==='('){ consume('('); const v=parseExpression(); consume(')'); return v; }
        throw new Error('Unexpected token');
      }
      function parseFunctionCall(){
        const name=consume('id').value; consume('(');
        const args=[]; if(peek() && peek().type!==')'){ do{ args.push(parseExpression()); if(peek() && peek().type===',') consume(','); else break; }while(true); }
        consume(')'); const fn=fnMap[name]; if(!fn) throw new Error(`Unknown function ${name}`);
        return fn(args);
      }
      const result=parseExpression();
      if(i<tokens.length) throw new Error('Unexpected token');
      if(Array.isArray(result)) throw new Error('Invalid range in expression');
      return result;
    }

    function setError(r,c,msg){ errMap.set(`${r},${c}`, msg); }
    function clearError(r,c){ errMap.delete(`${r},${c}`); }

    function valueAt(r,c, visited){
      const top = !visited;
      visited = visited || new Set();
      const key = `${r},${c}`;
      if(visited.has(key)){
        setError(r,c, '#CIRC!');
        return CIRC_ERROR;
      }
      visited.add(key);
      try{
        const raw = rawValue(r,c);
        if(typeof raw !== 'string') { clearError(r,c); return raw ?? ''; }
        if(isBlankFormula(raw)) { clearError(r,c); return raw; }
        if(raw.startsWith('=')){
          try{
            const v = evaluateFormula(raw.slice(1), visited);
            if(isErr(v)) setError(r,c, v.error);
            else clearError(r,c);
            return v;
          }catch(e){
            setError(r,c, String(e.message||e));
            return VALUE_ERROR;
          }
        }
        clearError(r,c);
        return raw;
      } finally {
        visited.delete(key);
        if(top) visited.clear();
      }
    }
    function displayValue(r,c){
      const raw = rawValue(r,c);
      if(isBlankFormula(raw)) return raw; // show '=' while user is typing
      if(typeof raw === 'string' && raw.startsWith('=')) {
        const v = valueAt(r,c);
        if(isErr(v)) return v.error;
        return String(v);
      }
      return raw;
    }

    function applyErrorDecorations(){
      for(const el of tbody.querySelectorAll('.cell')){
        const r = +el.dataset.r, c = +el.dataset.c;
        const key = `${r},${c}`;
        if(errMap.has(key)){
          el.classList.add('err');
          el.title = `#VALUE!: ${errMap.get(key)}`;
        } else {
          el.classList.remove('err');
          el.removeAttribute('title');
        }
      }
    }

      // Editing
      function onCellFocus(e){
        const el = e.currentTarget;
        setActiveCell(+el.dataset.r, +el.dataset.c);
      }
      function onEdit(e){
        const el = e.currentTarget;
        const r = +el.dataset.r, c = +el.dataset.c;
        pushUndo();
        data[r][c].value = el.textContent;
        if(document.activeElement === el){
          formulaBar.value = el.textContent;
        }
        recalc();
      }
      function onBlurNormalize(e){
        const el = e.currentTarget;
        // keep "=   " while user is composing a formula
        if (/^=\s*$/.test(el.textContent)) return;
        el.textContent = el.textContent.replace(/\s+/g,' ').trim();
      }

      formulaBar.addEventListener('input', () => {
        const r = +formulaBar.dataset.r, c = +formulaBar.dataset.c;
        if (isNaN(r) || isNaN(c)) return;
        pushUndo();
        data[r][c].value = formulaBar.value;
        const cell = tbody.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
        if (cell && document.activeElement !== cell) cell.textContent = formulaBar.value;
        recalc();
      });

      boldBtn?.addEventListener('click', () => {
        const {r,c} = activeCell;
        const cell = data[r][c];
        cell.bold = !cell.bold;
        const el = tbody.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
        if(el) applyCellStyles(el, cell);
      });
      italicBtn?.addEventListener('click', () => {
        const {r,c} = activeCell;
        const cell = data[r][c];
        cell.italic = !cell.italic;
        const el = tbody.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
        if(el) applyCellStyles(el, cell);
      });
      fillColorInput?.addEventListener('input', () => {
        const {r,c} = activeCell;
        const cell = data[r][c];
        cell.bgColor = fillColorInput.value;
        const el = tbody.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
        if(el) applyCellStyles(el, cell);
      });

      undoBtn?.addEventListener('click', undo);
      redoBtn?.addEventListener('click', redo);
      document.addEventListener('keydown', (e) => {
        const key = e.key.toLowerCase();
        if ((e.ctrlKey || e.metaKey) && !e.shiftKey && key === 'z') {
          e.preventDefault();
          undo();
        } else if ((e.ctrlKey || e.metaKey) && key === 'y') {
          e.preventDefault();
          redo();
        }
      });

    // Throttled recalc (avoid flood while typing quickly)
    let recalcTimer = 0;
    function recalc(){
      calcState.textContent = 'Calculating…';
      clearTimeout(recalcTimer);
      recalcTimer = setTimeout(()=>{
        try{
          refreshAllDisplay();
          calcState.textContent = 'Ready';
        }catch(err){
          calcState.textContent = 'Error';
          log('Recalc error:', err.message||err);
        }
      }, 16); // ~60fps
    }

    // Row/Col ops
    function modifyGrid(type){
      pushUndo();
      switch(type){
        case 'addRow':
          data.push(Array.from({length:cols},()=>createCell()));
          rows++;
          rowHeights.push(null);
          break;
        case 'addCol':
          for(const r of data) r.push(createCell());
          cols++;
          colWidths.push(null);
          break;
        case 'delRow':
          if(rows>1){
            data.pop();
            rows--;
            rowHeights.pop();
          }
          break;
        case 'delCol':
          if(cols>1){
            for(const r of data) r.pop();
            cols--;
            colWidths.pop();
          }
          break;
      }
      renderHeader();
      renderBody();
      recalc();
    }

    document.getElementById('addRow').addEventListener('click', () => modifyGrid('addRow'));
    document.getElementById('addCol').addEventListener('click', () => modifyGrid('addCol'));
    document.getElementById('delRow').addEventListener('click', () => modifyGrid('delRow'));
    document.getElementById('delCol').addEventListener('click', () => modifyGrid('delCol'));
    document.getElementById('newSheet').onclick = ()=>{
      sheets.length=0;
      sheets.push(createSheet('Sheet1'));
      loadSheet(0);
      renderTabs();
      fileInfo.textContent='';
    };

    // CSV helpers
    function toCSV(d=data){
      return d.map(row=>
        row.map(cell=>{
          const s = String(cell.value ?? '');
          if(/[,"\n]/.test(s)) return `"${s.replace(/"/g,'""')}"`;
          return s;
        }).join(',')
      ).join('\n');
    }
    function parseCSV(text){
      const out = []; let row=[]; let i=0; let cur=''; let inQ=false;
      const pushCell =()=>{ row.push(cur); cur=''; };
      const pushRow =()=>{ out.push(row); row=[]; };
      while(i<text.length){
        const ch = text[i++];
        if(inQ){
          if(ch === '"'){
            if(text[i]==='"'){ cur+='"'; i++; } else { inQ=false; }
          }else cur += ch;
        }else{
          if(ch === '"'){ inQ=true; }
          else if(ch === ','){ pushCell(); }
          else if(ch === '\n'){ pushCell(); pushRow(); }
          else if(ch === '\r'){ /* ignore */ }
          else cur += ch;
        }
      }
      pushCell(); pushRow();
      const maxC = Math.max(...out.map(r=>r.length));
      return out.map(r=>r.concat(Array(Math.max(0,maxC-r.length)).fill('')));
    }

    function loadArrayInto(sheet, arr){
      sheet.rows = arr.length;
      sheet.cols = Math.max(...arr.map(r=>r.length));
      sheet.data = createEmpty(sheet.rows, sheet.cols);
      sheet.colWidths = Array(sheet.cols).fill(null);
      sheet.rowHeights = Array(sheet.rows).fill(null);
      for(let r=0;r<sheet.rows;r++) for(let c=0;c<sheet.cols;c++) sheet.data[r][c].value = arr[r][c]??'';
    }

    // File open
    document.getElementById('fileInput').addEventListener('change', async (ev)=>{
      const file = ev.target.files[0]; if(!file) return; const name = file.name.toLowerCase();
      try{
        if(name.endsWith('.csv')){
          let txt = await file.text();
          if (txt.charCodeAt(0) === 0xFEFF) txt = txt.slice(1);
          const arr = parseCSV(txt);
          sheets.length=0;
          const sh = createSheet(file.name.replace(/\.csv$/i,''));
          loadArrayInto(sh, arr);
          sheets.push(sh);
          loadSheet(0);
          renderTabs();
          fileInfo.textContent = `Loaded CSV (${file.name})`;
        }else{
          const lib = await ensureXLSX();
          if(!lib){ fileInfo.textContent = 'XLSX library unavailable — please open CSV instead'; return; }
          const buf = await file.arrayBuffer();
          const wb = lib.read(buf, {type:'array'});
          sheets.length=0;
          wb.SheetNames.forEach(sn=>{
            const ws = wb.Sheets[sn];
            const arr = lib.utils.sheet_to_json(ws,{header:1, blankrows:true, defval:''});
            const sh = createSheet(sn);
            loadArrayInto(sh, arr);
            sheets.push(sh);
          });
          loadSheet(0);
          renderTabs();
          fileInfo.textContent = `Loaded XLSX (${file.name})`;
        }
      }catch(err){
        fileInfo.textContent = 'Failed to open file';
        calcState.textContent = 'Error';
        log('Open error:', err.message||err);
      } finally {
        ev.target.value='';
      }
    });

    // Save CSV
    document.getElementById('saveCSV').onclick = async ()=>{
      saveActiveState();
      if(sheets.length===1){
        const blob = new Blob([toCSV(sheets[0].data)], {type:'text/csv;charset=utf-8'});
        const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${sheets[0].name}.csv`; a.click(); URL.revokeObjectURL(a.href);
      }else{
        const zipLib = await ensureJSZip();
        if(!zipLib){ fileInfo.textContent = 'ZIP export unavailable'; return; }
        const zip = new zipLib();
        sheets.forEach(sh=> zip.file(`${sh.name}.csv`, toCSV(sh.data)) );
        const blob = await zip.generateAsync({type:'blob'});
        const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'sheets.zip'; a.click(); URL.revokeObjectURL(a.href);
      }
    };

    // Save XLSX (lazy-load lib to prevent ReferenceError)
    saveXLSXBtn.onclick = async ()=>{
      saveActiveState();
      const lib = XLSXRef || await ensureXLSX();
      if(!lib){ fileInfo.textContent = 'XLSX export unavailable — library not loaded'; return; }
      const wb = lib.utils.book_new();
      sheets.forEach(sh=>{
        const aoa = sh.data.map(row=>row.map(cell=>cell.value));
        const ws = lib.utils.aoa_to_sheet(aoa);
        for(let r=0;r<sh.rows;r++){
          for(let c=0;c<sh.cols;c++){
            const cell = sh.data[r][c];
            if(!cell.bold && !cell.italic && !cell.bgColor) continue;
            const addr = lib.utils.encode_cell({r,c});
            ws[addr] = ws[addr] || { t:'s', v: cell.value || '' };
            ws[addr].s = ws[addr].s || {};
            if(cell.bold || cell.italic){
              ws[addr].s.font = ws[addr].s.font || {};
              if(cell.bold) ws[addr].s.font.bold = true;
              if(cell.italic) ws[addr].s.font.italic = true;
            }
            if(cell.bgColor){
              ws[addr].s.fill = { patternType:'solid', fgColor:{ rgb:'FF'+cell.bgColor.slice(1).toUpperCase() } };
            }
          }
        }
        lib.utils.book_append_sheet(wb, ws, sh.name);
      });
      lib.writeFile(wb, 'sheets.xlsx');
    };

    // ===== Keyboard navigation (Enter/Shift+Enter, Tab/Shift+Tab, arrows) =====
      function focusCell(r,c){
        r = Math.max(0, Math.min(rows-1, r));
        c = Math.max(0, Math.min(cols-1, c));
        const el = tbody.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
        if (el){ el.focus(); placeCaretEnd(el); }
        setActiveCell(r,c);
      }
    function placeCaretEnd(el){
      const range = document.createRange();
      const node = el.firstChild || el;
      const len = (node.textContent||'').length;
      range.setStart(node, len); range.setEnd(node, len);
      const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
    }
    // Virtual cursor for formula navigation
    let formulaVirtualCursor = { r: 0, c: 0, active: false };

    tbody.addEventListener('keydown', (e)=>{
      const el = e.target.closest('.cell'); if (!el) return;
      const r = +el.dataset.r, c = +el.dataset.c;
      
      // Check if we're editing a formula (starts with =)
      const isEditingFormula = el.textContent.startsWith('=');
      
      // Reset virtual cursor when starting formula edit
      if (isEditingFormula && !formulaVirtualCursor.active) {
        formulaVirtualCursor = { r, c, active: true };
      } else if (!isEditingFormula) {
        formulaVirtualCursor.active = false;
      }
      
      const go = (nr, nc)=>{ 
        e.preventDefault(); 
        // Update current cell display before moving
        if (e.key === 'Enter') {
          el.textContent = displayValue(r, c);
          formulaVirtualCursor.active = false; // Reset on Enter
        }
        focusCell(nr, nc); 
      };
      
      // Excel-like formula navigation: insert cell references when editing formulas
      const insertCellRef = (deltaR, deltaC) => {
        // Compute target location before clamping
        const targetR = formulaVirtualCursor.r + deltaR;
        const targetC = formulaVirtualCursor.c + deltaC;

        const newR = Math.max(0, Math.min(rows - 1, targetR));
        const newC = Math.max(0, Math.min(cols - 1, targetC));

        // If clamping results in no movement, handle edges specially
        if (newR === formulaVirtualCursor.r && newC === formulaVirtualCursor.c) {
          if (deltaC !== 0) {
            const selection = window.getSelection();
            if (deltaC < 0 && selection.rangeCount > 0) {
              const range = selection.getRangeAt(0);
              // Prevent default to avoid leaving the cell when at the formula start
              if (range.startOffset <= 1) {
                e.preventDefault();
                return;
              }
            }
            // Allow caret navigation horizontally by deactivating the virtual cursor
            formulaVirtualCursor.active = false;
          } else {
            // Vertical edges: keep focus in cell to avoid row header selection
            e.preventDefault();
          }
          return;
        }

        e.preventDefault();

        // Move virtual cursor
        formulaVirtualCursor.r = newR;
        formulaVirtualCursor.c = newC;

        const ref = colLabel(formulaVirtualCursor.c) + (formulaVirtualCursor.r + 1);
        const currentText = el.textContent;
        const selection = window.getSelection();

        if (selection.rangeCount > 0) {
          const range = selection.getRangeAt(0);
          // Prevent replacing the leading '=' when caret is before it
          const start = Math.max(1, range.startOffset);
          const end = Math.max(1, range.endOffset);
          const newText = currentText.slice(0, start) + ref + currentText.slice(end);
          el.textContent = newText;

          // Update data and formula bar
          data[r][c].value = newText;
          formulaBar.value = newText;

          // Select the inserted reference so subsequent arrow presses replace it
          const selStart = start;
          const selEnd = selStart + ref.length;
          const textNode = el.firstChild || el;
          const newRange = document.createRange();
          newRange.setStart(textNode, selStart);
          newRange.setEnd(textNode, Math.min(selEnd, textNode.textContent?.length || 0));
          selection.removeAllRanges();
          selection.addRange(newRange);

          recalc();
        }
      };
      
      if (e.key === 'Enter') return go(r + (e.shiftKey?-1:1), c);
      if (e.key === 'Tab')   return go(r, c + (e.shiftKey?-1:1));
      
      // Arrow key behavior: insert cell references when editing formulas
      if (isEditingFormula && !e.shiftKey && !e.ctrlKey && !e.metaKey) {
        if (e.key === 'ArrowDown') return insertCellRef(1, 0);
        if (e.key === 'ArrowUp') return insertCellRef(-1, 0);
        if (e.key === 'ArrowLeft') return insertCellRef(0, -1);
        if (e.key === 'ArrowRight') return insertCellRef(0, 1);
      }
      
      // Normal navigation when not editing formulas or with modifiers
      if (e.key === 'ArrowDown' && !e.shiftKey) return go(r+1, c);
      if (e.key === 'ArrowUp'   && !e.shiftKey) return go(r-1, c);
      if (e.key === 'ArrowLeft' && (e.ctrlKey||e.metaKey)) return go(r, c-1);
      if (e.key === 'ArrowRight'&& (e.ctrlKey||e.metaKey)) return go(r, c+1);
    });

    tbody.addEventListener('copy', (e)=>{
      const el = document.activeElement?.closest('.cell');
      copyOrigin = el ? {r:+el.dataset.r, c:+el.dataset.c} : null;
    });
    tbody.addEventListener('cut', (e)=>{
      const el = document.activeElement?.closest('.cell');
      copyOrigin = el ? {r:+el.dataset.r, c:+el.dataset.c} : null;
    });
    // ===== Paste from Excel/Sheets (tab/newline grid paste) =====
    tbody.addEventListener('paste', (e)=>{
      const el = e.target.closest('.cell'); if (!el) return;
      e.preventDefault();
      const text = (e.clipboardData || window.clipboardData).getData('text');
      if (!text) return;
      pushUndo();
      const rowsClip = text.replace(/\r/g,'').split('\n').map(r=>r.split('\t'));
      const r0 = +el.dataset.r, c0 = +el.dataset.c;
      const dr = copyOrigin ? r0 - copyOrigin.r : 0;
      const dc = copyOrigin ? c0 - copyOrigin.c : 0;
      for (let i=0;i<rowsClip.length;i++){
        for (let j=0;j<rowsClip[i].length;j++){
          const rr = r0+i, cc = c0+j;
          if (rr<rows && cc<cols){
            let cellText = rowsClip[i][j];
            if(copyOrigin && typeof cellText==='string' && cellText.startsWith('=')){
              cellText = '=' + shiftFormulaRefs(cellText.slice(1), dr, dc);
            }
            data[rr][cc].value = cellText;
          }
        }
      }
      renderBody(); recalc();
      copyOrigin = null;
    });

    // ===== Auto-fit columns (lightweight) =====
    function autofitColumns(maxWidth=360){
      const ctx = document.createElement('canvas').getContext('2d');
      // Approximate table cell font (inherits from body)
      const bodyStyle = getComputedStyle(document.body);
      ctx.font = `${bodyStyle.fontSize} ${bodyStyle.fontFamily}`;
      for(let c=0;c<cols;c++){
        if(colWidths[c]!=null){
          applyColWidth(c, colWidths[c]);
          continue;
        }
        let w = ctx.measureText(colLabel(c)).width + 24;
        for(let r=0;r<Math.min(rows, 50); r++){ // sample first 50 rows for speed
          const txt = String(displayValue(r,c));
          w = Math.max(w, ctx.measureText(txt).width + 24);
          if (w >= maxWidth) break;
        }
        const finalW = Math.min(Math.max(80, Math.ceil(w)), maxWidth); // clamp 80..maxWidth
        applyColWidth(c, finalW);
      }
    }

    // ===== Runtime self-tests =====
    document.getElementById('runTests').onclick = async ()=>{
      const results = [];
      // Existing Test 1: formula evaluation
      const bak = data.map(r=>r.map(cell=>({...cell}))); const bakRows=rows, bakCols=cols;
      rows=2; cols=3; data=createEmpty(rows,cols);
      data[0][0].value='1'; data[0][1].value='2'; data[0][2].value='=A1+B1*3';
      const val = valueAt(0,2);
      results.push(val===7 ? '✓ Formula (=A1+B1*3) == 7' : `✗ Formula expected 7 got ${val}`);

      // Existing Test 2: CSV roundtrip
      const csv = toCSV();
      const arr = parseCSV(csv);
      const ok2 = arr[1][3]===undefined && arr[0][0]==='1' && arr[0][1]==='2' && arr[0][2]==='=A1+B1*3';
      results.push(ok2 ? '✓ CSV roundtrip' : '✗ CSV roundtrip failed');

      // Restore sheet
      rows=bakRows; cols=bakCols; data=bak; renderHeader(); renderBody();

      // Existing Test 3: XLSX availability (no failure if offline)
      results.push((XLSXRef||window.XLSX) ? '✓ XLSX present' : '• XLSX not loaded (CSV still OK)');

      // Added Test 4: SUM(range & multi-range) evaluation
      rows=3; cols=3; data=createEmpty(rows,cols);
      data[0][0].value='1'; data[0][1].value='2'; data[1][0].value='3'; data[1][1].value='4';
      data[0][2].value='5'; data[1][2].value='6';
      data[2][2].value='=SUM(A1:B2)';
      const sumVal = valueAt(2,2);
      results.push(sumVal===10 ? '✓ SUM(A1:B2) == 10' : `✗ SUM(A1:B2) expected 10 got ${sumVal}`);
      data[2][2].value='=SUM(A1:B1,B2:C2)';
      const sumVal2 = valueAt(2,2);
      results.push(sumVal2===13 ? '✓ SUM(A1:B1,B2:C2) == 13' : `✗ SUM multi-range expected 13 got ${sumVal2}`);

      // Added Test 5: A1 parse + colLabel check
      const p = parseA1('AA10');
      const lbl = colLabel(27); // 0-based 27 -> AB
      const ok5 = p && p.r===9 && p.c===26 && lbl==='AB';
      results.push(ok5 ? '✓ A1 parse & colLabel' : '✗ A1 parse/label failed');

      // Added Test 6: Unsafe formula handled
      data[0][0].value='=A1+BADFUNC(1)';
      const unsafe = valueAt(0,0);
      results.push(isErr(unsafe) ? '✓ Unsafe formula -> #VALUE!' : `✗ Unsafe formula not blocked (${unsafe})`);

      // Added Test 7: Blank '=' preserved
      data[0][0].value='='; const blank = valueAt(0,0);
      results.push(blank==='=' ? '✓ Blank formula (=) kept' : `✗ '=' should be kept, got ${blank}`);

      // Added Test 8: Whitespace-only after '=' preserved
      data[0][0].value='=   '; const blank2 = valueAt(0,0);
      results.push(/^=\s*$/.test(blank2) ? '✓ Whitespace-only formula kept' : `✗ '=   ' should be kept, got ${blank2}`);

      // Added Test 9: Nested functions & MIN/MAX/AVERAGE
      rows=2; cols=2; data=createEmpty(rows,cols);
      data[0][0].value='5'; data[0][1].value='15';
      data[1][0].value='=SUM(MIN(A1:B1), MAX(A1:B1), AVERAGE(A1:B1))';
      const nested = valueAt(1,0);
      results.push(nested===30 ? '✓ Nested MIN/MAX/AVERAGE' : `✗ Nested functions expected 30 got ${nested}`);

      // Added Test 10: Parentheses precedence
      data[1][1].value='=(1+2)*3';
      const prec = valueAt(1,1);
      results.push(prec===9 ? '✓ (1+2)*3 == 9' : `✗ (1+2)*3 expected 9 got ${prec}`);

      // Added Test 11: Error propagation for non-numeric
      rows=1; cols=3; data=createEmpty(rows,cols);
      data[0][0].value='a';
      data[0][1].value='=1+A1';
      const err1 = valueAt(0,1);
      results.push(err1 && err1.error==='#VALUE!' ? '✓ 1+A1 with A1="a" -> #VALUE!' : `✗ 1+A1 expected #VALUE! got ${String(err1)}`);
      data[0][2].value='=SUM(1,A1)';
      const err2 = valueAt(0,2);
      results.push(err2 && err2.error==='#VALUE!' ? '✓ SUM(1,A1) -> #VALUE!' : `✗ SUM(1,A1) expected #VALUE! got ${String(err2)}`);

      // Added Test 12: Blank cell reference -> #VALUE!
      data=createEmpty(rows,cols);
      data[0][1].value='=1+A1';
      const errBlank = valueAt(0,1);
      results.push(errBlank && errBlank.error==='#VALUE!' ? '✓ 1+A1 with A1 blank -> #VALUE!' : `✗ 1+A1 blank expected #VALUE! got ${String(errBlank)}`);

      // Added Test 13: Absolute reference shifting
      rows=3; cols=3; data=createEmpty(rows,cols);
      data[0][0].value='1'; data[1][0].value='2'; data[0][1].value='3'; data[1][1].value='4';
      const baseFormula='=$A$1+$A1+A$1+A1';
      const shifted='='+shiftFormulaRefs(baseFormula.slice(1),1,1);
      const okAbs = shifted==='=$A$1+$A2+B$1+B2';
      data[2][2].value=shifted;
      const absVal = valueAt(2,2);
      results.push(okAbs && absVal===10 ? '✓ $A$1/$A1/A$1 refs' : `✗ $ refs failed (${shifted} -> ${absVal})`);

      // Restore again
      rows=bakRows; cols=bakCols; data=bak; renderHeader(); renderBody();

      fileInfo.textContent = results.join(' | ');
      log('Test results:', results.join(' | '));
    };

    // ===== Dropdown menus (Export, Advanced) =====
    (function setupMenus(){
      const menus = Array.from(document.querySelectorAll('.menu'));
      const closeAll = ()=>{
        menus.forEach(m=>{
          m.classList.remove('open');
          const t = m.querySelector('.menu-trigger');
          if(t) t.setAttribute('aria-expanded','false');
        });
      };
      menus.forEach(m=>{
        const trigger = m.querySelector('.menu-trigger');
        const items = m.querySelector('.menu-items');
        if(!trigger || !items) return;
        trigger.addEventListener('click', (e)=>{
          e.preventDefault();
          const wasOpen = m.classList.contains('open');
          closeAll();
          m.classList.toggle('open', !wasOpen);
          trigger.setAttribute('aria-expanded', (!wasOpen).toString());
        });
        items.addEventListener('click', (e)=>{
          if(e.target.closest('button')) closeAll();
        });
      });
      document.addEventListener('click', (e)=>{
        if(!e.target.closest('.menu')) closeAll();
      });
      document.addEventListener('keydown', (e)=>{
        if(e.key==='Escape') closeAll();
      });
    })();

    // ===== Hamburger menu toggle =====
    const hamburgerBtn = document.getElementById('hamburgerBtn');
    const headerNav = document.getElementById('headerNav');
    
    if (hamburgerBtn && headerNav) {
      hamburgerBtn.addEventListener('click', (e) => {
        e.preventDefault();
        const isOpen = headerNav.classList.contains('open');
        headerNav.classList.toggle('open', !isOpen);
        hamburgerBtn.setAttribute('aria-expanded', (!isOpen).toString());
        hamburgerBtn.classList.toggle('active', !isOpen);
      });

      // Close menu when clicking outside
      document.addEventListener('click', (e) => {
        if (!e.target.closest('header')) {
          headerNav.classList.remove('open');
          hamburgerBtn.setAttribute('aria-expanded', 'false');
          hamburgerBtn.classList.remove('active');
        }
      });

      // Close menu on Escape key
      document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && headerNav.classList.contains('open')) {
          headerNav.classList.remove('open');
          hamburgerBtn.setAttribute('aria-expanded', 'false');
          hamburgerBtn.classList.remove('active');
        }
      });
    }

    // ===== Init =====
    renderHeader(); renderBody(); recalc(); renderTabs();
});

