document.addEventListener('DOMContentLoaded', () => {
    const gridEl = document.getElementById('grid');
    const thead = gridEl.querySelector('thead');
    const tbody = gridEl.querySelector('tbody');
    const calcState = document.getElementById('calcState');
    const fileInfo = document.getElementById('fileInfo');
    const xlsxState = document.getElementById('xlsxState');
    const pickerWrap = document.getElementById('sheetPicker');
    const sheetSelect = document.getElementById('sheetSelect');
    const saveXLSXBtn = document.getElementById('saveXLSX');
    const debugBox = document.getElementById('debugBox');
    const debugOut = document.getElementById('debugOut');

    // In-memory sheet data model
    let rows = 30, cols = 12;
    let data = createEmpty(rows, cols); // stores raw strings (including formulas)
    let copyOrigin = null; // track source cell for copy/paste

    // Error map: key "r,c" -> message
    const errMap = new Map();

    // XLSX lib reference (avoid global window access)
    let XLSXRef = null; // set after ensureXLSX()

    // For XLSX multi-sheet handling
    let currentWB = null; // Workbook object

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
      await navigator.clipboard.writeText(report);
      fileInfo.textContent = 'Debug report copied to clipboard';
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

    // Utilities
    function createEmpty(r,c){ return Array.from({length:r},()=>Array.from({length:c},()=>'')); }
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

    // Rendering
    function renderHeader(){
      const tr = document.createElement('tr');
      tr.appendChild(document.createElement('th')); // corner
      for(let c=0;c<cols;c++){
        const th = document.createElement('th');
        th.textContent = colLabel(c);
        tr.appendChild(th);
      }
      thead.innerHTML='';
      thead.appendChild(tr);
    }
    function renderBody(){
      tbody.innerHTML='';
      for(let r=0;r<rows;r++){
        const tr = document.createElement('tr');
        const rowTh = document.createElement('th');
        rowTh.textContent = r+1;
        tr.appendChild(rowTh);
        for(let c=0;c<cols;c++){
          const td = document.createElement('td');
          const div = document.createElement('div');
          div.className='cell';
          div.contentEditable = true;
          div.dataset.r = r;
          div.dataset.c = c;
          div.textContent = displayValue(r,c);
          div.addEventListener('input', onEdit);
          div.addEventListener('blur', onBlurNormalize);
          td.appendChild(div);
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
      applyErrorDecorations();
      // initial autofit for the newly rendered table
      autofitColumns();
    }

    // Caret helpers for flicker-free refresh
    function getCaret() {
      const el = document.activeElement;
      if (!el || !el.classList || !el.classList.contains('cell')) return null;
      const sel = window.getSelection();
      if (!sel || sel.rangeCount === 0) return null;
      const range = sel.getRangeAt(0);
      return { el, r: +el.dataset.r, c: +el.dataset.c, start: range.startOffset, end: range.endOffset };
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
        if (el === active) continue; // keep user input while editing
        const r = +el.dataset.r, c = +el.dataset.c;
        const newText = displayValue(r,c);
        if (el.textContent !== newText) el.textContent = newText;
      }
      applyErrorDecorations();
      setCaret(snap);
    }

    // Formula evaluation (simple & safe-ish)
    function rawValue(r,c){ return data[r]?.[c] ?? ''; }
    function isErr(v){ return v && typeof v==='object' && v.error; }
    function numeric(v){
      if(isErr(v)) return v;
      const n = Number(v);
      return isFinite(n) ? n : {error:'#VALUE!'};
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

    function evaluateFormula(expr){
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
        if(t.type==='cell'){ consume('cell'); return numeric(valueAt(t.pos.r,t.pos.c)); }
        if(t.type==='range'){
          consume('range'); const out=[];
          const r1=Math.min(t.start.r,t.end.r), r2=Math.max(t.start.r,t.end.r);
          const c1=Math.min(t.start.c,t.end.c), c2=Math.max(t.start.c,t.end.c);
          for(let r=r1;r<=r2;r++) for(let c=c1;c<=c2;c++) out.push(numeric(valueAt(r,c)));
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

    function valueAt(r,c){
      const raw = rawValue(r,c);
      if(typeof raw !== 'string') { clearError(r,c); return raw ?? ''; }
      if(isBlankFormula(raw)) { clearError(r,c); return raw; }
      if(raw.startsWith('=')){
        try{
          const v = evaluateFormula(raw.slice(1));
          clearError(r,c);
          return v;
        }catch(e){
          setError(r,c, String(e.message||e));
          return 'ERR';
        }
      }
      clearError(r,c);
      return raw;
    }
    function displayValue(r,c){
      const raw = rawValue(r,c);
      if(isBlankFormula(raw)) return raw; // show '=' while user is typing
      if(typeof raw === 'string' && raw.startsWith('=')) {
        const v = valueAt(r,c);
        if(v === 'ERR') return 'ERR';
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
          el.title = `ERR: ${errMap.get(key)}`;
        } else {
          el.classList.remove('err');
          el.removeAttribute('title');
        }
      }
    }

    // Editing
    function onEdit(e){
      const el = e.currentTarget;
      const r = +el.dataset.r, c = +el.dataset.c;
      data[r][c] = el.textContent;
      recalc();
    }
    function onBlurNormalize(e){
      const el = e.currentTarget;
      // keep "=   " while user is composing a formula
      if (/^=\s*$/.test(el.textContent)) return;
      el.textContent = el.textContent.replace(/\s+/g,' ').trim();
    }

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
      switch(type){
        case 'addRow':
          data.push(Array.from({length:cols},()=>''));
          rows++;
          break;
        case 'addCol':
          for(const r of data) r.push('');
          cols++;
          break;
        case 'delRow':
          if(rows>1){
            data.pop();
            rows--;
          }
          break;
        case 'delCol':
          if(cols>1){
            for(const r of data) r.pop();
            cols--;
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
      rows = 30; cols = 12; data = createEmpty(rows, cols);
      currentWB = null; pickerWrap.classList.remove('active'); sheetSelect.innerHTML=''; fileInfo.textContent = '';
      renderHeader(); renderBody(); recalc();
    };

    // CSV helpers
    function toCSV(){
      return data.map(row=>
        row.map(cell=>{
          const s = String(cell ?? '');
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

    function loadArray(arr){
      rows = arr.length;
      cols = Math.max(...arr.map(r=>r.length));
      data = createEmpty(rows, cols);
      for(let r=0;r<rows;r++) for(let c=0;c<cols;c++) data[r][c]=arr[r][c]??'';
      renderHeader(); renderBody(); recalc();
    }

    // File open
    document.getElementById('fileInput').addEventListener('change', async (ev)=>{
      const file = ev.target.files[0]; if(!file) return; const name = file.name.toLowerCase();
      try{
        if(name.endsWith('.csv')){
          let txt = await file.text();
          // Strip UTF-8 BOM if present
          if (txt.charCodeAt(0) === 0xFEFF) txt = txt.slice(1);
          loadArray(parseCSV(txt));
          fileInfo.textContent = `Loaded CSV (${file.name})`;
          currentWB = null; pickerWrap.classList.remove('active'); sheetSelect.innerHTML='';
        }else{
          const lib = await ensureXLSX();
          if(!lib){ fileInfo.textContent = 'XLSX library unavailable — please open CSV instead'; return; }
          const buf = await file.arrayBuffer();
          const wb = lib.read(buf, {type:'array'});
          currentWB = wb;
          sheetSelect.innerHTML = '';
          wb.SheetNames.forEach((sn, i)=>{
            const opt = document.createElement('option');
            opt.value = sn; opt.textContent = sn;
            if(i===0) opt.selected = true;
            sheetSelect.appendChild(opt);
          });
          pickerWrap.classList.toggle('active', wb.SheetNames.length > 1);
          loadSheetByName(wb.SheetNames[0], lib);
          fileInfo.textContent = `Loaded XLSX (${file.name}) — ${wb.SheetNames.length>1? 'Select sheet':''}`;
        }
      }catch(err){
        fileInfo.textContent = 'Failed to open file';
        calcState.textContent = 'Error';
        log('Open error:', err.message||err);
      } finally {
        ev.target.value='';
      }
    });

    function loadSheetByName(name, lib){
      const L = lib || XLSXRef || window.XLSX; if(!currentWB || !L) return;
      const ws = currentWB.Sheets[name];
      const arr = L.utils.sheet_to_json(ws, {header:1, blankrows:true, defval:''});
      loadArray(arr); // render + autofit inside
    }
    sheetSelect.addEventListener('change', async ()=>{
      const lib = XLSXRef || await ensureXLSX();
      loadSheetByName(sheetSelect.value, lib);
    });

    // Save CSV
    document.getElementById('saveCSV').onclick = ()=>{
      const blob = new Blob([toCSV()], {type:'text/csv;charset=utf-8'});
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'sheet.csv'; a.click(); URL.revokeObjectURL(a.href);
    };

    // Save XLSX (lazy-load lib to prevent ReferenceError)
    saveXLSXBtn.onclick = async ()=>{
      const lib = XLSXRef || await ensureXLSX();
      if(!lib){ fileInfo.textContent = 'XLSX export unavailable — library not loaded'; return; }
      const aoa = data.map(row=>row.slice());
      const ws = lib.utils.aoa_to_sheet(aoa);
      const wb = lib.utils.book_new();
      lib.utils.book_append_sheet(wb, ws, 'Sheet1');
      lib.writeFile(wb, 'sheet.xlsx');
    };

    // ===== Keyboard navigation (Enter/Shift+Enter, Tab/Shift+Tab, arrows) =====
    function focusCell(r,c){
      r = Math.max(0, Math.min(rows-1, r));
      c = Math.max(0, Math.min(cols-1, c));
      const el = tbody.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
      if (el){ el.focus(); placeCaretEnd(el); }
    }
    function placeCaretEnd(el){
      const range = document.createRange();
      const node = el.firstChild || el;
      const len = (node.textContent||'').length;
      range.setStart(node, len); range.setEnd(node, len);
      const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
    }
    tbody.addEventListener('keydown', (e)=>{
      const el = e.target.closest('.cell'); if (!el) return;
      const r = +el.dataset.r, c = +el.dataset.c;
      const go = (nr, nc)=>{ e.preventDefault(); focusCell(nr, nc); };
      if (e.key === 'Enter') return go(r + (e.shiftKey?-1:1), c);
      if (e.key === 'Tab')   return go(r, c + (e.shiftKey?-1:1));
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
            data[rr][cc] = cellText;
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
      const headerCells = thead.querySelectorAll('th');
      for(let c=0;c<cols;c++){
        let w = ctx.measureText(colLabel(c)).width + 24;
        for(let r=0;r<Math.min(rows, 50); r++){ // sample first 50 rows for speed
          const txt = String(displayValue(r,c));
          w = Math.max(w, ctx.measureText(txt).width + 24);
          if (w >= maxWidth) break;
        }
        const finalW = Math.min(Math.max(80, Math.ceil(w)), maxWidth); // clamp 80..maxWidth
        // nth-child: +2 because first column is row header <th>
        const tdList = tbody.querySelectorAll(`td:nth-child(${c+2})`);
        tdList.forEach(td => { td.style.width = finalW + 'px'; td.style.maxWidth = finalW + 'px'; });
        if (headerCells[c+1]){
          headerCells[c+1].style.width = finalW + 'px';
          headerCells[c+1].style.maxWidth = finalW + 'px';
        }
      }
    }

    // ===== Runtime self-tests =====
    document.getElementById('runTests').onclick = async ()=>{
      const results = [];
      // Existing Test 1: formula evaluation
      const bak = data.map(r=>r.slice()); const bakRows=rows, bakCols=cols;
      rows=2; cols=3; data=createEmpty(rows,cols);
      data[0][0]='1'; data[0][1]='2'; data[0][2]='=A1+B1*3';
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
      data[0][0]='1'; data[0][1]='2'; data[1][0]='3'; data[1][1]='4';
      data[0][2]='5'; data[1][2]='6';
      data[2][2]='=SUM(A1:B2)';
      const sumVal = valueAt(2,2);
      results.push(sumVal===10 ? '✓ SUM(A1:B2) == 10' : `✗ SUM(A1:B2) expected 10 got ${sumVal}`);
      data[2][2]='=SUM(A1:B1,B2:C2)';
      const sumVal2 = valueAt(2,2);
      results.push(sumVal2===13 ? '✓ SUM(A1:B1,B2:C2) == 13' : `✗ SUM multi-range expected 13 got ${sumVal2}`);

      // Added Test 5: A1 parse + colLabel check
      const p = parseA1('AA10');
      const lbl = colLabel(27); // 0-based 27 -> AB
      const ok5 = p && p.r===9 && p.c===26 && lbl==='AB';
      results.push(ok5 ? '✓ A1 parse & colLabel' : '✗ A1 parse/label failed');

      // Added Test 6: Unsafe formula handled
      data[0][0]='=A1+BADFUNC(1)';
      const unsafe = valueAt(0,0);
      results.push(unsafe==='ERR' ? '✓ Unsafe formula -> ERR' : `✗ Unsafe formula not blocked (${unsafe})`);

      // Added Test 7: Blank '=' preserved
      data[0][0]='='; const blank = valueAt(0,0);
      results.push(blank==='=' ? '✓ Blank formula (=) kept' : `✗ '=' should be kept, got ${blank}`);

      // Added Test 8: Whitespace-only after '=' preserved
      data[0][0]='=   '; const blank2 = valueAt(0,0);
      results.push(/^=\s*$/.test(blank2) ? '✓ Whitespace-only formula kept' : `✗ '=   ' should be kept, got ${blank2}`);

      // Added Test 9: Nested functions & MIN/MAX/AVERAGE
      rows=2; cols=2; data=createEmpty(rows,cols);
      data[0][0]='5'; data[0][1]='15';
      data[1][0]='=SUM(MIN(A1:B1), MAX(A1:B1), AVERAGE(A1:B1))';
      const nested = valueAt(1,0);
      results.push(nested===30 ? '✓ Nested MIN/MAX/AVERAGE' : `✗ Nested functions expected 30 got ${nested}`);

      // Added Test 10: Parentheses precedence
      data[1][1]='=(1+2)*3';
      const prec = valueAt(1,1);
      results.push(prec===9 ? '✓ (1+2)*3 == 9' : `✗ (1+2)*3 expected 9 got ${prec}`);

      // Added Test 11: Error propagation for non-numeric
      rows=1; cols=3; data=createEmpty(rows,cols);
      data[0][0]='a';
      data[0][1]='=1+A1';
      const err1 = valueAt(0,1);
      results.push(err1 && err1.error==='#VALUE!' ? '✓ 1+A1 with A1="a" -> #VALUE!' : `✗ 1+A1 expected #VALUE! got ${String(err1)}`);
      data[0][2]='=SUM(1,A1)';
      const err2 = valueAt(0,2);
      results.push(err2 && err2.error==='#VALUE!' ? '✓ SUM(1,A1) -> #VALUE!' : `✗ SUM(1,A1) expected #VALUE! got ${String(err2)}`);

      // Added Test 12: Absolute reference shifting
      rows=3; cols=3; data=createEmpty(rows,cols);
      data[0][0]='1'; data[1][0]='2'; data[0][1]='3'; data[1][1]='4';
      const baseFormula='=$A$1+$A1+A$1+A1';
      const shifted='='+shiftFormulaRefs(baseFormula.slice(1),1,1);
      const okAbs = shifted==='=$A$1+$A2+B$1+B2';
      data[2][2]=shifted;
      const absVal = valueAt(2,2);
      results.push(okAbs && absVal===10 ? '✓ $A$1/$A1/A$1 refs' : `✗ $ refs failed (${shifted} -> ${absVal})`);

      // Restore again
      rows=bakRows; cols=bakCols; data=bak; renderHeader(); renderBody();

      fileInfo.textContent = results.join(' | ');
      log('Test results:', results.join(' | '));
    };

    // ===== Init =====
    renderHeader(); renderBody(); recalc();
});

