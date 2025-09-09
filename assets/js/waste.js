// assets/js/waste.js
// 廃棄分析ロジック本体（file:// 直開き対応 / CP932 既定）
// - CSV読込: 既定エンコーディング CP932（Shift_JIS）。失敗時は UTF-8 にフォールバック。
// - 列参照: 列名優先（登録日/商品コード/POS原価/数量）。見つからない場合は列番号フォールバック（10/19/27/29列目）。
// - 集計: 月単位(JAN×YYYY-MM) / 全期間(JAN)。
// - 代表値: 商品名称・部門名称は JAN 単位の最頻値を採用。
// - 販売データが存在すれば (JAN,YYYY-MM) / JAN で粗利JOINし、手残り（粗利−廃棄原価）を算出。
// - 表描画 / TOP20棒グラフ / CSV(Shift_JIS優先) エクスポート。
// 依存: PapaParse, Plotly（いずれも CDN で index.html 側が読み込み済み想定）

(function(){
  'use strict';

  const WASTE = {
    raw: [],
    fields: [],
    aggrMonth: {},  // { `${jan}|${ym}`: { jan, ym, qty, wasteCost, name, dept, gp:0, net:0 } }
    aggrTotal: {},  // { jan: { jan, qty, wasteCost, name, dept, gp:0, net:0 } }
    meta: { months: [] },
  };

  // ========== utility ==========
  const el = (id)=> document.getElementById(id);
  const toNum = (v)=> {
    if (v == null || v === '') return 0;
    const n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/[,\s]/g,''));
    return Number.isFinite(n) ? n : 0;
  };
  const yyyymm = (dstr)=>{
    if(!dstr) return '';
    const s = String(dstr).trim();
    let m = s.match(/^(\d{4})[-/](\d{1,2})/);
    if(!m){
      m = s.match(/^(\d{4})(\d{2})/);
      if(!m) return '';
      const y = m[1], mm = m[2];
      return `${y}-${mm}`;
    }
    const y = m[1], mm = String(m[2]).padStart(2,'0');
    return `${y}-${mm}`;
  };
  const normKey = (k)=> String(k||'').replace(/^﻿/,'').trim(); // BOM除去 + trim
  const pickCol = (cols, candidates)=>{
    const set = new Set(cols);
    for(const c of candidates){ if(set.has(c)) return c; }
    for(const col of cols){
      for(const c of candidates){
        if(col && c && col.indexOf(c) !== -1) return col;
      }
    }
    return null;
  };

  // ========== CSV読み込み（CP932既定） ==========
  function readAsTextWithEncoding(file, enc){
    return new Promise((resolve, reject)=>{
      try{
        const fr = new FileReader();
        fr.onerror = ()=> reject(fr.error || new Error('FileReader error'));
        fr.onload = ()=> resolve(String(fr.result || ''));
        fr.readAsText(file, enc);
      }catch(err){ reject(err); }
    });
  }
  async function loadWasteCSV(file){
    // 1) Shift_JIS で読む（既定）
    let txt = '';
    try{
      txt = await readAsTextWithEncoding(file, 'shift_jis');
    }catch(_){
      // 2) UTF-8 にフォールバック
      txt = await readAsTextWithEncoding(file, 'utf-8');
    }
    // 3) PapaParse で parse（ヘッダあり）
    const res = Papa.parse(txt, { header:true, skipEmptyLines:true });
    // ヘッダ正規化
    const fields = (res.meta && Array.isArray(res.meta.fields) ? res.meta.fields : []).map(normKey);
    const data = (res.data || []).map(row=>{
      const out = {};
      for(const k in row){
        out[normKey(k)] = row[k];
      }
      return out;
    });
    return { data, fields };
  }

  // ========== 名称・部門 代表値（最頻値） ==========
  function buildNameDeptMaps(rows, cJAN, cNAME, cDEPT){
    const nameCount = new Map();
    const deptCount = new Map();
    for(const r of rows){
      const jan = String(r[cJAN] ?? '').trim();
      if(!jan) continue;
      if(cNAME){
        const name = String(r[cNAME] ?? '').trim();
        if(name){
          if(!nameCount.has(jan)) nameCount.set(jan, new Map());
          const m = nameCount.get(jan);
          m.set(name, (m.get(name)||0)+1);
        }
      }
      if(cDEPT){
        const dept = String(r[cDEPT] ?? '').trim();
        if(dept){
          if(!deptCount.has(jan)) deptCount.set(jan, new Map());
          const m = deptCount.get(jan);
          m.set(dept, (m.get(dept)||0)+1);
        }
      }
    }
    const nameRep = new Map(), deptRep = new Map();
    for(const [jan, m] of nameCount){
      let best=null, cnt=-1;
      for(const [name, c] of m){ if(c>cnt){ best=name; cnt=c; } }
      if(best) nameRep.set(jan, best);
    }
    for(const [jan, m] of deptCount){
      let best=null, cnt=-1;
      for(const [dept, c] of m){ if(c>cnt){ best=dept; cnt=c; } }
      if(best) deptRep.set(jan, best);
    }
    return { nameRep, deptRep };
  }

  // ========== 集計 ==========
  function aggregate(){
    WASTE.aggrMonth = {};
    WASTE.aggrTotal = {};
    const rows = WASTE.raw;
    if(rows.length === 0) return;

    const cols = Object.keys(rows[0]||{});
    const fields = WASTE.fields || [];

    // 列名優先・番号フォールバック
    const cDATE = pickCol(cols, ['登録日']) || fields[9]  || null;   // 10列目
    const cJAN  = pickCol(cols, ['商品コード','JAN','ＪＡＮ','JANコード','商品CD']) || fields[18] || null; // 19列目
    const cPOS  = pickCol(cols, ['POS原価','原価単価','原価','売単価']) || fields[26] || null; // 27列目
    const cQTY  = pickCol(cols, ['数量','数','個数']) || fields[28] || null; // 29列目
    const cNAME = pickCol(cols, ['商品名称','商品名','品名']) || null;
    const cDEPT = pickCol(cols, ['部門名称','部門名']) || null;

    if(!cDATE || !cJAN || !cPOS || !cQTY){
      console.warn('必要列が見つかりません', {cDATE,cJAN,cPOS,cQTY, cols, fields});
    }

    const { nameRep, deptRep } = buildNameDeptMaps(rows, cJAN, cNAME, cDEPT);
    const months = new Set();

    for(const r of rows){
      const jan = String(r[cJAN] ?? '').trim();
      const ym  = yyyymm(r[cDATE]);
      const qty = toNum(r[cQTY]);
      const pos = toNum(r[cPOS]);
      if(!jan || !ym) continue;

      const name = nameRep.get(jan) || '';
      const dept = deptRep.get(jan) || '';
      const wasteCost = qty * pos;
      months.add(ym);

      const k = `${jan}|${ym}`;
      if(!WASTE.aggrMonth[k]) WASTE.aggrMonth[k] = { jan, ym, qty:0, wasteCost:0, name, dept, gp:0, net:0 };
      WASTE.aggrMonth[k].qty += qty;
      WASTE.aggrMonth[k].wasteCost += wasteCost;

      if(!WASTE.aggrTotal[jan]) WASTE.aggrTotal[jan] = { jan, qty:0, wasteCost:0, name, dept, gp:0, net:0 };
      WASTE.aggrTotal[jan].qty += qty;
      WASTE.aggrTotal[jan].wasteCost += wasteCost;
    }

    // 販売データ（任意）JOIN
    const SALES = window.SALES || null;
    if(SALES && (SALES.gpByJanMonth || SALES.gpByJanTotal)){
      if(SALES.gpByJanMonth){
        for(const k in WASTE.aggrMonth){
          const rec = WASTE.aggrMonth[k];
          const gp = toNum(SALES.gpByJanMonth[k]);
          rec.gp = gp;
          rec.net = gp - rec.wasteCost;
        }
      }
      if(SALES.gpByJanTotal){
        for(const jan in WASTE.aggrTotal){
          const rec = WASTE.aggrTotal[jan];
          const gp = toNum(SALES.gpByJanTotal[jan]);
          rec.gp = gp;
          rec.net = gp - rec.wasteCost;
        }
      }
    }

    WASTE.meta.months = Array.from(months).sort();
  }

  // ========== 描画 ==========
  function renderTable(mode){
    const wrap = el('wasteTable');
    if(!wrap) return;
    const SALES = window.SALES || null;
    const hasMonthGP = !!(SALES && SALES.gpByJanMonth);
    const hasTotalGP = !!(SALES && SALES.gpByJanTotal);
    let html = '';

    if(mode === 'month'){
      const rows = Object.values(WASTE.aggrMonth).sort((a,b)=> (a.ym===b.ym)? a.jan.localeCompare(b.jan) : a.ym.localeCompare(b.ym));
      html += '<table class="table"><thead><tr>'
           +  '<th>JAN</th><th>商品名称</th><th>部門名称</th><th>年月</th><th class="num">廃棄数量</th><th class="num">廃棄原価合計</th>';
      if(hasMonthGP){ html += '<th class="num">粗利額合計</th><th class="num">手残り利益</th>'; }
      html += '</tr></thead><tbody>';
      for(const r of rows){
        html += `<tr><td>${r.jan}</td><td>${r.name||''}</td><td>${r.dept||''}</td><td>${r.ym}</td>`
             +  `<td class="num">${r.qty}</td><td class="num">${Math.round(r.wasteCost)}</td>`;
        if(hasMonthGP){ html += `<td class="num">${Math.round(r.gp||0)}</td><td class="num">${Math.round(r.net||0)}</td>`; }
        html += '</tr>';
      }
      html += '</tbody></table>';
    }else{
      const rows = Object.values(WASTE.aggrTotal).sort((a,b)=> a.jan.localeCompare(b.jan));
      html += '<table class="table"><thead><tr>'
           +  '<th>JAN</th><th>商品名称</th><th>部門名称</th><th class="num">廃棄数量(全期間)</th><th class="num">廃棄原価合計(全期間)</th>';
      if(hasTotalGP){ html += '<th class="num">粗利額合計(全期間)</th><th class="num">手残り利益(全期間)</th>'; }
      html += '</tr></thead><tbody>';
      for(const r of rows){
        html += `<tr><td>${r.jan}</td><td>${r.name||''}</td><td>${r.dept||''}</td>`
             +  `<td class="num">${r.qty}</td><td class="num">${Math.round(r.wasteCost)}</td>`;
        if(hasTotalGP){ html += `<td class="num">${Math.round(r.gp||0)}</td><td class="num">${Math.round(r.net||0)}</td>`; }
        html += '</tr>';
      }
      html += '</tbody></table>';
    }

    wrap.innerHTML = html;
    const exp = el('wasteExportBtn');
    if(exp) exp.disabled = !html.includes('<tbody>');
  }

  function renderPlot(mode){
    const div = el('wastePlot');
    if(!div) return;
    if(!(window.Plotly && window.Plotly.newPlot)){ div.innerHTML=''; return; }

    let labels=[], values=[], title='';
    if(mode === 'month'){
      const yms = WASTE.meta.months;
      if(!yms || yms.length===0){ window.Plotly.purge(div); return; }
      const last = yms[yms.length-1];
      const arr = Object.values(WASTE.aggrMonth).filter(r=>r.ym===last).sort((a,b)=> b.wasteCost - a.wasteCost).slice(0,20);
      labels = arr.map(r=> (r.name? `${r.name}
(${r.jan})` : r.jan) );
      values = arr.map(r=> Math.round(r.wasteCost));
      title = `廃棄原価TOP20（${last}）`;
    }else{
      const arr = Object.values(WASTE.aggrTotal).sort((a,b)=> b.wasteCost - a.wasteCost).slice(0,20);
      labels = arr.map(r=> (r.name? `${r.name}
(${r.jan})` : r.jan) );
      values = arr.map(r=> Math.round(r.wasteCost));
      title = '廃棄原価TOP20（全期間）';
    }
    window.Plotly.newPlot(div, [{type:'bar', x:labels, y:values}], {title, margin:{t:40,l:40,r:10,b:100}});
  }

  const currentMode = ()=> {
    const r = document.querySelector('input[name="wasteAgg"]:checked');
    return r ? r.value : 'month';
  };
  const refresh = ()=>{
    const mode = currentMode();
    renderTable(mode);
    renderPlot(mode);
  };

  // ========== CSVエクスポート（CP932優先） ==========
  function toSJISBlob(csv){
    try{
      // @ts-ignore
      const enc = new TextEncoder('shift_jis');
      return new Blob([enc.encode(csv)], {type:'text/csv; charset=shift_jis'});
    }catch(_){}
    if(window.Encoding && window.Encoding.convert){
      const u8 = window.Encoding.convert(csv, { to: 'SJIS', from: 'UNICODE', type: 'array' });
      return new Blob([new Uint8Array(u8)], {type:'text/csv; charset=shift_jis'});
    }
    return new Blob(['﻿'+csv], {type:'text/csv; charset=utf-8'});
  }
  function exportCSV(){
    const mode = currentMode();
    const SALES = window.SALES || null;
    const hasMonthGP = !!(SALES && SALES.gpByJanMonth);
    const hasTotalGP = !!(SALES && SALES.gpByJanTotal);
    let headers=[], rows=[];
    if(mode==='month'){
      headers = ['JAN','商品名称','部門名称','年月','廃棄数量','廃棄原価合計'];
      if(hasMonthGP){ headers.push('粗利額合計','手残り利益'); }
      rows = Object.values(WASTE.aggrMonth).sort((a,b)=> (a.ym===b.ym)? a.jan.localeCompare(b.jan) : a.ym.localeCompare(b.ym))
        .map(r=>{
          const rec = [r.jan, r.name||'', r.dept||'', r.ym, r.qty, Math.round(r.wasteCost)];
          if(hasMonthGP){ rec.push(Math.round(r.gp||0), Math.round(r.net||0)); }
          return rec;
        });
    }else{
      headers = ['JAN','商品名称','部門名称','廃棄数量(全期間)','廃棄原価合計(全期間)'];
      if(hasTotalGP){ headers.push('粗利額合計(全期間)','手残り利益(全期間)'); }
      rows = Object.values(WASTE.aggrTotal).sort((a,b)=> a.jan.localeCompare(b.jan))
        .map(r=>{
          const rec = [r.jan, r.name||'', r.dept||'', r.qty, Math.round(r.wasteCost)];
          if(hasTotalGP){ rec.push(Math.round(r.gp||0), Math.round(r.net||0)); }
          return rec;
        });
    }
    const esc = (v)=> {
      const s = (v==null ? '' : String(v));
      return /[",
]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
    };
    const csv = [headers, ...rows].map(r=> r.map(esc).join(',')).join('
');
    const blob = toSJISBlob(csv);
    const a = document.createElement('a');
    const yms = WASTE.meta.months;
    const last = yms && yms.length ? yms[yms.length-1] : '';
    a.download = (mode==='month' && last) ? `waste_report_${last}.csv` : 'waste_report_all.csv';
    a.href = URL.createObjectURL(blob);
    document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 0);
  }

  // ========== 配線 ==========
  function wire(){
    const pickBtn = el('wastePickBtn');
    const input = el('wasteFile');
    if(pickBtn && input){
      pickBtn.addEventListener('click', ()=> input.click());
      input.addEventListener('change', async (e)=>{
        const file = e.target.files && e.target.files[0];
        if(!file) return;
        const st = el('wasteStatus');
        if(st) st.textContent = `読み込み中… (${file.name})`;
        try{
          const res = await loadWasteCSV(file); // CP932既定で読み込み
          WASTE.raw = (res.data || []);
          WASTE.fields = (res.fields || []);
          if(st) st.textContent = `取り込み完了：${WASTE.raw.length}行`;
          aggregate();
          refresh();
          window.WASTE = WASTE; // デバッグ用公開
        }catch(err){
          console.error(err);
          if(st) st.textContent = '読み込み失敗';
        }
      });
    }
    document.querySelectorAll('input[name="wasteAgg"]').forEach(r=> r.addEventListener('change', refresh));
    const exp = el('wasteExportBtn');
    if(exp) exp.addEventListener('click', exportCSV);
  }

  if(document.readyState === 'loading'){ document.addEventListener('DOMContentLoaded', wire); }
  else { wire(); }

})();