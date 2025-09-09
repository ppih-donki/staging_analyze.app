
    /** 固定カテゴリ（小分類）マスタURL */
    const CATEGORY_TXT_URL = "https://ppih-donki.github.io/analyze.app/data/category.txt";

    // ===== ユーティリティ =====
    const $ = (id) => document.getElementById(id);

    function toNarrow(str){
      if (str == null) return "";
      return String(str)
        .replace(/\ufeff/g,"")
        .replace(/\u3000/g," ")
        .replace(/[\uFF01-\uFF5E]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
        .trim();
    }
    function keyNorm(s){ return toNarrow(s).toLowerCase().replace(/\s+/g,""); }

    function parseCSVFile(file){
      return new Promise((resolve,reject)=>{
        const enc = /商品マスタ|ﾏｽﾀ|マスタ/i.test(file.name) ? "Shift_JIS" : "UTF-8";
        Papa.parse(file, { header:true, skipEmptyLines:true, encoding:enc,
          complete: (res)=>{
            const headers = res.meta.fields.map(h=>toNarrow(h));
            const data = res.data.map(row=>{
              const o={};
              res.meta.fields.forEach((raw,i)=>{
                const normalized = headers[i];
                o[normalized] = row[raw];
              });
              return o;
            });
            resolve({ data, headers });
          }, error: reject
        });
      });
    }

    function normalizeDigits(s){ if(s==null) return ""; return String(toNarrow(s)).replace(/\D+/g, ""); }

    function resolveHeader(headers, candidates, indexFallback){
      const norm = (s)=> toNarrow(s).toLowerCase().replace(/\s+/g, "");
      const list = headers.map(h=>norm(h));
      for(const c of (candidates||[])){
        const k = norm(c); const i = list.indexOf(k); if(i>=0) return headers[i];
      }
      if(Number.isInteger(indexFallback) && headers.length>indexFallback) return headers[indexFallback];
      return null;
    }

    function toNumber(x, def=0){
      if (x === null || x === undefined) return def;
      let s = String(x);
      s = toNarrow(s).replace(/[^\d.\-]/g, "");
      if (!s) return def;
      const v = Number(s);
      return isNaN(v) ? def : v;
    }

    function parseDate(s){
      if(!s) return null;
      const t = toNarrow(s).replace(/\s+/g," ");
      const d = new Date(t);
      return isNaN(d.getTime()) ? null : d;
    }
    const fmtDate = (d)=> !d ? "-" : `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
    const fmtDateSlash = (d)=> !d ? "-" : `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()}`;
    const minDate = (arr)=> arr.reduce((a,r)=>(!a||r.生成時間<a? r.生成時間:a), null);
    const maxDate = (arr)=> arr.reduce((a,r)=>(!a||r.生成時間>a? r.生成時間:a), null);

    function renderTable(rows, elId, cols=null, max=null){
      const el = $(elId);
      if (!el) return;
      if (!rows || rows.length===0){ el.innerHTML = "<p class='muted'>データがありません。</p>"; return; }
      const headers = cols || Object.keys(rows[0]);
      let html = "<table><thead><tr>";
      headers.forEach(h=> html += `<th>${h}</th>`); html += "</tr></thead><tbody>";
      const view = max ? rows.slice(0, max) : rows;
      view.forEach(r=>{
        html += "<tr>";
        headers.forEach(h=> html += `<td>${r[h] ?? ""}</td>`);
        html += "</tr>";
      });
      html += "</tbody></table>";
      el.innerHTML = html;
    }

    // ===== 状態 =====
    let els = {};
    let joined = [];
    let masterMap = null;
    let categoryMap = null;

    // ベスレポ保持
    let lastBestList = null;
    let lastBestMeta = null;

    // 時間帯保持
    let lastTimeRows = null;
    let lastTimeMeta = null;

    // RFM保持
    let lastRfmRows = null;
    let lastRfmMeta = null;

    // セグメント保持
    let lastSegMeta = null;
    let lastSegTime = null;
    let lastSegProd = null;
    let lastSegCat = null;

    // リピート/併売 保持
    let lastRepeatRows = null;
    let lastRepeatMeta = null;
    let lastPairRows = null;
    let lastPairMeta = null;

    // ===== カテゴリマスタ自動読込 =====
    async function autoLoadCategory(){
      const badge = els.catStatus;
      try{
        const res = await fetch(CATEGORY_TXT_URL, { mode:"cors", cache:"no-store" });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const text = await res.text();
        categoryMap = parseCategoryTxt(text);
        if (badge){ badge.className="muted"; badge.textContent = "カテゴリマスタ：サーバーから取り込み済み"; }
      }catch(e){
        if (badge){ badge.className="error"; badge.textContent = "カテゴリtxtの取得に失敗: " + (e?.message||e); }
      }
    }
    function parseCategoryTxt(text){
      const lines = String(text).split(/\r?\n/);
      const map = new Map();
      for (const line of lines){
        const raw = toNarrow(line);
        if (!raw) continue;
        let parts = raw.split("\t");
        if (parts.length < 2) parts = raw.split(/\s+/);
        if (parts.length < 2) continue;
        const code = parts[0];
        const name = parts.slice(1).join(" ");
        if (!/^\d{3,}$/.test(code)) continue;
        map.set(code, name);
      }
      return map;
    }

    function findHeader(headers, candidates){
      const normHeaders = headers.map(h=>keyNorm(h));
      for (const cand of candidates){
        const k = keyNorm(cand);
        const idx = normHeaders.indexOf(k);
        if (idx !== -1) return headers[idx];
      }
      return null;
    }

    function buildMasterMap(pm){
      const { data: pmRows, headers } = pm;
      const map = new Map();

      const codeCands = ["商品コード","商品cd","jan","janコード","バーコード","コード"];
      const catCands  = ["商品分類","小分類","小分類コード","分類"];
      const nameCands = ["商品名","商品名規格","品名","商品名称"];
      const costCands = ["pos原価","pos 原価","ｐｏｓ原価","原価","仕入原価","pos原価(税抜)","原価(税抜)"];

      const codeHdr = findHeader(headers, codeCands);
      const catHdr  = findHeader(headers, catCands);
      const nameHdr = findHeader(headers, nameCands);
      let   costHdr = resolveHeader(headers, costCands, 12);

      for (const r of pmRows){
        const codeRaw = toNarrow(codeHdr ? r[codeHdr] : "").trim();
        const code = normalizeDigits(codeRaw);
        if (!code) continue;
        const nm   = toNarrow(nameHdr ? r[nameHdr] : "") || "";
        const cat  = toNarrow(catHdr ? r[catHdr] : "") || "";
        let cost = 0; let 未登録 = false;
        if (costHdr && r.hasOwnProperty(costHdr)) {
          const rawCost = String(r[costHdr] ?? "").trim();
          if (rawCost === "**********") { cost = 0; 未登録 = true; }
          else { cost = toNumber(rawCost, 0); }
        }
        map.set(code, { 商品名: nm, 商品分類: cat, POS原価: cost, 未登録 });
      }
      return map;
    }

    // ===== 取り込み & サマリー =====
    async function runImport(){
      const statusEl = els.status; const runBtn = els.runBtn;
      if (runBtn) runBtn.disabled = true;
      if (statusEl){ statusEl.className = "muted"; statusEl.textContent = "取り込み中…"; }

      try{
        const [tx, ln] = await Promise.all([
          parseCSVFile(els.txInput.files[0]),
          parseCSVFile(els.lnInput.files[0])
        ]);
        let pm = null;
        if (els.pmInput.files.length===1) pm = await parseCSVFile(els.pmInput.files[0]);

        const txRows = tx.data, lnRows = ln.data;

        const needTx = new Set(["CGID","オーダーID","生成時間"].map(toNarrow));
        const needLn = new Set(["CGID","オーダーID","JANコード","商品名","商品数","商品価格","生成時間"].map(toNarrow));
        const hasAll=(rows,need)=> rows.length && [...need].every(c=> Object.prototype.hasOwnProperty.call(rows[0],c));
        if(!hasAll(txRows,needTx) || !hasAll(lnRows,needLn)){
          if (statusEl){ statusEl.className="error"; statusEl.textContent="必須列が不足（取引毎: CGID, オーダーID, 生成時間 / 商品毎: CGID, オーダーID, JANコード, 商品名, 商品数, 商品価格, 生成時間）"; }
          return;
        }

        txRows.forEach(r=>{
          r.CGID = toNarrow(r.CGID);
          r.オーダーID = toNarrow(r.オーダーID);
          r.生成時間 = parseDate(r.生成時間);
        });
        lnRows.forEach(r=>{
          r.CGID = toNarrow(r.CGID);
          r.オーダーID = toNarrow(r.オーダーID);
          r.JANコード = toNarrow(r.JANコード);
          r.商品名 = toNarrow(r.商品名);
          r.生成時間 = parseDate(r.生成時間);
          r.商品数 = Math.trunc(toNumber(r.商品数, 0));
          r.商品価格 = toNumber(r.商品価格, 0);
        });

        masterMap = null;
        let masterInfoText = "商品マスタ：未読込（原価=0扱い）";
        if (pm){
          masterMap = buildMasterMap(pm);
          masterInfoText = `商品マスタ：${pm.data.length.toLocaleString()}行（JANキー: ${masterMap.size.toLocaleString()}件）`;
        }

        const txMap = new Map();
        txRows.forEach(r => txMap.set(`${r.CGID}||${r.オーダーID}`, r));

        joined = [];
        let matched = 0;
        const unmatchedJAN = new Map();

        for(const r of lnRows){
          const key = `${r.CGID}||${r.オーダーID}`;
          if(!txMap.has(key)) continue;
          const txRec = txMap.get(key);

          let 商品名2 = r.商品名;
          let 商品分類 = "";
          let 小分類名 = "";
          let 原価 = 0;

          if (masterMap && masterMap.has(r.JANコード)){
            const m = (()=>{ const janRaw = toNarrow(r.JANコード); const key = normalizeDigits(janRaw); return masterMap.get(key); })();
            if (m.商品名) 商品名2 = m.商品名;
            商品分類 = m.商品分類 || "";
            原価 = toNumber(m.POS原価, 0);
                      if (m && m.未登録) { 原価 = toNumber(r.商品価格, 0); }
matched++;
          } else if (masterMap){
            unmatchedJAN.set(r.JANコード, (unmatchedJAN.get(r.JANコード)||0)+1);
          }

          if (categoryMap && 商品分類){
            小分類名 = categoryMap.get(商品分類) || "";
          }

          const 明細金額 = (r.商品価格 || 0) * (r.商品数 || 0);
          const 明細粗利 = ((r.商品価格 || 0) - (原価 || 0)) * (r.商品数 || 0);

          joined.push({
            CGID: r.CGID,
            オーダーID: r.オーダーID,
            生成時間: txRec.生成時間 || r.生成時間,
            JANコード: r.JANコード,
            商品名: 商品名2,
            商品分類,
            小分類名,
            商品数: r.商品数,
            商品価格: r.商品価格,
            POS原価: 原価,
            明細金額,
            明細粗利
          });
        }

        const uniqUsers = new Set(joined.map(x=>x.CGID)).size;
        const receipts  = new Set(joined.map(x=>x.オーダーID)).size;
        const totalQty  = joined.reduce((s,x)=> s + (x.商品数||0), 0);
        const totalAmt  = joined.reduce((s,x)=> s + (x.明細金額||0), 0);
        const totalGp   = joined.reduce((s,x)=> s + (x.明細粗利||0), 0);
        const dMin = minDate(joined), dMax = maxDate(joined);

        $("summary").innerHTML = `
          <div class="metric">期間: <span class="kpi-strong">${fmtDateSlash(dMin)} ~ ${fmtDateSlash(dMax)}</span></div>
          <div class="metric">ユニークユーザー数: <span class="kpi-strong">${uniqUsers.toLocaleString()}</span></div>
          <div class="metric">レシート件数: <span class="kpi-strong">${receipts.toLocaleString()}</span></div>
          <div class="metric">売上点数合計: <span class="kpi-strong">${totalQty.toLocaleString()}</span></div>
          <div class="metric">売上金額合計: <span class="kpi-strong">${totalAmt.toLocaleString()}</span></div>
          <div class="metric">粗利合計: <span class="kpi-strong">${totalGp.toLocaleString()}</span></div>
        `;

        // 日付UIの初期値（seg/repeat/pair も含めて設定）
        if (dMin && dMax){
          ["best","time","prod","rfm","seg","repeat","pair"].forEach(prefix=>{
            $(prefix+"Start").min = $(prefix+"End").min = fmtDate(dMin);
            $(prefix+"Start").max = $(prefix+"End").max = fmtDate(dMax);
            $(prefix+"Start").value = fmtDate(dMin);
            $(prefix+"End").value   = fmtDate(dMax);
          });
        }

        const diag = $("masterDiag");
        const rate = joined.length ? Math.round((matched / joined.length) * 100) : 0;
        let topUnmatched = "";
        if (unmatchedJAN.size){
          const arr = Array.from(unmatchedJAN.entries()).sort((a,b)=> b[1]-a[1]).slice(0,10);
          topUnmatched = "<br>未マッチJAN上位: " + arr.map(([jan,c])=> `${jan}(${c})`).join(", ");
        }
        diag.style.display = "block";
        diag.innerHTML = `${masterInfoText} / 明細行: ${joined.length.toLocaleString()} / マスタ適用: ${matched.toLocaleString()}（${rate}%）${topUnmatched}`;

        renderTable(
          joined,
          "preview",
          ["CGID","オーダーID","生成時間","JANコード","商品名","商品分類","小分類名","商品数","商品価格","POS原価","明細金額","明細粗利"],
          50
        );

        if (statusEl){ statusEl.className = "muted"; statusEl.textContent = "完了！"; }
      }catch(err){
        console.error(err);
        if (statusEl){ statusEl.className = "error"; statusEl.textContent = "エラー: " + (err?.message || String(err)); }
      }finally{
        if (runBtn) runBtn.disabled = false;
      }
    }

    // ===== ベスレポ =====
    function computeBestReport(){
      if (!joined.length){
        $("bestInfo").textContent = "先にCSVを取り込んでください。";
        $("bestTable").innerHTML = "";
        $("bestExport").disabled = true;
        lastBestList = null; lastBestMeta = null;
        return;
      }
      const s = $("bestStart").value;
      const e = $("bestEnd").value;
      const metric = $("bestMetric").value;

      const sDate = s ? new Date(s+"T00:00:00") : null;
      const eDate = e ? new Date(e+"T23:59:59") : null;

      const rows = joined.filter(r=>{
        if (!r.生成時間) return false;
        if (sDate && r.生成時間 < sDate) return false;
        if (eDate && r.生成時間 > eDate) return false;
        return true;
      });

      const byJan = new Map();
      for (const r of rows){
        const k = r.JANコード || "-";
        if (!byJan.has(k)){
          byJan.set(k, {
            JANコード: k, 商品名: r.商品名 || "", 小分類名: r.小分類名 || "",
            売上金額合計: 0, 売上点数合計: 0, 粗利合計: 0,
            ユーザーset: new Set(), userCnt: new Map(),
          });
        }
        const o = byJan.get(k);
        o.売上金額合計 += (r.明細金額 || 0);
        o.売上点数合計 += (r.商品数 || 0);
        o.粗利合計   += (r.明細粗利 || 0);
        if (r.CGID){
          o.ユーザーset.add(r.CGID);
          o.userCnt.set(r.CGID, (o.userCnt.get(r.CGID) || 0) + (r.商品数 || 0));
        }
      }

      const list = [];
      for (const o of byJan.values()){
        const repeatUsers = Array.from(o.userCnt.values()).filter(cnt => cnt >= 2).length;
        list.push({
          JANコード: o.JANコード, 商品名: o.商品名, 小分類名: o.小分類名,
          売上金額合計: Math.round(o.売上金額合計),
          売上点数合計: Math.round(o.売上点数合計),
          粗利合計: Math.round(o.粗利合計),
          購買ユニークユーザー数: o.ユーザーset.size,
          リピートユーザー数: repeatUsers
        });
      }

      const keyMap = { amt:"売上金額合計", qty:"売上点数合計", gp:"粗利合計", uniq:"購買ユニークユーザー数", repeat:"リピートユーザー数" };
      const metricLabelMap = { amt:"売上ベスレポ", qty:"点数ベスレポ", gp:"粗利ベスレポ", uniq:"ユニークユーザー数ベスレポ", repeat:"リピートベスレポ" };
      const sortKey = keyMap[metric] || "売上金額合計";
      const label  = metricLabelMap[metric] || "売上ベスレポ";
      list.sort((a,b)=> b[sortKey] - a[sortKey]);

      $("bestInfo").textContent = `期間: ${s || "-"} ~ ${e || "-"} / 件数: ${list.length.toLocaleString()}`;
      renderTable(list, "bestTable",
        ["JANコード","商品名","小分類名","売上金額合計","売上点数合計","粗利合計","購買ユニークユーザー数","リピートユーザー数"], null);

      lastBestList = list;
      lastBestMeta = { start: s || "", end: e || "", label };
      $("bestExport").disabled = !lastBestList || lastBestList.length===0;
    }
    function exportBestCSV(){
      if (!lastBestList || !lastBestList.length) return;
      const headers = ["JANコード","商品名","小分類名","売上金額合計","売上点数合計","粗利合計","購買ユニークユーザー数","リピートユーザー数"];
      const ymd = (s)=> s ? s.replaceAll("-", "/") : "-";
      const preface = `${ymd(lastBestMeta.start)} ~ ${ymd(lastBestMeta.end)} ${lastBestMeta.label}`;
      const lines = [preface, headers.join(",")];
      for (const r of lastBestList){
        const row = headers.map(h=>{ let v=r[h]; if (v==null) v=""; v=String(v).replaceAll('"','""'); if (/[",\n]/.test(v)) v=`"${v}"`; return v; });
        lines.push(row.join(","));
      }
      const csv="\uFEFF"+lines.join("\n");
      const fn = `${lastBestMeta.start || "all"}_${lastBestMeta.end || "all"}_${lastBestMeta.label}.csv`;
      const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob);
      const a=document.createElement("a"); a.href=url; a.download=fn; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ===== 時間帯集計 =====
    function computeTimeReport(){
      if (!joined.length){
        $("timeInfo").textContent = "先にCSVを取り込んでください。";
        $("timeTable").innerHTML = "";
        $("timeExport").disabled = true;
        $("timePng").disabled = true;
        lastTimeRows = null; lastTimeMeta = null;
        Plotly.purge("timeChart");
        return;
      }
      const s = $("timeStart").value;
      const e = $("timeEnd").value;
      const dow = $("timeDow").value;
      const metric = $("timeMetric").value;

      const sDate = s ? new Date(s+"T00:00:00") : null;
      const eDate = e ? new Date(e+"T23:59:59") : null;

      const rows = joined.filter(r=>{
        if (!r.生成時間) return false;
        const wd = r.生成時間.getDay(); // 0:Sun
        const isWeekend = (wd===0 || wd===6);
        if (dow==="weekday" && isWeekend) return false;
        if (dow==="weekend" && !isWeekend) return false;
        if (sDate && r.生成時間 < sDate) return false;
        if (eDate && r.生成時間 > eDate) return false;
        return true;
      });

      const buckets = Array.from({length:24}, (_,h)=>({
        時間帯: `${String(h).padStart(2,"0")}:00`,
        売上金額: 0,
        売上点数: 0,
        粗利: 0,
        _receipts: new Set(),
        _users: new Set(),
      }));

      for (const r of rows){
        const h = r.生成時間.getHours();
        const b = buckets[h];
        b.売上金額 += (r.明細金額 || 0);
        b.売上点数 += (r.商品数 || 0);
        b.粗利    += (r.明細粗利 || 0);
        if (r.オーダーID) b._receipts.add(r.オーダーID);
        if (r.CGID) b._users.add(r.CGID);
      }

      const tableRows = buckets.map(b=>{
        const rc = b._receipts.size || 0;
        const basket = rc ? (b.売上点数/rc) : 0;
        const aov    = rc ? (b.売上金額/rc) : 0;
        return {
          時間帯: b.時間帯,
          売上金額: Math.round(b.売上金額),
          売上点数: Math.round(b.売上点数),
          粗利: Math.round(b.粗利),
          レシート件数: rc,
          ユニークユーザー数: b._users.size,
          バスケットサイズ: +basket.toFixed(2),
          平均客単価: +aov.toFixed(2)
        };
      });

      const yKeyMap = { amt:"売上金額", receipts:"レシート件数", basket:"バスケットサイズ", aov:"平均客単価" };
      const yKey = yKeyMap[metric] || "売上金額";
      const label = yKey;

      Plotly.newPlot("timeChart", [{
        type:"scatter", mode:"lines+markers",
        x: buckets.map(b=> b.時間帯),
        y: tableRows.map(r=> r[yKey]),
        name: label
      }], {
        margin:{l:50,r:10,t:10,b:40},
        xaxis:{title:"時間帯（時）"},
        yaxis:{title: label, rangemode:"tozero"},
        displayModeBar:false
      }, {responsive:true});

      const ranking = [...tableRows].sort((a,b)=> b[yKey] - a[yKey]);

      $("timeInfo").textContent = `期間: ${s || "-"} ~ ${e || "-"} / 曜日: ${({"all":"すべて","weekday":"平日","weekend":"土日"}[dow])} / 指標: ${label}`;
      renderTable(ranking, "timeTable",
        ["時間帯","売上金額","売上点数","粗利","レシート件数","ユニークユーザー数","バスケットサイズ","平均客単価"], null);

      lastTimeRows = tableRows;
      lastTimeMeta = { start:s||"", end:e||"", label:`時間帯集計（${label}）`, dow };
      $("timeExport").disabled = lastTimeRows.length===0;
      $("timePng").disabled = false;
    }
    async function exportTimePNG(){ await Plotly.downloadImage("timeChart", {format:"png", filename:"time_series"}); }
    function exportTimeCSV(){
      if (!lastTimeRows || !lastTimeRows.length) return;
      const headers = ["時間帯","売上金額","売上点数","粗利","レシート件数","ユニークユーザー数","バスケットサイズ","平均客単価"];
      const ymd = (s)=> s ? s.replaceAll("-", "/") : "-";
      const dowLabel = {"all":"すべて","weekday":"平日","weekend":"土日"}[lastTimeMeta.dow] || "すべて";
      const preface = `${ymd(lastTimeMeta.start)} ~ ${ymd(lastTimeMeta.end)} ${lastTimeMeta.label}（${dowLabel}）`;
      const lines = [preface, headers.join(",")];
      for (const r of lastTimeRows){
        const row = headers.map(h=>{ let v=r[h]; if (v==null) v=""; v=String(v).replaceAll('"','""'); if (/[",\n]/.test(v)) v=`"${v}"`; return v; });
        lines.push(row.join(","));
      }
      const csv="\uFEFF"+lines.join("\n");
      const fn = `${lastTimeMeta.start || "all"}_${lastTimeMeta.end || "all"}_${lastTimeMeta.label}.csv`;
      const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob);
      const a=document.createElement("a"); a.href=url; a.download=fn; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ===== 単品分析 =====
    function computeProductReport(){
      if (!joined.length){
        $("prodHeader").textContent = "先にCSVを取り込んでください。";
        $("prodSummary").innerHTML = "";
        Plotly.purge("prodChart");
        $("prodDiag").style.display = "none";
        return;
      }
      const s = $("prodStart").value;
      const e = $("prodEnd").value;
      const janInput = toNarrow($("prodJan").value).replace(/\s+/g,"");
      if (!janInput){
        $("prodHeader").textContent = "JANコードを入力してください。";
        $("prodSummary").innerHTML = "";
        Plotly.purge("prodChart");
        $("prodDiag").style.display = "none";
        return;
      }

      const sDate = s ? new Date(s+"T00:00:00") : null;
      const eDate = e ? new Date(e+"T23:59:59") : null;

      const rows = joined.filter(r=>{
        if (!r.生成時間) return false;
        if (sDate && r.生成時間 < sDate) return false;
        if (eDate && r.生成時間 > eDate) return false;
        return (toNarrow(r.JANコード).replace(/\s+/g,"") === janInput);
      });

      const prodName = rows[0]?.商品名 || (masterMap?.get(janInput)?.商品名) || "";
      const smallCat = rows[0]?.小分類名 || "";
      $("prodHeader").innerHTML = `期間: ${s || "-"} ~ ${e || "-"} / 対象JAN: <span class="badge">${janInput}</span> ${prodName ? " / 商品名: "+prodName : ""} ${smallCat ? " / 小分類: "+smallCat : ""}`;

      if (!rows.length){
        $("prodSummary").innerHTML = "<p class='muted'>該当データがありません。</p>";
        Plotly.purge("prodChart");
        $("prodDiag").style.display = "none";
        return;
      }

      const totalQty = rows.reduce((s,x)=> s + (x.商品数||0), 0);
      const totalAmt = rows.reduce((s,x)=> s + (x.明細金額||0), 0);
      const totalGp  = rows.reduce((s,x)=> s + (x.明細粗利||0), 0);
      const uuSet    = new Set(rows.map(x=> x.CGID).filter(Boolean));
      const cntByUser = new Map();
      rows.forEach(r=> cntByUser.set(r.CGID, (cntByUser.get(r.CGID)||0) + (r.商品数||0)));
      const repeatUsers = Array.from(cntByUser.values()).filter(v=> v>=2).length;

      $("prodSummary").innerHTML = `
        <div class="metric">売上点数合計: <span class="kpi-strong">${Math.round(totalQty).toLocaleString()}</span></div>
        <div class="metric">売上金額合計: <span class="kpi-strong">${Math.round(totalAmt).toLocaleString()}</span></div>
        <div class="metric">粗利額合計: <span class="kpi-strong">${Math.round(totalGp).toLocaleString()}</span></div>
        <div class="metric">ユニーク購入者数: <span class="kpi-strong">${uuSet.size.toLocaleString()}</span></div>
        <div class="metric">リピート購入者数: <span class="kpi-strong">${repeatUsers.toLocaleString()}</span></div>
      `;

      const buckets = Array.from({length:24}, (_,h)=>({h, qty:0}));
      rows.forEach(r=>{
        const h = r.生成時間.getHours();
        buckets[h].qty += (r.商品数||0);
      });
      Plotly.newPlot("prodChart", [{
        type:"bar",
        x: buckets.map(b=> String(b.h).padStart(2,"0")+":00"),
        y: buckets.map(b=> Math.round(b.qty)),
        name:"売上点数（時間帯）"
      }], {
        margin:{l:50,r:10,t:10,b:40},
        xaxis:{title:"時間帯（時）"},
        yaxis:{title:"売上点数", rangemode:"tozero"},
        displayModeBar:false
      }, {responsive:true});

      const receiptSet = new Set(rows.map(r=> r.オーダーID));
      $("prodDiag").style.display = "block";
      $("prodDiag").innerHTML = `明細行: ${rows.length.toLocaleString()} / レシート: ${receiptSet.size.toLocaleString()} / 最小日時: ${fmtDateSlash(minDate(rows))} / 最大日時: ${fmtDateSlash(maxDate(rows))}`;
    }

    // ===== RFM =====
    function quantiles(arr, qs){
      const a = arr.slice().sort((x,y)=>x-y);
      const res = [];
      for (const q of qs){
        if (a.length===0){ res.push(0); continue; }
        const pos = (a.length-1)*q;
        const base = Math.floor(pos), rest = pos - base;
        const v = a[base] + (a[Math.min(base+1,a.length-1)] - a[base]) * rest;
        res.push(v);
      }
      return res;
    }
    function scoreByQuantile(value, cuts, reverse=false){
      if (cuts.length===2){
        if (!reverse){ if (value <= cuts[0]) return 1; if (value <= cuts[1]) return 3; return 5; }
        else         { if (value <= cuts[0]) return 5; if (value <= cuts[1]) return 3; return 1; }
      }
      if (!reverse){
        if (value <= cuts[0]) return 1;
        if (value <= cuts[1]) return 2;
        if (value <= cuts[2]) return 3;
        if (value <= cuts[3]) return 4;
        return 5;
      }else{
        if (value <= cuts[0]) return 5;
        if (value <= cuts[1]) return 4;
        if (value <= cuts[2]) return 3;
        if (value <= cuts[3]) return 2;
        return 1;
      }
    }
    function rfmSegmentName(Rs, Fs, Ms){
      if (Rs>=4 && Fs>=4 && Ms>=3) return "ヘビーユーザー";
      if (Rs>=4 && Fs<=2)          return "ライトユーザー";
      if (Rs<=2)                   return "離反顧客";
      if (Rs>=3 && Fs>=3)          return "一般ユーザー";
      return "離反傾向";
    }
    function computeRFM(){
      if (!joined.length){
        $("rfmInfo").textContent="先にCSVを取り込んでください。";
        $("rfmTable").innerHTML=""; $("rfmExport").disabled=true;
        Plotly.purge("rfmPie"); Plotly.purge("rfmScatter");
        lastRfmRows=null; lastRfmMeta=null; return;
      }
      const s=$("rfmStart").value, e=$("rfmEnd").value, taxRateSel=Number($("rfmTax").value||"0");
      const sDate=s?new Date(s+"T00:00:00"):null, eDate=e?new Date(e+"T23:59:59"):null;

      const rows=joined.filter(r=> r.生成時間 && (!sDate||r.生成時間>=sDate) && (!eDate||r.生成時間<=eDate));
      if (!rows.length){
        $("rfmInfo").textContent="該当データがありません。";
        $("rfmTable").innerHTML=""; $("rfmExport").disabled=true;
        Plotly.purge("rfmPie"); Plotly.purge("rfmScatter");
        lastRfmRows=[]; return;
      }

      const receiptMap = new Map(); // key: CGID||オーダーID
      for (const r of rows){
        const key=`${r.CGID}||${r.オーダーID}`;
        if (!receiptMap.has(key)) receiptMap.set(key,{ CGID:r.CGID, オーダーID:r.オーダーID, 金額:0, 時刻:r.生成時間 });
        const o=receiptMap.get(key);
        o.金額 += (r.明細金額||0);
        if (r.生成時間 && (!o.時刻 || r.生成時間>o.時刻)) o.時刻=r.生成時間;
      }

      const byUser = new Map();
      const taxDiv = taxRateSel>0 ? (1+taxRateSel/100) : 1;
      for (const r of rows){
        const k=r.CGID||"-"; if (!byUser.has(k)) byUser.set(k,{ CGID:k, M:0, GM:0, last:null });
        const u=byUser.get(k);
        const lineEx = (r.明細金額||0) / taxDiv;
        u.M += lineEx;
        u.GM += (r.明細粗利||0);
        if (!u.last || r.生成時間>u.last) u.last=r.生成時間;
      }
      for (const rec of receiptMap.values()){
        if (!byUser.has(rec.CGID)) byUser.set(rec.CGID,{CGID:rec.CGID,M:0,GM:0,last:null});
        const u=byUser.get(rec.CGID);
        u.F = (u.F||0)+1;
      }

      const baseDate = eDate || maxDate(rows) || new Date();
      const list=[];
      for (const u of byUser.values()){
        if (!u.F) u.F=0;
        if (!u.last) continue;
        const days = Math.max(0, Math.floor((baseDate - u.last)/(1000*60*60*24)));
        list.push({ CGID:u.CGID, R_days:days, F:u.F, M: u.M, GM: u.GM, last_date: u.last });
      }
      if (!list.length){
        $("rfmInfo").textContent="対象顧客がありません。";
        $("rfmTable").innerHTML=""; $("rfmExport").disabled=true;
        Plotly.purge("rfmPie"); Plotly.purge("rfmScatter"); lastRfmRows=[]; return;
      }

      const n=list.length;
      let cutsR, cutsF, cutsM;
      if (n<100){
        cutsR = quantiles(list.map(x=>x.R_days), [1/3, 2/3]);
        cutsF = quantiles(list.map(x=>x.F),      [1/3, 2/3]);
        cutsM = quantiles(list.map(x=>x.M),      [1/3, 2/3]);
      }else{
        cutsR = quantiles(list.map(x=>x.R_days), [0.2,0.4,0.6,0.8]);
        cutsF = quantiles(list.map(x=>x.F),      [0.2,0.4,0.6,0.8]);
        cutsM = quantiles(list.map(x=>x.M),      [0.2,0.4,0.6,0.8]);
      }

      const rowsOut=[];
      for (const x of list){
        const Rs = scoreByQuantile(x.R_days, cutsR, true);
        const Fs = scoreByQuantile(x.F,      cutsF, false);
        const Ms = scoreByQuantile(x.M,      cutsM, false);
        const seg = rfmSegmentName(Rs, Fs, Ms);
        rowsOut.push({
          CGID: x.CGID,
          "R（日数）": x.R_days,
          "F（回）": x.F,
          "M（税抜合計）": Math.round(x.M),
          "GM（粗利合計）": Math.round(x.GM),
          Rスコア: Rs, Fスコア: Fs, Mスコア: Ms,
          セグメント: seg,
          最終購入日: fmtDateSlash(x.last_date)
        });
      }

      $("rfmInfo").textContent=`期間: ${s||"-"} ~ ${e||"-"} / 顧客数: ${rowsOut.length.toLocaleString()} / 税率: ${taxRateSel}%`;

      const segCount = {};
      for (const r of rowsOut){ segCount[r.セグメント]=(segCount[r.セグメント]||0)+1; }
      const segLabels = Object.keys(segCount);
      const segValues = segLabels.map(k=>segCount[k]);
      Plotly.newPlot("rfmPie", [{ type:"pie", labels:segLabels, values:segValues, hole:0.35 }], { margin:{l:10,r:10,t:20,b:10}, showlegend:true }, {responsive:true});

      Plotly.newPlot("rfmScatter", [{
        x: rowsOut.map(r=> r["R（日数）"]),
        y: rowsOut.map(r=> r["F（回）"]),
        text: rowsOut.map(r=> `${r.CGID} / M=${r["M（税抜合計）"]}`),
        mode:"markers",
        marker:{ size: rowsOut.map(r=> Math.max(6, Math.min(24, r["M（税抜合計）"]/Math.max(1, (cutsM[0]||1)))) ) },
        type:"scatter",
        transforms:[{ type:"groupby", groups: rowsOut.map(r=> r.セグメント) }],
      }], { margin:{l:50,r:10,t:10,b:40}, xaxis:{title:"R（日数・小さいほど良）", rangemode:"tozero"}, yaxis:{title:"F（回）", rangemode:"tozero"} }, {responsive:true});

      renderTable(rowsOut, "rfmTable",
        ["CGID","R（日数）","F（回）","M（税抜合計）","GM（粗利合計）","Rスコア","Fスコア","Mスコア","セグメント","最終購入日"], null);

      lastRfmRows = rowsOut;
      lastRfmMeta = { start:s||"", end:e||"", label:"RFM（税抜）", tax: taxRateSel };
      $("rfmExport").disabled = rowsOut.length===0;
    }
    function exportRFM(){
      if (!lastRfmRows || !lastRfmRows.length) return;
      const headers = ["CGID","R（日数）","F（回）","M（税抜合計）","GM（粗利合計）","Rスコア","Fスコア","Mスコア","セグメント","最終購入日"];
      const ymd=(s)=> s ? s.replaceAll("-","/") : "-";
      const preface = `${ymd(lastRfmMeta.start)} ~ ${ymd(lastRfmMeta.end)} ${lastRfmMeta.label}（税率=${lastRfmMeta.tax}%）`;
      const lines=[preface, headers.join(",")];
      for (const r of lastRfmRows){
        const row=headers.map(h=>{ let v=r[h]; if (v==null) v=""; v=String(v).replaceAll('"','""'); if (/[",\n]/.test(v)) v=`"${v}"`; return v; });
        lines.push(row.join(","));
      }
      const csv="\uFEFF"+lines.join("\n"); const fn=`${lastRfmMeta.start||"all"}_${lastRfmMeta.end||"all"}_RFM.csv`;
      const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob);
      const a=document.createElement("a"); a.href=url; a.download=fn; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ===== 共通: 指定期間・RFM基準でセグメントを付与し、対象ユーザー集合を返す =====
    function buildSegmentationMap(startStr, endStr, taxRate=10){
      const sDate = startStr ? new Date(startStr+"T00:00:00") : null;
      const eDate = endStr   ? new Date(endStr+"T23:59:59") : null;
      const rows = joined.filter(r=> r.生成時間 && (!sDate || r.生成時間>=sDate) && (!eDate || r.生成時間<=eDate));
      if (!rows.length) return { segOf:new Map(), users:new Set(), receipts:new Set(), rows:[] };

      const receiptMap = new Map();
      for (const r of rows){
        const key = `${r.CGID}||${r.オーダーID}`;
        if (!receiptMap.has(key)) receiptMap.set(key, {CGID:r.CGID, 時刻:r.生成時間});
        const o = receiptMap.get(key);
        if (r.生成時間 && (!o.時刻 || r.生成時間>o.時刻)) o.時刻=r.生成時間;
      }

      const byUser = new Map();
      const taxDiv = taxRate>0 ? (1+taxRate/100) : 1;
      for (const r of rows){
        const k=r.CGID||"-"; if (!byUser.has(k)) byUser.set(k,{M:0,last:null});
        const u=byUser.get(k);
        u.M += (r.明細金額||0)/taxDiv;
        if (!u.last || r.生成時間>u.last) u.last=r.生成時間;
      }
      for (const rec of receiptMap.values()){
        if (!byUser.has(rec.CGID)) byUser.set(rec.CGID,{M:0,last:null});
        byUser.get(rec.CGID).F = (byUser.get(rec.CGID).F||0)+1;
      }

      const baseDate = eDate || maxDate(rows) || new Date();
      const list=[];
      for (const [cgid,u] of byUser.entries()){
        if (!u.last){ continue; }
        const days = Math.max(0, Math.floor((baseDate - u.last)/(1000*60*60*24)));
        list.push({CGID:cgid, R_days:days, F:u.F||0, M:u.M});
      }

      let cutsR, cutsF, cutsM;
      if (list.length<100){
        cutsR = quantiles(list.map(x=>x.R_days), [1/3,2/3]);
        cutsF = quantiles(list.map(x=>x.F),      [1/3,2/3]);
        cutsM = quantiles(list.map(x=>x.M),      [1/3,2/3]);
      }else{
        cutsR = quantiles(list.map(x=>x.R_days), [0.2,0.4,0.6,0.8]);
        cutsF = quantiles(list.map(x=>x.F),      [0.2,0.4,0.6,0.8]);
        cutsM = quantiles(list.map(x=>x.M),      [0.2,0.4,0.6,0.8]);
      }

      const segOf = new Map();
      for (const x of list){
        const Rs = scoreByQuantile(x.R_days, cutsR, true);
        const Fs = scoreByQuantile(x.F,      cutsF, false);
        const Ms = scoreByQuantile(x.M,      cutsM, false);
        segOf.set(x.CGID, rfmSegmentName(Rs, Fs, Ms));
      }

      const users = new Set(rows.map(r=> r.CGID).filter(Boolean));
      const receipts = new Set(rows.map(r=> r.オーダーID).filter(Boolean));
      return { segOf, users, receipts, rows };
    }

    // ===== セグメント分析 =====
    function computeSegReport(){
      if (!joined.length){
        $("segInfo").textContent="先にCSVを取り込んでください。";
        $("segKpi").innerHTML=""; $("segCatTable").innerHTML=""; $("segProdTable").innerHTML="";
        Plotly.purge("segTimeChart");
        $("segExport").disabled=true; return;
      }
      const s = $("segStart").value, e=$("segEnd").value;
      const checked = Array.from(document.querySelectorAll(".segChk:checked")).map(x=> x.value);
      if (!checked.length){
        $("segInfo").textContent="セグメントを1つ以上選択してください。";
        return;
      }

      const { segOf, rows } = buildSegmentationMap(s, e, 10);
      if (!rows.length){
        $("segInfo").textContent="該当データがありません。";
        $("segKpi").innerHTML=""; $("segCatTable").innerHTML=""; $("segProdTable").innerHTML="";
        Plotly.purge("segTimeChart");
        $("segExport").disabled=true; return;
      }

      const targetUsers = new Set();
      for (const [cgid, seg] of segOf.entries()){
        if (checked.includes(seg)) targetUsers.add(cgid);
      }
      const targetRows = rows.filter(r=> targetUsers.has(r.CGID));
      if (!targetRows.length){
        $("segInfo").textContent="選択したセグメントに該当する購買がありません。";
        $("segKpi").innerHTML=""; $("segCatTable").innerHTML=""; $("segProdTable").innerHTML="";
        Plotly.purge("segTimeChart");
        $("segExport").disabled=true; return;
      }

      const uu = new Set(targetRows.map(r=>r.CGID)).size;
      const rc = new Set(targetRows.map(r=>r.オーダーID)).size;
      const qty = targetRows.reduce((s,x)=> s+(x.商品数||0), 0);
      const amt = targetRows.reduce((s,x)=> s+(x.明細金額||0), 0);

      $("segKpi").innerHTML = `
        <div class="metric">ユニークユーザー数: <span class="kpi-strong">${uu.toLocaleString()}</span></div>
        <div class="metric">レシート数合計: <span class="kpi-strong">${rc.toLocaleString()}</span></div>
        <div class="metric">レシート/ユーザー: <span class="kpi-strong">${(uu? (rc/uu):0).toFixed(2)}</span></div>
        <div class="metric">購買点数合計: <span class="kpi-strong">${Math.round(qty).toLocaleString()}</span></div>
        <div class="metric">点数/ユーザー: <span class="kpi-strong">${(uu? (qty/uu):0).toFixed(2)}</span></div>
        <div class="metric">購買額合計: <span class="kpi-strong">${Math.round(amt).toLocaleString()}</span></div>
        <div class="metric">購買額/ユーザー: <span class="kpi-strong">${(uu? (amt/uu):0).toFixed(2)}</span></div>
      `;
      $("segInfo").textContent = `期間: ${s || "-"} ~ ${e || "-"} / セグメント: ${checked.join(", ")}`;

      // 商品 Top100
      const byJan = new Map();
      for (const r of targetRows){
        const k=r.JANコード||"-";
        if (!byJan.has(k)) byJan.set(k,{ JANコード:k, 商品名:r.商品名||"", 小分類名:r.小分類名||"", 金額:0, 点数:0, _users:new Set() });
        const o=byJan.get(k);
        o.金額 += (r.明細金額||0);
        o.点数 += (r.商品数||0);
        if (r.CGID) o._users.add(r.CGID);
      }
      const prodList = Array.from(byJan.values()).map(o=>({
        JANコード:o.JANコード, 商品名:o.商品名, 小分類名:o.小分類名,
        売上金額合計: Math.round(o.金額), 売上点数合計: Math.round(o.点数), ユーザー数:o._users.size
      })).sort((a,b)=> b.売上金額合計 - a.売上金額合計).slice(0,100);
      renderTable(prodList, "segProdTable", ["JANコード","商品名","小分類名","売上金額合計","売上点数合計","ユーザー数"], null);

      // カテゴリ Top20（小分類名ベース）
      const byCat = new Map();
      for (const r of targetRows){
        const k=r.小分類名 || (r.商品分類 || "-");
        if (!byCat.has(k)) byCat.set(k,{ カテゴリ:k, 金額:0, 点数:0, _users:new Set() });
        const o=byCat.get(k);
        o.金額 += (r.明細金額||0);
        o.点数 += (r.商品数||0);
        if (r.CGID) o._users.add(r.CGID);
      }
      const catList = Array.from(byCat.values()).map(o=>({
        カテゴリ:o.カテゴリ, 売上金額合計: Math.round(o.金額), 売上点数合計: Math.round(o.点数), ユーザー数:o._users.size
      })).sort((a,b)=> b.売上金額合計 - a.売上金額合計).slice(0,20);
      renderTable(catList, "segCatTable", ["カテゴリ","売上金額合計","売上点数合計","ユーザー数"], null);

      // 時間帯線グラフ（売上金額）
      const buckets = Array.from({length:24}, (_,h)=>({h, amt:0}));
      for (const r of targetRows){
        const h=r.生成時間.getHours();
        buckets[h].amt += (r.明細金額||0);
      }
      Plotly.newPlot("segTimeChart", [{
        type:"scatter", mode:"lines+markers",
        x: buckets.map(b=> String(b.h).padStart(2,"0")+":00"),
        y: buckets.map(b=> Math.round(b.amt)),
        name:"売上金額"
      }], {
        margin:{l:50,r:10,t:10,b:40},
        xaxis:{title:"時間帯（時）"},
        yaxis:{title:"売上金額", rangemode:"tozero"},
        displayModeBar:false
      }, {responsive:true});

      lastSegMeta = { start:s||"", end:e||"", segs:checked.slice() };
      lastSegTime = buckets.map(b=>({ 時間帯: String(b.h).padStart(2,"0")+":00", 売上金額: Math.round(b.amt) }));
      lastSegProd = prodList.slice();
      lastSegCat  = catList.slice();
      $("segExport").disabled = false;
    }
    function exportSegCSV(){
      if (!lastSegMeta) return;
      const ymd=(s)=> s ? s.replaceAll("-","/") : "-";
      const head = `${ymd(lastSegMeta.start)} ~ ${ymd(lastSegMeta.end)} セグメント: ${lastSegMeta.segs.join(" / ")}`;
      // 3シート風にまとめて1CSV（見出し行でブロック分け）
      const lines = ["\uFEFF"+head];

      // KPI は値のみ簡易
      // 省略（表示のみ）—必要であればここに追加可能

      // 時間帯
      lines.push("時間帯別,");
      lines.push(["時間帯","売上金額"].join(","));
      lastSegTime.forEach(r=>{
        const row=["時間帯","売上金額"].map(h=> r[h]);
        lines.push(row.join(","));
      });

      // カテゴリ
      lines.push("");
      lines.push("カテゴリTop20,");
      const catHeaders=["カテゴリ","売上金額合計","売上点数合計","ユーザー数"];
      lines.push(catHeaders.join(","));
      lastSegCat.forEach(r=>{
        const row=catHeaders.map(h=> String(r[h]).replaceAll('"','""'));
        lines.push(row.join(","));
      });

      // 商品
      lines.push("");
      lines.push("商品Top100,");
      const prodHeaders=["JANコード","商品名","小分類名","売上金額合計","売上点数合計","ユーザー数"];
      lines.push(prodHeaders.join(","));
      lastSegProd.forEach(r=>{
        const row=prodHeaders.map(h=>{ let v=String(r[h]??"").replaceAll('"','""'); if (/[",\n]/.test(v)) v=`"${v}"`; return v; });
        lines.push(row.join(","));
      });

      const csv = lines.join("\n");
      const fn = `${lastSegMeta.start||"all"}_${lastSegMeta.end||"all"}_segment.csv`;
      const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob);
      const a=document.createElement("a"); a.href=url; a.download=fn; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ===== リピート分析 =====
    function computeRepeatReport(){
      if (!joined.length){
        $("repeatInfo").textContent="先にCSVを取り込んでください。";
        $("repeatKpi").innerHTML=""; $("repeatTable").innerHTML=""; $("repeatExport").disabled=true;
        return;
      }
      const s=$("repeatStart").value, e=$("repeatEnd").value;
      const checked = Array.from(document.querySelectorAll(".repeatChk:checked")).map(x=> x.value);
      if (!checked.length){
        $("repeatInfo").textContent="セグメントを1つ以上選択してください。";
        return;
      }

      const { segOf, rows } = buildSegmentationMap(s, e, 10);
      const targetUsers = new Set();
      for (const [cgid, seg] of segOf.entries()){
        if (checked.includes(seg)) targetUsers.add(cgid);
      }
      const targetRows = rows.filter(r=> targetUsers.has(r.CGID));
      if (!targetRows.length){
        $("repeatInfo").textContent="選択したセグメントに該当する購買がありません。";
        $("repeatKpi").innerHTML=""; $("repeatTable").innerHTML=""; $("repeatExport").disabled=true; return;
      }

      // ユーザー×JAN の購入回数
      const userJanCnt = new Map(); // key: CGID||JAN -> count
      const byJan = new Map();      // JAN集計
      for (const r of targetRows){
        const k = `${r.CGID}||${r.JANコード}`;
        userJanCnt.set(k, (userJanCnt.get(k)||0) + (r.商品数||0));

        if (!byJan.has(r.JANコード)){
          byJan.set(r.JANコード, { JANコード:r.JANコード, 商品名:r.商品名||"", 小分類名:r.小分類名||"", ユーザー: new Map(), 金額:0, 点数:0 });
        }
        const o = byJan.get(r.JANコード);
        o.金額 += (r.明細金額||0);
        o.点数 += (r.商品数||0);
        o.ユーザー.set(r.CGID, (o.ユーザー.get(r.CGID)||0) + (r.商品数||0));
      }

      // リピート定義：同一JANを2回以上購入したユーザー
      const list=[];
      for (const o of byJan.values()){
        let repeatUsers=0;
        for (const cnt of o.ユーザー.values()){
          if (cnt>=2) repeatUsers++;
        }
        list.push({
          JANコード:o.JANコード, 商品名:o.商品名, 小分類名:o.小分類名,
          リピートユーザー数: repeatUsers,
          売上点数合計: Math.round(o.点数),
          売上金額合計: Math.round(o.金額)
        });
      }
      list.sort((a,b)=> b.リピートユーザー数 - a.リピートユーザー数 || b.売上金額合計 - a.売上金額合計);
      const top = list.slice(0,30);
      renderTable(top, "repeatTable", ["JANコード","商品名","小分類名","リピートユーザー数","売上点数合計","売上金額合計"], null);

      // 1ユーザーあたりの平均リピート回数（>=2のみ/全体の両方を表示）
      const userRepeatCounts = new Map(); // CGID -> 合計(各JANの max(0, cnt-1))
      for (const [key,cnt] of userJanCnt.entries()){
        const [cgid] = key.split("||");
        if (!userRepeatCounts.has(cgid)) userRepeatCounts.set(cgid,0);
        if (cnt>=2) userRepeatCounts.set(cgid, userRepeatCounts.get(cgid) + (cnt-1));
      }
      const uuAll = new Set(targetRows.map(r=> r.CGID)).size;
      const usersWithRepeat = Array.from(userRepeatCounts.values()).filter(v=> v>0);
      const avgAll = uuAll ? (Array.from(userRepeatCounts.values()).reduce((a,b)=>a+b,0) / uuAll) : 0;
      const avgAmongRepeaters = usersWithRepeat.length ? (usersWithRepeat.reduce((a,b)=>a+b,0) / usersWithRepeat.length) : 0;

      $("repeatKpi").innerHTML = `
        <div class="metric">ユニークユーザー数: <span class="kpi-strong">${uuAll.toLocaleString()}</span></div>
        <div class="metric">リピート経験者数: <span class="kpi-strong">${usersWithRepeat.length.toLocaleString()}</span></div>
        <div class="metric">平均リピート回数（全ユーザー）: <span class="kpi-strong">${avgAll.toFixed(2)}</span></div>
        <div class="metric">平均リピート回数（リピート者）: <span class="kpi-strong">${avgAmongRepeaters.toFixed(2)}</span></div>
      `;
      $("repeatInfo").textContent = `期間: ${s || "-"} ~ ${e || "-"} / セグメント: ${checked.join(", ")}`;

      lastRepeatRows = top;
      lastRepeatMeta = { start:s||"", end:e||"", segs:checked.slice() };
      $("repeatExport").disabled = !lastRepeatRows.length;
    }
    function exportRepeatCSV(){
      if (!lastRepeatRows || !lastRepeatRows.length) return;
      const headers = ["JANコード","商品名","小分類名","リピートユーザー数","売上点数合計","売上金額合計"];
      const ymd=(s)=> s ? s.replaceAll("-","/") : "-";
      const preface = `${ymd(lastRepeatMeta.start)} ~ ${ymd(lastRepeatMeta.end)} リピート上位（セグメント: ${lastRepeatMeta.segs.join(" / ")}）`;
      const lines=[preface, headers.join(",")];
      for (const r of lastRepeatRows){
        const row=headers.map(h=>{ let v=r[h]; if (v==null) v=""; v=String(v).replaceAll('"','""'); if (/[",\n]/.test(v)) v=`"${v}"`; return v; });
        lines.push(row.join(","));
      }
      const csv="\uFEFF"+lines.join("\n");
      const fn = `${lastRepeatMeta.start||"all"}_${lastRepeatMeta.end||"all"}_repeat.csv`;
      const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob);
      const a=document.createElement("a"); a.href=url; a.download=fn; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ===== 併売分析 =====
    function computePairReport(){
      if (!joined.length){
        $("pairInfo").textContent="先にCSVを取り込んでください。";
        $("pairTable").innerHTML=""; $("pairExport").disabled=true; return;
      }
      const s=$("pairStart").value, e=$("pairEnd").value;
      const checked = Array.from(document.querySelectorAll(".pairChk:checked")).map(x=> x.value);
      if (!checked.length){
        $("pairInfo").textContent="セグメントを1つ以上選択してください。";
        return;
      }
      const minSupp = Math.max(1, Math.trunc(toNumber($("pairMinSupp").value||"5",5)));
      const sortKey = $("pairSortKey").value;

      const { segOf, rows } = buildSegmentationMap(s, e, 10);
      const targetUsers = new Set();
      for (const [cgid, seg] of segOf.entries()){
        if (checked.includes(seg)) targetUsers.add(cgid);
      }
      const targetRows = rows.filter(r=> targetUsers.has(r.CGID));
      if (!targetRows.length){
        $("pairInfo").textContent="選択したセグメントに該当する購買がありません。";
        $("pairTable").innerHTML=""; $("pairExport").disabled=true; return;
      }

      // レシート -> ユニークJAN集合
      const recToSet = new Map();
      for (const r of targetRows){
        const key=r.オーダーID;
        if (!recToSet.has(key)) recToSet.set(key, new Set());
        recToSet.get(key).add(r.JANコード);
      }
      const recCount = recToSet.size;

      // 単品出現回数
      const itemCount = new Map(); // JAN -> 出現レシート数
      recToSet.forEach(set=>{
        for (const jan of set){
          itemCount.set(jan, (itemCount.get(jan)||0)+1);
        }
      });

      // ペア出現回数（同一レシート内の組合せ）
      const pairCount = new Map(); // "A||B" (A<B) -> 出現レシート数
      recToSet.forEach(set=>{
        const arr = Array.from(set).sort();
        for (let i=0;i<arr.length;i++){
          for (let j=i+1;j<arr.length;j++){
            const a=arr[i], b=arr[j];
            const key=`${a}||${b}`;
            pairCount.set(key, (pairCount.get(key)||0)+1);
          }
        }
      });

      // 指標計算
      const rowsOut=[];
      for (const [key, supp] of pairCount.entries()){
        if (supp < minSupp) continue;
        const [a,b] = key.split("||");
        const cntA = itemCount.get(a)||1;
        const cntB = itemCount.get(b)||1;
        const confAB = supp / cntA;
        const confBA = supp / cntB;
        const lift = (supp * recCount) / (cntA * cntB);
        const aName = masterMap?.get(a)?.商品名 || (joined.find(r=> r.JANコード===a)?.商品名) || "";
        const bName = masterMap?.get(b)?.商品名 || (joined.find(r=> r.JANコード===b)?.商品名) || "";
        rowsOut.push({
          商品A_JAN:a, 商品A名:aName,
          商品B_JAN:b, 商品B名:bName,
          同時購買レシート数: supp,
          リフト: +lift.toFixed(3),
          信頼度AtoB: +confAB.toFixed(3),
          信頼度BtoA: +confBA.toFixed(3)
        });
      }

      // 並び順
      if (sortKey==="support"){
        rowsOut.sort((x,y)=> y.同時購買レシート数 - x.同時購買レシート数 || y.リフト - x.リフト);
      }else if (sortKey==="conf"){
        rowsOut.sort((x,y)=> Math.max(y.信頼度AtoB,y.信頼度BtoA) - Math.max(x.信頼度AtoB,x.信頼度BtoA) || y.同時購買レシート数 - x.同時購買レシート数);
      }else{ // lift
        rowsOut.sort((x,y)=> y.リフト - x.リフト || y.同時購買レシート数 - x.同時購買レシート数);
      }

      $("pairInfo").textContent = `期間: ${s || "-"} ~ ${e || "-"} / セグメント: ${checked.join(", ")} / レシート: ${recCount.toLocaleString()} / 最小同時購買数: ${minSupp}`;
      renderTable(rowsOut, "pairTable", ["商品A_JAN","商品A名","商品B_JAN","商品B名","同時購買レシート数","リフト","信頼度AtoB","信頼度BtoA"], null);

      lastPairRows = rowsOut;
      lastPairMeta = { start:s||"", end:e||"", segs:checked.slice(), minSupp, sortKey, recCount };
      $("pairExport").disabled = rowsOut.length===0;
    }
    function exportPairCSV(){
      if (!lastPairRows || !lastPairRows.length) return;
      const headers = ["商品A_JAN","商品A名","商品B_JAN","商品B名","同時購買レシート数","リフト","信頼度AtoB","信頼度BtoA"];
      const ymd=(s)=> s ? s.replaceAll("-","/") : "-";
      const preface = `${ymd(lastPairMeta.start)} ~ ${ymd(lastPairMeta.end)} 併売（セグメント: ${lastPairMeta.segs.join(" / ")} / minSupp=${lastPairMeta.minSupp} / sort=${lastPairMeta.sortKey})`;
      const lines=[preface, headers.join(",")];
      for (const r of lastPairRows){
        const row=headers.map(h=>{ let v=r[h]; if (v==null) v=""; v=String(v).replaceAll('"','""'); if (/[",\n]/.test(v)) v=`"${v}"`; return v; });
        lines.push(row.join(","));
      }
      const csv="\uFEFF"+lines.join("\n");
      const fn = `${lastPairMeta.start||"all"}_${lastPairMeta.end||"all"}_pair.csv`;
      const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob);
      const a=document.createElement("a"); a.href=url; a.download=fn; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // ===== UI配線 =====
    function updateButton(){
      const ok = (els.txInput?.files?.length===1 && els.lnInput?.files?.length===1);
      if (els.runBtn) els.runBtn.disabled = !ok;
      if (els.status){
        els.status.className = "muted";
        els.status.textContent = ok ? "準備OK。取り込み＆集計できます。" : "CSVを2つ選択してください";
      }
      $("fnameTx").textContent = els.txInput?.files?.[0]?.name || "未選択";
      $("fnameTx").classList.toggle("muted", !els.txInput?.files?.length);
      $("fnameLn").textContent = els.lnInput?.files?.[0]?.name || "未選択";
      $("fnameLn").classList.toggle("muted", !els.lnInput?.files?.length);
      $("fnamePm").textContent = els.pmInput?.files?.[0]?.name || "未選択";
      $("fnamePm").classList.toggle("muted", !els.pmInput?.files?.length);
    }
    function onTabClick(t){
      document.querySelectorAll(".tab,.tabcontent").forEach(x=>x.classList.remove("active"));
      t.classList.add("active");
      const pane = document.getElementById(t.dataset.tab);
      if (pane) pane.classList.add("active");
    }
    function bindModals(){
      document.querySelectorAll(".info-btn").forEach(btn=>{
        btn.addEventListener("click", ()=>{ const id=btn.getAttribute("data-modal"); const m=$(id); if(m) m.style.display="block"; });
      });
      document.querySelectorAll(".modal .close").forEach(x=>{
        x.addEventListener("click", ()=>{ const id=x.getAttribute("data-close"); const m=$(id); if(m) m.style.display="none"; });
      });
      window.addEventListener("click", (e)=>{ if (e.target.classList.contains("modal")) e.target.style.display="none"; });
    }

    // ===== 初期化 =====
    document.addEventListener("DOMContentLoaded", async ()=>{
      els = {
        pickTx: $("pickTx"),
        pickLn: $("pickLn"),
        pickPm: $("pickPm"),
        txInput: $("txFile"),
        lnInput: $("lnFile"),
        pmInput: $("pmFile"),
        runBtn: $("runBtn"),
        status: $("status"),
        catStatus: $("catStatus"),
      };

      // ファイルピッカー
      els.pickTx?.addEventListener("click", (e)=>{ e.preventDefault(); e.stopPropagation(); els.txInput?.click(); });
      els.pickLn?.addEventListener("click", (e)=>{ e.preventDefault(); e.stopPropagation(); els.lnInput?.click(); });
      els.pickPm?.addEventListener("click", (e)=>{ e.preventDefault(); e.stopPropagation(); els.pmInput?.click(); });
      els.txInput?.addEventListener("change", updateButton);
      els.lnInput?.addEventListener("change", updateButton);
      els.pmInput?.addEventListener("change", updateButton);

      // タブ
      document.querySelectorAll(".tab").forEach(t=> t.addEventListener("click", ()=> onTabClick(t)));
      bindModals();

      // カテゴリマスタ
      await autoLoadCategory();

      // 取り込み
      els.runBtn?.addEventListener("click", runImport);

      // ベスレポ
      $("bestRun")?.addEventListener("click", computeBestReport);
      $("bestExport")?.addEventListener("click", exportBestCSV);

      // 時間帯
      $("timeRun")?.addEventListener("click", computeTimeReport);
      $("timeExport")?.addEventListener("click", exportTimeCSV);
      $("timePng")?.addEventListener("click", exportTimePNG);

      // 単品
      $("prodRun")?.addEventListener("click", computeProductReport);

      // RFM
      $("rfmRun")?.addEventListener("click", computeRFM);
      $("rfmExport")?.addEventListener("click", exportRFM);

      // セグメント
      $("segRun")?.addEventListener("click", computeSegReport);
      $("segExport")?.addEventListener("click", exportSegCSV);

      // リピート
      $("repeatRun")?.addEventListener("click", computeRepeatReport);
      $("repeatExport")?.addEventListener("click", exportRepeatCSV);

      // 併売
      $("pairRun")?.addEventListener("click", computePairReport);
      $("pairExport")?.addEventListener("click", exportPairCSV);

      // 初期UI
      updateButton();
    });
  