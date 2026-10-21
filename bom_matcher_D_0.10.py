#!/usr/bin/env python3
"""
bom_matcher_D_0.10.py  v0.10  —  D 案：pywebview 原生視窗版

啟動：
    run_D_webview.bat          （雙擊）
    python bom_matcher_D_0.10.py
    python bom_matcher_D_0.10.py --dump-html   # 只輸出 UI 的 HTML（供瀏覽器檢視）

工作目錄固定為「程式所在目錄」，不受啟動時的 cwd 影響。
"""

import sys
import json
import importlib.util
import threading
import subprocess
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent

_spec = importlib.util.spec_from_file_location(
    'bom_core', APP_DIR / 'bom_core_0.10.py')
core = importlib.util.module_from_spec(_spec)
sys.modules['bom_core'] = core
_spec.loader.exec_module(core)


# ════════════════════════════════════════════════════════════════════════════
# HTML / CSS / JS（單檔內嵌，方便發佈）
# ════════════════════════════════════════════════════════════════════════════

HTML = r"""
<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<title>BOM Matcher</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%}
body{font-family:"Inter","Segoe UI","Microsoft JhengHei",sans-serif;background:#F7F8FA;
     font-size:13px;color:#1f2430;-webkit-font-smoothing:antialiased;overflow:hidden}
.app{display:flex;height:100vh}
button{font-family:inherit;cursor:pointer}
input,select,textarea{font-family:inherit;font-size:inherit;color:inherit}

/* sidebar */
.sb{width:212px;background:#fff;border-right:1px solid #E8EAED;padding:18px 12px;
    flex:none;display:flex;flex-direction:column}
.brand{display:flex;align-items:center;gap:9px;padding:0 4px 14px}
.brand .m{width:30px;height:30px;border-radius:8px;background:#4F46E5;color:#fff;
          font-weight:700;display:flex;align-items:center;justify-content:center;font-size:14px}
.brand b{font-size:13.5px;display:block}
.brand span{font-size:11px;color:#8b90a0}
.grp{font-size:10px;letter-spacing:.1em;color:#9aa0ae;font-weight:600;padding:14px 8px 6px}
.nav{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:8px;
     font-size:13px;color:#3c4252;margin-bottom:2px;border:none;background:none;width:100%;
     text-align:left}
.nav:hover{background:#F2F4F7}
.nav .ic{width:15px;text-align:center;opacity:.7}
.nav.on{background:#4F46E5;color:#fff;font-weight:600}
.nav .bdg{margin-left:auto;font-size:10.5px;background:#EEF0F4;color:#5b6070;
          padding:1px 6px;border-radius:9px}
.nav.on .bdg{background:rgba(255,255,255,.22);color:#fff}
.sbfoot{margin-top:auto;font-size:10.5px;color:#9aa0ae;padding:8px 6px;line-height:1.6;
        word-break:break-all}

/* content */
.ct{flex:1;padding:20px 24px 0;min-width:0;display:flex;flex-direction:column;overflow:hidden}
.ct-hd{display:flex;align-items:flex-end;margin-bottom:14px;flex:none}
.ct-hd h2{font-size:20px;font-weight:700;letter-spacing:-.01em}
.ct-hd p{font-size:12px;color:#8b90a0;margin-top:2px}
.ct-hd .rt{margin-left:auto;display:flex;gap:8px}

.btn{border:1px solid #DDE0E6;background:#fff;border-radius:8px;padding:7px 14px;
     font-size:12.5px;color:#3c4252;font-weight:500}
.btn:hover{background:#F7F8FA}
.btn.p{background:#4F46E5;border-color:#4F46E5;color:#fff;font-weight:600}
.btn.p:hover{background:#4238d6}
.btn.g{background:#0f9d58;border-color:#0f9d58;color:#fff;font-weight:600}
.btn:disabled{opacity:.45;cursor:default}

.card{background:#fff;border:1px solid #E8EAED;border-radius:12px;
      box-shadow:0 1px 2px rgba(16,24,40,.04)}
.cards{display:flex;gap:12px;margin-bottom:12px;flex:none}
.stat{flex:1;padding:11px 14px}
.stat .k{font-size:11px;color:#8b90a0}
.stat .v{font-size:23px;font-weight:700;margin-top:2px;letter-spacing:-.02em}
.stat .s{font-size:11px;margin-top:1px;color:#0f9d58}
.stat .s.r{color:#e5484d}
.stat .s.o{color:#c77700}

.split{display:flex;gap:12px;flex:1;min-height:0}
.left{flex:1;min-width:0;display:flex;flex-direction:column}
.right{width:300px;flex:none;display:flex;flex-direction:column;gap:12px;overflow:auto}

.card-hd{padding:10px 14px;border-bottom:1px solid #EEF0F4;display:flex;
         align-items:center;gap:8px;flex:none;flex-wrap:wrap}
.card-hd b{font-size:13px;white-space:nowrap}
.dcell{display:flex;align-items:center;gap:0}
.card-hd .sub{font-size:11px;color:#9aa0ae;white-space:nowrap;flex:none}
.card-hd .rt{margin-left:auto;display:flex;gap:6px;align-items:center;flex-wrap:wrap;
             justify-content:flex-end}
.chip{font-size:10.5px;padding:3px 8px;border-radius:9px;border:1px solid #E2E5EA;
      color:#5b6070;background:#fff;white-space:nowrap;flex:none}
.chip.on{background:#EEF0FF;border-color:#c9cdf5;color:#4F46E5;font-weight:600}
.search{border:1px solid #E2E5EA;border-radius:7px;padding:4px 9px;font-size:11.5px;
        width:150px;flex:none}

.tb{overflow:auto;flex:1}
table{border-collapse:separate;border-spacing:0;width:100%;font-size:12px}
th{position:sticky;top:0;background:#FAFBFC;text-align:left;padding:8px 10px;font-size:11px;
   font-weight:600;color:#6b7180;border-bottom:1px solid #E8EAED;white-space:nowrap;z-index:2}
td{padding:5px 10px;border-bottom:1px solid #F2F4F7;vertical-align:middle}
td.c{text-align:center}
tr.r-nc  td{background:#FFFBEF}
tr.r-alt td{background:#F3F9F1}
tr.r-kn  td{background:#FFF6EC}
tr.r-bad td{background:#FFF1F1}
tr.r-chk td{background:#F5F3FF}
.unkmsg{color:#c74343;font-size:10px;font-weight:700;margin-top:2px}
tr.sel td{background:#F1F2FF!important;box-shadow:inset 2px 0 0 #4F46E5}
tr:hover td{filter:brightness(.985)}
.pill{font-size:10px;font-weight:700;padding:1.5px 7px;border-radius:5px}
.pill.rc{background:#E7F1FE;color:#1a63c7}
.pill.ec{background:#F0F0F4;color:#5b6070}
.cellin{border:1px solid transparent;background:transparent;border-radius:5px;
        padding:2px 6px;width:100%;font-size:11.5px}
.cellin:hover{border-color:#E2E5EA;background:#fff}
.cellin:focus{outline:none;border-color:#4F46E5;background:#fff}
.mono{font-family:Consolas,monospace}
.miss{color:#c74343;font-style:italic}
.dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px}
.f-nc{background:#e8b931}.f-alt{background:#5aae62}.f-kn{background:#e8863a}
.f-chk{background:#D92D20}
.legend b{color:#3c4252}

.kv{padding:10px 14px;font-size:12px}
.kv .r{display:flex;justify-content:space-between;align-items:center;padding:4px 0;
       border-bottom:1px dashed #F0F2F5;gap:10px}
.kv .r:last-child{border:none}
.kv .k{color:#8b90a0;flex:none}
.kv .v{font-weight:600;text-align:right;word-break:break-all}
.kv input,.kv select{border:1px solid #E2E5EA;border-radius:6px;padding:3px 8px;
                     text-align:right;min-width:120px;max-width:170px}
.note{padding:11px 14px;font-size:12px;line-height:1.6}
.note .src{font-size:10.5px;color:#9aa0ae;margin-bottom:5px}
.note .body{background:#FFF6EC;border-left:3px solid #e8863a;border-radius:5px;
            padding:9px 11px;color:#6b4413;white-space:pre-wrap}
.note .act{display:flex;gap:7px;margin-top:9px}
.console{background:#171B24;border-radius:12px;padding:10px 13px;font-family:Consolas,monospace;
         font-size:11px;color:#c8cede;line-height:1.7;min-height:96px;max-height:190px;
         overflow:auto;white-space:pre-wrap}
.console .ok{color:#4ec97b}.console .wn{color:#e8b931}.console .er{color:#ff7b72}

.foot{border-top:1px solid #E8EAED;background:#fff;margin:12px -24px 0;padding:10px 24px;
      display:flex;align-items:center;gap:12px;flex:none}
.foot .path{font-size:11.5px;color:#6b7180;font-family:Consolas,monospace}
.foot .rt{margin-left:auto;display:flex;gap:8px;align-items:center}
.legend{display:flex;gap:12px;font-size:11px;color:#6b7180;align-items:center}
.legend span{white-space:nowrap}

.page{display:none;flex:1;min-height:0;flex-direction:column}
.page.on{display:flex}
.form{padding:16px 18px;display:grid;grid-template-columns:150px 1fr auto;gap:10px 12px;
      align-items:center}
.form label{font-size:12px;color:#5b6070;text-align:right}
.form input,.form select{border:1px solid #E2E5EA;border-radius:8px;padding:7px 10px;width:100%}
.form .hint{grid-column:2/4;font-size:11px;color:#9aa0ae;margin-top:-6px}
.banner{background:#FFF6EC;border:1px solid #f0d3ae;color:#6b4413;border-radius:8px;
        padding:8px 12px;font-size:12px;margin-bottom:12px}
.empty{padding:34px;text-align:center;color:#9aa0ae;font-size:13px}
.switch{display:flex;align-items:center;gap:8px;font-size:12px;color:#3c4252}
.switch.wide{grid-column:2/4;justify-content:flex-start}
/* .form input{width:100%} 連 checkbox 都套到，勾選框被撐成整行寬，
   看起來就變成「敘述靠左、勾選置中、說明靠右」。 */
.form input[type=checkbox]{width:auto;flex:none;margin:0}
.switch.wide span{color:#9aa0ae}

/* v0.10：ADATA LOGO（比照 Firmware Tool）。蜂鳥圖為橫向白底，用 contain 等比縮放。
   ⚠️ -webkit-user-drag / user-select 不是美化 —— 沒有它們，按住 Shift 拖圖會啟動
      瀏覽器原生的圖片拖曳，pointerdown 之後的連點就收不到。 */
.brand img.m{width:44px;height:30px;object-fit:contain;background:#fff;
             border:1px solid #E8EAED;border-radius:7px;padding:2px;
             -webkit-user-drag:none;user-select:none}
.egg{position:fixed;inset:0;background:rgba(15,23,42,.55);display:flex;
     align-items:center;justify-content:center;z-index:99}
.egg[hidden]{display:none}
.egg-box{background:#0b1220;padding:10px;border-radius:14px;
         box-shadow:0 20px 50px rgba(0,0,0,.4)}
.egg-hd{display:flex;justify-content:space-between;align-items:center;margin:0 2px 6px;
        font:600 11px Consolas,monospace;letter-spacing:2px;color:rgba(226,232,240,.55)}
.egg-hd button{background:none;border:none;color:rgba(226,232,240,.75);font-size:15px;
               cursor:pointer;padding:0 4px}
.egg canvas{display:block;border-radius:8px}
</style></head>
<body>
<div class="app">
  <aside class="sb">
    <div class="brand">__LOGO_TAG__
      <div><b>BOM Matcher</b><span id="ver">E-BOM workspace</span></div></div>
    <div class="grp">WORKFLOW</div>
    <button class="nav on" data-p="files"><span class="ic">▤</span>檔案設定<span class="bdg" id="bd-files">-</span></button>
    <button class="nav" data-p="rows"><span class="ic">◇</span>比對結果<span class="bdg" id="bd-rows">-</span></button>
    <button class="nav" data-p="kn"><span class="ic">◈</span>Knowledge 命中<span class="bdg" id="bd-kn">-</span></button>
    <button class="nav" data-p="hdr"><span class="ic">▦</span>表頭資訊</button>
    <button class="nav" data-p="out"><span class="ic">↓</span>產出 E-BOM</button>
    <div class="sbfoot" id="sbfoot"></div>
  </aside>
  <!-- v0.10 彩蛋：沒有開關、沒有提示（比照 Firmware Tool）。✕ 與 Esc 都能關。 -->
  <div class="egg" id="egg" hidden onpointerdown="if(event.target===this) eggClose()">
    <div class="egg-box"><div class="egg-hd"><span>BEE</span>
      <button onclick="eggClose()" title="關閉">✕</button></div>
      <canvas id="eggcv" width="420" height="520"></canvas></div></div>

  <section class="ct">
    <div class="ct-hd">
      <div><h2 id="ttl">檔案設定</h2><p id="sub"></p></div>
      <div class="rt" id="hdbtn"></div>
    </div>
    <div id="banner"></div>

    <!-- ── 檔案設定 ── -->
    <div class="page on" id="p-files">
      <div class="card"><div class="form" id="fileform"></div></div>
      <div style="height:12px"></div>
      <div class="card console" id="log1">（尚未執行）</div>
    </div>

    <!-- ── 比對結果 ── -->
    <div class="page" id="p-rows">
      <div class="cards" id="stats"></div>
      <div class="split">
        <div class="left"><div class="card" style="flex:1;display:flex;flex-direction:column;overflow:hidden">
          <div class="card-hd"><b>BOM 明細</b><span class="sub">點列可在右側檢視 · 欄位可直接編輯</span>
            <div class="rt" id="filters"></div></div>
          <div class="tb"><table id="tbl"><thead><tr>
            <th style="width:44px">Item</th><th style="width:58px">RC/EC</th>
            <th style="width:150px">P/N</th><th>Description</th>
            <th style="width:52px">Qty</th><th style="width:150px" id="thTop">Location-Top</th>
            <th style="width:150px;display:none" id="thBot">Location-Bot</th>
          </tr></thead><tbody id="tbody"></tbody></table></div>
        </div></div>
        <div class="right" id="detail"></div>
      </div>
      <div class="foot">
        <div class="legend">
          <span><span class="dot f-nc"></span>原本為 NC</span>
          <span><span class="dot f-alt"></span>含替代料</span>
          <span><span class="dot f-kn"></span>Knowledge 命中</span>
          <span id="lgChk" style="display:none"><span class="dot f-chk"></span>
            <b>R8 降壓選料</b></span>
          <span id="lgBad" style="display:none">&#9888;&#65039; Location 未知</span>
        </div>
        <div class="rt"><span class="path" id="outpath2"></span>
          <button class="btn p" onclick="doExport()">▶ 產出 E-BOM</button></div>
      </div>
    </div>

    <!-- ── Knowledge ── -->
    <div class="page" id="p-kn"><div style="overflow:auto;flex:1" id="knlist"></div></div>

    <!-- ── 表頭 ── -->
    <div class="page" id="p-hdr">
      <div class="card"><div class="form" id="hdrform"></div></div>
    </div>

    <!-- ── 產出 ── -->
    <div class="page" id="p-out">
      <div class="card" style="padding:16px 18px">
        <div style="font-size:12px;color:#8b90a0;margin-bottom:6px">輸出路徑</div>
        <div style="display:flex;gap:8px">
          <input id="outpath" class="cellin mono"
                 style="border:1px solid #E2E5EA;padding:7px 10px;border-radius:8px;font-size:12px">
          <button class="btn" onclick="browse('out')">變更…</button>
        </div>
        <div id="outinfo" style="font-size:12px;color:#8b90a0;margin-top:10px"></div>
        <div style="margin-top:14px;display:flex;gap:8px">
          <button class="btn p" onclick="doExport()">▶ 產出 Excel</button>
          <button class="btn" onclick="openOut()">開啟輸出資料夾</button>
        </div>
      </div>
      <div style="height:12px"></div>
      <div class="card console" id="log2">（尚未產出）</div>
    </div>
  </section>
</div>

<script>
let S = {rows:[], stats:{}, disc:{}, cfg:{}, hdr:{}, log:[], sel:0, filter:'all', q:''};
/* window.pywebview 是在頁面載入後才注入的，不能在腳本執行當下就判定，
   否則會永遠顯示「預覽模式」橫幅。改用即時檢查 + 延遲顯示橫幅。 */
const LIVE = () => !!(window.pywebview && window.pywebview.api);
const api = () => window.pywebview.api;

const esc = s => String(s==null?'':s).replace(/[&<>"]/g,
  c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

function page(p){
  document.querySelectorAll('.nav').forEach(b=>b.classList.toggle('on', b.dataset.p===p));
  document.querySelectorAll('.page').forEach(e=>e.classList.toggle('on', e.id==='p-'+p));
  const T={files:['檔案設定','預設值全部來自程式所在目錄及其 TXT\\ · ExcelDB\\ · screenshot\\ · output\\ 子資料夾'],
           rows:['比對結果','點任一列可在右側檢視與修正；底色與輸出 Excel 一致'],
           kn:['Knowledge 命中','逐筆確認；忽略後該列不標橘底，Note 也不寫入 Excel'],
           hdr:['表頭資訊','填入 Excel rows 3–8 的欄位，留空則維持空白'],
           out:['產出 E-BOM','']};
  document.getElementById('ttl').textContent = T[p][0];
  document.getElementById('sub').textContent = T[p][1];
  const hb = document.getElementById('hdbtn');
  hb.innerHTML = (p==='rows'||p==='files')
    ? '<button class="btn" onclick="doMatch()">↺ 重新比對</button>'
      + '<button class="btn p" onclick="doExport()">▶ 產出 E-BOM</button>' : '';
  if(p==='kn') renderKn();
  if(p==='hdr') renderHdr();
  if(p==='out') renderOut();
}
document.querySelectorAll('.nav').forEach(b=>b.onclick=()=>page(b.dataset.p));

function renderFiles(){
  const d=S.disc, c=S.cfg;
  const sel=(id,opts,val)=>`<select id="${id}" onchange="setCfg('${id}',this.value)">`
    + opts.map(o=>`<option value="${esc(o)}" ${o===val?'selected':''}>${esc(short(o))}</option>`).join('')
    + (val && !opts.includes(val) ? `<option selected value="${esc(val)}">${esc(short(val))}</option>`:'')
    + (opts.length?'':'<option value="">（未找到）</option>') + '</select>';
  document.getElementById('fileform').innerHTML =
     `<label>Text BOM</label>${sel('bom',d.txts,c.bom)}<button class="btn" onclick="browse('bom')">瀏覽…</button>
      <div class="hint">只列出內容含 Altium BOM 表頭的 .txt（自動排除 changelog 之類的檔案）</div>
      <label>零件資料庫</label>${sel('db',d.dbs,c.db)}<button class="btn" onclick="browse('db')">瀏覽…</button>
      <div class="hint">只列出第 1 列有 P/N + Description 的 xlsx（自動排除 E-BOM 產出檔）</div>
      <label>Knowledge DB</label>${sel('kn',d.kns,c.kn)}<button class="btn" onclick="browse('kn')">瀏覽…</button>
      <div class="hint">只列出第 1 列有 品名 + 包裝 的 xlsx</div>
      <label>PCB 圖片<br><span style="font-size:10px;color:#9aa0ae">單面板／TOP 面</span></label>${sel('pcb',d.imgs,c.pcb)}<button class="btn" onclick="browse('pcb')">瀏覽…</button>
      <label>Location TOP/BOT 對照表</label>${sel('sides',d.sides_all||[],c.sides)}<button class="btn" onclick="browse('sides')">瀏覽…</button>
      <div class="hint">選填。給了就切換成雙面板模式，輸出多一欄 Location-Bottom Layer。xlsx 為主、txt 為輔</div>
      <label>PCB 圖片（BOT 面）</label>${sel('pcb_bot',d.imgs,c.pcb_bot)}<button class="btn" onclick="browse('pcb_bot')">瀏覽…</button>
      <div class="hint">雙面板時提供兩張圖，圖片區會自動切成上下兩半。檔名含 top / bot 只是自動偵測的規則，手動指定任何檔案都可以；Top 與 Bot 選到同一張時會提醒確認</div>
      <label>輸出路徑</label><input id="outp" value="${esc(c.out||'')}"
             onchange="setCfg('out',this.value)"><button class="btn" onclick="browse('out')">變更…</button>
      <label></label>
      <div class="switch wide"><input type="checkbox" id="split" ${c.split?'checked':''}
           onchange="setCfg('split',this.checked)"><b>替代料拆列</b>
           <span>一列只有一個 P/N（關閉則回到 v0.06 的 a;b 單列）</span></div>
      <label></label><div><button class="btn p" onclick="doMatch()">▶ 執行比對</button>
        <button class="btn" onclick="rescan()">↺ 重新掃描目錄</button></div><div></div>`;
}
function short(p){ if(!p) return '（無）';
  const b=S.disc.base||''; return p.startsWith(b) ? p.slice(b.length).replace(/^[\\/]/,'') : p; }

function renderStats(){
  const lb=document.getElementById('lgBad');
  if(lb) lb.style.display = S.double ? '' : 'none';
  const lc=document.getElementById('lgChk');
  if(lc) lc.style.display = S.rows.some(r=>r.chk) ? '' : 'none';
  const s=S.stats; if(!s.items){document.getElementById('stats').innerHTML='';return;}
  const box=(k,v,note,cls)=>`<div class="card stat"><div class="k">${k}</div>
      <div class="v">${v}</div><div class="s ${cls||''}">${note}</div></div>`;
  document.getElementById('stats').innerHTML =
    box('總項次',s.items,'含 PCB 固定末列')
  + box('輸出列數',s.rows,'替代料 +'+(s.rows-s.items))
  + box('命中 P/N',s.hit,Math.round(s.hit/Math.max(1,s.items-1)*100)+'%')
  + box('未命中',s.miss,'需人工補料號','r')
  + box('Knowledge',s.kn,'橘底標示','o')
  + box('Total Qty',s.qty,'主列加總');
}

const FILTERS=[['all','全部'],['miss','未命中'],['nc','NC'],['alt','替代料'],
               ['kn','Knowledge'],['chk','需確認'],['unk','Location 未知']];
function keep(r){
  if(S.q){ const q=S.q.toLowerCase();
    if(!((r.pn||'')+' '+(r.desc||'')+' '+(r.ref||'')).toLowerCase().includes(q)) return false; }
  if(S.filter==='miss') return !r.pn;
  if(S.filter==='nc')   return r.nc;
  if(S.filter==='alt')  return r.group;
  if(S.filter==='kn')   return !!r.kn;
  if(S.filter==='chk')  return !!r.chk;
  if(S.filter==='unk')  return !!r.ref_unk;
  return true;
}
function renderFilters(){
  const n={all:S.rows.length,miss:S.rows.filter(r=>!r.pn).length,
           nc:S.rows.filter(r=>r.nc).length,alt:S.rows.filter(r=>r.group).length,
           kn:S.rows.filter(r=>r.kn).length,
           chk:S.rows.filter(r=>r.chk).length,
           unk:S.rows.filter(r=>r.ref_unk).length};
  document.getElementById('filters').innerHTML =
    FILTERS.map(([k,l])=>`<button class="chip ${S.filter===k?'on':''}"
        onclick="S.filter='${k}';renderRows()">${l} ${n[k]}</button>`).join('')
    + `<input class="search" placeholder="搜尋 P/N、Description、Location…" value="${esc(S.q)}"
        oninput="S.q=this.value;renderRows()">`;
}
function renderRows(){
  renderFilters();
  const tb=document.getElementById('tbody');
  if(!S.rows.length){ tb.innerHTML='<tr><td colspan="6" class="empty">尚未執行比對</td></tr>';
                      document.getElementById('detail').innerHTML=''; return; }
  tb.innerHTML = S.rows.map((r,i)=>{
    if(!keep(r)) return '';
    const cls=[r.ref_unk?'r-bad':r.chk?'r-chk':r.kn?'r-kn':r.nc?'r-nc':r.group?'r-alt':'',
               i===S.sel?'sel':''].join(' ');
    const dot=r.chk?'<span class="dot f-chk"></span>'
             :r.kn?'<span class="dot f-kn"></span>':r.nc?'<span class="dot f-nc"></span>'
             :r.group?'<span class="dot f-alt"></span>':'';
    return `<tr class="${cls}" onclick="S.sel=${i};renderRows();renderDetail()">
      <td class="c">${r.item==null?'':r.item}</td>
      <td class="c"><span class="pill ${r.rcec==='RC'?'rc':'ec'}">${r.rcec}</span></td>
      <td>${r.pn ? `<input class="cellin mono" value="${esc(r.pn)}"
             onchange="upd(${i},'pn',this.value)">`
          : `<input class="cellin miss" placeholder="— 無命中 —"
             onchange="upd(${i},'pn',this.value)">`}</td>
      <td><div class="dcell">${dot}<input class="cellin" value="${esc(r.desc)}"
             onchange="upd(${i},'desc',this.value)"></div></td>
      <td class="c">${r.qty==null?'':`<input class="cellin" style="text-align:center"
             value="${esc(r.qty)}" onchange="upd(${i},'qty',this.value)">`}</td>
      ${S.double
        ? `<td><input class="cellin mono" value="${esc(r.ref_top)}" onchange="upd(${i},'ref_top',this.value)"></td>
           <td><input class="cellin mono" value="${esc(r.ref_bot)}" onchange="upd(${i},'ref_bot',this.value)">
           ${r.ref_unk ? `<div class="unkmsg">⚠ ${esc(r.ref_unk)} 查不到 Top/Bot</div>` : ''}</td>`
        : `<td><input class="cellin mono" value="${esc(r.ref)}" onchange="upd(${i},'ref',this.value)"></td>`}
    </tr>`;}).join('');
  document.getElementById('thBot').style.display = S.double ? '' : 'none';
  document.getElementById('outpath2').textContent = short(S.cfg.out||'');
  renderDetail();
}
function renderDetail(){
  const r=S.rows[S.sel]; const d=document.getElementById('detail');
  if(!r){d.innerHTML='';return;}
  let h=`<div class="card"><div class="card-hd"><b>Item ${r.item==null?'↳ 替代料':r.item}</b>
    <span class="sub">${esc(r.part||'')}</span></div><div class="kv">
    <div class="r"><span class="k">Footprint</span><span class="v">${esc(r.fp||'—')}</span></div>
    <div class="r"><span class="k">RC/EC</span><span class="v">
      <select onchange="upd(${S.sel},'rcec',this.value)">
        <option ${r.rcec==='RC'?'selected':''}>RC</option>
        <option ${r.rcec==='EC'?'selected':''}>EC</option></select></span></div>
    <div class="r"><span class="k">P/N</span><span class="v">
      <input value="${esc(r.pn)}" onchange="upd(${S.sel},'pn',this.value)"></span></div>
    <div class="r"><span class="k">Qty</span><span class="v">${r.qty==null
        ? '<span style="color:#9aa0ae">（併入主列）</span>' : esc(r.qty)}</span></div>
    <div class="r"><span class="k">比對狀態</span><span class="v" style="color:${r.pn?'#0f9d58':'#c74343'}">
        ${r.pn?'已命中':'未命中'}</span></div></div></div>`;
  if(r.kn){ h+=`<div class="card"><div class="card-hd"><b>Knowledge 命中</b></div>
    <div class="note"><div class="src">來源：${esc(short(S.cfg.kn||''))}</div>
    <div class="body">${esc(r.kn)}</div>
    <div class="act"><button class="btn g" onclick="alert('已採用，將寫入 Excel')">✓ 採用</button>
      <button class="btn" onclick="upd(${S.sel},'kn','')">✕ 忽略此筆</button></div></div></div>`; }
  h+=`<div class="console">${S.log.map(fmt).join('\n')}</div>`;
  d.innerHTML=h;
}
const fmt = l => l.startsWith('✓')?`<span class="ok">${esc(l)}</span>`
              : l.startsWith('✗')?`<span class="er">${esc(l)}</span>`
              : l.startsWith('[')?`<span class="wn">${esc(l)}</span>`: esc(l);

function renderKn(){
  const hits=S.rows.map((r,i)=>[i,r]).filter(([,r])=>r.kn);
  document.getElementById('knlist').innerHTML = hits.length ? hits.map(([i,r])=>
    `<div class="card" style="margin-bottom:10px"><div class="card-hd">
      <b>Item ${r.item==null?'↳':r.item} · ${esc(r.pn||'（無 P/N）')}</b>
      <span class="sub">${esc(r.desc)}</span>
      <div class="rt"><button class="btn" onclick="upd(${i},'kn','');renderKn()">✕ 忽略此筆</button></div>
      </div><div class="note"><div class="body">${esc(r.kn)}</div></div></div>`).join('')
    : '<div class="empty">沒有任何 Knowledge 命中</div>';
}

const HF=[['pcb_no','PCB板號'],['pcb_pn','PCB料號'],['rcec_no','RC/EC NO.'],
          ['pcb_vendor','PCB供應商'],['pcb_spec','PCB 規格'],['serial_no','發行序號'],
          ['issue_date','發行(改版)日期'],['rev','目前版次'],['ic_type','Support IC Type']];
function renderHdr(){
  document.getElementById('hdrform').innerHTML = HF.map(([k,l])=>
    `<label>${l}</label><input value="${esc(S.hdr[k]||'')}"
       onchange="S.hdr['${k}']=this.value;push()"><div></div>`).join('');
}
function renderOut(){
  document.getElementById('outpath').value = S.cfg.out||'';
  const s=S.stats;
  document.getElementById('outinfo').textContent = s.items
    ? `${s.items} 項（含 PCB）→ ${s.rows} 列　|　Total Qty = ${s.qty}　|　PCB 圖片：${short(S.cfg.pcb||'')}`
    : '尚未執行比對';
  document.getElementById('log2').innerHTML = S.log.map(fmt).join('\n') || '（尚未產出）';
}

function badges(){
  document.getElementById('bd-files').textContent =
    (S.disc.txts||[]).length+(S.disc.dbs||[]).length+(S.disc.kns||[]).length+(S.disc.imgs||[]).length;
  document.getElementById('bd-rows').textContent = S.rows.length||'-';
  document.getElementById('bd-kn').textContent = S.rows.filter(r=>r.kn).length||'-';
  document.getElementById('sbfoot').textContent = S.disc.base||'';
  document.getElementById('ver').textContent = (S.launcher||'') + ' · v' + (S.ver||'');
}
function paintLog(){
  document.getElementById('log1').innerHTML = S.log.map(fmt).join('\n')||'（尚未執行）';
}
function renderAll(){ renderFiles(); renderStats(); renderRows(); badges(); paintLog(); }

/* ── bridge ─────────────────────────────────────────────────────────── */
function apply(st){ Object.assign(S,st); renderAll(); }
function setCfg(k,v){ S.cfg[k]=v; if(LIVE()) api().set_cfg(k, v); if(k==='out') renderRows(); }
function push(){ if(LIVE()) api().set_header(S.hdr); }
function upd(i,f,v){
  S.rows[i][f]= (f==='qty' ? (v===''?null:parseInt(v,10)) : v);
  if(LIVE()) api().update_row(i,f,v).then(st=>{ S.stats=st.stats; renderStats(); });
  if(f==='kn'||f==='rcec'){ renderRows(); }
}
function busy(on){ document.querySelectorAll('.btn.p').forEach(b=>b.disabled=on); }
function doMatch(){ if(!LIVE()) return; busy(true);
  api().run_match().then(st=>{ apply(st); busy(false); page('rows'); }); }
function doExport(force){ if(!LIVE()) return;
  // 直接讀輸入框的現值，不要只靠 onchange ——
  // 使用者改完檔名直接按按鈕時，change 事件不一定趕得及送出去。
  const el = document.getElementById('outp');
  const p  = el && el.value.trim() ? el.value.trim() : (S.cfg.out||'');
  api().path_exists(p).then(exists=>{
    if(!force && exists && !confirm('這個檔案已經存在，要覆蓋嗎？\n\n' + p)) return;
    busy(true);
    api().export_excel(p).then(st=>{ apply(st); busy(false); page('out');
      renderOut(); });
  });
}
function browse(kind){ if(!LIVE()) return;
  api().browse(kind).then(st=>{ if(st) apply(st); }); }
function rescan(){ if(!LIVE()) return; api().rescan().then(apply); }
function openOut(){ if(LIVE()) api().open_output(); }

/* 先畫一個空殼，避免載入瞬間空白 */
S.ver='__VER__'; S.disc={base:'',txts:[],dbs:[],kns:[],imgs:[]};
S.cfg={bom:'',db:'',kn:'',pcb:'',out:'',split:true}; S.hdr={}; renderAll();

/* ── v0.10 彩蛋：按住 Shift 連點三次 LOGO ──────────────────────────────
   ⚠️ 監聽 pointerdown 不是 click：<img> 的原生拖曳會把 click 吃掉。 */
/*__BEE_GAME__*/
let eggN=0, eggT=0, eggStop=null;
function eggOpen(){
  document.getElementById('egg').hidden=false;
  if(eggStop) eggStop();
  eggStop = bmBeeGame(document.getElementById('eggcv'), eggClose);
}
function eggClose(){
  if(eggStop){ eggStop(); eggStop=null; }   // 收掉動畫與鍵盤監聽，否則方向鍵會繼續被吃
  document.getElementById('egg').hidden=true;
}
const LOGO = document.getElementById('logo');
if(LOGO) LOGO.addEventListener('pointerdown', e=>{
  if(!e.shiftKey){ eggN=0; return; }
  e.preventDefault();                       // Shift+點擊會選取文字
  const now=Date.now();
  if(now-eggT > __EGG_MS__) eggN=0;
  eggT=now;
  if(++eggN >= __EGG_N__){ eggN=0; eggOpen(); }
});

window.addEventListener('pywebviewready', ()=>{
  document.getElementById('banner').innerHTML='';
  api().get_state().then(apply);
});
setTimeout(()=>{ if(!LIVE()) document.getElementById('banner').innerHTML =
  '<div class="banner">預覽模式：這是用瀏覽器開啟的靜態畫面，未連上 Python。'
  + '實際使用請執行 <b>run_D_webview.bat</b>。</div>'; }, 1500);
</script></body></html>
"""


# ════════════════════════════════════════════════════════════════════════════
# Python 端 API
# ════════════════════════════════════════════════════════════════════════════

# pywebview 會遞迴列舉 js_api 物件的「公開屬性」並序列化給 JS。
# 若把 window 物件掛成 self.window，會走進 window.native.AccessibilityObject...
# 造成 maximum recursion depth exceeded 並讓視窗當場關閉。
# → window 一律放模組層變數，不掛在 Api 上。
WIN = None

# v0.10：標題列帶上啟動用的批次檔名稱，同時開 C/D 兩個介面時才分得出來
LAUNCHER = 'run_D_webview.bat'
APP_TITLE = f'BOM Matcher v{core.VERSION}  [{LAUNCHER}]'


class Api:
    def __init__(self):
        self.disc = core.discover(APP_DIR)
        self.cfg = {
            'bom':   self.disc['bom'],
            'db':    self.disc['db'],
            'kn':    self.disc['knowledge'],
            'pcb':   self.disc['pcb'],
            'pcb_bot': self.disc['pcb_bot'],
            'sides': self.disc['sides'],
            'out':   self.disc['out'],
            'split': True,
        }
        self.hdr  = core.empty_header()
        self.rows = []
        self.log  = []

    # ── helpers ────────────────────────────────────────────────────────────
    def _log(self, msg):
        self.log.append(str(msg))

    def _rows_js(self):
        return [{
            'item': r['item'], 'rcec': r['rcec'], 'pn': r['pn'], 'desc': r['desc'],
            'qty': r['qty'], 'ref': r['ref'], 'nc': r['nc'], 'alt': r['alt'],
            'group': r['group'], 'kn': r['kn'], 'fp': r['fp'], 'part': r['part'],
            # v0.10：畫面也顯示 (T)/(B)，與輸出 Excel 一致（tag_side 冪等）
            'ref_top': core.tag_side(r.get('ref_top', ''), 'T'),
            'ref_bot': core.tag_side(r.get('ref_bot', ''), 'B'),
            'ref_unk': r.get('ref_unk', ''), 'chk': bool(r.get('chk')),
        } for r in self.rows]

    def _state(self):
        return {
            'ver':   core.VERSION,
            'launcher': LAUNCHER,
            'disc':  self.disc,
            'cfg':   self.cfg,
            'hdr':   self.hdr,
            'rows':  self._rows_js(),
            'stats': core.stats(self.rows) if self.rows else {},
            'double': any(r.get('ref_bot') or r.get('ref_unk')
                          for r in self.rows) if self.rows else False,
            'log':   self.log,
        }

    # ── exposed ────────────────────────────────────────────────────────────
    def get_state(self):
        return self._state()

    def set_cfg(self, key, value):
        if key == 'split':
            self.cfg['split'] = bool(value)
        else:
            self.cfg[key] = value
        if key == 'bom' and value:
            self.cfg['out'] = core.default_out_name(value, APP_DIR)
        return self._state()

    def set_header(self, hdr):
        self.hdr.update({k: str(v or '') for k, v in (hdr or {}).items()})
        return True

    def rescan(self):
        self.disc = core.discover(APP_DIR)
        for k, src in (('bom', 'bom'), ('db', 'db'),
                       ('kn', 'knowledge'), ('pcb', 'pcb'), ('out', 'out')):
            if not self.cfg.get(k):
                self.cfg[k] = self.disc[src]
        self._log(f"✓ 重新掃描 {APP_DIR}：TXT {len(self.disc['txts'])} · "
                  f"DB {len(self.disc['dbs'])} · Knowledge {len(self.disc['kns'])} · "
                  f"圖片 {len(self.disc['imgs'])}")
        return self._state()

    def browse(self, kind):
        import webview
        if WIN is None:
            return None
        types = {
            'bom': ('Text BOM (*.txt)', 'All files (*.*)'),
            'db':  ('Excel (*.xlsx)', 'All files (*.*)'),
            'kn':  ('Excel (*.xlsx)', 'All files (*.*)'),
            'pcb': ('Image (*.jpg;*.jpeg;*.png)', 'All files (*.*)'),
            'pcb_bot': ('Image (*.jpg;*.jpeg;*.png)', 'All files (*.*)'),
            'sides': ('Excel (*.xlsx)', 'Text (*.txt)', 'All files (*.*)'),
            'out': ('Excel (*.xlsx)',),
        }[kind]
        if kind == 'out':
            r = WIN.create_file_dialog(
                webview.SAVE_DIALOG, directory=str(core.output_dir(APP_DIR)),
                save_filename=Path(self.cfg['out']).name, file_types=types)
            path = r if isinstance(r, str) else (r[0] if r else '')
        else:
            r = WIN.create_file_dialog(
                webview.OPEN_DIALOG, directory=str(APP_DIR), file_types=types)
            path = r[0] if r else ''
        if not path:
            return None
        return self.set_cfg(kind, path)

    def run_match(self):
        self.log = []
        try:
            if not self.cfg['bom'] or not self.cfg['db']:
                raise ValueError('請先選擇 Text BOM 與零件資料庫')
            self.rows, *_ = core.run_match(
                self.cfg['bom'], self.cfg['db'], self.cfg['kn'],
                split_alt=self.cfg['split'], log=self._log,
                sides_path=self.cfg.get('sides', ''))
        except Exception as e:
            self.rows = []
            self._log(f'✗ 錯誤：{e}')
        return self._state()

    def update_row(self, idx, field, value):
        if 0 <= idx < len(self.rows):
            if field == 'qty':
                v = str(value).strip()
                self.rows[idx]['qty'] = int(float(v)) if v.replace('.', '', 1).isdigit() else None
            else:
                self.rows[idx][field] = str(value)
        return {'stats': core.stats(self.rows) if self.rows else {}}

    def path_exists(self, p):
        try:
            return bool(p) and Path(core.force_xlsx(str(p))).is_file()
        except Exception:
            return False

    def export_excel(self, out_path=None):
        self.log = []
        try:
            if not self.rows:
                raise ValueError('尚未執行比對')
            if out_path and str(out_path).strip():
                self.cfg['out'] = str(out_path).strip()
            core.write_excel(self.rows, self.cfg['out'],
                             self.cfg['pcb'] or None, self.hdr, log=self._log,
                             pcb_bot=self.cfg.get('pcb_bot') or None)
        except Exception as e:
            self._log(f'✗ 錯誤：{e}')
        return self._state()

    def open_output(self):
        d = Path(self.cfg['out']).parent
        if d.is_dir():
            subprocess.Popen(['explorer', str(d)])
        return True


# ════════════════════════════════════════════════════════════════════════════
# 入口
# ════════════════════════════════════════════════════════════════════════════

def _selftest():
    """--selftest：開視窗、用 JS 觸發整條流程、印出結果後自動關閉。"""
    import time
    out = []

    def probe():
      try:
        time.sleep(3)
        WIN.evaluate_js('doMatch()')
        time.sleep(4)
        out.append('rows      = ' + str(WIN.evaluate_js('S.rows.length')))
        out.append('stats     = ' + str(WIN.evaluate_js('JSON.stringify(S.stats)')))
        out.append('tbody rows= ' + str(WIN.evaluate_js(
            'document.querySelectorAll("#tbody tr").length')))
        out.append('banner    = ' + repr(WIN.evaluate_js(
            'document.getElementById("banner").innerText')))
        WIN.evaluate_js('page("rows"); S.sel=1; renderRows()')
        time.sleep(1)
        out.append('detail    = ' + repr(WIN.evaluate_js(
            'document.getElementById("detail").innerText.slice(0,60)')))
        WIN.evaluate_js('upd(0,"qty","9")')
        time.sleep(1)
        out.append('after edit= ' + str(WIN.evaluate_js('JSON.stringify(S.stats)')))
        WIN.evaluate_js('doExport(true)')   # 略過覆蓋確認對話框
        time.sleep(5)
        out.append('log       = ' + str(WIN.evaluate_js('JSON.stringify(S.log)')))
      except Exception as e:
        import traceback
        out.append('EXC: ' + traceback.format_exc())
      finally:
        (APP_DIR / '_selftest_D.txt').write_text('\n'.join(out), encoding='utf-8')
        WIN.destroy()

    threading.Thread(target=probe, daemon=True).start()
    return out


def page_html():
    """HTML 裡的佔位字串換成實際內容（LOGO、版號、彩蛋遊戲）。"""
    logo = core.app_logo_uri()
    tag = (f'<img class="m" id="logo" src="{logo}" draggable="false" alt="ADATA">'
           if logo else '<div class="m">B</div>')
    return (HTML.replace('__LOGO_TAG__', tag)
                .replace('__VER__', core.VERSION)
                .replace('__EGG_MS__', str(core.EGG_WINDOW_MS))
                .replace('__EGG_N__', str(core.EGG_CLICKS))
                .replace('/*__BEE_GAME__*/', core.BEE_GAME_JS))


def main():
    if '--dump-html' in sys.argv:
        out = APP_DIR / 'preview_D_ui.html'
        out.write_text(page_html(), encoding='utf-8')
        print(f'✓ UI HTML 已輸出: {out}')
        return 0

    global WIN
    import webview
    api = Api()
    WIN = webview.create_window(
        APP_TITLE, html=page_html(), js_api=api,
        width=1280, height=820, min_size=(1080, 680))

    result = _selftest() if '--selftest' in sys.argv else None
    webview.start()
    if result is not None:
        print('\n'.join(result))
    return 0


if __name__ == '__main__':
    sys.exit(main())
