# -*- coding: utf-8 -*-
"""
空白图检测 · Excel 批量插件（一键启动 / 自包含 / 零 token / HTML 界面）
================================================================
把纯白图（白底图）识别封装成「双击即用」的桌面插件：

  - 一键启动：双击 exe → 自动打开浏览器 → 卡片式向导操作
  - 内含环境：PyInstaller 单文件打包，已内置 Python + numpy/Pillow/openpyxl，
    无需在用户机器安装任何环境
  - 零 token：纯像素分析（white_detect），不调 AI 模型 / 不调 Dify / 不联网推理，
    完全离线判定，不产生任何 token 费用
  - 自定义字段：上传 Excel 后从表头勾选「图片地址列」（可多选）
  - 双维度同时返回（v3 新增）：
        A. 全部图片均为空白才算不合规
        B. 任意一张图片空白即不合规
    一次扫描把两种维度的判定结果都写到 Excel 不同列，互不冲突
  - 返回 Excel：结果写回新 Excel（检测结果 + 汇总 两个工作表）

两种运行形态：
  1) 有参数且为 .xlsx  → 无界面（headless）批处理，便于脚本/命令行调用
  2) 无参数（双击）    → 启动本地服务并弹出浏览器 HTML 向导

用法：
  WhiteImageExcel.exe                      # 浏览器 HTML 向导
  WhiteImageExcel.exe 商品.xlsx            # 无界面处理，默认两种维度同时输出
  WhiteImageExcel.exe 商品.xlsx --rule any # 仅输出"任意一张即不合规"维度
  WhiteImageExcel.exe 商品.xlsx --rule all # 仅输出"全部空白才不合规"维度
================================================================
"""

from __future__ import annotations

import os
import sys
import time
import json
import base64
import threading
import webbrowser
import tempfile

# ── 必须在 import numpy 之前固化线程数，避免 OpenBLAS 在冻结 exe 里派生大量
#    自旋线程抢占 GIL，导致像素分析慢 100 倍甚至看似“卡死”。──
os.environ.setdefault("OMP_NUM_THREADS", "1")
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("MKL_NUM_THREADS", "1")
os.environ.setdefault("NUMEXPR_NUM_THREADS", "1")

# ── 冻结(exe)模式下把 stdout/stderr 重定向到调试文件（行缓冲，进程被杀也不丢）──
if getattr(sys, "frozen", False):
    _dbg_path = os.path.join(os.path.dirname(sys.executable), "debug_trace.log")
    try:
        _dbg = open(_dbg_path, "a", buffering=1, encoding="utf-8", errors="replace")
        sys.stdout = _dbg
        sys.stderr = _dbg
    except Exception:
        pass
    print(f"[boot] {time.strftime('%H:%M:%S')} exe started, pid={os.getpid()}", flush=True)

# batch_excel 已包含全部 Excel 读写与判定逻辑（纯像素，无 token）
try:
    from batch_excel import (run_batch, run_batch_chunked, count_data_rows,
                             export_from_progress)
except ImportError:
    _HERE = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, _HERE)
    from batch_excel import (run_batch, run_batch_chunked, count_data_rows,  # type: ignore
                             export_from_progress)


# ════════════════════════════════════════════════════════════════
#  HTML 向导界面（完全静态，交互通过 fetch 调用本地 API）
# ════════════════════════════════════════════════════════════════
PAGE_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>空白图检测 · Excel 批量插件</title>
<style>
  :root{
    --bg:#f4f6fb; --card:#ffffff; --primary:#2563eb; --primary-d:#1d4ed8;
    --text:#1f2937; --muted:#6b7280; --line:#e5e7eb; --ok:#16a34a; --warn:#dc2626;
    --chip:#eef2ff;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;
       background:var(--bg);color:var(--text);line-height:1.5}
  .wrap{max-width:880px;margin:0 auto;padding:28px 18px 60px}
  header.top{margin-bottom:18px}
  header.top h1{font-size:22px;margin:0 0 4px}
  header.top p{margin:0;color:var(--muted);font-size:13px}
  .badge{display:inline-block;background:var(--chip);color:var(--primary);
         font-size:12px;padding:2px 10px;border-radius:999px;margin-right:6px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:14px;
        padding:18px 20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(16,24,40,.04)}
  .step{display:flex;align-items:center;gap:10px;margin-bottom:12px}
  .step .num{width:24px;height:24px;border-radius:50%;background:var(--primary);color:#fff;
             font-size:13px;display:flex;align-items:center;justify-content:center;flex:0 0 auto}
  .step h2{font-size:16px;margin:0}
  .hint{color:var(--muted);font-size:12.5px;margin:0 0 12px}
  .drop{border:2px dashed var(--line);border-radius:12px;padding:26px;text-align:center;
        cursor:pointer;transition:.15s;background:#fafbff}
  .drop:hover{border-color:var(--primary);background:#f0f5ff}
  .drop.drag{border-color:var(--primary);background:#e8f0ff}
  .drop b{color:var(--primary)}
  .fname{margin-top:10px;font-size:13px;color:var(--text)}
  .field-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px}
  .chk{display:flex;align-items:center;gap:8px;border:1px solid var(--line);border-radius:9px;
       padding:8px 10px;font-size:13px;cursor:pointer;user-select:none}
  .chk:hover{border-color:var(--primary)}
  .chk input{width:16px;height:16px;accent-color:var(--primary)}
  .chk.on{border-color:var(--primary);background:var(--chip)}
  .rule-box{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:4px}
  .opt{border:1px solid var(--line);border-radius:11px;padding:12px 14px;cursor:pointer;transition:.15s;position:relative}
  .opt:hover{border-color:var(--primary)}
  .opt.on{border-color:var(--primary);background:var(--chip)}
  .opt .t{font-weight:600;font-size:14px;padding-left:22px;display:block}
  .opt .d{font-size:12px;color:var(--muted);margin-top:3px;padding-left:22px}
  .opt input{position:absolute;left:12px;top:13px;width:16px;height:16px;accent-color:var(--primary);margin:0}
  .row{display:flex;gap:14px;flex-wrap:wrap;align-items:center}
  .row label{font-size:13px;color:var(--muted)}
  .row input[type=number]{width:90px;padding:7px 9px;border:1px solid var(--line);
       border-radius:8px;font-size:13px}
  .btn{background:var(--primary);color:#fff;border:0;border-radius:10px;
       font-size:15px;font-weight:600;padding:13px 18px;cursor:pointer;width:100%;
       transition:.15s;font-family:inherit}
  .btn:hover{background:var(--primary-d)}
  .btn:disabled{background:#9db4e8;cursor:not-allowed}
  .prog{height:10px;background:#eef2f7;border-radius:999px;overflow:hidden;margin:10px 0}
  .prog>i{display:block;height:100%;width:0;background:var(--primary);transition:width .25s}
  .log{background:#0f172a;color:#cbd5e1;font-family:Consolas,Menlo,monospace;font-size:12px;
       border-radius:10px;padding:12px;height:170px;overflow:auto;white-space:pre-wrap;margin:8px 0 0}
  .result{display:none;border:1px solid var(--ok);background:#f0fdf4;border-radius:12px;padding:16px}
  .result h3{margin:0 0 8px;color:var(--ok);font-size:16px}
  .result table{width:100%;border-collapse:collapse;font-size:13px}
  .result td{padding:5px 8px;border-bottom:1px solid #d7f0de}
  .result td:first-child{color:var(--muted);width:42%}
  .dl{display:inline-block;margin-top:12px;background:var(--ok);color:#fff;text-decoration:none;
      padding:11px 18px;border-radius:10px;font-weight:600;font-size:14px;cursor:pointer}
  /* ── 分批（v4）── */
  .batchbar{font-size:13px;color:var(--primary);font-weight:600;margin:2px 0 6px;min-height:18px}
  .chunkbox{margin-top:12px;border:1px solid var(--line);border-radius:11px;padding:10px 12px;background:#fbfcff}
  .chunkbox.in-result{border-color:#d7f0de;background:#f7fdf9}
  .chunk-title{display:flex;align-items:baseline;gap:8px;font-size:13px;font-weight:600;margin-bottom:8px}
  .chunk-item{display:flex;align-items:center;gap:10px;flex-wrap:wrap;
              padding:7px 9px;border:1px solid var(--line);border-radius:9px;
              background:#fff;margin-bottom:6px;font-size:12.5px}
  .chunk-item:last-child{margin-bottom:0}
  .chunk-item .nm{font-weight:600;color:var(--text)}
  .chunk-item .meta{color:var(--muted)}
  .chunk-item .tag{background:#eef2ff;color:var(--primary);border-radius:999px;
                   padding:1px 8px;font-size:11.5px}
  .chunk-item .tag.warn{background:#fef2f2;color:var(--warn)}
  .chunk-item .tag.ok{background:#f0fdf4;color:var(--ok)}
  .chunk-item .sp{flex:1 1 auto}
  .dl-sm{background:var(--ok);color:#fff;border:0;border-radius:8px;padding:6px 12px;
         font-size:12.5px;font-weight:600;cursor:pointer;font-family:inherit;white-space:nowrap}
  .dl-sm:hover{filter:brightness(.95)}
  .dl-sm:disabled{background:#9db4e8;cursor:not-allowed}
  .dlrow{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
  .partialrow{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:10px;
              padding-top:10px;border-top:1px dashed var(--line)}
  .err{color:var(--warn);font-size:13px;margin-top:8px}
  .muted{color:var(--muted)}
  .hidden{display:none}
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <h1>空白图（白底图）批量识别</h1>
    <p><span class="badge">纯像素分析</span><span class="badge">离线 · 零 token</span>
       <span class="badge">结果写回 Excel</span></p>
  </header>

  <!-- Step 1 上传 -->
  <div class="card">
    <div class="step"><div class="num">1</div><h2>选择 Excel</h2></div>
    <p class="hint">支持 .xlsx / .xls。解析与检测全部在本机完成，不会上传任何服务器。</p>
    <div class="drop" id="drop">
      <div>拖拽文件到此处，或 <b>点击选择</b></div>
      <div class="fname" id="fname">未选择文件</div>
    </div>
    <input type="file" id="file" accept=".xlsx,.xls" class="hidden"/>
  </div>

  <!-- Step 2 字段 + 维度 -->
  <div class="card" id="step2">
    <div class="step"><div class="num">2</div><h2>字段与判定维度</h2></div>
    <p class="hint">勾选包含图片地址的列（可多选，即“自定义字段”）。再选择不合规判定维度。</p>
    <div style="font-size:13px;font-weight:600;margin:4px 0 6px">图片地址列</div>
    <div class="field-grid" id="cols"></div>
    <div id="colsErr" class="err hidden">请至少勾选一个图片地址列。</div>

    <div style="font-size:13px;font-weight:600;margin:16px 0 6px">不合规判定维度（可多选，默认同时输出两种）</div>
    <div class="rule-box">
      <label class="opt on" data-rule="any">
        <input type="checkbox" name="rule" value="any" checked/>
        <div class="t">B · 一张即不合规</div>
        <div class="d">任意一张图片为空白，整行判为不合规</div>
      </label>
      <label class="opt on" data-rule="all">
        <input type="checkbox" name="rule" value="all" checked/>
        <div class="t">A · 全部空白才不合规</div>
        <div class="d">仅当该行所有图片均为空白，才判为不合规</div>
      </label>
    </div>
  </div>

  <!-- Step 3 参数 -->
  <div class="card">
    <div class="step"><div class="num">3</div><h2>处理参数</h2></div>
    <div class="row">
      <label>并发线程数</label>
      <input type="number" id="workers" value="8" min="1" max="32"/>
      <label>近白像素阈值</label>
      <input type="number" id="threshold" value="0.99" min="0.5" max="1" step="0.01"/>
    </div>
    <div class="row" style="margin-top:10px">
      <label>分批行数上限</label>
      <input type="number" id="chunkSize" value="1000" min="50" step="100"/>
      <span class="muted" style="font-size:12.5px">
        超过该行数时自动分批执行：每批完成即可下载该批结果，全部完成后再合并为全量 Excel。
      </span>
    </div>
  </div>

  <button class="btn" id="start" disabled>开始检测</button>

  <!-- 进度 -->
  <div class="card hidden" id="progCard">
    <div class="step"><div class="num">›</div><h2>处理中</h2></div>
    <div class="prog"><i id="bar"></i></div>
    <div class="batchbar" id="batchbar"></div>
    <div class="log" id="log"></div>

    <!-- 运行中随时导出已完成部分 -->
    <div class="partialrow">
      <button class="dl-sm" id="btnPartial" onclick="downloadPartial()">导出已完成部分</button>
      <span class="muted" style="font-size:12px">
        不必等全部跑完 —— 立即导出一份含「已完成行」的 Excel（未完成的行标为「未处理」）
      </span>
    </div>

    <!-- 已完成批次（可逐批下载） -->
    <div class="chunkbox hidden" id="chunkBox">
      <div class="chunk-title">
        <span>已完成的批次</span>
        <span class="muted" style="font-weight:400;font-size:12px">点「下载本批」即可立即取走该批结果</span>
      </div>
      <div id="chunkList"></div>
    </div>
  </div>

  <!-- 结果 -->
  <div class="result" id="result">
    <h3>✓ 检测完成</h3>
    <table id="sumTab"></table>
    <div class="dlrow">
      <a class="dl" id="dlFinal" href="#" onclick="return false;">下载全量结果 Excel</a>
    </div>
    <div id="resultChunks" class="chunkbox in-result hidden">
      <div class="chunk-title">
        <span>分批结果</span>
        <span class="muted" style="font-weight:400;font-size:12px">按批次单独下载</span>
      </div>
      <div id="resultChunkList"></div>
    </div>
  </div>
</div>

<script>
const $ = s => document.querySelector(s);
let fileB64 = null, fileName = "", headers = [];
let chunkSig = "";   // 批次列表渲染签名，避免轮询重复重绘

const drop = $("#drop"), file = $("#file");
drop.onclick = () => file.click();
file.onchange = e => { if (e.target.files[0]) loadFile(e.target.files[0]); };
["dragover","dragenter"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("drag");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("drag");}));
drop.addEventListener("drop", e=>{ const f=e.dataTransfer.files[0]; if(f) loadFile(f); });

function loadFile(f){
  $("#fname").textContent = f.name + "（" + (f.size/1024).toFixed(0) + " KB）";
  const rd = new FileReader();
  rd.onload = () => {
    fileB64 = rd.result.split(",")[1];
    fileName = f.name;
    prepare();
  };
  rd.readAsDataURL(f);
}

async function prepare(){
  setErr("");
  const r = await fetch("/api/prepare",{method:"POST",
    headers:{"Content-Type":"application/json"},
    body: JSON.stringify({filename:fileName, data:fileB64})});
  const j = await r.json();
  if(j.error){ setErr("解析失败：" + j.error); return; }
  headers = j.headers || [];
  renderCols(headers);
}

function renderCols(list){
  const box = $("#cols"); box.innerHTML = "";
  list.forEach((h,i)=>{
    const lab = document.createElement("label");
    lab.className = "chk";
    const key = (h||"").toLowerCase();
    const hit = /图片|图|url|链接|link|image|主图|详情图|pic|img/.test(key);
    if(hit) lab.classList.add("on");
    lab.innerHTML = '<input type="checkbox" '+(hit?"checked":"")+'/> <span></span>';
    lab.querySelector("span").textContent = (h===null||h===""?"（空列名"+(i+1)+"）":h);
    lab.querySelector("input").onchange = ()=> lab.classList.toggle("on", lab.querySelector("input").checked);
    box.appendChild(lab);
  });
  updateStart();
}

function selCols(){
  return [...document.querySelectorAll("#cols .chk input")].map((c,i)=>c.checked?headers[i]:null).filter(x=>x!=null);
}
function selRules(){ return [...document.querySelectorAll('input[name=rule]:checked')].map(x=>x.value); }

document.querySelectorAll(".opt").forEach(o=>{
  o.onclick = (e)=>{
    // 阻止 label 默认的 input toggle 行为，自己手动控制避免双触发
    if(e.target.tagName !== "INPUT"){
      const cb = o.querySelector("input");
      cb.checked = !cb.checked;
    }
    o.classList.toggle("on", o.querySelector("input").checked);
    updateStart();
  };
  // 同步初始 .on 状态（首次默认两个都勾上时）
  o.classList.toggle("on", o.querySelector("input").checked);
});

function updateStart(){ $("#start").disabled = !(fileB64 && selCols().length>0 && selRules().length>0); }
$("#cols").addEventListener("change", updateStart);
document.querySelectorAll('input[name=rule]').forEach(c=>c.addEventListener("change", ()=>{
  c.parentElement.classList.toggle("on", c.checked);
  updateStart();
}));

function setErr(m){ const e=$("#colsErr"); if(m){e.textContent=m;e.classList.remove("hidden");}else e.classList.add("hidden"); }

$("#start").onclick = async ()=>{
  const cols = selCols();
  const rules = selRules();
  if(!cols.length){ setErr("请至少勾选一个图片地址列。"); return; }
  if(!rules.length){ setErr("请至少勾选一种判定维度。"); return; }
  setErr("");
  $("#progCard").classList.remove("hidden");
  $("#result").style.display="none";
  const log = $("#log"); log.textContent = "";
  $("#bar").style.width = "0%";
  $("#start").disabled = true;

  const body = {
    filename:fileName, data:fileB64,
    sheet:null, cols:cols, rules:rules,
    workers:+$("#workers").value||8,
    threshold:+$("#threshold").value||0.99,
    chunk_size:+$("#chunkSize").value||1000
  };
  const r = await fetch("/api/run",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(body)});
  const j = await r.json();
  if(j.error){ log.textContent += "\n[错误] " + j.error; $("#start").disabled=false; return; }
  chunkSig = ""; $("#chunkBox").classList.add("hidden");
  poll();
};

async function poll(){
  const log = $("#log");
  try{
    const r = await fetch("/api/progress");
    const j = await r.json();
    if(j.msg) log.textContent += "\n" + j.msg;
    log.scrollTop = log.scrollHeight;
    const pct = j.total ? Math.min(100, Math.round(100*j.done/j.total)) : (j.status==="done"?100:0);
    $("#bar").style.width = pct + "%";

    // 分批进度条
    if((j.chunk_total||0) > 1){
      $("#batchbar").textContent =
        "分批执行：已完成 " + (j.current_chunk||0) + "/" + j.chunk_total + " 批" +
        (j.chunk_size ? "（每批 " + j.chunk_size + " 行）" : "");
    } else if(j.status === "running"){
      $("#batchbar").textContent = "";
    }

    // 每完成一批就刷新「可下载」列表
    const chunks = j.chunks || [];
    const sig = chunks.map(c=>c.index+":"+c.rows).join(",");
    if(sig !== chunkSig){
      chunkSig = sig;
      renderChunks("#chunkList", chunks, false);
    }

    if(j.status==="done"){
      showResult(j);
      return;
    }
    if(j.status==="error"){
      log.textContent += "\n[错误] " + (j.error||"未知错误");
      $("#start").disabled=false;
      return;
    }
    setTimeout(poll, 700);
  }catch(e){ log.textContent += "\n[轮询失败] " + e; setTimeout(poll, 1500); }
}

// ── 分批结果渲染 + 下载（v4）──
function renderChunks(sel, chunks, inResult){
  const box = $(sel);
  const wrap = inResult ? $("#resultChunks") : $("#chunkBox");
  if(!chunks.length){ wrap.classList.add("hidden"); return; }
  wrap.classList.remove("hidden");
  box.innerHTML = chunks.map(c=>{
    const seg = "源表第 " + c.start_row + "-" + c.end_row + " 行 · " + c.rows + " 行";
    return '<div class="chunk-item">' +
      '<span class="nm">批次 ' + c.index + '/' + c.total + '</span>' +
      '<span class="meta">' + seg + '</span>' +
      '<span class="tag">一张即不合规 ' + c.any + ' 行</span>' +
      '<span class="tag ok">全部空白才不合规 ' + c.all + ' 行</span>' +
      '<span class="sp"></span>' +
      '<button class="dl-sm" data-chunk="' + c.index + '" data-name="' + (c.name||"") + '">下载本批</button>' +
      '</div>';
  }).join("");
  box.querySelectorAll("button[data-chunk]").forEach(b=>{
    b.onclick = ()=> downloadChunk(b.dataset.chunk, b.dataset.name, b);
  });
}

// 运行中导出「已完成部分」
async function downloadPartial(){
  const btn = $("#btnPartial");
  const old = btn ? btn.textContent : "";
  if(btn){ btn.disabled = true; btn.textContent = "导出中…"; }
  try{
    const r = await fetch("/api/export-partial");
    const j = await r.json();
    if(j.error){ alert(j.error); return; }
    const a = document.createElement("a");
    a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + j.b64;
    a.download = j.name || "已完成部分.xlsx";
    document.body.appendChild(a); a.click(); a.remove();
    const s = j.summary || {};
    const log = $("#log");
    log.textContent += "\n[导出] 已完成 " + s.rows_done + " 行 / 未处理 " + s.pending_rows +
      " 行（any " + s.noncompliant_rows_any + " / all " + s.noncompliant_rows_all + "）";
    log.scrollTop = log.scrollHeight;
  }catch(e){ alert("导出失败：" + e); }
  finally{ if(btn){ btn.disabled = false; btn.textContent = old; } }
}

async function downloadChunk(kind, fallbackName, btn){
  const old = btn ? btn.textContent : "";
  if(btn){ btn.disabled = true; btn.textContent = "准备中…"; }
  try{
    const r = await fetch("/api/download?chunk=" + encodeURIComponent(kind));
    const j = await r.json();
    if(j.error){ alert(j.error); return; }
    const a = document.createElement("a");
    a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + j.b64;
    a.download = j.name || fallbackName || "空白图结果.xlsx";
    document.body.appendChild(a); a.click(); a.remove();
  }catch(e){ alert("下载失败：" + e); }
  finally{ if(btn){ btn.disabled = false; btn.textContent = old; } }
}

function showResult(j){
  const s = j.summary || {};
  const rows = [
    ["输出文件", s.output || ""],
    ["批次数", (s["chunks"]||1) + (s["chunk_size"] ? ("（每批 " + s["chunk_size"] + " 行）") : "")],
    ["判定维度", (s["rules"]||[]).map(r => r==="any"?"B · 一张即不合规":"A · 全部空白才不合规").join("  +  ")],
  ];
  // 按维度分别展示不合规行数（任何一种维度被勾选都显示）
  if((s["rules"]||[]).includes("any"))
    rows.push(["不合规行数（B · 一张即不合规）", s["noncompliant_rows_any"] ?? s["white_rows"] ?? ""]);
  if((s["rules"]||[]).includes("all"))
    rows.push(["不合规行数（A · 全部空白才不合规）", s["noncompliant_rows_all"] ?? ""]);
  rows.push(
    ["纯白图数量(张)", s["white_imgs"] ?? ""],
    ["正常图数量(张)", s["normal_imgs"] ?? ""],
    ["检测失败(张)", s["fail_imgs"] ?? ""],
    ["耗时(秒)", s["elapsed"] ?? ""],
  );
  $("#sumTab").innerHTML = rows.map(r=>`<tr><td>${r[0]}</td><td>${r[1]}</td></tr>`).join("");

  $("#result").style.display = "block";
  const f = $("#dlFinal");
  f.textContent = "下载全量结果 Excel";
  f.onclick = ()=> downloadChunk("final", s.out_name, f);
  // 分批结果列表（全量下方）
  renderChunks("#resultChunkList", j.chunks || [], true);
  // 顶部进度区的批次列表也保留，方便对照
  renderChunks("#chunkList", j.chunks || [], false);
  $("#start").disabled = false;
}
</script>
</body>
</html>"""


# ════════════════════════════════════════════════════════════════
#  本地服务（stdlib http.server，无需 FastAPI，打包体积小）
# ════════════════════════════════════════════════════════════════
_STATE = {"done": 0, "total": 0, "msg": "", "status": "idle",
          "result_b64": None, "summary": None, "error": None,
          # ── 分批（v4）──
          "chunks": [],          # [{index,start_row,end_row,rows,any,all,white,normal,fail,name,b64}]
          "chunk_total": 0,
          "current_chunk": 0,
          "chunk_size": 0,
          "final_name": "", "final_b64": None,
          "tmpdir": "", "input_path": "", "rules": ["any", "all"], "cols": []}
_STATE_LOCK = threading.Lock()


def _set_state(**kw):
    with _STATE_LOCK:
        _STATE.update(kw)


def _append_chunk(info):
    """线程安全地追加一个已完成批次。"""
    with _STATE_LOCK:
        _STATE["chunks"] = list(_STATE.get("chunks") or []) + [info]


def _get_state():
    with _STATE_LOCK:
        st = dict(_STATE)
    # 剥离大体积 base64，避免每次轮询传输；下载时走 /api/download
    st["chunks"] = [{k: v for k, v in c.items() if k != "b64"}
                    for c in (st.get("chunks") or [])]
    st.pop("final_b64", None)
    st.pop("result_b64", None)
    return st


def _get_b64(kind):
    """kind='final' 取全量，或整数批次号取该批。"""
    with _STATE_LOCK:
        if kind == "final":
            return _STATE.get("final_b64"), _STATE.get("final_name")
        try:
            idx = int(kind)
        except Exception:
            return None, None
        for c in (_STATE.get("chunks") or []):
            if c.get("index") == idx:
                return c.get("b64"), c.get("name")
    return None, None


def _export_partial_now():
    """
    基于当前任务**最新写出的进度文件**，导出「已完成部分」结果 Excel。
    用于长任务执行中途先取走已完成的行，不必等全部跑完。
    返回 (name, b64, summary)；失败抛异常。
    """
    import glob
    with _STATE_LOCK:
        tmpdir = _STATE.get("tmpdir") or ""
        in_path = _STATE.get("input_path") or ""
        rules = tuple(_STATE.get("rules") or ("any", "all"))
    if not tmpdir or not os.path.isdir(tmpdir):
        raise RuntimeError("当前没有运行中的任务")
    if not os.path.exists(in_path):
        raise RuntimeError("任务输入文件已清理（任务可能已结束），请直接下载完整结果")
    files = glob.glob(os.path.join(tmpdir, "*.progress.json"))
    if not files:
        raise RuntimeError("尚无进度文件（任务刚启动或进度尚未落盘），请稍后再试")
    latest = max(files, key=os.path.getmtime)
    out = os.path.join(tmpdir, "已完成部分.xlsx")
    s = export_from_progress(in_path, latest, output=out, rules=rules)
    with open(out, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    name = f"空白图结果_已完成{s['rows_done']}行.xlsx"
    return name, b64, s


def _find_free_port(start=7862, max_tries=20):
    import socket
    for p in range(start, start + max_tries):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(("127.0.0.1", p))
            s.close()
            return p
        except OSError:
            continue
    return start


def _b64_to_bytes(data: str) -> bytes:
    return base64.b64decode(data)


def _run_detect(filename, data_b64, sheet, cols, rules, workers, threshold,
                chunk_size=1000):
    # 清理上一次任务留下的临时目录（批次文件仅供当次下载）
    with _STATE_LOCK:
        prev = _STATE.get("tmpdir") or ""
    if prev and os.path.isdir(prev):
        try:
            import shutil
            shutil.rmtree(prev, ignore_errors=True)
        except Exception:
            pass

    tmpdir = tempfile.mkdtemp(prefix="whimg_")
    in_path = os.path.join(tmpdir, "input.xlsx")
    out_path = os.path.join(tmpdir, "空白图结果_全量.xlsx")
    src_name = os.path.splitext(os.path.basename(filename or "input"))[0]
    with open(in_path, "wb") as f:
        f.write(_b64_to_bytes(data_b64))
    # rules 可能是 list 或 str（兼容旧客户端），统一为 tuple
    if isinstance(rules, str):
        rules = tuple(r.strip() for r in rules.split(",") if r.strip() in ("any", "all"))
    rules = tuple(rules or ("any", "all"))
    try:
        chunk_size = int(chunk_size or 0) or 1000
        total_rows = count_data_rows(in_path, sheet=sheet)
        chunk_total = (total_rows + chunk_size - 1) // chunk_size if total_rows else 1
        chunk_total = max(1, chunk_total)
        _set_state(tmpdir=tmpdir, input_path=in_path, rules=list(rules),
                   cols=list(cols or []), chunks=[], chunk_total=chunk_total,
                   current_chunk=0, chunk_size=chunk_size, final_b64=None,
                   status="running",
                   msg=f"读取到 {total_rows} 行，按每 {chunk_size} 行共 {chunk_total} 批执行")

        def cb(d, t, m):
            _set_state(done=d, total=t, msg=m, status="running")

        def chunk_done(ci, ct, s, path):
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            start_row = s.get("offset", 0) + 1
            end_row = s.get("offset", 0) + s.get("total_rows", 0)
            name = (f"{src_name}_空白图结果_批次{ci}.xlsx" if ct > 1
                    else f"{src_name}_空白图结果.xlsx")
            _append_chunk({
                "index": ci, "total": ct,
                "start_row": start_row, "end_row": end_row,
                "rows": s.get("rows_done", 0),
                "any": s.get("noncompliant_rows_any", 0),
                "all": s.get("noncompliant_rows_all", 0),
                "white": s.get("white_imgs", 0),
                "normal": s.get("normal_imgs", 0),
                "fail": s.get("fail_imgs", 0),
                "name": name, "b64": b64,
            })
            _set_state(current_chunk=ci, chunk_total=ct,
                       msg=f"✓ 第 {ci}/{ct} 批已完成（源表第 {start_row}-{end_row} 行），可立即下载")

        s = run_batch_chunked(
            in_path, output=out_path, sheet=sheet,
            cols=",".join(cols) if cols else None,
            workers=workers, threshold=threshold, rules=rules,
            chunk_size=chunk_size, progress_cb=cb, chunk_cb=chunk_done,
        )

        with open(s["output"], "rb") as f:
            final_b64 = base64.b64encode(f.read()).decode("ascii")
        final_name = f"{src_name}_空白图结果_全量.xlsx" if s.get("chunks", 1) > 1 \
            else f"{src_name}_空白图结果.xlsx"
        _set_state(done=s["rows_done"], total=s["rows_done"], status="done",
                   final_b64=final_b64, final_name=final_name,
                   current_chunk=s.get("chunks", 1),
                   summary={"output": s["output"], "rules": s["rules"],
                            "chunks": s.get("chunks", 1),
                            "chunk_size": s.get("chunk_size", chunk_size),
                            "noncompliant_rows_any": s["noncompliant_rows_any"],
                            "noncompliant_rows_all": s["noncompliant_rows_all"],
                            "white_rows": s["white_rows"], "white_imgs": s["white_imgs"],
                            "normal_imgs": s["normal_imgs"], "fail_imgs": s["fail_imgs"],
                            "elapsed": s["elapsed"], "out_name": final_name},
                   msg=f"全部完成 ✓ 共 {s.get('chunks', 1)} 批 / {s['rows_done']} 行 | "
                       f"any 不合规 {s['noncompliant_rows_any']} 行 | "
                       f"all 不合规 {s['noncompliant_rows_all']} 行 —— 可下载全量结果")
    except Exception as e:
        _set_state(status="error", error=str(e), msg=f"[错误] {e}")
    finally:
        try:
            os.remove(in_path)
        except Exception:
            pass


def _launch_web():
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    class Handler(BaseHTTPRequestHandler):
        def _send(self, code, payload, ctype="application/json"):
            if isinstance(payload, (dict, list)):
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                ctype = "application/json; charset=utf-8"
            elif isinstance(payload, str) and payload.startswith(("<!DOCTYPE", "<html")):
                body = payload.encode("utf-8")
                ctype = "text/html; charset=utf-8"
            else:
                body = payload if isinstance(payload, bytes) else str(payload).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _body(self):
            n = int(self.headers.get("Content-Length", 0))
            return json.loads(self.rfile.read(n) or b"{}")

        def do_GET(self):
            if self.path in ("/", "/index.html"):
                self._send(200, PAGE_HTML)
            elif self.path == "/api/progress":
                self._send(200, _get_state())
            elif self.path.startswith("/api/download"):
                # /api/download?chunk=2  → 第 2 批；/api/download?chunk=final → 全量
                from urllib.parse import urlparse, parse_qs
                q = parse_qs(urlparse(self.path).query)
                kind = (q.get("chunk") or ["final"])[0]
                b64, name = _get_b64(kind)
                if not b64:
                    self._send(404, {"error": "结果不存在或已过期，请重新执行检测"})
                else:
                    self._send(200, {"name": name or "空白图结果.xlsx", "b64": b64})
            elif self.path == "/api/export-partial":
                # 运行中随时导出「已完成部分」（基于最新进度文件）
                try:
                    name, b64, s = _export_partial_now()
                    self._send(200, {
                        "name": name, "b64": b64,
                        "summary": {
                            "rows_done": s["rows_done"], "pending_rows": s["pending_rows"],
                            "total_rows": s["total_rows"],
                            "noncompliant_rows_any": s["noncompliant_rows_any"],
                            "noncompliant_rows_all": s["noncompliant_rows_all"],
                            "white_imgs": s["white_imgs"],
                            "normal_imgs": s["normal_imgs"],
                            "fail_imgs": s["fail_imgs"],
                        }})
                except Exception as e:
                    self._send(400, {"error": str(e)})
            else:
                self._send(404, {"error": "not found"})

        def do_POST(self):
            try:
                if self.path == "/api/prepare":
                    d = self._body()
                    data = _b64_to_bytes(d.get("data", ""))
                    import io, openpyxl
                    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
                    ws = wb[wb.sheetnames[0]]
                    first = next(ws.iter_rows(values_only=True), None)
                    wb.close()
                    headers = [(c if c is not None else "") for c in (first or [])]
                    self._send(200, {"headers": headers})
                    return
                if self.path == "/api/run":
                    d = self._body()
                    _set_state(done=0, total=0, msg="开始…", status="running",
                               result_b64=None, summary=None, error=None,
                               chunks=[], chunk_total=0, current_chunk=0,
                               final_b64=None, final_name="")
                    t = threading.Thread(
                        target=_run_detect,
                        args=(d.get("filename"), d.get("data"), d.get("sheet"),
                              d.get("cols", []), d.get("rules", ["any", "all"]),
                              int(d.get("workers", 8)), float(d.get("threshold", 0.99)),
                              int(d.get("chunk_size", 1000) or 1000)),
                        daemon=True)
                    t.start()
                    self._send(200, {"ok": True})
                    return
            except Exception as e:
                self._send(500, {"error": str(e)})
                return
            self._send(404, {"error": "not found"})

        def log_message(self, *a):
            pass

    port = _find_free_port()
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://localhost:{port}/"
    print(f"[web] 服务已启动: {url}", flush=True)
    try:
        webbrowser.open(url)
    except Exception:
        pass
    print(f"[web] 若浏览器未自动打开，请手动访问: {url}", flush=True)
    srv.serve_forever()


# ════════════════════════════════════════════════════════════════
#  无界面模式 + 入口
# ════════════════════════════════════════════════════════════════
def _export_partial(input_path, progress_path, output=None, rule="any,all",
                    cols=None, sheet=None):
    """不跑检测：从进度文件导出「已完成部分」结果 Excel。"""
    rules = tuple(r.strip() for r in str(rule).split(",") if r.strip() in ("any", "all"))
    if not rules:
        rules = ("any", "all")
    print(f"== 从进度文件导出「已完成部分」==\n输入: {input_path}\n进度: {progress_path}")
    try:
        s = export_from_progress(input_path, progress_path, output=output,
                                 rules=rules, cols=cols, sheet=sheet)
    except Exception as e:
        print(f"[错误] {e}")
        return 1
    print(f"\n[完成] 导出: {s['output']}")
    print(f"  已完成 {s['rows_done']} 行 / 未处理 {s['pending_rows']} 行"
          f"（共 {s['total_rows']} 行）")
    print(f"  any 不合规 {s['noncompliant_rows_any']} 行 | "
          f"all 不合规 {s['noncompliant_rows_all']} 行 | "
          f"纯白图 {s['white_imgs']} 张 | 正常图 {s['normal_imgs']} 张 | 失败 {s['fail_imgs']} 张")
    return 0


def _headless(input_path, output=None, limit=None, workers=8, sheet=None,
              resume=False, cols=None, rule="any,all", chunk_size=1000,
              persist_every=300):
    # rule 字符串 → tuple，如 "any,all" / "any" / "all"
    rules = tuple(r.strip() for r in str(rule).split(",") if r.strip() in ("any", "all"))
    if not rules:
        rules = ("any", "all")
    print(f"== 空白图检测（headless）==\n输入: {input_path} | 维度: {rules} | 分批: 每 {chunk_size} 行")
    try:
        def _chunk_done(ci, ct, s, path):
            print(f"[批次完成] 第 {ci}/{ct} 批 → {path} "
                  f"（any {s['noncompliant_rows_any']} / all {s['noncompliant_rows_all']}）",
                  flush=True)

        s = run_batch_chunked(input_path, output=output, limit=limit or 0,
                              workers=workers, sheet=sheet, resume=resume, cols=cols,
                              rules=rules, chunk_size=int(chunk_size or 1000),
                              persist_every=int(persist_every or 0),
                              progress_cb=lambda d, t, m: print(m, flush=True),
                              chunk_cb=_chunk_done)
    except Exception as e:
        print(f"[错误] {e}")
        return 1
    print(f"\n[完成] 全量结果: {s['output']}")
    if s.get("chunks", 1) > 1:
        print(f"[完成] 共 {s['chunks']} 批（每批 {s['chunk_size']} 行），分批文件：")
        for p in s.get("chunk_paths", []):
            print(f"        - {p}")
    rules_label = "、".join(
        "任意一张空白即不合规" if r == "any" else "全部图片均为空白才算不合规"
        for r in s["rules"])
    print(f"  判定维度: {rules_label}")
    print(f"  已处理 {s['rows_done']} 行 | "
          f"any 不合规 {s['noncompliant_rows_any']} 行 | "
          f"all 不合规 {s['noncompliant_rows_all']} 行 | "
          f"纯白图 {s['white_imgs']} 张 | 正常图 {s['normal_imgs']} 张 | 失败 {s['fail_imgs']} 张")
    print(f"  耗时 {s['elapsed']}s")
    return 0


def main():
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith((".xlsx", ".xls")):
        import argparse
        p = argparse.ArgumentParser(description="空白图（白底图）批量识别 · Excel 插件")
        p.add_argument("input", help="输入 Excel 路径")
        p.add_argument("-o", "--output", help="输出 Excel 路径（默认 <输入>_空白图结果.xlsx）")
        p.add_argument("-l", "--limit", type=int, default=None, help="仅处理前 N 行（调试用）")
        p.add_argument("-w", "--workers", type=int, default=8, help="并发线程数（默认 8）")
        p.add_argument("-s", "--sheet", default=None, help="指定工作表名（默认第一个）")
        p.add_argument("-c", "--cols", default=None, help="指定 URL 列名（逗号分隔，默认自动识别）")
        p.add_argument("-r", "--resume", action="store_true", help="断点续跑")
        p.add_argument("--chunk-size", type=int, default=1000,
                       help="分批行数上限（默认 1000；超过则分批跑并逐批产出 Excel，最后合并全量）")
        p.add_argument("--persist-every", type=int, default=300,
                       help="运行中每完成 N 个 URL 落盘一次进度（默认 300；0=只在结束时落盘）")
        p.add_argument("--export-partial", default=None, metavar="PROGRESS_JSON",
                       help="不跑检测：从指定的 *.progress.json 导出「已完成部分」结果 Excel")
        p.add_argument("--rule", default="any,all",
                       help="不合规维度（逗号分隔，可多选）：any=一张即不合规 / all=全部空白才不合规；"
                            "默认 any,all（两种维度同时输出到 Excel 不同列）")
        args = p.parse_args(sys.argv[1:])
        if args.export_partial:
            sys.exit(_export_partial(args.input, args.export_partial, args.output,
                                     args.rule, args.cols, args.sheet))
        sys.exit(_headless(args.input, args.output, args.limit, args.workers,
                           args.sheet, args.resume, args.cols, args.rule,
                           args.chunk_size, args.persist_every))
    _launch_web()


if __name__ == "__main__":
    main()
