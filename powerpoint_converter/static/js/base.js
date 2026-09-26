"use strict";
// 共通の道具（要素の作成・サーバとのやり取り・通知・メニュー・ダイアログ）と、画面の状態 S

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const sleep = ms => new Promise(r => setTimeout(r, ms));
const EMU_CM = 360000;
const EMU_PT = 12700;

function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === "class") el.className = v;
    else if (k === "style" && typeof v === "object") Object.assign(el.style, v);
    else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
    else if (k === "html") el.innerHTML = v;
    else if (k === "value") el.value = v;
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const c of kids.flat()) if (c != null && c !== false) el.append(c.nodeType ? c : document.createTextNode(c));
  return el;
}

async function api(url, opt = {}) {
  // 書き換える要求には X-PC を付ける（サーバはこれが無い POST を断る。ほかのサイトから送らせないため）
  if (opt.method === "POST") opt.headers = { ...(opt.headers || {}), "X-PC": "1" };
  const r = await fetch(url, opt);
  let j = null;
  try { j = await r.json(); } catch { }
  if (!r.ok) { const e = new Error(j?.error || r.statusText); e.conflict = j?.conflict; e.status = r.status; throw e; }
  return j;
}
const post = (url, body) => api(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body || {}) });

let toastTimer;
function toast(msg, err = false, ms = 3200) {
  const t = $("#toast");
  t.textContent = msg; t.className = err ? "err" : ""; clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add("hidden"), ms);
}
function fmtTime(t) {
  const d = new Date(t * 1000), now = new Date(), p = n => String(n).padStart(2, "0");
  return d.toDateString() === now.toDateString() ? `今日 ${p(d.getHours())}:${p(d.getMinutes())}` : `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())}`;
}
function fmtSize(n) { return n > 1e6 ? (n / 1e6).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1e3)) + " KB"; }
const store = {
  get(k, d) { try { const v = localStorage.getItem("pc." + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem("pc." + k, JSON.stringify(v)); } catch { } },
};
const isTyping = t => t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable);

// ---------------------------------------------------------------- 画面の状態
const S = {
  view: "", listPath: "", info: null,
  deck: null, path: "", rev: -1, cur: 0,
  slide: null, lastRender: {},       // part → {fp, meta, frames, map}
  thumbs: {},                        // part → 最後に描けたサムネイルの URL
  sel: [], group: null,
  editing: null, pane: store.get("pane", "format"),
  zoom: "fit", scale: 1, W: 0, H: 0,
  pollToken: 0, slideToken: 0, animSel: null,
  clipToken: null,                   // この画面でコピーした印（システムのクリップボードにも書く）
  fonts: null, find: { q: "", results: [], i: -1, case: false },
};

// ---------------------------------------------------------------- メニュー
// items: {label, key, disabled, checked, do, sub: [...]} または "-"（区切り）・{head: "見出し"}・{el: 要素}
function closeMenus(level = 0) { $$(".menu").forEach(m => { if (+m.dataset.level >= level) m.remove(); }); }
function menu(x, y, items, level = 0) {
  closeMenus(level);
  const m = h("div", { class: "menu", "data-level": level });
  for (const it of items) {
    if (!it) continue;
    if (it === "-") { m.append(h("hr")); continue; }
    if (it.head) { m.append(h("div", { class: "hd" }, it.head)); continue; }
    if (it.el) { m.append(it.el); continue; }
    const b = h("button", { disabled: it.disabled },
      h("span", { style: { width: "14px", display: "inline-block" } }, it.checked ? "✓" : ""), it.label,
      it.sub ? h("span", { class: "k" }, "▸") : it.key ? h("span", { class: "k" }, it.key) : null);
    if (it.sub) {
      const open = () => {
        $$(".menu button.open", m).forEach(x => x.classList.remove("open")); b.classList.add("open");
        const r = b.getBoundingClientRect(); menu(r.right - 2, r.top - 6, it.sub, level + 1);
      };
      b.addEventListener("mouseenter", open); b.addEventListener("click", open);
    } else {
      b.addEventListener("mouseenter", () => { closeMenus(level + 1); $$(".menu button.open", m).forEach(x => x.classList.remove("open")); });
      b.addEventListener("click", () => { closeMenus(0); it.do?.(); });
    }
    m.append(b);
  }
  document.body.append(m);
  const r = m.getBoundingClientRect();
  m.style.left = Math.max(4, Math.min(x, innerWidth - r.width - 6)) + "px";
  m.style.top = Math.max(4, Math.min(y, innerHeight - r.height - 6)) + "px";
  return m;
}
document.addEventListener("pointerdown", e => { if (!e.target.closest(".menu")) closeMenus(0); }, true);
function menuBelow(btn, items) { const r = btn.getBoundingClientRect(); return menu(r.left, r.bottom + 2, items); }

// ---------------------------------------------------------------- ダイアログ
function dialog(title, body, buttons = [{ label: "閉じる" }], opts = {}) {
  const close = () => { bg.remove(); opts.onclose?.(); };
  const ft = h("div", { class: "ft" }, buttons.map(b => h("button", { class: b.primary ? "primary" : "", onclick: async () => {
    if (b.do && (await b.do()) === false) return;
    close();
  } }, b.label)));
  const dlg = h("div", { class: "dlg", style: opts.width ? { width: opts.width } : null },
    h("div", { class: "hd" }, title, h("span", { class: "sp" }), h("button", { onclick: close, title: "閉じる (Esc)" }, "✕")),
    h("div", { class: "bd" }, body), ft);
  const bg = h("div", { class: "modal", onpointerdown: e => { if (e.target === bg) close(); } }, dlg);
  bg.addEventListener("keydown", e => { if (e.key === "Escape") { e.stopPropagation(); close(); } });
  document.body.append(bg);
  return { close, el: dlg };
}
const dialogOpen = () => !!$(".modal");
