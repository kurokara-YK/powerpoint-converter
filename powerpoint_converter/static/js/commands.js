"use strict";
// 操作をサーバに送る・リボン・コピーと貼り付け・配置（整列・回転・グループ・重なり）・キーボード

// ---------------------------------------------------------------- 操作を送る
async function op(name, args, opt = {}) {
  if (!S.deck?.editable && name !== "copy") { toast("この資料は見るだけ（直せるのは .pptx）", true); return null; }
  try {
    const r = await post("/api/op", { op: name, args, version: S.deck.version });
    S.deck.version = r.version;
    if (!r.copied) $("#saved").textContent = `保存した（${r.label}）`;
    if (r.select) S.sel = r.select;
    if (r.slide != null) S.cur = r.slide;
    return r;
  } catch (e) {
    toast(e.message, true, 5000);
    if (e.conflict) { S.slide = null; api(`/api/poll?since=-1&cur=${S.cur}`).then(applyModel); }
    if (!opt.keepVisual) drawSlide(); else loadSlide();
    return null;
  }
}
async function uploadImage(file, replaceId = null) {
  if (!S.deck?.editable) return;
  try {
    const q = `slide=${S.cur}&name=${encodeURIComponent(file.name || "画像")}&version=${encodeURIComponent(S.deck.version)}` +
      (replaceId != null ? `&replace=${replaceId}` : "");
    const r = await api(`/api/image?${q}`, { method: "POST", body: file });
    S.deck.version = r.version; if (r.select) S.sel = r.select;
    $("#saved").textContent = `保存した（${r.label}）`;
  } catch (e) { toast(e.message, true); }
}
function pickImage(replaceId = null) {
  const f = $("#imageFile");
  f.onchange = e => { const x = e.target.files[0]; if (x) uploadImage(x, replaceId); e.target.value = ""; };
  f.click();
}
async function undo() {
  if (S.editing) await commitEdit();
  try {
    const r = await post("/api/undo", { version: S.deck.version }); S.deck.version = r.version; toast("元に戻した: " + r.label);
    if (r.slide != null && r.slide !== S.cur && r.slide < S.deck.slides.length) gotoSlide(r.slide);
  } catch (e) { toast(e.message, true); }
}
async function redo() {
  try { const r = await post("/api/redo", { version: S.deck.version }); S.deck.version = r.version; toast("やり直した: " + r.label); }
  catch (e) { toast(e.message, true); }
}
function duplicate() { if (S.sel.length) op("duplicate", { slide: S.cur, ids: S.sel }); }
function delSel() { if (S.sel.length) { const ids = S.sel; S.sel = []; op("delete", { slide: S.cur, ids }); } }

// ---------------------------------------------------------------- コピーと貼り付け
// この画面でコピーするとき、システムのクリップボードに目印の文字も書く。貼り付けのとき、目印ならこの画面で
// コピーした図形を、画像なら画像を、ほかの文字ならテキスト ボックスを入れる（PowerPoint と同じ）
const CLIP_MARK = "[powerpoint-converter:";
async function copySel(cut, token = null) {
  // token はコピーのイベントから来たとき（クリップボードにはイベントが書く）。ボタンからなら自分で書く
  if (!S.sel.length) return;
  if (!token) { token = `${CLIP_MARK}${Date.now()}]`; navigator.clipboard?.writeText(token).catch(() => { }); }
  const n = S.sel.length;
  const r = await op(cut ? "cut" : "copy", { slide: S.cur, ids: S.sel }, { keepVisual: true });
  if (!r) return;
  S.clipToken = token;
  if (cut) S.sel = []; else toast(`${n} 個の図形をコピーした`);
}
async function pasteInternal() {
  if (!S.clipToken) { toast("コピーした図形が無い", true); return; }
  await op("paste", { slide: S.cur });
}
async function pasteText(t) {
  const lines = t.replace(/\r\n?/g, "\n").replace(/\n+$/, "").split("\n");
  const r = await op("add", { slide: S.cur, kind: "textbox", text: lines[0] });
  if (r && lines.length > 1) await op("text", { slide: S.cur, id: r.select[0], text: lines });
}
async function pasteButton() {
  let t = null;
  try { t = await navigator.clipboard.readText(); } catch { }
  if (t && !t.startsWith(CLIP_MARK) && t.trim()) return pasteText(t);
  return pasteInternal();
}
const canClip = e => S.view === "deck" && S.deck && !isTyping(e.target) && !S.editing && !dialogOpen();
document.addEventListener("copy", e => {
  if (!canClip(e) || !S.sel.length) return;
  e.preventDefault(); const tok = `${CLIP_MARK}${Date.now()}]`;
  e.clipboardData.setData("text/plain", tok); copySel(false, tok);
});
document.addEventListener("cut", e => {
  if (!canClip(e) || !S.sel.length || !S.deck.editable) return;
  e.preventDefault(); const tok = `${CLIP_MARK}${Date.now()}]`;
  e.clipboardData.setData("text/plain", tok); copySel(true, tok);
});
document.addEventListener("paste", e => {
  if (!canClip(e)) return;
  const f = [...(e.clipboardData?.files || [])].find(f => f.type.startsWith("image/"));
  if (f) { e.preventDefault(); return uploadImage(f); }
  const t = e.clipboardData?.getData("text/plain") || "";
  if (t.startsWith(CLIP_MARK)) { e.preventDefault(); return pasteInternal(); }
  if (t.trim()) { e.preventDefault(); return pasteText(t); }
  if (S.clipToken) { e.preventDefault(); pasteInternal(); }
});

// ---------------------------------------------------------------- 配置
function framesFor(fn) {
  const ss = selFrames();
  if (!ss.length) return;
  const frames = ss.map(s => ({ id: s.id, ...fn(s, ss) })).map(f => ({ id: f.id, x: Math.round(f.x), y: Math.round(f.y), w: f.w, h: f.h }));
  return op("frame", { slide: S.cur, frames, move: true });
}
function align(kind) {
  // 1つなら スライドに、2つ以上なら 選んだ図形の範囲に揃える（PowerPoint と同じ）
  const ss = selFrames(); if (!ss.length) return;
  const ref = ss.length === 1 ? { x: 0, y: 0, w: S.deck.size.cx, h: S.deck.size.cy } : union(ss.map(s => s.frame));
  return framesFor(s => {
    const f = { ...s.frame };
    if (kind === "l") f.x = ref.x; if (kind === "c") f.x = ref.x + (ref.w - f.w) / 2; if (kind === "r") f.x = ref.x + ref.w - f.w;
    if (kind === "t") f.y = ref.y; if (kind === "m") f.y = ref.y + (ref.h - f.h) / 2; if (kind === "b") f.y = ref.y + ref.h - f.h;
    return f;
  });
}
function distribute(axis) {
  const ss = selFrames(); if (ss.length < 3) return;
  const P = axis === "h" ? ["x", "w"] : ["y", "h"];
  const sorted = [...ss].sort((a, b) => a.frame[P[0]] - b.frame[P[0]]);
  const first = sorted[0].frame, last = sorted.at(-1).frame;
  const space = last[P[0]] + last[P[1]] - first[P[0]] - sorted.reduce((a, s) => a + s.frame[P[1]], 0);
  const gap = space / (sorted.length - 1);
  const pos = {}; let at = first[P[0]];
  for (const s of sorted) { pos[s.id] = at; at += s.frame[P[1]] + gap; }
  return framesFor(s => ({ ...s.frame, [P[0]]: pos[s.id] }));
}
function alignItems() {
  const n = selFrames().length;
  return [
    { head: n <= 1 ? "スライドに揃える" : "選んだ図形どうしで揃える" },
    { label: "左揃え", key: "", do: () => align("l") }, { label: "左右中央揃え", do: () => align("c") }, { label: "右揃え", do: () => align("r") },
    "-",
    { label: "上揃え", do: () => align("t") }, { label: "上下中央揃え", do: () => align("m") }, { label: "下揃え", do: () => align("b") },
    "-",
    { label: "左右に整列", disabled: n < 3, do: () => distribute("h") }, { label: "上下に整列", disabled: n < 3, do: () => distribute("v") },
  ];
}
function rotateItems() {
  const rot = d => op("rotate", { slide: S.cur, ids: S.sel, deg: ((shapeById(S.sel[0])?.rot || 0) + d + 360) % 360 });
  return [
    { label: "右へ 90 度回転", do: () => rot(90) }, { label: "左へ 90 度回転", do: () => rot(-90) },
    { label: "回転をなくす", do: () => op("rotate", { slide: S.cur, ids: S.sel, deg: 0 }) },
    "-",
    { label: "上下反転", do: () => op("flip", { slide: S.cur, ids: S.sel, axis: "v" }) },
    { label: "左右反転", do: () => op("flip", { slide: S.cur, ids: S.sel, axis: "h" }) },
  ];
}
function orderItems() {
  const z = where => op("zorder", { slide: S.cur, ids: S.sel, where });
  return [{ label: "最前面へ移動", key: "Ctrl+Shift+]", do: () => z("front") }, { label: "前面へ移動", do: () => z("forward") },
    { label: "背面へ移動", do: () => z("backward") }, { label: "最背面へ移動", key: "Ctrl+Shift+[", do: () => z("back") }];
}
function groupSel() { if (S.sel.length >= 2) op("group", { slide: S.cur, ids: S.sel }); }
function ungroupSel() {
  const gs = selShapes().filter(s => s.kind === "grpSp").map(s => s.id);
  if (gs.length) { S.group = null; op("ungroup", { slide: S.cur, ids: gs }); }
}
const ANIM_EFFECTS = [["fade", "開始: フェード"], ["appear", "開始: アピール"], ["fly", "開始: スライドイン（下から）"],
  ["wipe", "開始: ワイプ（下から）"], ["fadeout", "終了: フェード"], ["disappear", "終了: クリア"]];

// ---------------------------------------------------------------- リボン
function renderRibbon() {
  const m = S.deck; if (!m) return;
  const ed = m.editable, has = S.sel.length > 0;
  $("#bUndo").disabled = !m.can_undo; $("#bUndo").title = m.can_undo ? `元に戻す: ${m.undo_label} (Ctrl+Z)` : "元に戻す";
  $("#bRedo").disabled = !m.can_redo; $("#bRedo").title = m.can_redo ? `やり直し: ${m.redo_label} (Ctrl+Y)` : "やり直し";
  for (const b of ["#bNewSlide", "#bTextbox", "#bShape", "#bTable", "#bImage", "#bPaste"]) $(b).disabled = !ed;
  for (const b of ["#bDup", "#bDel", "#bCut"]) $(b).disabled = !ed || !has;
  $("#bCopy").disabled = !has;
  $("#bArrange").disabled = !ed || !has;
}
const SHAPE_GALLERY = [
  ["線", [["line", "╱", "直線"], ["arrow", "↗", "矢印"]]],
  ["四角形", [["rect", "▭", "正方形/長方形"], ["roundRect", "▢", "角丸四角形"], ["snip1Rect", "⬔", "1 つの角を切り取った四角形"]]],
  ["基本図形", [["ellipse", "◯", "楕円"], ["triangle", "△", "二等辺三角形"], ["rtTriangle", "◺", "直角三角形"], ["diamond", "◇", "ひし形"],
    ["pentagon", "⬠", "五角形"], ["hexagon", "⬡", "六角形"], ["octagon", "⯃", "八角形"], ["parallelogram", "▱", "平行四辺形"],
    ["trapezoid", "⏢", "台形"], ["can", "⛁", "円柱"], ["cube", "⧈", "直方体"], ["donut", "◎", "ドーナツ"], ["plus", "✚", "十字形"],
    ["heart", "♡", "ハート"], ["sun", "☼", "太陽"], ["moon", "☾", "月"], ["cloud", "☁", "雲"], ["smileyFace", "☺", "スマイル"],
    ["noSmoking", "⊘", "禁止"], ["leftBrace", "{", "中かっこ（左）"], ["rightBrace", "}", "中かっこ（右）"],
    ["leftBracket", "[", "大かっこ（左）"], ["rightBracket", "]", "大かっこ（右）"]]],
  ["ブロック矢印", [["rightArrow", "➡", "右矢印"], ["leftArrow", "⬅", "左矢印"], ["upArrow", "⬆", "上矢印"], ["downArrow", "⬇", "下矢印"],
    ["leftRightArrow", "⬌", "左右矢印"], ["chevron", "❯", "山形"], ["homePlate", "⭆", "五方向"], ["curvedRightArrow", "↪", "右カーブ矢印"]]],
  ["吹き出し", [["wedgeRectCallout", "🗆", "四角形の吹き出し"], ["wedgeRoundRectCallout", "💬", "角丸四角形の吹き出し"],
    ["wedgeEllipseCallout", "🗨", "円形の吹き出し"], ["cloudCallout", "💭", "雲形の吹き出し"]]],
  ["星", [["star5", "★", "星 5 pt"], ["star8", "✴", "星 8 pt"]]],
  ["フローチャート", [["flowChartProcess", "▭", "処理"], ["flowChartDecision", "◇", "判断"], ["flowChartTerminator", "⬭", "端子"],
    ["flowChartDocument", "🗎", "書類"], ["flowChartMagneticDisk", "⛁", "磁気ディスク"]]],
];
$("#bUndo").onclick = undo; $("#bRedo").onclick = redo;
$("#bPaste").onclick = () => pasteButton();
$("#bCut").onclick = () => copySel(true);
$("#bCopy").onclick = () => copySel(false);
$("#bNewSlide").onclick = e => {
  const cur = S.slide ? S.deck.layouts.find(l => l.name === S.slide.layout) : null;
  menuBelow(e.currentTarget, [{ label: "今のスライドと同じレイアウト", key: "Ctrl+M", do: () => newSlide() }, "-", { head: "レイアウト" },
    ...S.deck.layouts.map(l => ({ label: l.name, checked: cur && l.part === cur.part, do: () => newSlide(l.part) }))]);
};
$("#bTextbox").onclick = () => op("add", { slide: S.cur, kind: "textbox" }).then(async r => { if (!r) return; await waitShape(r.select[0]); startEdit(r.select[0]); });
$("#bShape").onclick = e => {
  const items = [];
  for (const [cat, list] of SHAPE_GALLERY) {
    items.push({ head: cat });
    items.push({ el: h("div", { class: "gallery" }, list.map(([k, g, name]) => h("button", { title: name, onclick: () => { closeMenus(0); op("add", { slide: S.cur, kind: k }); } }, g))) });
  }
  menuBelow(e.currentTarget, items);
};
$("#bTable").onclick = e => {
  const lab = h("div", { class: "hd" }, "表の大きさを選ぶ");
  const grid = h("div", { class: "tpick" });
  for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
    const cell = h("div", { "data-r": r, "data-c": c });
    cell.addEventListener("mouseenter", () => { $$("div", grid).forEach(x => x.classList.toggle("on", +x.dataset.r <= r && +x.dataset.c <= c)); lab.textContent = `${r + 1} 行 × ${c + 1} 列`; });
    cell.addEventListener("click", () => { closeMenus(0); op("add", { slide: S.cur, kind: "table", rows: r + 1, cols: c + 1 }); });
    grid.append(cell);
  }
  menuBelow(e.currentTarget, [{ el: lab }, { el: grid }]);
};
$("#bImage").onclick = () => pickImage();
$("#bArrange").onclick = e => menuBelow(e.currentTarget, [
  { label: "配置", sub: alignItems() }, { label: "回転と反転", sub: rotateItems() }, { label: "重なりの順番", sub: orderItems() }, "-",
  { label: "グループ化", key: "Ctrl+G", disabled: S.sel.length < 2, do: groupSel },
  { label: "グループ解除", key: "Ctrl+Shift+G", disabled: !selShapes().some(s => s.kind === "grpSp"), do: ungroupSel },
]);
$("#bDup").onclick = () => duplicate();
$("#bDel").onclick = () => delSel();
$("#bFind").onclick = () => openFind(false);
$("#bShowStart").onclick = () => startShow(false);
$("#bShowCur").onclick = () => startShow(true);
$("#bDownload").onclick = e => menuBelow(e.currentTarget, [
  { label: "PowerPoint（.pptx）", do: () => { location.href = "/download"; } },
  { label: "PDF（非表示のスライドは除く）", do: () => { toast("PDF を作っている（大きい資料は十数秒かかる）"); location.href = "/download?format=pdf"; } },
  { label: "このスライドの画像（PNG）", disabled: S.deck.slides[S.cur]?.status !== "ready", do: () => saveSlideImage() },
]);
async function waitShape(id) { for (let i = 0; i < 50; i++) { if (shapeById(id)) return; await sleep(100); } }

// ---------------------------------------------------------------- キーボード
let nudge = null;
document.addEventListener("keydown", async e => {
  if (!$("#show").classList.contains("hidden")) { if (e.key === "Escape") endShow(); return; }
  if (S.view !== "deck" || !S.deck || dialogOpen()) return;
  if (isTyping(e.target)) return;
  const ctrl = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
  if (e.key === "F5") { e.preventDefault(); return startShow(e.shiftKey); }
  if (ctrl && k === "z") { e.preventDefault(); return e.shiftKey ? redo() : undo(); }
  if (ctrl && k === "y") { e.preventDefault(); return redo(); }
  if (ctrl && k === "m") { e.preventDefault(); return newSlide(); }
  if (ctrl && k === "d") { e.preventDefault(); return duplicate(); }
  if (ctrl && k === "f") { e.preventDefault(); return openFind(false); }
  if (ctrl && k === "h") { e.preventDefault(); return openFind(true); }
  if (ctrl && k === "g") { e.preventDefault(); return e.shiftKey ? ungroupSel() : groupSel(); }
  if (ctrl && k === "a") { e.preventDefault(); return select(topShapes().filter(s => s.frame && !s.hidden).map(s => s.id)); }
  if (ctrl && S.sel.length && ["b", "i", "u"].includes(k)) { e.preventDefault(); return applyText({ [k]: "toggle" }); }
  if (ctrl && S.sel.length && (e.key === "]" || e.key === "[" || e.key === "}" || e.key === "{")) {
    e.preventDefault();
    if (e.shiftKey) return op("zorder", { slide: S.cur, ids: S.sel, where: e.key === "]" || e.key === "}" ? "front" : "back" });
    return applyText({ grow: e.key === "]" ? 2 : -2 });
  }
  if (ctrl && S.sel.length && ["e", "l", "r", "j"].includes(k)) { e.preventDefault(); return applyText({ algn: { e: "ctr", l: "l", r: "r", j: "just" }[k] }); }
  if (e.key === "Tab" && S.slide) {
    e.preventDefault();
    const list = topShapes().filter(s => s.frame && !s.hidden);
    if (!list.length) return;
    const i = list.findIndex(s => s.id === S.sel[0]);
    return select([list[(i + (e.shiftKey ? -1 : 1) + list.length) % list.length].id]);
  }
  if (e.key === "Escape" && S.editing) return commitEdit();
  if (e.key === "Escape") { if (S.group != null) { S.sel = [S.group]; S.group = null; } else S.sel = []; drawSelection(); renderPane(); return; }
  if ((e.key === "Delete" || e.key === "Backspace") && S.sel.length) { e.preventDefault(); return delSel(); }
  if ((e.key === "Enter" || e.key === "F2") && S.sel.length === 1) {
    const s = shapeById(S.sel[0]);
    if (s?.table) { e.preventDefault(); return openTableEditor(s.id); }
    if (canText(s)) { e.preventDefault(); return startEdit(S.sel[0]); }
  }
  if (e.key === "PageDown" || (!S.sel.length && (e.key === "ArrowDown" || e.key === "ArrowRight"))) { e.preventDefault(); return gotoSlide(S.cur + 1); }
  if (e.key === "PageUp" || (!S.sel.length && (e.key === "ArrowUp" || e.key === "ArrowLeft"))) { e.preventDefault(); return gotoSlide(S.cur - 1); }
  if (e.key === "Home" && !S.sel.length) { e.preventDefault(); return gotoSlide(0); }
  if (e.key === "End" && !S.sel.length) { e.preventDefault(); return gotoSlide(S.deck.slides.length - 1); }
  if (S.sel.length && e.key.startsWith("Arrow") && S.deck.editable) {
    e.preventDefault();
    const step = ctrl ? 3600 : e.shiftKey ? 180000 : 36000;   // 0.01 / 0.5 / 0.1 cm
    const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
    if (!nudge) nudge = { ids: [...S.sel], dx: 0, dy: 0 };
    nudge.dx += d[0]; nudge.dy += d[1];
    moveVisual(nudge.ids, nudge.dx, nudge.dy);
    clearTimeout(nudge.t);
    nudge.t = setTimeout(async () => {
      const n = nudge; nudge = null;
      const shapes = n.ids.map(shapeById).filter(s => s?.frame);
      const frames = shapes.map(s => ({ id: s.id, x: s.frame.x + n.dx, y: s.frame.y + n.dy, w: s.frame.w, h: s.frame.h }));
      for (const s of shapes) s.frame = { ...s.frame, x: s.frame.x + n.dx, y: s.frame.y + n.dy };
      await op("frame", { slide: S.cur, frames, move: true }, { keepVisual: true });
      drawSlide();
    }, 450);
  }
});

// ---------------------------------------------------------------- 画像のドロップ・拡大縮小・パネルの開閉
$("#stage").addEventListener("dragover", e => { if ([...e.dataTransfer.items].some(i => i.kind === "file")) e.preventDefault(); });
$("#stage").addEventListener("drop", e => { e.preventDefault(); const f = [...e.dataTransfer.files].find(f => f.type.startsWith("image/")); if (f) uploadImage(f); });
$("#stage").addEventListener("wheel", e => {
  if (!e.ctrlKey) return;
  e.preventDefault();
  const base = S.zoom === "fit" ? fitScale() : S.zoom;
  S.zoom = Math.max(fitScaleBase() * .2, Math.min(fitScaleBase() * 4, base * (e.deltaY < 0 ? 1.1 : 1 / 1.1)));
  drawSlide();
}, { passive: false });
$("#zIn").onclick = () => { S.zoom = (S.zoom === "fit" ? fitScale() : S.zoom) * 1.2; drawSlide(); };
$("#zOut").onclick = () => { S.zoom = (S.zoom === "fit" ? fitScale() : S.zoom) / 1.2; drawSlide(); };
$("#zFit").onclick = () => { S.zoom = "fit"; drawSlide(); };
function togglePanel(cls, key) {
  const w = $("#work"); w.classList.toggle(cls); store.set(key, w.classList.contains(cls));
  if (S.zoom === "fit") setTimeout(drawSlide, 0);
}
$("#tThumbs").onclick = () => togglePanel("nothumbs", "nothumbs");
$("#tPane").onclick = () => togglePanel("nopane", "nopane");
if (store.get("nothumbs", false)) $("#work").classList.add("nothumbs");
if (store.get("nopane", false)) $("#work").classList.add("nopane");
let rsz; window.addEventListener("resize", () => { clearTimeout(rsz); rsz = setTimeout(() => { if (S.view === "deck" && S.zoom === "fit") drawSlide(); }, 80); });
