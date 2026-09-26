"use strict";
// スライドの表示（LibreOffice が描いた背景と図形の絵を重ねる）と、マウスでの選択・移動・大きさ・回転

// ---------------------------------------------------------------- スライドの中身
function framesOf(shapes) {
  const f = {};
  const walk = (ms, top) => ms.forEach(s => { f[s.id] = { frame: s.frame, rot: s.rot || 0, top: top ?? s.id }; walk(s.children || [], top ?? s.id); });
  walk(shapes);
  return f;
}
function mapRender(meta, shapes) {
  // LibreOffice の図形（上から i 番目）→ PPTX の最上位の図形の id
  const frames = framesOf(shapes), map = [];
  const idOf = s => { const m = /^pc:(\d+)$/.exec(s.desc || ""); return m ? +m[1] : null; };
  const firstId = s => idOf(s) ?? (s.children || []).map(firstId).find(x => x != null) ?? null;
  meta.shapes.forEach((s, i) => { const id = firstId(s); map[i] = id != null && frames[id] ? frames[id].top : null; });
  return map;
}
async function loadSlide() {
  const my = ++S.slideToken, i = S.cur;
  let sm;
  try { sm = await api("/api/slide?i=" + i); } catch (e) { if (my === S.slideToken) toast(e.message, true); return; }
  if (my !== S.slideToken || i !== S.cur) return;
  if (sm.render && sm.status === "ready") {
    S.lastRender[sm.part] = { fp: sm.fp, meta: sm.render, frames: framesOf(sm.shapes), map: mapRender(sm.render, sm.shapes) };
  }
  S.slide = sm;
  const ids = new Set(allShapes().map(s => s.id));
  S.sel = S.sel.filter(id => ids.has(id));
  if (S.group != null && !ids.has(S.group)) S.group = null;
  drawSlide(); renderPane(); renderNotes(); renderStatus();
  if (typeof onSlideLoaded === "function") onSlideLoaded();
}
function allShapes(list = S.slide?.shapes || []) { const out = []; const w = ms => ms.forEach(s => { out.push(s); w(s.children || []); }); w(list); return out; }
function shapeById(id) { return allShapes().find(s => s.id === id); }
function topShapes() { return S.group != null ? (shapeById(S.group)?.children || []) : (S.slide?.shapes || []); }
function topOf(id) { return framesOf(S.slide?.shapes || [])[id]?.top ?? id; }

// ---------------------------------------------------------------- 大きさ
function fitScale() {
  const st = $("#stage"), m = S.deck;
  if (!m) return 1;
  const aw = st.clientWidth - 48, ah = st.clientHeight - 48;
  return Math.max(1e-6, Math.min(aw / m.size.cx, ah / m.size.cy));
}
function fitScaleBase() { const m = S.deck; return m ? 960 / m.size.cx : 1; }
function layout() {
  const m = S.deck; if (!m) return;
  S.scale = S.zoom === "fit" ? fitScale() : S.zoom;
  S.W = Math.round(m.size.cx * S.scale); S.H = Math.round(m.size.cy * S.scale);
  const sl = $("#slide"); sl.style.width = S.W + "px"; sl.style.height = S.H + "px";
  $("#stZoom").textContent = Math.round(S.scale / fitScaleBase() * 100) + "%";
}
const px = emu => emu * S.scale;
function box(f) { return { left: px(f.x) + "px", top: px(f.y) + "px", width: Math.max(1, px(f.w)) + "px", height: Math.max(1, px(f.h)) + "px" }; }
function moveBox(b, of, nf) {
  const sx = of.w ? nf.w / of.w : 1, sy = of.h ? nf.h / of.h : 1;
  return { x: nf.x + (b.x - of.x) * sx, y: nf.y + (b.y - of.y) * sy, w: b.w * sx, h: b.h * sy };
}
function union(fs) {
  const x0 = Math.min(...fs.map(f => f.x)), y0 = Math.min(...fs.map(f => f.y));
  const x1 = Math.max(...fs.map(f => f.x + f.w)), y1 = Math.max(...fs.map(f => f.y + f.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

// ---------------------------------------------------------------- 描く
function drawSlide() {
  const sl = $("#slide");
  const pv = $("#preview");
  // 文字を編集中なら、描き直しの前後で入力欄のフォーカスと選択範囲を保つ（外して付け直すと失われるため）
  const ed = S.editing, focused = ed && document.activeElement === ed.ta;
  const selRange = ed ? [ed.ta.selectionStart, ed.ta.selectionEnd] : null;
  layout();
  sl.innerHTML = "";
  const sm = S.slide;
  if (!sm) return;
  const clip = h("div", { class: "clip" }); sl.append(clip);
  const R = S.lastRender[sm.part];
  const cur = framesOf(sm.shapes);
  const ready = sm.status === "ready" && R && R.fp === sm.fp;
  $("#rendering").classList.toggle("hidden", ready || sm.status.startsWith?.("error"));
  if (R) {
    const emuPer = S.deck.size.cx / R.meta.page.w;
    clip.append(h("img", { class: "layer", src: `/render/${R.fp}/bg.png`, style: { left: 0, top: 0, width: S.W + "px", height: S.H + "px" } }));
    R.meta.shapes.forEach((s, i) => {
      const id = R.map[i];
      if (!s.png) return;
      if (id != null && (!cur[id] || shapeById(id)?.hidden)) return;      // 消された・隠された図形
      let b = { x: s.bound.x * emuPer, y: s.bound.y * emuPer, w: s.bound.w * emuPer, h: s.bound.h * emuPer };
      const of = id != null ? R.frames[id]?.frame : null, nf = id != null ? cur[id]?.frame : null;
      if (!ready && of && nf) b = moveBox(b, of, nf);    // 描き直しを待つ間は、前の絵を新しい位置に置く
      clip.append(h("img", { class: "layer shp", src: `/render/${R.fp}/${s.png}`, "data-id": id ?? "",
        style: { left: px(b.x) + "px", top: px(b.y) + "px", width: px(b.w) + "px", height: px(b.h) + "px" } }));
    });
    // 空のプレースホルダー（PowerPoint と同じく点線の枠と案内を出す）
    R.meta.shapes.forEach((s, i) => {
      const id = R.map[i];
      if (!s.empty || id == null || !cur[id]?.frame) return;
      const sh = shapeById(id); if (!sh || (sh.text || []).some(t => t)) return;
      const t = sh.ph?.type;
      const hint = t === "title" || t === "ctrTitle" ? "クリックしてタイトルを入力" : t === "pic" ? "図" : "クリックしてテキストを入力";
      clip.append(h("div", { class: "ph", "data-id": id, style: box(cur[id].frame) }, hint));
    });
  }
  // まだ描けていない新しい図形は、点線の枠と文字で仮に出す
  const drawn = new Set(R ? R.map.filter(x => x != null) : []);
  for (const s of sm.shapes) {
    if (!s.frame || drawn.has(s.id) || s.hidden || (R && R.fp === sm.fp)) continue;
    clip.append(h("div", { class: "ghost", "data-id": s.id, style: box(s.frame) }, (s.text || []).join("\n").replace(/\v/g, "\n")));
  }
  // 隠した図形は、選択ウィンドウで選んだときだけ点線で見せる
  for (const s of sm.shapes) if (s.hidden && s.frame && S.sel.includes(s.id)) clip.append(h("div", { class: "hiddenbox", style: box(s.frame) }));
  if (!R) clip.append(h("div", { class: "ph", style: { left: 0, top: 0, width: "100%", height: "100%", border: "none", fontSize: "16px" } },
    sm.status.startsWith?.("error") ? "描けなかった: " + sm.status : "LibreOffice で描いている…"));
  if (R && ready) R.meta.shapes.forEach((s, i) => {     // 文字のあふれ
    if (!s.overflow || R.map[i] == null) return;
    const f = cur[R.map[i]]?.frame; if (!f) return;
    clip.append(h("div", { class: "warnbox", title: `文字が枠から ${(s.overflow / 100).toFixed(1)} mm あふれている`, style: box(f) }));
  });
  drawAnnotations(sl);
  drawSelection();
  if (pv) { Object.assign(pv.style, { width: S.W + "px", height: S.H + "px" }); sl.append(pv); }
  if (S.editing) {
    repositionEdit();
    if (focused) { S.editing.ta.focus(); S.editing.ta.setSelectionRange(...selRange); }
  }
}
function drawAnnotations(sl) {
  if (S.pane !== "anim" || !S.slide) return;
  const seen = {}, frames = framesOf(S.slide.shapes);
  for (const e of S.slide.anims) {
    const s = shapeById(e.spid); if (!s) continue;
    const f = s.frame || shapeById(frames[e.spid]?.top)?.frame; if (!f) continue;
    seen[e.spid] = seen[e.spid] || 0;
    sl.append(h("div", { class: "anno" + (S.animSel === e.n ? " sel" : ""), style: { left: px(f.x) + "px", top: px(f.y) + seen[e.spid] * 17 + "px" } }, e.step || "0"));
    seen[e.spid]++;
  }
}

// ---------------------------------------------------------------- 選択の枠
function selShapes() { return S.sel.map(shapeById).filter(Boolean); }
function selFrames() { return selShapes().filter(s => s.frame); }
function drawSelection() {
  const sl = $("#slide");
  $$(".selbox,.grpbox", sl).forEach(e => e.remove());
  if (S.group != null) { const g = shapeById(S.group); if (g?.frame) sl.append(h("div", { class: "grpbox", style: box(g.frame) })); }
  const ss = selFrames();
  if (ss.length === 1) {
    const s = ss[0], f = s.frame;
    const el = h("div", { class: "selbox", style: { ...box(f), transform: s.rot ? `rotate(${s.rot}deg)` : "" } });
    const hs = s.rot ? [] : [["nw", 0, 0], ["n", .5, 0], ["ne", 1, 0], ["e", 1, .5], ["se", 1, 1], ["s", .5, 1], ["sw", 0, 1], ["w", 0, .5]];
    const cursors = { nw: "nwse-resize", se: "nwse-resize", ne: "nesw-resize", sw: "nesw-resize", n: "ns-resize", s: "ns-resize", e: "ew-resize", w: "ew-resize" };
    if (S.deck.editable && !S.editing) {
      for (const [d, fx, fy] of hs) {
        const hd = h("div", { class: "h", "data-h": d, style: { left: fx * 100 + "%", top: fy * 100 + "%", cursor: cursors[d] } });
        hd.addEventListener("pointerdown", e => startResize(e, d));
        el.append(hd);
      }
      if (s.kind !== "graphicFrame") {
        el.append(h("div", { class: "rotline" }));
        const r = h("div", { class: "h rot", style: { left: "50%", top: "-22px" } });
        r.addEventListener("pointerdown", e => startRotate(e));
        el.append(r);
      }
    }
    sl.append(el);
  } else {
    for (const s of ss) sl.append(h("div", { class: "selbox multi", style: { ...box(s.frame), transform: s.rot ? `rotate(${s.rot}deg)` : "" } }));
  }
  renderRibbon();
}
function select(ids) { S.sel = ids; drawSelection(); renderPane(); }

// ---------------------------------------------------------------- 当たり判定
const alphaCache = new Map();
function alphaAt(img, x, y) {
  // 図形の絵の、その点の不透明度（0〜255）。透明なところをクリックしたら下の図形を選ぶため
  let c = alphaCache.get(img.src);
  if (!c) {
    if (!img.complete || !img.naturalWidth) return 255;
    const cv = document.createElement("canvas"); cv.width = img.naturalWidth; cv.height = img.naturalHeight;
    const ctx = cv.getContext("2d", { willReadFrequently: true }); ctx.drawImage(img, 0, 0);
    try { c = ctx.getImageData(0, 0, cv.width, cv.height); } catch { return 255; }
    alphaCache.set(img.src, c);
    if (alphaCache.size > 300) alphaCache.delete(alphaCache.keys().next().value);
  }
  const r = img.getBoundingClientRect();
  const ix = Math.floor((x - r.left) / r.width * c.width), iy = Math.floor((y - r.top) / r.height * c.height);
  let best = 0;
  for (let dy = -3; dy <= 3; dy++) for (let dx = -3; dx <= 3; dx++) {
    const xx = ix + dx, yy = iy + dy;
    if (xx < 0 || yy < 0 || xx >= c.width || yy >= c.height) continue;
    best = Math.max(best, c.data[(yy * c.width + xx) * 4 + 3]);
  }
  return best;
}
function inFrame(f, rot, ex, ey) {
  if (!f) return false;
  let x = ex, y = ey;
  if (rot) {
    const cx = f.x + f.w / 2, cy = f.y + f.h / 2, a = -rot * Math.PI / 180, dx = ex - cx, dy = ey - cy;
    x = cx + dx * Math.cos(a) - dy * Math.sin(a); y = cy + dx * Math.sin(a) + dy * Math.cos(a);
  }
  const pad = 4 / S.scale;
  return x >= f.x - pad && x <= f.x + f.w + pad && y >= f.y - pad && y <= f.y + f.h + pad;
}
function toEmu(e) { const r = $("#slide").getBoundingClientRect(); return { x: (e.clientX - r.left) / S.scale, y: (e.clientY - r.top) / S.scale }; }
function hitTest(e) {
  const p = toEmu(e), list = topShapes();
  for (const id of S.sel) { const s = shapeById(id); if (s && !s.hidden && inFrame(s.frame, s.rot, p.x, p.y) && list.includes(s)) return s; }
  for (let i = list.length - 1; i >= 0; i--) {
    const s = list[i];
    if (!s.frame || s.hidden) continue;
    if (!inFrame(s.frame, s.rot, p.x, p.y)) continue;
    const hasText = (s.text || []).some(t => t) || s.txbox || s.ph || s.table;
    if (hasText || S.group != null) return s;
    const imgs = $$(`#slide img.shp[data-id="${s.id}"]`);
    if (!imgs.length) return s;
    if (imgs.some(img => { const r = img.getBoundingClientRect(); return e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom && alphaAt(img, e.clientX, e.clientY) > 20; })) return s;
  }
  return null;
}

// ---------------------------------------------------------------- マウス
let drag = null;
$("#slide").addEventListener("pointerdown", async e => {
  if (e.button !== 0 || !S.slide) return;
  if (e.target.closest("textarea.edit")) return;
  $$("#slide .hover").forEach(x => x.remove());
  if (S.editing) await commitEdit();
  const hit = hitTest(e);
  const sl = $("#slide");
  if (!hit) {
    if (S.group != null) S.group = null;
    if (!e.shiftKey && !e.ctrlKey) S.sel = [];
    drawSelection(); renderPane();
    const p = toEmu(e);
    drag = { kind: "band", x0: p.x, y0: p.y, base: [...S.sel], el: h("div", { class: "band" }) };
    sl.append(drag.el); sl.setPointerCapture(e.pointerId);
    return;
  }
  if (e.shiftKey || e.ctrlKey) { select(S.sel.includes(hit.id) ? S.sel.filter(x => x !== hit.id) : [...S.sel, hit.id]); return; }
  if (!S.sel.includes(hit.id)) select([hit.id]);
  if (!S.deck.editable) return;
  const p = toEmu(e);
  drag = { kind: "move", x0: p.x, y0: p.y, cx: e.clientX, cy: e.clientY, started: false, ids: [...S.sel], dx: 0, dy: 0 };
  sl.setPointerCapture(e.pointerId);
});
$("#slide").addEventListener("pointermove", e => {
  if (!drag) { hoverAt(e); return; }
  const p = toEmu(e);
  if (drag.kind === "band") {
    const f = { x: Math.min(drag.x0, p.x), y: Math.min(drag.y0, p.y), w: Math.abs(p.x - drag.x0), h: Math.abs(p.y - drag.y0) };
    Object.assign(drag.el.style, box(f));
    const inside = topShapes().filter(s => s.frame && !s.hidden && s.frame.x >= f.x && s.frame.y >= f.y && s.frame.x + s.frame.w <= f.x + f.w && s.frame.y + s.frame.h <= f.y + f.h).map(s => s.id);
    S.sel = [...new Set([...drag.base, ...inside])];
    drawSelection();
    return;
  }
  if (drag.kind === "move") {
    if (!drag.started && Math.hypot(e.clientX - drag.cx, e.clientY - drag.cy) < 4) return;
    drag.started = true;
    let dx = p.x - drag.x0, dy = p.y - drag.y0;
    if (e.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
    [dx, dy] = e.altKey ? [dx, dy] : snap(drag.ids, dx, dy);
    drag.dx = dx; drag.dy = dy;
    moveVisual(drag.ids, dx, dy);
    return;
  }
  if (drag.kind === "resize") return resizeMove(e, p);
  if (drag.kind === "rotate") return rotateMove(e, p);
});
$("#slide").addEventListener("pointerup", async () => {
  const d = drag; drag = null;
  clearGuides();
  if (!d) return;
  if (d.kind === "band") { d.el.remove(); renderPane(); return; }
  if (d.kind === "move" && d.started && (d.dx || d.dy)) {
    const shapes = d.ids.map(shapeById).filter(s => s?.frame);
    const frames = shapes.map(s => ({ id: s.id, x: Math.round(s.frame.x + d.dx), y: Math.round(s.frame.y + d.dy), w: s.frame.w, h: s.frame.h }));
    for (const s of shapes) s.frame = { ...s.frame, x: Math.round(s.frame.x + d.dx), y: Math.round(s.frame.y + d.dy) };
    await op("frame", { slide: S.cur, frames, move: true }, { keepVisual: true });
    drawSlide();
  }
  if (d.kind === "resize" && d.frame) {
    shapeById(d.id).frame = { ...shapeById(d.id).frame, ...d.frame };
    await op("frame", { slide: S.cur, frames: [{ id: d.id, ...d.frame }] }, { keepVisual: true });
    drawSlide();
  }
  if (d.kind === "rotate" && d.deg != null) {
    shapeById(d.id).rot = d.deg;
    await op("rotate", { slide: S.cur, ids: [d.id], deg: d.deg }, { keepVisual: true });
    drawSlide();
  }
});
$("#slide").addEventListener("dblclick", e => {
  if (!S.slide || !S.deck.editable) return;
  const hit = hitTest(e);
  if (!hit) return;
  if (hit.kind === "grpSp" && S.group == null) {
    // グループの中に入る（PowerPoint と同じ）。中の図形をクリックで選べる
    S.group = hit.id; S.sel = [];
    const inner = hitTest(e);
    select(inner ? [inner.id] : []);
    if (inner && canText(inner)) startEdit(inner.id);
    return;
  }
  if (hit.table) return openTableEditor(hit.id);
  if (canText(hit)) startEdit(hit.id);
});
$("#slide").addEventListener("contextmenu", e => {
  if (!S.slide) return;
  e.preventDefault();
  const hit = hitTest(e);
  if (hit && !S.sel.includes(hit.id)) select([hit.id]);
  if (!hit && !e.shiftKey) select([]);
  shapeMenu(e.clientX, e.clientY);
});
$("#slide").addEventListener("pointerleave", () => { $$("#slide .hover").forEach(x => x.remove()); });
function hoverAt(e) {
  const sl = $("#slide");
  $$(".hover", sl).forEach(x => x.remove());
  const hit = hitTest(e);
  if (hit && !S.sel.includes(hit.id) && hit.frame) sl.append(h("div", { class: "hover", style: { ...box(hit.frame), transform: hit.rot ? `rotate(${hit.rot}deg)` : "" } }));
  sl.style.cursor = hit ? "move" : "default";
}
function layersOf(id) { return $$(`#slide img.shp[data-id="${id}"], #slide .ghost[data-id="${id}"], #slide .ph[data-id="${id}"]`); }
function moveVisual(ids, dx, dy) {
  const tx = px(dx), ty = px(dy);
  if (S.group == null) for (const id of ids) layersOf(id).forEach(el => el.style.transform = `translate(${tx}px,${ty}px)`);
  $$("#slide .selbox").forEach(el => {
    const base = el.dataset.base ?? el.style.transform ?? ""; el.dataset.base = base;
    el.style.transform = `translate(${tx}px,${ty}px) ${base}`;
  });
}

// ---- 吸着（スマート ガイド）----
function clearGuides() { $$("#slide .guide").forEach(x => x.remove()); }
function snap(ids, dx, dy) {
  clearGuides();
  const fs = ids.map(shapeById).filter(s => s?.frame).map(s => s.frame);
  if (!fs.length) return [dx, dy];
  const b = union(fs), m = S.deck.size, th = 6 / S.scale;
  const others = topShapes().filter(s => s.frame && !s.hidden && !ids.includes(s.id)).map(s => s.frame);
  const xs = [0, m.cx / 2, m.cx, ...others.flatMap(f => [f.x, f.x + f.w / 2, f.x + f.w])];
  const ys = [0, m.cy / 2, m.cy, ...others.flatMap(f => [f.y, f.y + f.h / 2, f.y + f.h])];
  const best = (cands, vals) => { let r = null; for (const v of vals) for (const c of cands) { const d = c - v; if (Math.abs(d) < th && (!r || Math.abs(d) < Math.abs(r.d))) r = { d, at: c }; } return r; };
  const bx = best(xs, [b.x + dx, b.x + b.w / 2 + dx, b.x + b.w + dx]);
  const by = best(ys, [b.y + dy, b.y + b.h / 2 + dy, b.y + b.h + dy]);
  const sl = $("#slide");
  if (bx) { dx += bx.d; sl.append(h("div", { class: "guide", style: { left: px(bx.at) + "px", top: 0, width: "1px", height: "100%" } })); }
  if (by) { dy += by.d; sl.append(h("div", { class: "guide", style: { top: px(by.at) + "px", left: 0, height: "1px", width: "100%" } })); }
  return [dx, dy];
}

// ---- 大きさ・回転 ----
function startResize(e, dir) {
  e.stopPropagation(); e.preventDefault();
  const s = shapeById(S.sel[0]); if (!s?.frame) return;
  drag = { kind: "resize", id: s.id, dir, f0: { ...s.frame }, keep: s.kind === "pic" };
  $("#slide").setPointerCapture(e.pointerId);
}
function resizeMove(e, p) {
  const { f0, dir } = drag;
  let x0 = f0.x, y0 = f0.y, x1 = f0.x + f0.w, y1 = f0.y + f0.h;
  if (dir.includes("w")) x0 = Math.min(p.x, x1 - 1);
  if (dir.includes("e")) x1 = Math.max(p.x, x0 + 1);
  if (dir.includes("n")) y0 = Math.min(p.y, y1 - 1);
  if (dir.includes("s")) y1 = Math.max(p.y, y0 + 1);
  if ((drag.keep !== e.shiftKey) && dir.length === 2 && f0.w && f0.h) {    // 画像は縦横比を保つ（Shift で逆）
    const k = Math.max((x1 - x0) / f0.w, (y1 - y0) / f0.h), w = f0.w * k, hh = f0.h * k;
    if (dir.includes("w")) x0 = x1 - w; else x1 = x0 + w;
    if (dir.includes("n")) y0 = y1 - hh; else y1 = y0 + hh;
  }
  const f = { x: Math.round(x0), y: Math.round(y0), w: Math.round(x1 - x0), h: Math.round(y1 - y0) };
  drag.frame = f;
  // 絵を引き伸ばして仮に見せる（離したら LibreOffice で描き直す）
  const sx = f0.w ? f.w / f0.w : 1, sy = f0.h ? f.h / f0.h : 1;
  for (const img of layersOf(drag.id)) {
    const L = parseFloat(img.style.left) / S.scale, T = parseFloat(img.style.top) / S.scale;
    const nx = f.x + (L - f0.x) * sx, ny = f.y + (T - f0.y) * sy;
    img.style.transform = `translate(${px(nx - L)}px,${px(ny - T)}px) scale(${sx},${sy})`;
  }
  const sb = $("#slide .selbox"); if (sb) Object.assign(sb.style, box(f));
}
function startRotate(e) {
  e.stopPropagation(); e.preventDefault();
  const s = shapeById(S.sel[0]); if (!s?.frame) return;
  drag = { kind: "rotate", id: s.id, f: s.frame, rot0: s.rot || 0 };
  $("#slide").setPointerCapture(e.pointerId);
}
function rotateMove(e, p) {
  const f = drag.f, cx = f.x + f.w / 2, cy = f.y + f.h / 2;
  let deg = (Math.atan2(p.x - cx, -(p.y - cy)) * 180 / Math.PI + 360) % 360;
  if (!e.altKey) deg = e.shiftKey ? Math.round(deg / 15) * 15 % 360 : (Math.abs(deg % 90) < 3 || Math.abs(deg % 90) > 87 ? Math.round(deg / 90) * 90 % 360 : Math.round(deg));
  drag.deg = deg;
  const sb = $("#slide .selbox"); if (sb) sb.style.transform = `rotate(${deg}deg)`;
  for (const img of layersOf(drag.id)) {
    const ox = px(cx) - parseFloat(img.style.left), oy = px(cy) - parseFloat(img.style.top);
    img.style.transformOrigin = `${ox}px ${oy}px`; img.style.transform = `rotate(${deg - drag.rot0}deg)`;
  }
}

// ---------------------------------------------------------------- 右クリックのメニュー
function shapeMenu(x, y) {
  const ed = S.deck.editable, ss = selShapes(), one = ss.length === 1 ? ss[0] : null, has = ss.length > 0;
  if (!has) {
    return menu(x, y, [
      { label: "貼り付け", key: "Ctrl+V", disabled: !ed || !S.clipToken, do: () => pasteInternal() },
      "-",
      { label: "新しいスライド", key: "Ctrl+M", disabled: !ed, do: () => newSlide() },
      { label: "背景の書式…", disabled: !ed, do: () => setPane("format") },
      { label: "画面切り替え…", do: () => setPane("trans") },
      { label: "すべて選択", key: "Ctrl+A", do: () => select(topShapes().filter(s => s.frame && !s.hidden).map(s => s.id)) },
      "-",
      { label: "スライドを画像で保存", do: () => saveSlideImage() },
      { label: "このスライドから再生", key: "Shift+F5", do: () => startShow(true) },
    ]);
  }
  menu(x, y, [
    { label: "切り取り", key: "Ctrl+X", disabled: !ed, do: () => copySel(true) },
    { label: "コピー", key: "Ctrl+C", do: () => copySel(false) },
    { label: "貼り付け", key: "Ctrl+V", disabled: !ed || !S.clipToken, do: () => pasteInternal() },
    { label: "複製", key: "Ctrl+D", disabled: !ed, do: () => duplicate() },
    { label: "削除", key: "Delete", disabled: !ed, do: () => delSel() },
    "-",
    one && canText(one) ? { label: "文字の編集", key: "Enter", disabled: !ed, do: () => startEdit(one.id) } : null,
    one?.table ? { label: "表の編集…", disabled: !ed, do: () => openTableEditor(one.id) } : null,
    one?.image ? { label: "図の変更…", disabled: !ed, do: () => pickImage(one.id) } : null,
    { label: "グループ化", key: "Ctrl+G", disabled: !ed || ss.length < 2, do: () => groupSel() },
    { label: "グループ解除", key: "Ctrl+Shift+G", disabled: !ed || !ss.some(s => s.kind === "grpSp"), do: () => ungroupSel() },
    { label: "重なりの順番", disabled: !ed, sub: orderItems() },
    { label: "配置", disabled: !ed, sub: alignItems() },
    { label: "回転と反転", disabled: !ed, sub: rotateItems() },
    "-",
    { label: "アニメーションの追加", disabled: !ed, sub: ANIM_EFFECTS.map(([v, l]) => ({ label: l, do: () => addAnim(v) })) },
    { label: "アニメーションのプレビュー", disabled: !S.slide.anims.length, do: () => previewSlide() },
    { label: "書式の設定…", do: () => setPane("format") },
    { label: one?.hidden ? "表示する" : "隠す", disabled: !ed, do: () => op("visibility", { slide: S.cur, ids: S.sel, hidden: !one?.hidden }) },
  ]);
}
