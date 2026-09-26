"use strict";
// 資料を開く・サーバの変化を待つ・スライド一覧（左）・帯と状態表示・ノート

async function openDeck(path) {
  S.view = "deck"; S.path = path;
  $("#list").classList.add("hidden"); $("#editor").classList.remove("hidden");
  const parts = path.split("/");
  crumbs(parts.slice(0, -1), parts.at(-1));
  document.title = parts.at(-1) + " – PowerPoint Converter";
  let info;
  try { info = await post("/api/open", { path }); }
  catch (e) { toast(e.message, true, 6000); return; }
  const same = S.deck && S.deck.name === info.deck.name;
  S.sel = []; S.group = null;
  if (!same) { S.lastRender = {}; S.thumbs = {}; }
  S.cur = same ? S.cur : Math.min(store.get("cur:" + path, 0), info.deck.slides.length - 1);
  S.rev = -1; S.slide = null;
  applyModel(info.deck);
  pollLoop();
}

async function pollLoop() {
  const my = ++S.pollToken;
  while (my === S.pollToken && S.view === "deck") {
    try {
      const m = await api(`/api/poll?since=${S.rev}&cur=${S.cur}`);
      if (my !== S.pollToken) return;
      applyModel(m);
    } catch (e) {
      if (my !== S.pollToken) return;
      $("#saved").textContent = "サーバにつながらない"; await sleep(1500);
    }
  }
}

function applyModel(m) {
  const prev = S.deck;
  S.deck = m; S.rev = m.rev;
  if (S.cur >= m.slides.length) S.cur = Math.max(0, m.slides.length - 1);
  document.documentElement.style.setProperty("--ratio", `${m.size.cx}/${m.size.cy}`);
  renderThumbs(); renderBanner(); renderStatus(); renderRibbon();
  if (m.notice && (!prev || !prev.notice || prev.notice.t !== m.notice.t)) toast(m.notice.text);
  const s = m.slides[S.cur];
  if (s && (!S.slide || S.slide.fp !== s.fp || S.slide.status !== s.status || S.slide.index !== S.cur)) loadSlide();
  if (!s) { S.slide = null; drawSlide(); }
}

// ---------------------------------------------------------------- スライド一覧（左）
let dragFrom = null;
function thumbSrc(s) {
  // 描き直しを待つ間は、そのスライドの最後に描けた絵を出しておく
  if (s.status === "ready") return (S.thumbs[s.part] = `/render/${s.fp}/full.png`);
  const lr = S.lastRender[s.part];
  return S.thumbs[s.part] || (lr ? `/render/${lr.fp}/full.png` : null);
}
function renderThumbs() {
  const box = $("#thumbs"), m = S.deck;
  const keep = box.scrollTop;
  const olds = new Map($$(".th", box).map(el => [el.dataset.key, el]));
  box.innerHTML = "";
  m.slides.forEach((s, i) => {
    let el = olds.get(s.part);
    if (!el) el = makeThumb(box);
    el.dataset.key = s.part; el.dataset.i = i;
    el.classList.toggle("cur", i === S.cur);
    el.classList.toggle("hid", s.hidden);
    if (s.changed && !el.classList.contains("changed")) { el.classList.add("changed"); setTimeout(() => el.classList.remove("changed"), 7000); }
    $(".no span", el).textContent = s.number;
    $(".no .star", el).textContent = s.anims ? "★" : "";
    $(".no .star", el).title = s.anims ? `アニメーション ${s.anims} 個` : "";
    $(".no .trans", el).textContent = s.transition && s.transition !== "none" ? "⇢" : "";
    $(".no .trans", el).title = s.transition && s.transition !== "none" ? "画面切り替えあり" : "";
    const img = $("img", el), src = thumbSrc(s);
    if (src && img.getAttribute("src") !== src) img.src = src;
    img.style.visibility = src ? "visible" : "hidden";
    $(".spin", el).classList.toggle("hidden", s.status === "ready" || s.status.startsWith("error"));
    el.title = (s.title || "（タイトルなし）") + (s.hidden ? "（非表示）" : "") + (s.status.startsWith("error") ? "\n描けなかった: " + s.status : "");
    box.append(el);
  });
  box.scrollTop = keep;
}
function makeThumb(box) {
  const el = h("div", { class: "th", draggable: "true" },
    h("div", { class: "no" }, h("span", {}), h("i", { class: "star" }), h("i", { class: "trans" })),
    h("div", { class: "pic" }, h("img", { alt: "" }), h("span", { class: "spin hidden" })));
  el.addEventListener("click", () => gotoSlide(+el.dataset.i));
  el.addEventListener("contextmenu", e => { e.preventDefault(); gotoSlide(+el.dataset.i); slideMenu(e.clientX, e.clientY); });
  el.addEventListener("dragstart", e => { dragFrom = +el.dataset.i; e.dataTransfer.effectAllowed = "move"; e.dataTransfer.setData("text/plain", "slide"); });
  el.addEventListener("dragover", e => {
    if (dragFrom == null) return; e.preventDefault();
    const r = el.getBoundingClientRect(), after = e.clientY > r.top + r.height / 2;
    $$(".th", box).forEach(x => x.classList.remove("dragover", "dragover-after"));
    el.classList.add(after ? "dragover-after" : "dragover");
  });
  el.addEventListener("dragleave", () => el.classList.remove("dragover", "dragover-after"));
  el.addEventListener("drop", e => {
    e.preventDefault(); const after = el.classList.contains("dragover-after");
    $$(".th", box).forEach(x => x.classList.remove("dragover", "dragover-after"));
    if (dragFrom == null) return;
    let to = +el.dataset.i + (after ? 1 : 0); if (to > dragFrom) to--;
    const from = dragFrom; dragFrom = null;
    if (to !== from) op("slide_move", { from, to }).then(r => r && gotoSlideLater(to));
  });
  el.addEventListener("dragend", () => { dragFrom = null; $$(".th", box).forEach(x => x.classList.remove("dragover", "dragover-after")); });
  return el;
}
function scrollThumbIntoView() { $$(".th")[S.cur]?.scrollIntoView({ block: "nearest" }); }

async function gotoSlide(i) {
  if (!S.deck) return;
  i = Math.max(0, Math.min(S.deck.slides.length - 1, i));
  if (S.editing) await commitEdit();
  await flushNotes();
  if (i === S.cur && S.slide) return;
  stopPreview();
  S.cur = i; S.sel = []; S.group = null; S.animSel = null;
  store.set("cur:" + S.path, i);
  $$(".th").forEach((el, k) => el.classList.toggle("cur", k === i));
  scrollThumbIntoView();
  api(`/api/poll?since=-1&cur=${i}`).catch(() => { });   // 描く順番をこのスライド優先にする
  loadSlide();
}
function gotoSlideLater(i) { S.cur = i; S.slide = null; S.sel = []; store.set("cur:" + S.path, i); }
function newSlide(layout) { return op("slide_new", { after: S.cur, layout }).then(r => r && gotoSlideLater(r.slide)); }

function slideMenu(x, y) {
  const s = S.deck.slides[S.cur], ed = S.deck.editable;
  menu(x, y, [
    { label: "新しいスライド", key: "Ctrl+M", disabled: !ed, do: () => newSlide() },
    { label: "スライドの複製", disabled: !ed, do: () => op("slide_dup", { index: S.cur }).then(r => r && gotoSlideLater(r.slide)) },
    { label: "スライドの削除", disabled: !ed, do: () => confirm(`${S.cur + 1} 枚目を消す（元に戻すで戻せる）`) && op("slide_delete", { index: S.cur }) },
    "-",
    { label: s.hidden ? "スライドを再表示" : "非表示スライドに設定", disabled: !ed, do: () => op("slide_hide", { index: S.cur, hidden: !s.hidden }) },
    { label: "上へ移動", disabled: !ed || S.cur === 0, do: () => op("slide_move", { from: S.cur, to: S.cur - 1 }).then(r => r && gotoSlideLater(S.cur - 1)) },
    { label: "下へ移動", disabled: !ed || S.cur >= S.deck.slides.length - 1, do: () => op("slide_move", { from: S.cur, to: S.cur + 1 }).then(r => r && gotoSlideLater(S.cur + 1)) },
    "-",
    { label: "背景の書式…", disabled: !ed, do: () => { S.sel = []; setPane("format"); } },
    { label: "画面切り替え…", do: () => setPane("trans") },
    { label: "スライドを画像で保存", disabled: s.status !== "ready", do: () => saveSlideImage() },
    "-",
    { label: "このスライドから再生", key: "Shift+F5", do: () => startShow(true) },
  ]);
}
function saveSlideImage() {
  const s = S.deck.slides[S.cur];
  const a = h("a", { href: `/render/${s.fp}/full.png`, download: `${S.deck.name.replace(/\.[^.]+$/, "")}_${S.cur + 1}.png` });
  document.body.append(a); a.click(); a.remove();
}

// ---------------------------------------------------------------- 帯・状態
function renderBanner() {
  const b = $("#banner"), m = S.deck; b.innerHTML = "";
  if (m.error) b.append(h("div", { class: "msg err" }, "読み込めない: " + m.error));
  if (!m.editable) b.append(h("div", { class: "msg info" }, `${m.name} は見るだけ（LibreOffice で PPTX にして表示している）。直すには PowerPoint か LibreOffice で .pptx に保存する`));
}
function renderStatus() {
  const m = S.deck; if (!m) return;
  const s = m.slides[S.cur];
  $("#stSlide").textContent = s ? `スライド ${S.cur + 1} / ${m.slides.length}${s.hidden ? "（非表示）" : ""}` : "スライドが無い";
  const pend = m.slides.filter(x => x.status !== "ready" && !x.status.startsWith("error")).length;
  const err = m.slides.filter(x => x.status.startsWith("error")).length;
  $("#stRender").innerHTML = pend ? `<span class="spin"></span> 描画 残り ${pend} 枚` : err ? `描けなかった ${err} 枚` : "✓ 描画済み";
  const bad = (m.fonts || []).filter(f => !f.same_width);
  $("#stFonts").innerHTML = bad.length ? `<a href="#" id="fontLink" style="color:var(--warn)">⚠ フォントの置き換え ${bad.length} 件</a>` : "";
  const fl = $("#fontLink"); if (fl) fl.onclick = e => { e.preventDefault(); S.sel = []; setPane("format"); drawSelection(); };
}

// ---------------------------------------------------------------- ノート
let notesTimer = null, notesDirty = false;
function renderNotes() {
  const t = $("#notes"), sm = S.slide;
  if (document.activeElement === t && notesDirty) return;
  t.value = (sm?.notes || []).join("\n").replace(/\v/g, "\n");
  t.disabled = !S.deck?.editable || (!sm?.notes && !sm?.has_notes_master);
  t.placeholder = t.disabled && S.deck?.editable ? "この資料にはノートの型が無いので、ノートを書けない" : "クリックしてノートを入力";
  notesDirty = false;
}
$("#notes").addEventListener("input", () => { notesDirty = true; clearTimeout(notesTimer); notesTimer = setTimeout(flushNotes, 1200); });
$("#notes").addEventListener("blur", () => flushNotes());
$("#notes").addEventListener("keydown", e => e.stopPropagation());
async function flushNotes() {
  clearTimeout(notesTimer);
  if (!notesDirty || !S.slide) return;
  notesDirty = false;
  await op("notes", { slide: S.slide.index, text: $("#notes").value.split("\n") }, { keepVisual: true });
}
{
  const bar = $("#notesbar");
  bar.addEventListener("pointerdown", e => {
    const y0 = e.clientY, h0 = $("#notes").offsetHeight; bar.setPointerCapture(e.pointerId);
    const mv = ev => { const v = Math.max(30, Math.min(innerHeight * .6, h0 - (ev.clientY - y0))); document.documentElement.style.setProperty("--notes-h", v + "px"); if (S.zoom === "fit") drawSlide(); };
    bar.addEventListener("pointermove", mv);
    bar.addEventListener("pointerup", () => { bar.removeEventListener("pointermove", mv); store.set("notesH", $("#notes").offsetHeight); }, { once: true });
  });
  const nh = store.get("notesH", null); if (nh) document.documentElement.style.setProperty("--notes-h", nh + "px");
}
