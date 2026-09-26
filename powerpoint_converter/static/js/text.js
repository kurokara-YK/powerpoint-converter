"use strict";
// 図形の文字の編集（スライドの上に重ねた入力欄）と、選んだ文字だけへの書式

function canText(s) { return s && s.kind === "sp" && S.deck.editable; }
// 画面の入力欄では、段落の中の改行（\v）を「↵ + 改行」で見せる
const toEdit = paras => (paras || [""]).map(p => p.replace(/\v/g, "↵\n")).join("\n");
const fromEdit = v => v.replace(/↵\n/g, "\v").replace(/↵/g, "\v").split("\n");
// 入力欄の中の位置を、サーバの「通しの文字位置」（段落の区切り・段落内の改行はどちらも1文字）にする
const modelOffset = (v, pos) => v.slice(0, pos).replace(/↵\n/g, "\v").replace(/↵/g, "\v").length;

function editStyle(s) {
  const st = s.style || {};
  const fs = (st.sz || 18) * EMU_PT * S.scale;
  const fam = [st.font_ea, st.font_latin].filter(Boolean).map(f => `"${f}"`).join(",");
  return { fontSize: fs + "px", fontFamily: (fam ? fam + "," : "") + '"IPAexGothic","Noto Sans CJK JP",sans-serif',
    textAlign: { ctr: "center", r: "right", just: "justify" }[st.algn] || "left", color: st.color || "#000",
    fontWeight: st.b ? 700 : 400, fontStyle: st.i ? "italic" : "normal", transform: s.rot ? `rotate(${s.rot}deg)` : "" };
}
function startEdit(id) {
  const s = shapeById(id); if (!s || !s.frame) return;
  if (S.editing) return;
  const ta = h("textarea", { class: "edit", spellcheck: "false" });
  ta.value = toEdit(s.text);
  Object.assign(ta.style, box(s.frame), editStyle(s), { minHeight: px(s.frame.h) + "px" });
  layersOf(S.group != null ? S.group : id).forEach(img => img.style.opacity = .12);
  $("#slide").append(ta);
  const grow = () => { ta.style.height = "auto"; ta.style.height = Math.max(px(s.frame.h), ta.scrollHeight + 4) + "px"; };
  ta.addEventListener("input", grow);
  ta.addEventListener("keydown", e => {
    e.stopPropagation();
    const ctrl = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    if (e.key === "Escape") { e.preventDefault(); commitEdit(); return; }
    if (e.key === "Enter" && e.shiftKey) { e.preventDefault(); const a = ta.selectionStart; ta.setRangeText("↵\n", a, ta.selectionEnd, "end"); grow(); return; }
    if (e.key === "Enter" && ctrl) { e.preventDefault(); commitEdit(); return; }
    if (ctrl && ["b", "i", "u"].includes(k)) { e.preventDefault(); applyText({ [k]: "toggle" }); return; }
    if (ctrl && (e.key === "]" || e.key === "[")) { e.preventDefault(); applyText({ grow: e.key === "]" ? 2 : -2 }); return; }
  });
  ta.addEventListener("select", () => renderPane());
  ta.addEventListener("pointerdown", e => e.stopPropagation());
  ta.addEventListener("contextmenu", e => e.stopPropagation());
  S.editing = { id, ta, orig: ta.value };
  grow(); ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length);
  drawSelection(); renderPane();
}
function repositionEdit() {
  const ed = S.editing, s = ed && shapeById(ed.id);
  if (!s?.frame) return;
  if (!ed.ta.isConnected) $("#slide").append(ed.ta);
  Object.assign(ed.ta.style, box(s.frame), editStyle(s), { minHeight: px(s.frame.h) + "px" });
  layersOf(S.group != null ? S.group : ed.id).forEach(img => img.style.opacity = .12);
}
async function saveEditText() {
  // 入力欄の文字をファイルに書く（書式を当てる前にも呼ぶ）。失敗したら false
  const ed = S.editing; if (!ed) return true;
  const v = ed.ta.value;
  if (v === ed.orig) return true;
  const s = shapeById(ed.id); if (s) s.text = fromEdit(v);
  const r = await op("text", { slide: S.cur, id: ed.id, text: fromEdit(v) }, { keepVisual: true });
  if (!r) {
    navigator.clipboard?.writeText(v).catch(() => { });
    toast("文字を当てられなかったので、入力した文字をクリップボードに写した", true, 6000);
    return false;
  }
  ed.orig = v;
  return true;
}
async function commitEdit() {
  const ed = S.editing; if (!ed) return;
  await saveEditText();
  S.editing = null;
  ed.ta.remove();
  drawSlide(); renderPane();
}
function editRange() {
  // 文字を編集中で、一部を選んでいれば、その範囲（通しの文字位置）
  const ed = S.editing; if (!ed) return null;
  const a = ed.ta.selectionStart, b = ed.ta.selectionEnd;
  if (a === b) return null;
  return { start: modelOffset(ed.ta.value, a), end: modelOffset(ed.ta.value, b), a, b };
}
async function applyText(props) {
  // 書式を当てる。文字を編集中で一部を選んでいればその文字だけ、そうでなければ選んだ図形全体
  const rng = editRange();
  if (rng) {
    const ed = S.editing;
    if (!(await saveEditText())) return;
    await op("text_style", { slide: S.cur, id: ed.id, start: rng.start, end: rng.end, props }, { keepVisual: true });
    if (S.editing === ed) { ed.ta.focus(); ed.ta.setSelectionRange(rng.a, rng.b); }
    return;
  }
  const ids = S.editing ? [S.editing.id] : S.sel;
  if (!ids.length) return;
  if (S.editing && !(await saveEditText())) return;
  await op("style", { slide: S.cur, ids, props }, { keepVisual: true });
}
