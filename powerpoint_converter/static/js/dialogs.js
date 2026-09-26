"use strict";
// 表の編集・検索と置換

// スライドを読み直したあとにすること（表の編集画面の作り直し・検索で移った先の図形を選ぶ）
function onSlideLoaded() {
  if (tableDlg?.dirty) tableDlg.render();
  if (S.pendingSelect && S.slide && S.slide.index === S.pendingSelect.slide) {
    const id = S.pendingSelect.id; S.pendingSelect = null;
    if (id != null && shapeById(id)) { const top = topOf(id); S.group = top !== id ? top : null; select([id]); drawSlide(); }
  }
}

// ---------------------------------------------------------------- 表の編集
let tableDlg = null;
function openTableEditor(id) {
  const s0 = shapeById(id); if (!s0?.table) return;
  if (S.editing) commitEdit();
  const ed = S.deck.editable;
  const focus = { r: 0, c: 0 };
  const grid = h("div");
  const note = h("div", { class: "note", style: { marginBottom: "8px" } });
  const act = (label, fn, title) => h("button", { title, onclick: fn }, label);
  const tools = h("div", { class: "chips", style: { marginBottom: "8px" } },
    act("↑ 上に行を挿入", () => structure("table_row", { row: focus.r, where: "above" })),
    act("↓ 下に行を挿入", () => structure("table_row", { row: focus.r, where: "below" })),
    act("行を削除", () => structure("table_row", { row: focus.r, where: "delete" })),
    act("← 左に列を挿入", () => structure("table_col", { col: focus.c, where: "left" })),
    act("→ 右に列を挿入", () => structure("table_col", { col: focus.c, where: "right" })),
    act("列を削除", () => structure("table_col", { col: focus.c, where: "delete" })));
  const body = h("div", {}, tools, note, grid);
  const dlg = dialog(`表の編集「${s0.name}」`, body, [{ label: "閉じる", primary: true }], { onclose: () => { tableDlg = null; } });
  async function structure(name, args) {
    await flushCell();
    const r = await op(name, { slide: S.cur, id, ...args });
    if (r) tableDlg.dirty = true;
  }
  let pending = null;
  async function flushCell() {
    if (!pending) return;
    const p = pending; pending = null;
    await op("table_text", { slide: S.cur, id, row: p.r, col: p.c, text: p.ta.value.split("\n") }, { keepVisual: true });
  }
  function render() {
    tableDlg.dirty = false;
    const t = shapeById(id)?.table;
    if (!t) { dlg.close(); return; }
    note.textContent = t.has_merge ? "結合したセルがある表なので、行・列の追加と削除はできない（文字は直せる）。" :
      "セルを選んでから、上のボタンで行・列を足し引きする。セルから離れると保存する。";
    $$("button", tools).forEach(b => b.disabled = !ed || t.has_merge);
    const tb = h("table", { class: "tbl-edit" });
    t.rows.forEach((row, r) => {
      const tr = h("tr");
      row.cells.forEach((c, ci) => {
        const ta = h("textarea", { disabled: !ed || c.merged, title: c.merged ? "結合されたセル（左か上のセルに入力する）" : null });
        ta.value = c.text.join("\n").replace(/\v/g, "\n");
        const td = h("td", { colspan: c.span > 1 ? c.span : null }, ta);
        ta.addEventListener("focus", () => { $$("td", tb).forEach(x => x.classList.remove("focus")); td.classList.add("focus"); focus.r = r; focus.c = ci; });
        ta.addEventListener("input", () => { pending = { r, c: ci, ta }; });
        ta.addEventListener("blur", () => flushCell());
        ta.addEventListener("keydown", e => e.stopPropagation());
        tr.append(td);
      });
      tb.append(tr);
    });
    grid.innerHTML = ""; grid.append(tb);
  }
  tableDlg = { id, render, dirty: false };
  render();
  $("textarea", grid)?.focus();
}

// ---------------------------------------------------------------- 検索と置換
function openFind(replace) {
  const bar = $("#findbar");
  bar.classList.remove("hidden");
  const q = $("#findQ");
  if (replace && q.value) $("#findR").focus(); else { q.focus(); q.select(); }
  if (S.zoom === "fit") drawSlide();
}
function closeFind() { $("#findbar").classList.add("hidden"); S.find.results = []; S.find.i = -1; if (S.zoom === "fit") drawSlide(); }
async function runFind(dir = 1, fresh = false) {
  const q = $("#findQ").value;
  if (!q) { $("#findCount").textContent = ""; $("#findRes").textContent = ""; return; }
  if (fresh || q !== S.find.q || S.find.caseUsed !== S.find.case) {
    try { S.find.results = (await api(`/api/find?q=${encodeURIComponent(q)}&case=${S.find.case ? 1 : 0}`)).results; }
    catch (e) { toast(e.message, true); return; }
    S.find.q = q; S.find.caseUsed = S.find.case;
    if (!fresh) S.find.i = -1;
  }
  const n = S.find.results.length;
  if (!n) { $("#findCount").textContent = "見つからない"; $("#findRes").textContent = ""; return; }
  S.find.i = S.find.i < 0 ? (dir > 0 ? 0 : n - 1) : (S.find.i + dir + n) % n;
  S.find.i = Math.min(S.find.i, n - 1);
  const res = S.find.results[S.find.i], total = S.find.results.reduce((a, r) => a + r.count, 0);
  $("#findCount").textContent = `${S.find.i + 1} / ${n}（${total} か所）`;
  $("#findRes").textContent = `${res.slide + 1}枚目 ${res.where === "notes" ? "ノート" : res.where === "table" ? "表" : ""}: ${res.text}`;
  if (res.slide !== S.cur) { S.pendingSelect = { slide: res.slide, id: res.id }; await gotoSlide(res.slide); }
  else if (res.id != null && shapeById(res.id)) { const top = topOf(res.id); S.group = top !== res.id ? top : null; select([res.id]); drawSlide(); }
  if (res.where === "notes") $("#notes").focus();
}
async function replaceCurrent() {
  const res = S.find.results[S.find.i];
  if (!res) return runFind(1);
  const r = await op("replace", { find: $("#findQ").value, replace: $("#findR").value, case: S.find.case, slide: res.slide, id: res.where === "notes" ? "notes" : res.id });
  if (r) { toast(r.label); await sleep(300); await runFind(0, true); }
}
async function replaceAll() {
  if (!$("#findQ").value) return;
  const r = await op("replace", { find: $("#findQ").value, replace: $("#findR").value, case: S.find.case });
  if (r) { toast(r.label); await sleep(300); await runFind(0, true); }
}
$("#findQ").addEventListener("keydown", e => {
  e.stopPropagation();
  if (e.key === "Enter") { e.preventDefault(); runFind(e.shiftKey ? -1 : 1); }
  if (e.key === "Escape") closeFind();
});
$("#findR").addEventListener("keydown", e => {
  e.stopPropagation();
  if (e.key === "Enter") { e.preventDefault(); replaceCurrent(); }
  if (e.key === "Escape") closeFind();
});
$("#findNext").onclick = () => runFind(1);
$("#findPrev").onclick = () => runFind(-1);
$("#findRep").onclick = () => replaceCurrent();
$("#findAll").onclick = () => replaceAll();
$("#findClose").onclick = () => closeFind();
$("#findCase").onclick = e => { S.find.case = !S.find.case; e.currentTarget.classList.toggle("on", S.find.case); runFind(1, true); };
