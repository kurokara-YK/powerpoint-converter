"use strict";
// 右の作業ウィンドウ（書式・アニメーション・画面切り替え・選択）

function setPane(p) {
  S.pane = p; store.set("pane", p);
  $$("#pane .tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === p));
  if ($("#work").classList.contains("nopane")) togglePanel("nopane", "nopane");
  renderPane(); drawSlide();
}
$$("#pane .tabs button").forEach(b => b.onclick = () => setPane(b.dataset.tab));
function renderPane() {
  const body = $("#paneBody"), keep = body.scrollTop;
  body.innerHTML = "";
  $$("#pane .tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === S.pane));
  if (!S.slide || !S.deck) return;
  if (S.pane === "anim") renderAnimPane(body);
  else if (S.pane === "trans") renderTransPane(body);
  else if (S.pane === "select") renderSelectPane(body);
  else if (!selShapes().length && !S.editing) renderSlidePane(body);
  else renderFormatPane(body);
  body.scrollTop = keep;
}
const sec = (title, ...kids) => h("div", { class: "sec" }, h("h4", {}, title), ...kids);
const row = (label, ...kids) => h("div", { class: "row" }, label != null ? h("label", {}, label) : null, ...kids);
const seg = (items, cur, fn, dis) => h("div", { class: "seg" }, items.map(([v, l, t]) =>
  h("button", { class: cur === v ? "on" : "", title: t || null, disabled: dis, onclick: () => fn(v) }, l)));

// ---------------------------------------------------------------- 書式（図形を選んでいるとき）
function renderFormatPane(body) {
  const ed = S.deck.editable;
  const ss = S.editing ? [shapeById(S.editing.id)].filter(Boolean) : selShapes();
  const s = ss[0], one = ss.length === 1;
  if (!s) return;
  const rng = editRange();
  body.append(sec(one ? `${s.label}「${s.name}」` : `${ss.length} 個の図形`,
    one ? h("div", { class: "note" }, `id=${s.id}${s.ph ? ` ・ プレースホルダー（${s.ph.type}）` : ""}${s.frame_src === "inherited" ? " ・ 位置はレイアウトから継承" : ""}${s.hidden ? " ・ 隠している" : ""}`) : null,
    one ? h("div", { class: "chips", style: { marginTop: "4px" } }, h("button", { disabled: !ed, onclick: () => renameShape(s) }, "名前の変更")) : null));

  if (one && s.frame) {
    const cm = v => (v / EMU_CM).toFixed(2);
    const inp = (k, v) => h("input", { type: "number", step: "0.01", value: cm(v), disabled: !ed, onchange: e => {
      const f = { ...s.frame, [k]: Math.round(parseFloat(e.target.value) * EMU_CM) };
      op("frame", { slide: S.cur, frames: [{ id: s.id, ...f }] });
    } });
    body.append(sec("位置とサイズ（cm）",
      h("div", { class: "grid4" }, h("label", {}, "横"), inp("x", s.frame.x), h("label", {}, "縦"), inp("y", s.frame.y),
        h("label", {}, "幅"), inp("w", s.frame.w), h("label", {}, "高さ"), inp("h", s.frame.h)),
      row("回転", h("input", { type: "number", step: "1", value: Math.round(s.rot || 0), disabled: !ed,
        onchange: e => op("rotate", { slide: S.cur, ids: [s.id], deg: parseFloat(e.target.value) || 0 }) }), "°",
        h("button", { disabled: !ed, title: "左右反転", onclick: () => op("flip", { slide: S.cur, ids: [s.id], axis: "h" }) }, "⇋"),
        h("button", { disabled: !ed, title: "上下反転", onclick: () => op("flip", { slide: S.cur, ids: [s.id], axis: "v" }) }, "⇵"))));
  }

  const texty = ss.some(x => x.kind === "sp" || x.kind === "grpSp");
  if (texty) {
    const st = s.style || {};
    const fontIn = h("input", { list: "fontlist", class: "grow", value: st.font_ea || st.font_latin || "", placeholder: "テーマのフォント", disabled: !ed });
    fontIn.addEventListener("change", () => applyText({ font: fontIn.value.trim() }));
    loadFonts();
    body.append(sec(rng ? "文字（選んだ文字だけに当たる）" : S.editing ? "文字（文字を選ぶと、その文字だけに当たる）" : "文字",
      row("フォント", fontIn),
      row("サイズ", h("input", { type: "number", step: "0.5", min: "1", value: st.sz ?? "", placeholder: "既定", disabled: !ed, onchange: e => applyText({ sz: parseFloat(e.target.value) }) }), "pt",
        h("button", { title: "大きく (Ctrl+])", disabled: !ed, onclick: () => applyText({ grow: 2 }) }, "A⁺"),
        h("button", { title: "小さく (Ctrl+[)", disabled: !ed, onclick: () => applyText({ grow: -2 }) }, "A⁻")),
      row("書式",
        h("button", { class: st.b ? "on" : "", disabled: !ed, title: "太字 (Ctrl+B)", onclick: () => applyText({ b: "toggle" }), style: { fontWeight: 700 } }, "B"),
        h("button", { class: st.i ? "on" : "", disabled: !ed, title: "斜体 (Ctrl+I)", onclick: () => applyText({ i: "toggle" }), style: { fontStyle: "italic" } }, "I"),
        h("button", { class: st.u ? "on" : "", disabled: !ed, title: "下線 (Ctrl+U)", onclick: () => applyText({ u: "toggle" }), style: { textDecoration: "underline" } }, "U"),
        h("button", { disabled: !ed, title: "取り消し線", onclick: () => applyText({ strike: "toggle" }), style: { textDecoration: "line-through" } }, "S"),
        h("button", { disabled: !ed, title: "上付き", onclick: () => applyText({ baseline: 30000 }) }, "x²"),
        h("button", { disabled: !ed, title: "下付き", onclick: () => applyText({ baseline: -25000 }) }, "x₂"),
        h("button", { disabled: !ed, title: "上付き・下付きを戻す", onclick: () => applyText({ baseline: 0 }) }, "x")),
      row("色", h("input", { type: "color", class: "swatch", title: "文字の色", value: st.color || "#000000", disabled: !ed, onchange: e => applyText({ color: e.target.value }) }),
        h("span", { class: "note" }, "蛍光ペン"),
        h("input", { type: "color", class: "swatch", title: "蛍光ペンの色", value: "#ffff00", disabled: !ed, onchange: e => applyText({ highlight: e.target.value }) }),
        h("button", { disabled: !ed, title: "蛍光ペンを消す", onclick: () => applyText({ highlight: "none" }) }, "消す"))));
    body.append(sec("段落",
      row("揃え", seg([["l", "左", "左揃え (Ctrl+L)"], ["ctr", "中", "中央揃え (Ctrl+E)"], ["r", "右", "右揃え (Ctrl+R)"], ["just", "両", "両端揃え (Ctrl+J)"]], st.algn, v => applyText({ algn: v }), !ed)),
      row("上下", seg([["t", "上"], ["ctr", "中"], ["b", "下"]], s.anchor, v => op("style", { slide: S.cur, ids: ss.map(x => x.id), props: { anchor: v } }), !ed)),
      row("箇条書き", seg([["none", "なし"], ["char", "•"], ["number", "1."]], s.bullet, v => applyText({ bullet: v }), !ed),
        h("button", { disabled: !ed, title: "インデントを減らす", onclick: () => applyText({ lvl_delta: -1 }) }, "⇤"),
        h("button", { disabled: !ed, title: "インデントを増やす", onclick: () => applyText({ lvl_delta: 1 }) }, "⇥")),
      row("行間", h("select", { disabled: !ed, onchange: e => applyText({ line_spacing: parseFloat(e.target.value) }) },
        ...["1.0", "1.15", "1.2", "1.5", "2.0", "2.5", "3.0"].map(v => h("option", { value: v, selected: Math.abs((s.line_spacing || 1) - v) < .01 }, v)))),
      row("自動調整", h("select", { disabled: !ed, onchange: e => op("style", { slide: S.cur, ids: ss.map(x => x.id), props: { autofit: e.target.value } }) },
        ...[["none", "自動調整なし"], ["shrink", "はみ出す時だけ縮小"], ["shape", "図形を文字に合わせる"]].map(([v, l]) => h("option", { value: v, selected: s.autofit === v }, l)))),
      one && canText(s) && !S.editing ? h("div", { class: "row" }, h("button", { disabled: !ed, onclick: () => startEdit(s.id) }, "✎ 文字を編集（ダブルクリック）")) : null));
  }

  const shapeProps = props => op("style", { slide: S.cur, ids: ss.map(x => x.id), props });
  const fillable = ss.some(x => x.kind === "sp");
  const ln = s.line || {};
  const isLine = s.kind === "cxnSp";
  body.append(sec(isLine ? "線の書式" : "図形の書式",
    fillable ? row("塗り", h("input", { type: "color", class: "swatch", value: s.fill && s.fill !== "none" ? s.fill : "#ffffff", disabled: !ed, onchange: e => shapeProps({ fill: e.target.value }) }),
      h("button", { disabled: !ed, onclick: () => shapeProps({ fill: "none" }) }, "塗りなし")) : null,
    row("線", h("input", { type: "color", class: "swatch", value: ln.color || "#000000", disabled: !ed, onchange: e => shapeProps({ line: e.target.value }) }),
      h("button", { disabled: !ed, onclick: () => shapeProps({ line: "none" }) }, "線なし")),
    row("太さ", h("select", { disabled: !ed, onchange: e => shapeProps({ line_w: parseFloat(e.target.value) }) },
      h("option", { value: "" }, ln.w ? `${ln.w} pt` : "既定"), ...[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map(v => h("option", { value: v }, `${v} pt`)))),
    row("種類", h("select", { disabled: !ed, onchange: e => shapeProps({ line_dash: e.target.value }) },
      ...[["solid", "実線"], ["sysDot", "点線（丸）"], ["sysDash", "点線（角）"], ["dash", "破線"], ["dashDot", "一点鎖線"], ["lgDash", "長破線"], ["lgDashDot", "長鎖線"]]
        .map(([v, l]) => h("option", { value: v, selected: (ln.dash || "solid") === v }, l)))),
    isLine ? row("矢印", h("select", { disabled: !ed, onchange: e => shapeProps({ head: e.target.value }) },
      ...[["none", "始点: なし"], ["triangle", "始点: 矢印"], ["arrow", "始点: 開いた矢印"], ["oval", "始点: 円"]].map(([v, l]) => h("option", { value: v, selected: ln.head === v }, l))),
      h("select", { disabled: !ed, onchange: e => shapeProps({ tail: e.target.value }) },
        ...[["none", "終点: なし"], ["triangle", "終点: 矢印"], ["arrow", "終点: 開いた矢印"], ["oval", "終点: 円"]].map(([v, l]) => h("option", { value: v, selected: ln.tail === v }, l)))) : null,
    one && s.image ? row(null, h("button", { disabled: !ed, onclick: () => pickImage(s.id) }, "🖼 図の変更…")) : null,
    one && s.table ? row(null, h("button", { disabled: !ed, onclick: () => openTableEditor(s.id) }, "▦ 表の編集…")) : null));

  body.append(sec("配置",
    h("div", { class: "chips" }, ...[["l", "⇤", "左揃え"], ["c", "↔", "左右中央"], ["r", "⇥", "右揃え"], ["t", "⤒", "上揃え"], ["m", "↕", "上下中央"], ["b", "⤓", "下揃え"]]
      .map(([k, g, t]) => h("button", { title: t + (ss.length === 1 ? "（スライドに）" : ""), disabled: !ed, onclick: () => align(k) }, g)),
      h("button", { title: "左右に整列（3 つ以上）", disabled: !ed || ss.length < 3, onclick: () => distribute("h") }, "⋯"),
      h("button", { title: "上下に整列（3 つ以上）", disabled: !ed || ss.length < 3, onclick: () => distribute("v") }, "⋮")),
    h("div", { class: "chips", style: { marginTop: "5px" } }, ...[["front", "最前面へ"], ["forward", "前面へ"], ["backward", "背面へ"], ["back", "最背面へ"]].map(([w, l]) =>
      h("button", { disabled: !ed, onclick: () => op("zorder", { slide: S.cur, ids: S.sel, where: w }) }, l))),
    h("div", { class: "chips", style: { marginTop: "5px" } },
      h("button", { disabled: !ed || ss.length < 2, onclick: groupSel }, "グループ化"),
      h("button", { disabled: !ed || !ss.some(x => x.kind === "grpSp"), onclick: ungroupSel }, "グループ解除"),
      h("button", { disabled: !ed, onclick: duplicate }, "⧉ 複製"), h("button", { disabled: !ed, onclick: delSel }, "🗑 削除"),
      S.group != null ? h("button", { onclick: () => { S.group = null; select([]); } }, "グループから出る (Esc)") : null)));
  const n = S.slide.anims.filter(e => ss.some(x => x.id === e.spid)).length;
  body.append(sec("アニメーション", h("div", { class: "note" }, n ? `この図形には効果が ${n} 個ある。` : "効果は無い。"),
    h("button", { onclick: () => setPane("anim") }, "アニメーションの設定 →")));
}
async function loadFonts() {
  if (S.fonts) return;
  S.fonts = [];
  try { S.fonts = (await api("/api/fonts")).fonts; } catch { }
  let dl = $("#fontlist");
  if (!dl) { dl = h("datalist", { id: "fontlist" }); document.body.append(dl); }
  dl.innerHTML = ""; for (const f of S.fonts) dl.append(h("option", { value: f }));
}
function renameShape(s) {
  const n = prompt("図形の名前（選択ウィンドウに出る名前）", s.name);
  if (n != null && n !== s.name) op("rename", { slide: S.cur, id: s.id, name: n });
}

// ---------------------------------------------------------------- 書式（何も選んでいないとき＝スライド）
function renderSlidePane(body) {
  const sm = S.slide, s = S.deck.slides[S.cur], ed = S.deck.editable;
  body.append(sec(`スライド ${S.cur + 1}`,
    h("div", { class: "note" }, `レイアウト: ${sm.layout || "?"}`, h("br"), `図形 ${sm.shapes.length} 個 ・ アニメーション ${sm.anims.length} 個 ・ 画面切り替え: ${sm.transition?.label || "なし"}`),
    h("div", { class: "chips", style: { marginTop: "8px" } },
      h("button", { disabled: !ed, onclick: () => op("slide_hide", { index: S.cur, hidden: !s.hidden }) }, s.hidden ? "再表示する" : "非表示にする"),
      h("button", { disabled: !ed, onclick: () => op("slide_dup", { index: S.cur }).then(r => r && gotoSlideLater(r.slide)) }, "複製"),
      h("button", { disabled: !ed, onclick: () => confirm(`${S.cur + 1} 枚目を消す（元に戻すで戻せる）`) && op("slide_delete", { index: S.cur }) }, "削除"),
      h("button", { disabled: s.status !== "ready", onclick: saveSlideImage }, "画像で保存"))));
  const bg = sm.background;
  const color = h("input", { type: "color", class: "swatch", value: bg && bg.startsWith("#") ? bg : "#ffffff", disabled: !ed });
  const all = h("input", { type: "checkbox", id: "bgAll" });
  body.append(sec("背景",
    h("div", { class: "note" }, bg ? (bg.startsWith("#") ? `単色 ${bg}` : "このスライドだけの背景（画像やグラデーション）") : "スライド マスターの背景"),
    row("単色", color, h("button", { disabled: !ed, onclick: () => op("background", { slide: S.cur, color: color.value, all: all.checked }) }, "塗る"),
      h("button", { disabled: !ed || (!bg && !all.checked), onclick: () => op("background", { slide: S.cur, color: null, all: all.checked }) }, "マスターに戻す")),
    row(null, all, h("label", { for: "bgAll", style: { width: "auto" } }, "すべてのスライドに当てる")),
    h("div", { class: "note" }, "単色にすると、そのスライドの背景の画像（見出しの帯などが入っていることがある）は外れる。")));
  const warns = [];
  const R = S.lastRender[sm.part];
  if (R && R.fp === sm.fp) R.meta.shapes.forEach((x, i) => {
    const sh = R.map[i] != null ? shapeById(R.map[i]) : null, nm = sh ? (sh.text || []).join(" ").slice(0, 16) || sh.name : x.name;
    if (x.overflow) warns.push([`「${nm}」の文字が枠から ${(x.overflow / 100).toFixed(1)} mm あふれている`, sh?.id]);
    if (x.outside) warns.push([`「${nm}」がスライドの外に ${(x.outside / 100).toFixed(0)} mm はみ出している`, sh?.id]);
  });
  if (warns.length) body.append(sec("気になるところ", ...warns.map(([w, id]) =>
    h("div", { class: "fontwarn", style: { color: "var(--warn)", cursor: id != null ? "pointer" : "default" }, onclick: () => id != null && select([id]) }, "⚠ " + w))));
  const fonts = S.deck.fonts || [];
  if (fonts.length) body.append(sec("フォントの置き換え（この PC に無いフォント）",
    ...fonts.map(f => h("div", { class: "fontwarn" }, h("b", {}, f.font), " → ", f.replaced_by, h("span", { class: "muted" }, `（${f.uses} か所）`),
      f.same_width ? h("span", { class: "badge" }, "文字幅は同じ") : h("span", { class: "badge warn" }, "改行がずれうる"))),
    h("div", { class: "note", style: { marginTop: "6px" } }, "文字幅が違うフォントに置き換わると、PowerPoint で開いたときと改行の位置が変わることがある。元のフォントを ~/.fonts に入れて fc-cache -f すると同じになる。")));
  body.append(sec("操作", h("div", { class: "note", html:
    "クリックで選ぶ・ドラッグで動かす（<kbd>Shift</kbd> で縦横だけ、<kbd>Alt</kbd> で吸着なし）・右クリックでメニュー<br>" +
    "ダブルクリックで文字を編集（<kbd>Shift</kbd>+<kbd>Enter</kbd> で段落内の改行 ↵、<kbd>Esc</kbd> で確定）。文字を選んで <kbd>Ctrl</kbd>+<kbd>B</kbd> などで、その文字だけに書式<br>" +
    "グループはダブルクリックで中に入る・表はダブルクリックで編集<br>" +
    "<kbd>Ctrl</kbd>+<kbd>C</kbd>/<kbd>X</kbd>/<kbd>V</kbd> コピー・切り取り・貼り付け（ほかのアプリの文字・画像も）・<kbd>Ctrl</kbd>+<kbd>D</kbd> 複製<br>" +
    "<kbd>Ctrl</kbd>+<kbd>G</kbd> グループ化・<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>G</kbd> 解除・<kbd>Tab</kbd> 次の図形<br>" +
    "<kbd>Ctrl</kbd>+<kbd>F</kbd> 検索・<kbd>Ctrl</kbd>+<kbd>H</kbd> 置換・<kbd>Ctrl</kbd>+<kbd>Z</kbd> 元に戻す・<kbd>F5</kbd> スライドショー" })));
}

// ---------------------------------------------------------------- アニメーション
const CLS_COLOR = { entr: "var(--entr)", exit: "var(--exit)", emph: "var(--emph)", path: "var(--path)" };
function renderAnimPane(body) {
  const sm = S.slide, ed = S.deck.editable, sel = selShapes();
  const eff = h("select", {}, ANIM_EFFECTS.map(([v, l]) => h("option", { value: v }, l)));
  const trig = h("select", {}, h("option", { value: "clickEffect" }, "クリック時"), h("option", { value: "withEffect" }, "直前の動作と同時"), h("option", { value: "afterEffect" }, "直前の動作の後"));
  body.append(sec("アニメーションの追加",
    row(null, eff), row(null, trig,
      h("button", { class: "primary", disabled: !ed || !sel.length, onclick: () => addAnim(eff.value, trig.value) }, "＋ 追加")),
    h("div", { class: "note" }, sel.length ? `選んでいる ${sel.length} 個の図形に付ける` : "スライドの図形を選ぶと付けられる")));
  body.append(row(null, h("button", { class: "primary", disabled: !sm.anims.length, onclick: () => previewSlide() }, "▶ プレビュー"),
    h("button", { onclick: () => startShow(true) }, "▷ スライドショー")),
    h("div", { class: "note", style: { marginBottom: "8px" } }, "プレビューはスライドの上で効果を順に自動で再生する（クリックか Esc で止める）。スライドショーはクリック・→ で進む。"));
  if (!sm.anims.length) { body.append(h("div", { class: "note" }, "このスライドにアニメーションは無い。")); return; }
  const list = h("div", { class: "anim-list" });
  const groups = [];
  for (const e of sm.anims) { let g = groups.find(x => x.group === e.group); if (!g) groups.push(g = { group: e.group, step: e.step, click: e.click, items: [] }); g.items.push(e); }
  const clicks = groups.filter(x => x.click);
  groups.forEach(g => {
    const ci = clicks.indexOf(g);
    const head = h("div", { class: "anim-head" }, h("span", { class: "n" }, g.click ? g.step : "自動"),
      h("span", {}, g.click ? "クリック" : "スライドが出たとき"), h("span", { class: "sp" }),
      g.click && ed ? h("button", { title: "前へ", disabled: ci <= 0, onclick: () => op("anim_move", { slide: S.cur, group: g.group, delta: -1 }) }, "▲") : null,
      g.click && ed ? h("button", { title: "後へ", disabled: ci >= clicks.length - 1, onclick: () => op("anim_move", { slide: S.cur, group: g.group, delta: 1 }) }, "▼") : null);
    const step = h("div", { class: "anim-step" }, head);
    for (const e of g.items) {
      const tsel = h("select", { disabled: !ed, onclick: ev => ev.stopPropagation(), onchange: ev => op("anim_trigger", { slide: S.cur, n: e.n, trigger: ev.target.value }) },
        ...[["clickEffect", "クリック時"], ["withEffect", "同時"], ["afterEffect", "後"]].map(([v, l]) => h("option", { value: v, selected: e.trigger === v }, l)));
      step.append(h("div", { class: "anim-row" + (S.animSel === e.n ? " sel" : ""), onclick: () => {
        S.animSel = e.n; const top = topOf(e.spid); S.group = top !== e.spid ? top : null; S.sel = [e.spid]; drawSlide(); renderPane();
      } },
        h("span", { class: "dot", style: { background: CLS_COLOR[e.cls] || "#888" }, title: e.cls_label }),
        h("div", { class: "what" }, h("div", { class: "nm" }, `${e.name}: ${e.shape}`), h("div", { class: "tr" }, `${e.cls_label} ・ ${e.trigger_label}${e.dur ? ` ・ ${e.dur} 秒` : ""}`)),
        h("div", {}, tsel, ed ? h("button", { title: "削除", onclick: ev => { ev.stopPropagation(); op("anim_delete", { slide: S.cur, n: e.n }); } }, "✕") : null)));
    }
    list.append(step);
  });
  body.append(list);
}

async function addAnim(effect, trigger = "clickEffect") {
  // 効果を付けたら、PowerPoint と同じく自動でプレビューする（描き直しを待ってから）
  const r = await op("anim_add", { slide: S.cur, ids: S.sel, effect, trigger });
  if (!r) return;
  for (let i = 0; i < 40; i++) { await sleep(150); if (S.slide && S.deck.slides[S.cur]?.fp === S.slide.fp && S.slide.anims.length) break; }
  previewSlide();
}

// ---------------------------------------------------------------- 画面切り替え
const TRANSITIONS = [["none", "なし"], ["fade", "フェード"], ["push", "プッシュ"], ["wipe", "ワイプ"], ["cover", "カバー"], ["pull", "アンカバー"],
  ["split", "スプリット"], ["randomBar", "ランダム"], ["circle", "円"], ["diamond", "ひし形"], ["plus", "プラス"], ["wedge", "くさび形"],
  ["wheel", "ホイール"], ["blinds", "ブラインド"], ["checker", "チェッカー"], ["dissolve", "ディゾルブ"], ["zoom", "ズーム"], ["cut", "カット"]];
function renderTransPane(body) {
  const t = S.slide.transition || { effect: "none", speed: "med" }, ed = S.deck.editable;
  const speed = h("select", { disabled: !ed }, ...[["fast", "速く"], ["med", "普通"], ["slow", "遅く"]].map(([v, l]) => h("option", { value: v, selected: t.speed === v }, l)));
  const autoOn = h("input", { type: "checkbox", id: "advOn", checked: t.advance_after != null, disabled: !ed });
  const autoSec = h("input", { type: "number", step: "0.5", min: "0", value: t.advance_after ?? 5, disabled: !ed });
  const click = h("input", { type: "checkbox", id: "advClick", checked: t.advance_click !== false, disabled: !ed });
  const args = (effect, all) => ({ slide: S.cur, all, effect, speed: speed.value, advance_after: autoOn.checked ? parseFloat(autoSec.value) || 0 : null, advance_click: click.checked });
  body.append(sec("このスライドに入るときの効果",
    h("div", { class: "trans-grid" }, TRANSITIONS.map(([v, l]) => h("button", { class: t.effect === v ? "on" : "", disabled: !ed, onclick: () => op("transition", args(v, false)) }, l)))));
  body.append(sec("タイミング",
    row("速さ", speed),
    row(null, click, h("label", { for: "advClick", style: { width: "auto" } }, "クリックで次へ")),
    row(null, autoOn, h("label", { for: "advOn", style: { width: "auto" } }, "自動で次へ"), autoSec, "秒後"),
    h("div", { class: "chips" }, h("button", { disabled: !ed, onclick: () => op("transition", args(t.effect, false)) }, "このスライドに当てる"),
      h("button", { disabled: !ed, onclick: () => op("transition", args(t.effect, true)) }, "すべてのスライドに当てる"))));
  body.append(row(null, h("button", { onclick: () => startShow(true) }, "▶ このスライドから再生")));
  body.append(h("div", { class: "note" }, "効果は LibreOffice のスライドショーで再生する。PowerPoint でも同じ名前の効果になる。"));
}

// ---------------------------------------------------------------- 選択（PowerPoint の選択ウィンドウ）
function renderSelectPane(body) {
  const ed = S.deck.editable;
  body.append(h("div", { class: "note", style: { marginBottom: "8px" } }, "上ほど前面。クリックで選ぶ（Shift で追加）、👁 で表示・非表示、ダブルクリックで名前を変える。"));
  const list = h("div", { class: "sel-list" });
  const add = (s, depth, groupId) => {
    const txt = (s.text || []).join(" ").replace(/\v/g, " ").trim();
    const el = h("div", { class: "sel-row" + (S.sel.includes(s.id) ? " sel" : "") + (s.hidden ? " off" : ""), style: { paddingLeft: 6 + depth * 14 + "px" } },
      h("button", { title: s.hidden ? "表示する" : "隠す", disabled: !ed, onclick: ev => { ev.stopPropagation(); op("visibility", { slide: S.cur, ids: [s.id], hidden: !s.hidden }); } }, s.hidden ? "◌" : "👁"),
      h("div", { class: "nm", title: s.name }, txt ? `${txt.slice(0, 30)}` : s.name, h("span", { class: "kind" }, `  ${s.label}`)),
      depth === 0 ? h("div", {}, h("button", { title: "前面へ", disabled: !ed, onclick: ev => { ev.stopPropagation(); op("zorder", { slide: S.cur, ids: [s.id], where: "forward" }); } }, "▲"),
        h("button", { title: "背面へ", disabled: !ed, onclick: ev => { ev.stopPropagation(); op("zorder", { slide: S.cur, ids: [s.id], where: "backward" }); } }, "▼")) : h("div"));
    el.addEventListener("click", ev => {
      S.group = groupId ?? null;
      select(ev.shiftKey && (S.group ?? null) === (groupId ?? null) ? [...new Set([...S.sel, s.id])] : [s.id]);
      drawSlide();
    });
    el.addEventListener("dblclick", () => renameShape(s));
    list.append(el);
    if (s.children) [...s.children].reverse().forEach(c => add(c, depth + 1, s.id));
  };
  [...S.slide.shapes].reverse().forEach(s => add(s, 0, null));
  body.append(list);
  if (ed) body.append(h("div", { class: "chips", style: { marginTop: "8px" } },
    h("button", { onclick: () => op("visibility", { slide: S.cur, ids: S.slide.shapes.map(s => s.id), hidden: false }) }, "すべて表示"),
    h("button", { onclick: () => op("visibility", { slide: S.cur, ids: S.slide.shapes.map(s => s.id), hidden: true }) }, "すべて非表示")));
}
