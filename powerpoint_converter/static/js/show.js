"use strict";
// スライドショー。LibreOffice が書き出した SVG（LibreOffice 自身の再生の仕組み入り）を全画面で開く。
//
// Esc で終わるための工夫:
//   - 全画面のときは、最初の Esc をブラウザが「全画面の解除」に使い、ページに届かない → 全画面が解けたら終わる
//   - SVG の中の再生の仕組みは、Esc で LibreOffice のアプリ向けの終了（exitSlideShowInApp）を呼ぶだけで、
//     ブラウザでは何も起きない → SVG の中で Esc を先に受け取って終わる
//   - 最後のスライドの次へ進むと、再生の仕組みは同じ終了を呼んで 1 枚目に戻る
//     → PowerPoint と同じ「スライド ショーの最後です」を出す

const showOpen = () => !$("#show").classList.contains("hidden");
let showWin = null;

async function startShow(fromCurrent) {
  if (!S.deck) return;
  if (S.editing) await commitEdit();
  const show = $("#show"), wait = $("#showWait"), fr = $("#showFrame");
  // 読み込みの間も display: none にはしない（visibility で隠す）。display: none の中では図形の大きさが 0 と測られ、
  // 再生の仕組みが「元の大きさ」を 0 と覚えて、ズームなどが 0 で割って止まる・位置がずれるため
  show.classList.remove("hidden"); fr.classList.remove("hidden"); fr.style.visibility = "hidden";
  wait.classList.remove("hidden"); $("#showEnd").classList.add("hidden");
  wait.innerHTML = '<span class="spin"></span><br>スライドショーを準備している（LibreOffice で書き出し中。大きい資料は十数秒）<br><span style="font-size:12px;color:#888">Esc で戻る</span>';
  try { await show.requestFullscreen?.(); } catch { }
  let st;
  for (let i = 0; i < 600; i++) {
    if (!showOpen()) return;
    try { st = await api("/api/present"); } catch (e) { wait.textContent = e.message; return; }
    if (st.status === "ready") break;
    if (st.status === "error") { wait.textContent = "スライドショーを作れなかった（Esc で戻る）"; return; }
    await sleep(500);
  }
  if (!showOpen()) return;
  const target = visibleIndex(fromCurrent ? S.cur : 0);
  fr.onload = () => {
    const w = fr.contentWindow;
    if (!w?.aSlideShow || showWin === w) return;   // 空のページ（about:blank）の load と、2回目の load は無視する
    showWin = w;
    wait.classList.add("hidden"); fr.style.visibility = "visible";      // 見える状態にしてからスライドを出す
    try { w.aSlideShow.exitSlideShowInApp = () => showEnd(); } catch { }
    try { if (target > 0) w.aSlideShow.displaySlide(target, true); } catch { }
    w.addEventListener("keydown", e => {
      if (e.key === "Escape") { e.preventDefault(); e.stopImmediatePropagation(); endShow(); return; }
      if (!$("#showEnd").classList.contains("hidden")) {   // 最後の画面
        e.preventDefault(); e.stopImmediatePropagation();
        if (["ArrowLeft", "PageUp", "Backspace", "ArrowUp"].includes(e.key)) backFromEnd(); else endShow();
      }
    }, true);
    w.focus();
  };
  fr.src = "/present.svg?" + Date.now();
}
function showEnd() { $("#showEnd").classList.remove("hidden"); showWin?.focus(); }
function backFromEnd() {
  $("#showEnd").classList.add("hidden");
  try { const n = showWin.theMetaDoc.nNumberOfSlides; showWin.aSlideShow.displaySlide(n - 1, true); } catch { }
  showWin?.focus();
}
function visibleIndex(i) {
  // 非表示のスライドはスライドショーに出ないので、その分を詰めた番号にする
  const sl = S.deck.slides; let k = 0;
  for (let j = 0; j < sl.length; j++) { if (j >= i && !sl[j].hidden) return k; if (!sl[j].hidden) k++; }
  return Math.max(0, k - 1);
}
function endShow() {
  if (!showOpen()) return;
  $("#show").classList.add("hidden"); $("#showEnd").classList.add("hidden");
  $("#showFrame").src = "about:blank"; $("#showFrame").style.visibility = "hidden"; showWin = null;
  if (document.fullscreenElement) document.exitFullscreen().catch(() => { });
}
$("#showClose").onclick = endShow;
$("#showEnd").addEventListener("click", endShow);
document.addEventListener("fullscreenchange", () => { if (!document.fullscreenElement && showOpen()) endShow(); });

// ---------------------------------------------------------------- アニメーションのプレビュー（編集画面の上で再生）
// LibreOffice や PowerPoint と同じく、効果を付けたときや「プレビュー」を押したときに、スライドの上で動きを見せる。
// そのスライドだけのスライドショーの SVG を作り（非表示のスライドでもよい）、効果を1つずつ自動で進める。
let previewToken = 0;
const PREVIEW_HOLD = 1200;   // 自動再生で、1回のクリックの効果が終わってから次のクリックまで置く時間（ms）
let PV = null;               // 再生中のプレビュー {ss, my, n, clicks, auto, ended, step, btnAuto}
async function previewSlide(autoPlay = true) {
  if (!S.slide || !S.slide.anims.length) { toast("このスライドにアニメーションは無い"); return; }
  if (S.editing) await commitEdit();
  const my = ++previewToken, index = S.cur;
  const clicks = new Set(S.slide.anims.filter(e => e.click).map(e => e.group)).size;
  stopPreview(false);
  const step = h("span", { class: "step" });
  const btnAuto = h("button", { title: "自動再生・一時停止" }, "⏸");
  const bar = h("div", { class: "bar" }, step,
    btnAuto, h("button", { title: "次のクリックへ（スライドをクリックしても進む）", onclick: () => previewNext() }, "次へ ▶"),
    h("button", { title: "最初からもう一度", onclick: () => previewSlide(true) }, "⟲"),
    h("button", { title: "閉じる (Esc)", onclick: () => stopPreview() }, "✕"));
  const ov = h("div", { id: "preview", title: "クリックで次へ・Esc で閉じる" },
    h("div", { class: "msg" }, h("span", { class: "spin" }), " プレビューを準備している…"), bar);
  ov.addEventListener("pointerdown", e => e.stopPropagation());   // 下のスライドの図形を選ばない
  ov.addEventListener("click", e => { if (!e.target.closest(".bar")) previewNext(); });
  btnAuto.addEventListener("click", e => { e.stopPropagation(); if (!PV) return; PV.auto = !PV.auto; syncBar(); if (PV.auto) autoLoop(PV); });
  const sl = $("#slide"); Object.assign(ov.style, { width: sl.style.width, height: sl.style.height });
  sl.append(ov);
  let r;
  try { r = await api("/api/preview?i=" + index); } catch (e) { toast(e.message, true); stopPreview(); return; }
  if (my !== previewToken || S.cur !== index) return;
  // src を決めてから入れる。src の無い iframe を入れると、Chrome は空のページの load を先に知らせ、
  // その load でも再生を始めると、クリックを送る処理が2つ動いて、効果が次々に飛ばされる
  const fr = h("iframe", { title: "プレビュー", tabindex: "-1", src: r.url });
  fr.style.visibility = "hidden";     // display: none にはしない（図形の大きさが 0 と測られるため）
  let started = false;
  fr.addEventListener("load", () => {
    const w = fr.contentWindow;
    if (started || !w?.aSlideShow || my !== previewToken) return;   // スライドショーの SVG の load で1回だけ
    started = true;
    const ss = w.aSlideShow;
    try { ss.exitSlideShowInApp = () => { }; } catch { }
    fr.style.visibility = "visible";
    $(".msg", ov)?.remove();
    try { if (r.page > 0) ss.displaySlide(r.page, true); } catch { }
    PV = { ss, my, n: 0, clicks, auto: autoPlay, ended: false, step, btnAuto, busy: false };
    syncBar();
    if (PV.auto) autoLoop(PV);
  });
  ov.insertBefore(fr, bar);
}
const pvPlaying = pv => { try { return pv.ss.isAnyEffectPlaying() || pv.ss.isTransitionPlaying(); } catch { return false; } };
async function pvWaitDone(pv) { while (pv.my === previewToken && pvPlaying(pv)) await sleep(40); }
function syncBar() {
  if (!PV) return;
  PV.step.textContent = PV.ended ? `最後まで再生した（クリック ${PV.clicks} 回）` : `クリック ${PV.n} / ${PV.clicks}`;
  PV.btnAuto.textContent = PV.auto && !PV.ended ? "⏸" : "▶";
  PV.btnAuto.disabled = PV.ended;
}
async function pvStep(pv) {
  // 1クリック分進める。効果が動いている間は進めない（進めると、動いている効果が最後まで飛ばされる）。
  // 動いている間のクリックは1回分だけ覚えておき、終わってから進める
  if (pv.ended) return false;
  if (pv.busy) { pv.pending = true; return false; }
  pv.busy = true;
  let more;
  do {
    pv.pending = false;
    await pvWaitDone(pv);
    more = false;
    if (pv.my === previewToken) { try { more = pv.ss.nextEffect(); } catch { } }
    if (more) { pv.n++; syncBar(); await sleep(60); await pvWaitDone(pv); }
    else pv.ended = true;
  } while (more && pv.pending && pv.my === previewToken);
  pv.busy = false;
  if (pv.my === previewToken) syncBar();
  return more;
}
async function autoLoop(pv) {
  if (pv.looping) return;
  pv.looping = true;
  await pvWaitDone(pv);                          // スライドが出たときに自動で始まる効果
  while (pv.my === previewToken && pv.auto && !pv.ended) {
    await sleep(PREVIEW_HOLD);                   // 今の状態を見せてから、次のクリックへ
    if (pv.my !== previewToken || !pv.auto) break;
    await pvStep(pv);
  }
  pv.looping = false;
  syncBar();                                     // 最後まで再生したら、その状態のまま止まる（閉じない）
}
function previewNext() {
  if (!PV) return;
  PV.auto = false; syncBar();                    // 手で進めたら、自動再生は止める
  pvStep(PV);
}
function stopPreview(bump = true) {
  if (bump) previewToken++;
  PV = null;
  $("#preview")?.remove();
}
document.addEventListener("keydown", e => { if (e.key === "Escape" && $("#preview")) { e.stopPropagation(); stopPreview(); } }, true);
