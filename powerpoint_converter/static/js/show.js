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
  show.classList.remove("hidden"); fr.classList.add("hidden"); wait.classList.remove("hidden"); $("#showEnd").classList.add("hidden");
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
    const w = fr.contentWindow; showWin = w;
    try { w.aSlideShow.exitSlideShowInApp = () => showEnd(); } catch { }
    try { if (target > 0) w.aSlideShow.displaySlide(target, true); } catch { }
    w.addEventListener("keydown", e => {
      if (e.key === "Escape") { e.preventDefault(); e.stopImmediatePropagation(); endShow(); return; }
      if (!$("#showEnd").classList.contains("hidden")) {   // 最後の画面
        e.preventDefault(); e.stopImmediatePropagation();
        if (["ArrowLeft", "PageUp", "Backspace", "ArrowUp"].includes(e.key)) backFromEnd(); else endShow();
      }
    }, true);
    wait.classList.add("hidden"); fr.classList.remove("hidden"); w.focus();
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
  $("#showFrame").src = "about:blank"; showWin = null;
  if (document.fullscreenElement) document.exitFullscreen().catch(() => { });
}
$("#showClose").onclick = endShow;
$("#showEnd").addEventListener("click", endShow);
document.addEventListener("fullscreenchange", () => { if (!document.fullscreenElement && showOpen()) endShow(); });

// ---------------------------------------------------------------- アニメーションのプレビュー（編集画面の上で再生）
// LibreOffice や PowerPoint と同じく、効果を付けたときや「プレビュー」を押したときに、スライドの上で動きを見せる。
// そのスライドだけのスライドショーの SVG を作り（非表示のスライドでもよい）、効果を1つずつ自動で進める。
let previewToken = 0;
async function previewSlide() {
  if (!S.slide || !S.slide.anims.length) { toast("このスライドにアニメーションは無い"); return; }
  if (S.editing) await commitEdit();
  const my = ++previewToken, index = S.cur;
  stopPreview(false);
  const ov = h("div", { id: "preview", title: "クリックか Esc で止める" }, h("div", { class: "msg" }, h("span", { class: "spin" }), " プレビューを準備している…"));
  ov.addEventListener("pointerdown", e => e.stopPropagation());   // 下のスライドの図形を選ばない
  ov.addEventListener("click", () => stopPreview());
  const sl = $("#slide"); Object.assign(ov.style, { width: sl.style.width, height: sl.style.height });
  sl.append(ov);
  let r;
  try { r = await api("/api/preview?i=" + index); } catch (e) { toast(e.message, true); stopPreview(); return; }
  if (my !== previewToken || S.cur !== index) return;
  const fr = h("iframe", { title: "プレビュー", tabindex: "-1" });
  fr.onload = async () => {
    const w = fr.contentWindow;
    try { w.aSlideShow.exitSlideShowInApp = () => { }; } catch { }
    try { if (r.page > 0) w.aSlideShow.displaySlide(r.page, true); } catch { }
    $(".msg", ov)?.remove();
    fr.style.visibility = "visible";
    await sleep(500);
    for (let k = 0; k < 300 && my === previewToken; k++) {
      let playing = false;
      try { playing = w.aSlideShow.isAnyEffectPlaying() || w.aSlideShow.isTransitionPlaying(); } catch { break; }
      if (playing) { await sleep(60); continue; }
      let more = false;
      try { more = w.aSlideShow.nextEffect(); } catch { }
      if (!more) break;
      await sleep(150);
    }
    await sleep(900);
    if (my === previewToken) stopPreview();
  };
  fr.style.visibility = "hidden";
  ov.append(fr);
  fr.src = r.url;
}
function stopPreview(bump = true) {
  if (bump) previewToken++;
  $("#preview")?.remove();
}
document.addEventListener("keydown", e => { if (e.key === "Escape" && $("#preview")) { e.stopPropagation(); stopPreview(); } }, true);
