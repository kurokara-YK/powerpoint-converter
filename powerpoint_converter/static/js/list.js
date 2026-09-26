"use strict";
// data/ の一覧（フォルダと資料）。資料をドロップすると、開いているフォルダに取り込む

function crumbs(parts, last) {
  const c = $("#crumbs"); c.innerHTML = "";
  c.append(h("a", { href: "#list=" }, "data"));
  let acc = "";
  for (const p of parts) {
    acc = acc ? acc + "/" + p : p;
    c.append(h("span", { class: "sep" }, "›"), h("a", { href: "#list=" + encodeURIComponent(acc) }, p));
  }
  if (last) c.append(h("span", { class: "sep" }, "›"), h("span", { class: "cur" }, last));
}

async function route() {
  const hash = decodeURIComponent(location.hash.slice(1));
  if (hash.startsWith("deck=")) return openDeck(hash.slice(5));
  const path = hash.startsWith("list=") ? hash.slice(5) : (S.info?.start || "");
  return showList(path);
}
window.addEventListener("hashchange", route);

async function showList(path) {
  if (S.editing) await commitEdit();
  if (typeof flushNotes === "function") await flushNotes();
  S.pollToken++; S.deck = null; S.view = "list"; S.listPath = path;
  $("#editor").classList.add("hidden"); $("#list").classList.remove("hidden");
  $("#saved").textContent = "";
  crumbs(path ? path.split("/") : []);
  document.title = "PowerPoint Converter";
  let d;
  try { d = await api("/api/browse?path=" + encodeURIComponent(path)); }
  catch (e) { $("#listBody").innerHTML = ""; $("#listBody").append(h("p", {}, e.message)); return; }
  const body = $("#listBody"); body.innerHTML = "";
  if (d.folders.length) {
    body.append(h("h2", {}, path ? "フォルダ" : "ワークスペース"));
    body.append(h("div", { class: "cards" }, d.folders.map(f =>
      h("div", { class: "card folder", onclick: () => location.hash = "list=" + encodeURIComponent(f.path) },
        h("div", { class: "thumb" }, "📁"),
        h("div", { class: "meta" }, h("div", { class: "name" }, f.name), h("div", { class: "sub" }, `資料 ${f.count} 件 ・ ${fmtTime(f.mtime)}`))))));
  }
  body.append(h("h2", {}, "資料"));
  if (!d.decks.length) body.append(h("p", { class: "muted" }, path ? "このフォルダには資料が無い。下の枠にドロップして取り込む。" : "data/ の直下に資料は無い。フォルダ（ワークスペース）を開くか、作る。"));
  else body.append(h("div", { class: "cards" }, d.decks.map(f => {
    const img = h("img", { src: "/thumb?path=" + encodeURIComponent(f.path) + "&t=" + f.mtime, alt: "", onerror: e => { e.target.replaceWith("📊"); } });
    return h("div", { class: "card", onclick: () => location.hash = "deck=" + encodeURIComponent(f.path) },
      h("div", { class: "thumb" }, img),
      h("div", { class: "meta" }, h("div", { class: "name" }, f.name, f.editable ? null : h("span", { class: "badge" }, "見るだけ")),
        h("div", { class: "sub" }, `${fmtTime(f.mtime)} ・ ${fmtSize(f.size)}`)));
  })));
  const inp = h("input", { type: "text", placeholder: path ? "新しいフォルダの名前" : "新しいワークスペースの名前" });
  const make = async () => {
    if (!inp.value.trim()) return;
    try { const r = await post("/api/mkdir", { path, name: inp.value }); location.hash = "list=" + encodeURIComponent(r.path); }
    catch (e) { toast(e.message, true); }
  };
  inp.addEventListener("keydown", e => { if (e.key === "Enter") make(); });
  body.append(h("div", { class: "newws" }, inp, h("button", { class: "primary", onclick: make }, "作る")));
}

async function importFiles(files) {
  let last;
  for (const f of files) {
    try {
      const r = await api(`/api/import?path=${encodeURIComponent(S.listPath)}&name=${encodeURIComponent(f.name)}`, { method: "POST", body: f });
      last = r.path; toast(`${f.name} を取り込んだ`);
    } catch (e) { toast(e.message, true); }
  }
  if (last && files.length === 1) location.hash = "deck=" + encodeURIComponent(last); else showList(S.listPath);
}
{
  const d = $("#drop"), f = $("#dropFile");
  d.onclick = () => f.click();
  f.onchange = () => importFiles([...f.files]);
  $("#list").addEventListener("dragover", e => { e.preventDefault(); d.classList.add("over"); });
  $("#list").addEventListener("dragleave", e => { if (e.target === $("#list")) d.classList.remove("over"); });
  $("#list").addEventListener("drop", e => { e.preventDefault(); d.classList.remove("over"); importFiles([...e.dataTransfer.files]); });
}
