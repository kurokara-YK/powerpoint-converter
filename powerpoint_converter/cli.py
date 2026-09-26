"""powerpoint-converter — PowerPoint の資料をローカルのブラウザで見て直す。Claude Code が直した結果もすぐ見える。

  powerpoint-converter                          ブラウザで data/ の一覧を開く
  powerpoint-converter <資料 | フォルダ>        その資料（フォルダなら一覧）をブラウザで開く
  powerpoint-converter check   <資料>           全スライドを PNG にし、崩れ・あふれ・壊れを スライド番号 で出す
  powerpoint-converter outline <資料>           スライドと図形（id・名前・文字・アニメーション）の一覧
  powerpoint-converter op      <資料> <操作> <JSON>   画面と同じ操作を当てる（書式を保って文字を直すなど）
  powerpoint-converter render  <資料> [-o 出力先]    全スライドを PNG にする
  powerpoint-converter pdf     <資料> [-o 出力.pdf]  PDF にする
  powerpoint-converter import  <資料> [フォルダ]     data/ に取り込む

<資料> は .pptx（直せる）か、.ppt / .odp など（見るだけ）。
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from contextlib import contextmanager
from pathlib import Path

from .pptx import PptxError

DATA = Path(__file__).resolve().parent.parent / "data"


def _data(a) -> Path:
    return Path(a.data).expanduser().resolve() if getattr(a, "data", None) else DATA


def _deck(p: str) -> Path:
    path = Path(p).expanduser().resolve()
    if not path.is_file():
        raise PptxError(f"{p} が無い")
    return path


@contextmanager
def _office():
    """CLI 用の LibreOffice（serve とは別のプロファイル。使用中なら cli-2 … を使う）。"""
    from .office import Office
    o = Office("cli")
    try:
        yield o
    finally:
        o.release()


def _render_parallel(src: Path, out: Path, n_pages: int, width: int, checks: bool) -> list[dict]:
    """全ページを PNG にする。枚数が多ければ、複数の LibreOffice でページを分担する（35 枚で 16 秒 → 数秒）。"""
    import os
    import threading
    from .office import Office
    workers = max(1, min(4, (os.cpu_count() or 4) // 4, n_pages // 6))
    chunks = [list(range(k, n_pages, workers)) for k in range(workers)]
    results: list[dict] = []
    errors: list[Exception] = []

    def work(pages):
        o = Office("cli")
        try:
            results.extend(o.render_deck(src, out, width=width, checks=checks, pages=pages))
        except Exception as e:  # noqa: BLE001
            errors.append(e)
        finally:
            o.release()
    ts = [threading.Thread(target=work, args=(c,)) for c in chunks]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    if errors:
        raise errors[0]
    return sorted(results, key=lambda r: r["index"])


# ---------------------------------------------------------------- serve
def cmd_serve(a) -> None:
    from .server import serve
    data = _data(a)
    data.mkdir(parents=True, exist_ok=True)
    start, deck = "", None
    if a.path:
        p = Path(a.path).expanduser().resolve()
        if p.is_file():
            deck = p
            p = p.parent
        elif not p.is_dir():
            raise PptxError(f"{p} が無い")
        if not p.is_relative_to(data):   # data の外なら、その親を置き場として扱う
            data = p.parent if p.parent != p else p
        start = str(p.relative_to(data)) if p != data else ""
    serve(data, start, deck, a.port, not a.no_browser)


# ---------------------------------------------------------------- check
def validate(pkg) -> tuple[list[str], list[str]]:
    """PowerPoint で開いたときに「修復」を求められる原因になるものを探す。(エラー, 警告)"""
    from . import anim
    from .pptx import find_shape, kid, kids
    errors, warns = [], []
    for name in list(pkg.parts):
        if name.endswith(".rels"):
            owner = name.replace("_rels/", "")[:-5]
            for r in pkg.rels(owner if owner != "" else ""):
                if not r.external and r.target not in pkg.parts and not r.target.endswith("NULL"):
                    errors.append(f"{owner or '(パッケージ)'} の関係 {r.id} の先 {r.target} が無い")
        elif name != "[Content_Types].xml" and pkg.content_type(name) is None:
            errors.append(f"{name} の種類（[Content_Types].xml）が無い")
    ids = [s.getAttribute("id") for s in pkg.slide_ids()]
    if len(set(ids)) != len(ids):
        errors.append("presentation.xml のスライドの id が重なっている")
    secs = [e.getAttribute("id") for s in pkg.pres.documentElement.getElementsByTagNameNS("*", "section")
            for e in kids(kid(s, "sldIdLst") or s, "sldId")]
    if secs and sorted(secs) != sorted(ids):
        errors.append("セクションに入っているスライドと、スライドの一覧が合わない")
    for i, part in enumerate(pkg.slide_parts()):
        doc = pkg.xml(part)
        seen = {}
        for c in doc.documentElement.getElementsByTagNameNS("*", "cNvPr"):
            seen.setdefault(c.getAttribute("id"), []).append(c.getAttribute("name"))
        dup = {k: v for k, v in seen.items() if len(v) > 1}
        if dup:
            warns.append(f"{i + 1}枚目: 図形の id が重なっている（{', '.join(sorted(dup)[:5])}）")
        for e in anim.effects(doc):
            try:
                find_shape(doc, e["spid"])
            except PptxError:
                errors.append(f"{i + 1}枚目: アニメーションの対象 id={e['spid']} の図形が無い")
    return errors, warns


def cmd_check(a) -> int:
    from . import pptx, render
    from .pptx import Package
    path = _deck(a.deck)
    out = Path(a.out).expanduser().resolve() if a.out else render.CACHE / "check" / path.stem
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    report = {"file": str(path), "errors": [], "warnings": [], "slides": [], "fonts": []}
    names: dict[int, dict] = {}
    if pptx.is_editable(path):
        pkg = Package(path)
        e, w = validate(pkg)
        report["errors"] += e
        report["warnings"] += w
        for i in range(len(pkg.slide_parts())):
            m = pptx.slide_model(pkg, i)
            flat = {}

            def walk(ms):
                for s in ms:
                    flat[s["id"]] = s
                    walk(s.get("children", []))
            walk(m["shapes"])
            names[i] = flat
            report["slides"].append({"number": pkg.first_number() + i, "hidden": m["hidden"],
                                     "title": pptx.slide_title(pkg, m["part"]), "anims": len(m["anims"])})
        report["fonts"] = render.font_report(pptx.fonts_used(pkg))
    with tempfile.TemporaryDirectory(prefix="pc-") as d:
        src = Path(d) / ("deck" + path.suffix)
        if pptx.is_editable(path):
            # 描く用の写し: 非表示も出し、図形に id を書き込む（あふれた図形の名前を出すため）
            import zipfile
            pkg = Package(path)
            with zipfile.ZipFile(src, "w", zipfile.ZIP_STORED) as z:
                slides = set(pkg.slide_parts())
                for name in pkg.order:
                    data = pkg.parts[name]
                    if name in slides:
                        data = render.unhide(render.tag_shapes(data))
                    z.writestr(name, data)
        else:
            shutil.copy2(path, src)
        try:
            n_pages = len(report["slides"]) if report["slides"] else None
            if n_pages is None:
                with _office() as o:
                    n_pages = o.page_count(src)
            pages = _render_parallel(src, out, n_pages, 1280, True)
        except Exception as e:  # noqa: BLE001  LibreOffice が開けない・落ちた
            report["errors"].append(f"LibreOffice で描けない: {e}")
            pages = []
    for i, pg in enumerate(pages):
        if i < len(report["slides"]):
            report["slides"][i]["png"] = pg["png"]
        else:
            report["slides"].append({"number": i + 1, "png": pg["png"]})
        for s in pg["shapes"]:
            sid = int(s["desc"][3:]) if s["desc"].startswith("pc:") and s["desc"][3:].isdigit() else None
            info = names.get(i, {}).get(sid, {})
            label = info.get("name") or s["name"] or s["type"]
            text = " ".join(t for t in info.get("text", []) if t).replace("\v", " ")[:30]
            who = f"「{text}」" if text else label
            who += f"（id={sid}）" if sid is not None else ""
            if s.get("overflow"):
                report["warnings"].append(f"{i + 1}枚目 {who}: 文字が枠から {s['overflow'] / 100:.1f} mm あふれている")
            if s.get("outside"):
                report["warnings"].append(f"{i + 1}枚目 {who}: スライドの外に {s['outside'] / 100:.0f} mm はみ出している")
    if a.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
    else:
        n = len(report["slides"])
        hid = sum(1 for s in report["slides"] if s.get("hidden"))
        an = sum(1 for s in report["slides"] if s.get("anims"))
        print(f"{path.name}: {n} 枚（非表示 {hid} 枚・アニメーションあり {an} 枚）")
        for e in report["errors"]:
            print(f"エラー {e}")
        for w in report["warnings"]:
            print(f"警告 {w}")
        for f in report["fonts"]:
            note = "文字幅は同じ" if f["same_width"] else "文字幅が違うので、改行の位置が PowerPoint とずれることがある"
            print(f"フォント {f['font']} → {f['replaced_by']}（{f['uses']} か所。{note}）")
        if pages:
            print(f"画像: {out}/  （slide-01.png … 非表示のスライドも含む）")
        if not report["errors"] and not report["warnings"]:
            print("崩れ・壊れは見つからない")
    bad = report["errors"] or (a.strict and report["warnings"])
    return 1 if bad else 0


# ---------------------------------------------------------------- outline / op
def cmd_outline(a) -> None:
    from . import pptx
    from .pptx import Package
    pkg = Package(_deck(a.deck))
    cx, cy = pkg.slide_size()
    print(f"{Path(a.deck).name}: {len(pkg.slide_parts())} 枚  スライドの大きさ {cx / 360000:.2f} × {cy / 360000:.2f} cm"
          "  （座標は EMU。1 cm = 360000）")
    targets = [int(a.slide) - 1] if a.slide else range(len(pkg.slide_parts()))
    for i in targets:
        m = pptx.slide_model(pkg, i)
        flags = "（非表示）" if m["hidden"] else ""
        print(f"\n■ {i + 1}枚目{flags}  {m['part']}  レイアウト: {m['layout']}")

        def show(ms, depth):
            for s in ms:
                f = s["frame"]
                pos = f"x={f['x']} y={f['y']} w={f['w']} h={f['h']}" if f else "位置は継承"
                ph = f" [{s['ph']['type']}]" if s["ph"] else ""
                print(f"{'  ' * depth}  id={s['id']} {s['label']}{ph} 「{s['name']}」 {pos}")
                for t in s.get("text", []):
                    if t:
                        print(f"{'  ' * depth}      │ {t.replace(chr(11), '↵')}")
                show(s.get("children", []), depth + 1)
        show(m["shapes"], 0)
        for e in m["anims"]:
            print(f"    ▶ {e['step']}: {e['trigger_label']} {e['cls_label']} {e['name']} → id={e['spid']}")
        if m["notes"]:
            print("    ノート: " + " / ".join(t for t in m["notes"] if t)[:200])


def cmd_op(a) -> None:
    from . import edit
    from .pptx import Package
    path = _deck(a.deck)
    args = json.loads(a.args) if a.args else {}
    pkg = Package(path)
    r = edit.apply(pkg, a.op, args)
    if not pkg.flush():
        print("変更なし")
        return
    pkg.save()
    print(f"{r.get('label', a.op)}: {path.name} を保存した")


# ---------------------------------------------------------------- render / pdf / import
def cmd_render(a) -> None:
    from . import pptx, render
    from .pptx import Package
    path = _deck(a.deck)
    out = Path(a.out).expanduser().resolve() if a.out else path.with_suffix("")
    out = out if a.out else Path.cwd() / f"{path.stem}_png"
    with tempfile.TemporaryDirectory(prefix="pc-") as d:
        src = Path(d) / ("deck" + path.suffix)
        if pptx.is_editable(path):
            render.full_copy(Package(path), src, show_hidden=True)
        else:
            shutil.copy2(path, src)
        with _office() as o:
            n_pages = o.page_count(src)
        pages = _render_parallel(src, out, n_pages, a.width, False)
    print(f"{len(pages)} 枚を {out}/ に書いた")


def cmd_pdf(a) -> None:
    path = _deck(a.deck)
    out = Path(a.out).expanduser().resolve() if a.out else path.with_suffix(".pdf")
    with tempfile.TemporaryDirectory(prefix="pc-") as d:
        src = Path(d) / ("deck" + path.suffix)
        shutil.copy2(path, src)
        with _office() as o:
            o.export(src, out, "impress_pdf_Export")
    print(f"PDF: {out}")


def cmd_import(a) -> None:
    src = _deck(a.file)
    dest_dir = Path(a.dest).expanduser().resolve() if a.dest else _data(a) / "取り込み"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest, n = dest_dir / src.name, 2
    while dest.exists():
        dest, n = dest_dir / f"{src.stem}_{n}{src.suffix}", n + 1
    shutil.copy2(src, dest)
    print(f"取り込み: {dest}\n次は  powerpoint-converter '{dest}'")


def main(argv=None) -> None:
    argv = list(sys.argv[1:] if argv is None else argv)
    cmds = {"serve", "check", "outline", "op", "render", "pdf", "import"}
    if not argv or (argv[0] not in cmds and argv[0] not in ("-h", "--help")):
        argv = ["serve"] + argv
    ap = argparse.ArgumentParser(prog="powerpoint-converter", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd")
    s = sub.add_parser("serve", help="ブラウザで開く")
    s.add_argument("path", nargs="?")
    s.add_argument("--port", type=int, default=8766)
    s.add_argument("--no-browser", action="store_true")
    s.add_argument("--data", help="資料の置き場（既定は data/）")
    c = sub.add_parser("check", help="全スライドを描いて確かめる")
    c.add_argument("deck")
    c.add_argument("--json", action="store_true")
    c.add_argument("--out", help="PNG の出力先")
    c.add_argument("--strict", action="store_true", help="警告があっても終了コード 1 にする")
    o = sub.add_parser("outline", help="スライドと図形の一覧")
    o.add_argument("deck")
    o.add_argument("--slide", help="その番号のスライドだけ")
    p = sub.add_parser("op", help="画面と同じ操作を当てる")
    p.add_argument("deck")
    p.add_argument("op")
    p.add_argument("args", nargs="?", default="{}")
    r = sub.add_parser("render", help="全スライドを PNG にする")
    r.add_argument("deck")
    r.add_argument("-o", "--out")
    r.add_argument("--width", type=int, default=1600)
    f = sub.add_parser("pdf", help="PDF にする")
    f.add_argument("deck")
    f.add_argument("-o", "--out")
    i = sub.add_parser("import", help="data/ に取り込む")
    i.add_argument("file")
    i.add_argument("dest", nargs="?")
    i.add_argument("--data")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "check":
            sys.exit(cmd_check(a))
        {"serve": cmd_serve, "outline": cmd_outline, "op": cmd_op, "render": cmd_render, "pdf": cmd_pdf,
         "import": cmd_import}[a.cmd](a)
    except (PptxError, json.JSONDecodeError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(2)
    except Exception as e:
        from .office import OfficeError
        if isinstance(e, OfficeError):
            print(f"エラー: {e}", file=sys.stderr)
            sys.exit(2)
        raise
