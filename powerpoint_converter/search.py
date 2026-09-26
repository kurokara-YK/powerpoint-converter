"""検索と置換（Ctrl+F・Ctrl+H）。図形・グループの中・表のセル・ノートの文字を対象にする。

置換は、見つかった箇所を1つずつ置き換える（まとめて置き換えると、箇所と箇所の間の書式が失われるため）。
"""
from __future__ import annotations

from .pptx import (SHAPE_TAGS, PptxError, body_paragraphs, cnvpr, edit_paragraph, kids, notes_body_shape,
                   paragraph_text, text_body, top_shapes)


def _find(text: str, q: str, start: int, case: bool) -> int:
    return text.find(q, start) if case else text.lower().find(q.lower(), start)


def bodies(pkg, part: str):
    """スライドの中の文字の入れ物（テキスト本体）を (図形の id, 種類, テキスト本体) で返す。"""
    from .tables import tbl_of
    out = []

    def walk(shapes, top):
        for sh in shapes:
            sid = int(cnvpr(sh).getAttribute("id") or 0)
            if sh.localName == "grpSp":
                walk([k for k in kids(sh) if k.localName in SHAPE_TAGS], top or sid)
                continue
            tx = text_body(sh)
            if tx is not None:
                out.append((sid, "shape", tx))
            tbl = tbl_of(sh)
            if tbl is not None:
                for tc in tbl.getElementsByTagNameNS("*", "tc"):
                    t = [k for k in kids(tc) if k.localName == "txBody"]
                    if t:
                        out.append((sid, "table", t[0]))
    walk(top_shapes(pkg.xml(part)), None)
    return out


def _notes_body(pkg, part: str):
    notes = pkg.notes_of(part)
    if not notes or not pkg.has(notes):
        return None, None
    sh = notes_body_shape(pkg.xml(notes))
    return notes, (text_body(sh) if sh is not None else None)


def find_all(pkg, q: str, case: bool = False, notes: bool = True) -> list[dict]:
    """見つかった箇所の一覧。1つの図形に何か所あっても、図形ごとに1行（count に数）。"""
    if not q:
        return []
    out = []
    for i, part in enumerate(pkg.slide_parts()):
        targets = bodies(pkg, part)
        if notes:
            _, tx = _notes_body(pkg, part)
            if tx is not None:
                targets.append((None, "notes", tx))
        for sid, where, tx in targets:
            n, snippet = 0, ""
            for p in body_paragraphs(tx):
                t = paragraph_text(p)
                k = _find(t, q, 0, case)
                while k >= 0:
                    if not snippet:
                        snippet = t[max(0, k - 12):k + len(q) + 18].replace("\v", " ")
                    n += 1
                    k = _find(t, q, k + len(q), case)
            if n:
                out.append({"slide": i, "id": sid, "where": where, "count": n, "text": snippet})
    return out


def replace_in(tx, q: str, rep: str, case: bool) -> int:
    n = 0
    for p in body_paragraphs(tx):
        t = paragraph_text(p)
        k = _find(t, q, 0, case)
        while k >= 0:
            new = t[:k] + rep + t[k + len(q):]
            edit_paragraph(p, new)
            t = new
            n += 1
            k = _find(t, q, k + len(rep), case)
    return n


def op_replace(pkg, args):
    """置換する。slide（0 始まり）と id を渡すとその図形だけ、無ければ全部のスライド。notes でノートも。"""
    q, rep = str(args.get("find", "")), str(args.get("replace", ""))
    if not q:
        raise PptxError("検索する文字が空")
    case = bool(args.get("case"))
    only_slide = args.get("slide")
    only_id = args.get("id")                 # 図形の id か "notes"（ノートだけ）
    notes_only = only_id == "notes"
    total = 0
    for i, part in enumerate(pkg.slide_parts()):
        if only_slide is not None and i != int(only_slide):
            continue
        for sid, where, tx in bodies(pkg, part):
            if notes_only or (only_id is not None and sid != int(only_id)):
                continue
            c = replace_in(tx, q, rep, case)
            if c:
                pkg.touch(part)
                total += c
        if args.get("notes", True) and (only_id is None or notes_only):
            notes, tx = _notes_body(pkg, part)
            if tx is not None:
                c = replace_in(tx, q, rep, case)
                if c:
                    pkg.touch(notes)
                    total += c
    return {"label": f"置換（{total} か所）", "count": total}


OPS = {"replace": op_replace}
