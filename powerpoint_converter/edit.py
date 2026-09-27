"""画面（と CLI の op）からの操作を PPTX の XML に当てる。操作の一覧（OPS）はここにまとめる。

どの操作も、関係する部品の XML だけを書き換える。操作の結果は Package に溜まり、呼び出し側が保存する。
座標はすべて EMU（1 cm = 360000）。frame は {x, y, w, h}（スライドの左上から）。
  図形と文字 … このファイル      スライド・ノート・画面切り替え・背景 … slides.py
  コピー・貼り付け・グループ … clip.py      表 … tables.py      検索と置換 … search.py
"""
from __future__ import annotations

import struct

from . import anim, clip, search, slides, tables
from .oputil import (FILLS, LN_ORDER, RPR_ORDER, SHAPES, SP_ORDER, SPPR_ORDER, _esc, _expand, _ids, _refit_group,
                     _set_xfrm, _slide, append_shape, ensure_xfrm, frag)
from .pptx import (A, NS_R, RT_IMAGE, SHAPE_TAGS, Package, PptxError, all_ids, body_paragraphs, cnvpr, edit_body,
                   find_shape, from_slide, insert_ordered, kid, kids, paragraph_segments, prefix_for, read_xfrm,
                   set_text, text_body, xfrm_el)

PPR_ORDER = ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont",
             "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr", "extLst"]
BULLETS = ["buNone", "buAutoNum", "buChar", "buBlip"]


# ---------------------------------------------------------------- 位置・回転・反転・表示
def op_frame(pkg, args):
    """図形の枠を変える。frames: [{id, x, y, w, h}]（スライドの座標）。move: true なら大きさは変えない。"""
    part, doc = _slide(pkg, args)
    touched_groups = []
    for f in args["frames"]:
        sh, chain = find_shape(doc, int(f["id"]))
        cx = [read_xfrm(xfrm_el(g)) for g in chain]
        if any(c is None or c.get("rot") for c in cx):
            raise PptxError("回転したグループの中の図形は動かせない（グループごと動かす）")
        local = from_slide(cx, f)
        _set_xfrm(sh, local, keep_size=bool(args.get("move")))
        touched_groups.extend(reversed(chain))
    for g in touched_groups:
        _refit_group(g)
    pkg.touch(part)
    return {"label": "移動" if args.get("move") else "サイズ変更"}


def op_rotate(pkg, args):
    part, doc = _slide(pkg, args)
    for i in _ids(args):
        sh, _ = find_shape(doc, i)
        x = ensure_xfrm(pkg, part, sh)
        deg = float(args["deg"]) % 360
        if deg:
            x.setAttribute("rot", str(round(deg * 60000)))
        elif x.hasAttribute("rot"):
            x.removeAttribute("rot")
    pkg.touch(part)
    return {"label": "回転"}


def op_flip(pkg, args):
    """左右（axis: h）・上下（axis: v）に反転する。"""
    part, doc = _slide(pkg, args)
    attr = "flipH" if args.get("axis", "h") == "h" else "flipV"
    for i in _ids(args):
        sh, _ = find_shape(doc, i)
        x = ensure_xfrm(pkg, part, sh)
        if x.getAttribute(attr) in ("1", "true"):
            x.removeAttribute(attr)
        else:
            x.setAttribute(attr, "1")
    pkg.touch(part)
    return {"label": "左右反転" if attr == "flipH" else "上下反転"}


def op_visibility(pkg, args):
    """図形を隠す・出す（PowerPoint の選択ウィンドウの目のボタン）。"""
    part, doc = _slide(pkg, args)
    hidden = bool(args.get("hidden", True))
    for i in _ids(args):
        sh, _ = find_shape(doc, i)
        c = cnvpr(sh)
        if hidden:
            c.setAttribute("hidden", "1")
        elif c.hasAttribute("hidden"):
            c.removeAttribute("hidden")
    pkg.touch(part)
    return {"label": "非表示" if hidden else "表示"}


def op_rename(pkg, args):
    part, doc = _slide(pkg, args)
    sh, _ = find_shape(doc, int(args["id"]))
    cnvpr(sh).setAttribute("name", str(args["name"])[:200])
    pkg.touch(part)
    return {"label": "名前の変更"}


# ---------------------------------------------------------------- 文字
def op_text(pkg, args):
    """図形の文字を変える。text: 段落の列（段落の中の改行は \\v）。"""
    part, doc = _slide(pkg, args)
    sh, _ = find_shape(doc, int(args["id"]))
    tx = text_body(sh)
    if tx is None:
        if sh.localName != "sp":
            raise PptxError("この図形には文字を入れられない")
        tx = frag(doc, '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
                       '<a:p><a:pPr algn="ctr"/><a:endParaRPr lang="ja-JP" altLang="en-US"/></a:p></p:txBody>')[0]
        insert_ordered(sh, tx, SP_ORDER)
    edit_body(tx, [str(t) for t in args["text"]])
    pkg.touch(part)
    return {"label": "文字の編集"}


def _solid(doc, hexcolor: str):
    f = A(doc, "solidFill")
    f.appendChild(A(doc, "srgbClr", {"val": hexcolor.lstrip("#").upper()}))
    return f


def _set_fill(parent, color: str | None, order: list[str]) -> None:
    """塗り（solidFill など）を1つにする。color が 'none' なら塗りなし、None なら何もしない。"""
    if color is None:
        return
    for f in FILLS:
        for e in kids(parent, f):
            parent.removeChild(e)
    doc = parent.ownerDocument
    insert_ordered(parent, A(doc, "noFill") if color == "none" else _solid(doc, color), order)


def _run_rpr(n):
    """run（a:r・a:fld・a:br）の rPr。無ければ作る。"""
    r = kid(n, "rPr")
    if r is None:
        r = A(n.ownerDocument, "rPr", {"lang": "ja-JP", "altLang": "en-US"})
        n.insertBefore(r, n.firstChild)
    return r


def _set_or_remove(el, attr: str, value) -> None:
    if value:
        el.setAttribute(attr, str(value))
    elif el.hasAttribute(attr):
        el.removeAttribute(attr)


def _resolve_toggles(props: dict, rpr) -> dict:
    """b・i・u・strike が "toggle" なら、最初の文字の今の状態から、付けるか外すかを決める（PowerPoint と同じ）。"""
    if not any(v == "toggle" for v in props.values()):
        return props
    out = dict(props)
    for k in ("b", "i", "u", "strike"):
        if out.get(k) == "toggle":
            v = rpr.getAttribute(k) if rpr is not None else ""
            on = v in ("1", "true") if k in ("b", "i") else v not in ("", "none", "noStrike")
            out[k] = not on
    return out


def _first_rpr(tx, start: int = 0):
    for p, kind, el, s, e in _flat(tx):
        if kind in ("r", "fld") and e > start:
            return kid(el, "rPr")
    return None


def _apply_run(r, props: dict) -> None:
    """文字の書式（rPr・endParaRPr）。sz（pt）/ grow（±pt）/ b / i / u / strike / baseline / color / highlight / font"""
    doc = r.ownerDocument
    if props.get("sz"):
        r.setAttribute("sz", str(round(float(props["sz"]) * 100)))
    if props.get("grow"):
        cur = float(r.getAttribute("sz") or 1800) / 100
        r.setAttribute("sz", str(round(max(1, cur + float(props["grow"])) * 100)))
    for k in ("b", "i"):
        if k in props:
            r.setAttribute(k, "1" if props[k] else "0")
    if "u" in props:
        r.setAttribute("u", "sng" if props["u"] else "none")
    if "strike" in props:
        r.setAttribute("strike", "sngStrike" if props["strike"] else "noStrike")
    if "baseline" in props:        # 上付き 30000 / 下付き -25000 / 元に戻す 0
        _set_or_remove(r, "baseline", int(props["baseline"]))
    if props.get("color"):
        _set_fill(r, props["color"], RPR_ORDER)
    if "highlight" in props:
        for e in kids(r, "highlight"):
            r.removeChild(e)
        if props["highlight"] and props["highlight"] != "none":
            h = A(doc, "highlight")
            h.appendChild(A(doc, "srgbClr", {"val": props["highlight"].lstrip("#").upper()}))
            insert_ordered(r, h, RPR_ORDER)
    if "font" in props:            # "" ならテーマのフォントに戻す
        for tag in ("latin", "ea"):
            for e in kids(r, tag):
                r.removeChild(e)
            if props["font"]:
                insert_ordered(r, A(doc, tag, {"typeface": props["font"]}), RPR_ORDER)


def _ppr(p):
    ppr = kid(p, "pPr")
    if ppr is None:
        ppr = A(p.ownerDocument, "pPr")
        p.insertBefore(ppr, p.firstChild)
    return ppr


def _apply_para(p, props: dict) -> None:
    """段落の書式。algn / bullet（none・char・number）/ lvl_delta（±1）/ line_spacing（倍）/ space_before・after（pt）"""
    doc = p.ownerDocument
    if props.get("algn"):
        _ppr(p).setAttribute("algn", props["algn"])
    if props.get("bullet"):
        ppr = _ppr(p)
        for tag in BULLETS + ["buFont"]:
            for e in kids(ppr, tag):
                ppr.removeChild(e)
        b = props["bullet"]
        if b == "none":
            insert_ordered(ppr, A(doc, "buNone"), PPR_ORDER)
            for a in ("marL", "indent"):
                if ppr.hasAttribute(a):
                    ppr.removeAttribute(a)
        else:
            if b == "char":
                insert_ordered(ppr, A(doc, "buFont", {"typeface": "Arial", "panose": "020B0604020202020204",
                                                      "pitchFamily": "34", "charset": "0"}), PPR_ORDER)
                insert_ordered(ppr, A(doc, "buChar", {"char": "•"}), PPR_ORDER)
            else:
                insert_ordered(ppr, A(doc, "buAutoNum", {"type": "arabicPeriod"}), PPR_ORDER)
            lvl = int(ppr.getAttribute("lvl") or 0)
            ppr.setAttribute("marL", str(342900 + lvl * 457200))
            ppr.setAttribute("indent", "-342900")
    if props.get("lvl_delta"):
        ppr = _ppr(p)
        d = int(props["lvl_delta"])
        lvl = max(0, min(8, int(ppr.getAttribute("lvl") or 0) + d))
        _set_or_remove(ppr, "lvl", lvl)
        if ppr.hasAttribute("marL"):
            ppr.setAttribute("marL", str(max(0, int(ppr.getAttribute("marL")) + 457200 * d)))
    if props.get("line_spacing"):
        ppr = _ppr(p)
        for e in kids(ppr, "lnSpc"):
            ppr.removeChild(e)
        ls = A(doc, "lnSpc")
        ls.appendChild(A(doc, "spcPct", {"val": str(round(float(props["line_spacing"]) * 100000))}))
        insert_ordered(ppr, ls, PPR_ORDER)
    for key, tag in (("space_before", "spcBef"), ("space_after", "spcAft")):
        if key in props:
            ppr = _ppr(p)
            for e in kids(ppr, tag):
                ppr.removeChild(e)
            sp = A(doc, tag)
            sp.appendChild(A(doc, "spcPts", {"val": str(round(float(props[key]) * 100))}))
            insert_ordered(ppr, sp, PPR_ORDER)


def _apply_body(tx, props: dict) -> None:
    """テキスト本体。anchor（t・ctr・b）/ autofit（none・shape・shrink）/ vert（horz・eaVert）"""
    bp = kid(tx, "bodyPr")
    if bp is None:
        bp = A(tx.ownerDocument, "bodyPr")
        tx.insertBefore(bp, tx.firstChild)
    if props.get("anchor"):
        bp.setAttribute("anchor", props["anchor"])
    if props.get("vert"):
        bp.setAttribute("vert", props["vert"])
    if props.get("autofit"):
        for tag in ("noAutofit", "normAutofit", "spAutoFit"):
            for e in kids(bp, tag):
                bp.removeChild(e)
        tag = {"none": "noAutofit", "shrink": "normAutofit", "shape": "spAutoFit"}[props["autofit"]]
        el = A(tx.ownerDocument, tag)
        before = kid(bp, "scene3d") or kid(bp, "sp3d") or kid(bp, "flatTx") or kid(bp, "extLst")
        bp.insertBefore(el, before) if before is not None else bp.appendChild(el)


def _apply_shape(sh, props: dict) -> None:
    """図形の書式。fill / line（色）/ line_w（pt）/ line_dash / head・tail（矢印）"""
    doc = sh.ownerDocument
    pr = kid(sh, "spPr")
    if pr is None:
        return
    if props.get("fill") and sh.localName == "sp":
        _set_fill(pr, props["fill"], SPPR_ORDER)
    if any(k in props for k in ("line", "line_w", "line_dash", "head", "tail")):
        ln = kid(pr, "ln")
        if ln is None:
            ln = insert_ordered(pr, A(doc, "ln", {"w": "12700"}), SPPR_ORDER)
        if props.get("line"):
            _set_fill(ln, props["line"], LN_ORDER)
        if props.get("line_w"):
            ln.setAttribute("w", str(round(float(props["line_w"]) * 12700)))
        if props.get("line_dash"):
            for e in kids(ln, "prstDash") + kids(ln, "custDash"):
                ln.removeChild(e)
            if props["line_dash"] != "solid":
                insert_ordered(ln, A(doc, "prstDash", {"val": props["line_dash"]}), LN_ORDER)
        for end, tag in (("head", "headEnd"), ("tail", "tailEnd")):
            if end in props:
                for e in kids(ln, tag):
                    ln.removeChild(e)
                if props[end] and props[end] != "none":
                    insert_ordered(ln, A(doc, tag, {"type": props[end]}), LN_ORDER)


def op_style(pkg, args):
    """選んだ図形全体の書式（グループは中の図形に当てる）。props は _apply_run・_apply_para・_apply_body・_apply_shape"""
    part, doc = _slide(pkg, args)
    props = args.get("props", {})
    shapes = _expand(doc, _ids(args))
    first = next((text_body(sh) for sh in shapes if text_body(sh) is not None), None)
    props = _resolve_toggles(props, _first_rpr(first) if first is not None else None)
    for sh in shapes:
        tx = text_body(sh)
        if tx is not None:
            _apply_body(tx, props)
            for p in body_paragraphs(tx):
                _apply_para(p, props)
                for n in kids(p):
                    if n.localName in ("r", "fld", "br"):
                        _apply_run(_run_rpr(n), props)
                    elif n.localName == "endParaRPr":
                        _apply_run(n, props)
        _apply_shape(sh, props)
    pkg.touch(part)
    return {"label": "書式"}


def _flat(tx) -> list[tuple]:
    """テキスト本体の中の段落と run を、通しの文字位置で並べる（段落の区切りは1文字と数える）。
    戻り値は (段落, 種類, 要素, 開始, 終了) の列。段落そのものは種類 'p' で、その範囲を持つ。"""
    out, pos = [], 0
    for p in body_paragraphs(tx):
        start = pos
        for kind, t, el in paragraph_segments(p):
            out.append((p, kind, el, pos, pos + len(t)))
            pos += len(t)
        out.append((p, "p", p, start, pos))
        pos += 1
    return out


def _split_run(r, c: int):
    """run を c 文字目で2つに分ける（後ろ半分の run を返す）。"""
    t = kid(r, "t")
    s = "".join(n.data for n in t.childNodes if n.nodeType == n.TEXT_NODE)
    r2 = r.cloneNode(True)
    set_text(t, s[:c])
    set_text(kid(r2, "t"), s[c:])
    r.parentNode.insertBefore(r2, r.nextSibling)
    return r2


def style_range(tx, start: int, end: int, props: dict) -> None:
    """テキスト本体の start〜end の文字だけに書式を当てる（範囲の端で run を分ける）。"""
    props = _resolve_toggles(props, _first_rpr(tx, start))
    for cut in (start, end):
        for p, kind, el, s, e in _flat(tx):
            if kind == "r" and s < cut < e:
                _split_run(el, cut - s)
                break
    paras = []
    for p, kind, el, s, e in _flat(tx):
        if kind in ("r", "fld") and s >= start and e <= end and e > s:
            _apply_run(_run_rpr(el), props)
        if kind == "p" and s <= end and e >= start and p not in paras:
            paras.append(p)
    para_props = {k: v for k, v in props.items() if k in ("algn", "bullet", "lvl_delta", "line_spacing",
                                                           "space_before", "space_after")}
    for p in paras if para_props else []:
        _apply_para(p, para_props)


def op_text_style(pkg, args):
    """文字の一部だけの書式。start・end は通しの文字位置（段落の区切りは1文字、段落内の改行も1文字）。"""
    part, doc = _slide(pkg, args)
    sh, _ = find_shape(doc, int(args["id"]))
    tx = text_body(sh)
    if tx is None:
        raise PptxError("この図形には文字が無い")
    start, end = int(args["start"]), int(args["end"])
    if end <= start:
        raise PptxError("文字が選ばれていない")
    style_range(tx, start, end, args.get("props", {}))
    pkg.touch(part)
    return {"label": "文字の書式"}


# ---------------------------------------------------------------- 消す・複製・重なり
def op_delete(pkg, args):
    part, doc = _slide(pkg, args)
    gone = set()
    for i in _ids(args):
        try:
            sh, chain = find_shape(doc, i)
        except PptxError:
            continue          # 一緒に選んだグループの中で、もう消えている
        for c in sh.getElementsByTagNameNS("*", "cNvPr"):
            try:
                gone.add(int(c.getAttribute("id")))
            except ValueError:
                pass
        gone.add(i)
        sh.parentNode.removeChild(sh)
        for g in reversed(chain):
            _refit_group(g)
    anim.delete_for_shapes(doc, gone)
    pkg.touch(part)
    return {"label": "削除", "select": []}


def op_duplicate(pkg, args):
    part, doc = _slide(pkg, args)
    nid = max(all_ids(doc) | {1}) + 1
    d = int(args.get("offset", 180000))
    new = []
    for i in _ids(args):
        sh, chain = find_shape(doc, i)
        c = sh.cloneNode(True)
        new.append(nid)
        nid = clip.renumber(c, nid)
        x = xfrm_el(c)
        if x is not None and kid(x, "off") is not None:
            off = kid(x, "off")
            off.setAttribute("x", str(int(off.getAttribute("x")) + d))
            off.setAttribute("y", str(int(off.getAttribute("y")) + d))
        sh.parentNode.insertBefore(c, sh.nextSibling)
        for g in reversed(chain):
            _refit_group(g)
    pkg.touch(part)
    return {"label": "複製", "select": new}


def op_zorder(pkg, args):
    """重なりの順番。where: front（最前面）/ back（最背面）/ forward（前面へ）/ backward（背面へ）"""
    part, doc = _slide(pkg, args)
    where = args["where"]
    ids = _ids(args)
    if where in ("forward", "back"):     # 複数を動かすとき、互いの順番が崩れない順で動かす
        ids = list(reversed(ids))
    for i in ids:
        sh, _ = find_shape(doc, i)
        parent = sh.parentNode
        sibs = [n for n in kids(parent) if n.localName in SHAPE_TAGS]
        k = sibs.index(sh)
        if where == "front" and k < len(sibs) - 1:
            parent.insertBefore(sh, sibs[-1].nextSibling)
        elif where == "back" and k > 0:
            parent.insertBefore(sh, sibs[0])
        elif where == "forward" and k < len(sibs) - 1:
            parent.insertBefore(sh, sibs[k + 1].nextSibling)
        elif where == "backward" and k > 0:
            parent.insertBefore(sh, sibs[k - 1])
    pkg.touch(part)
    return {"label": "重なりの順番"}


# ---------------------------------------------------------------- 挿入
def _default_frame(pkg, w_ratio=0.3, h_ratio=0.15) -> dict:
    cx, cy = pkg.slide_size()
    w, h = round(cx * w_ratio), round(cy * h_ratio)
    return {"x": (cx - w) // 2, "y": (cy - h) // 2, "w": w, "h": h}


SQUARE = {"ellipse", "star5", "star8", "plus", "heart", "donut", "smileyFace", "sun", "moon", "cloud", "octagon",
          "hexagon", "flowChartConnector", "noSmoking", "blockArc"}


def op_add(pkg, args):
    """図形を入れる。kind は SHAPES のキー（textbox・rect・line・arrow・table など）"""
    part, doc = _slide(pkg, args)
    kind = args.get("kind", "textbox")
    if kind == "table":
        return tables.op_table_add(pkg, args)
    if kind not in SHAPES:
        raise PptxError(f"図形 {kind} は入れられない")
    if kind in SQUARE:
        cx, cy = pkg.slide_size()
        s = round(cy * 0.28)
        f = args.get("frame") or {"x": (cx - s) // 2, "y": (cy - s) // 2, "w": s, "h": s}
    else:
        f = args.get("frame") or _default_frame(pkg, 0.3, 0.12 if kind == "textbox" else 0.2)
    nid = max(all_ids(doc) | {1}) + 1
    name = f"{SHAPES[kind]} {nid}"
    text = _esc(args.get("text", "テキストを入力" if kind == "textbox" else ""))
    if kind in ("line", "arrow"):
        f = args.get("frame") or {**f, "h": 0}
    x = f'<a:xfrm><a:off x="{int(f["x"])}" y="{int(f["y"])}"/><a:ext cx="{int(f["w"])}" cy="{int(f["h"])}"/></a:xfrm>'
    if kind == "textbox":
        xml = (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr>{x}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
               f'<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'
               f'<a:p><a:r><a:rPr lang="ja-JP" altLang="en-US" sz="1800" dirty="0"/><a:t>{text}</a:t></a:r></a:p>'
               f'</p:txBody></p:sp>')
    elif kind in ("line", "arrow"):
        tail = '<a:tailEnd type="triangle"/>' if kind == "arrow" else ""
        xml = (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{nid}" name="{name}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
               f'<p:spPr>{x}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="28575">{tail}</a:ln></p:spPr>'
               f'<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0">'
               f'<a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/>'
               f'</a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>')
    else:
        body = (f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ja-JP" altLang="en-US" dirty="0"/><a:t>{text}</a:t></a:r></a:p>'
                if text else '<a:p><a:pPr algn="ctr"/><a:endParaRPr lang="ja-JP" altLang="en-US"/></a:p>')
        xml = (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr>{x}<a:prstGeom prst="{kind}"><a:avLst/></a:prstGeom></p:spPr>'
               f'<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'
               f'<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0">'
               f'<a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/>'
               f'</a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>{body}</p:txBody>'
               f'</p:sp>')
    append_shape(doc, frag(doc, xml)[0])
    pkg.touch(part)
    return {"label": "図形の挿入", "select": [nid]}


def image_size(data: bytes) -> tuple[str, int, int]:
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", data[16:24])
        return "png", w, h
    if data[:6] in (b"GIF87a", b"GIF89a"):
        w, h = struct.unpack("<HH", data[6:10])
        return "gif", w, h
    if data[:2] == b"\xff\xd8":
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            m = data[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return "jpeg", w, h
            seg = struct.unpack(">H", data[i + 2:i + 4])[0]
            i += 2 + seg
        return "jpeg", 0, 0
    raise PptxError("画像は PNG・JPEG・GIF だけ入れられる")


def add_media(pkg: Package, part: str, data: bytes) -> tuple[str, int, int]:
    """画像を部品として足し、スライドからの関係の id と、画像の大きさ（px）を返す。"""
    ext, pw, ph = image_size(data)
    media = pkg.new_part_name("ppt/media/pc_image{}." + ("jpg" if ext == "jpeg" else ext))
    pkg.put(media, data)
    pkg.ensure_default(media.rsplit(".", 1)[1], {"png": "image/png", "gif": "image/gif", "jpeg": "image/jpeg"}[ext])
    return pkg.add_rel(part, RT_IMAGE, media), pw, ph


def op_image(pkg, args, data: bytes):
    part, doc = _slide(pkg, args)
    rid, pw, ph = add_media(pkg, part, data)
    cx, cy = pkg.slide_size()
    w, h = (pw or 400) * 9525, (ph or 300) * 9525   # 96 dpi
    k = min(1.0, cx * 0.6 / w, cy * 0.6 / h)
    w, h = round(w * k), round(h * k)
    nid = max(all_ids(doc) | {1}) + 1
    name = _esc(args.get("name", "画像"))
    xml = (f'<p:pic><p:nvPicPr><p:cNvPr id="{nid}" name="図 {nid}" descr="{name}"/><p:cNvPicPr>'
           f'<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill>'
           f'<a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>'
           f'<a:xfrm><a:off x="{(cx - w) // 2}" y="{(cy - h) // 2}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
           f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    append_shape(doc, frag(doc, xml)[0])
    pkg.touch(part)
    return {"label": "画像の挿入", "select": [nid]}


def op_replace_image(pkg, args, data: bytes):
    """画像を差し替える（PowerPoint の「図の変更」）。幅は保ち、高さを新しい画像の縦横比に合わせる。"""
    part, doc = _slide(pkg, args)
    sh, _ = find_shape(doc, int(args["id"]))
    blips = sh.getElementsByTagNameNS("*", "blip")
    if sh.localName != "pic" or not blips:
        raise PptxError("画像ではない")
    rid, pw, ph = add_media(pkg, part, data)
    b = blips[0]
    attr = b.getAttributeNodeNS(NS_R, "embed")
    if attr is not None:
        attr.value = rid
    else:
        b.setAttributeNS(NS_R, prefix_for(doc, NS_R, "r") + ":embed", rid)
    src = kid(kid(sh, "blipFill"), "srcRect")
    if src is not None:                     # 切り抜きは新しい画像には合わないので外す
        src.parentNode.removeChild(src)
    x = ensure_xfrm(pkg, part, sh)
    if pw and ph:
        ext = kid(x, "ext")
        ext.setAttribute("cy", str(round(int(ext.getAttribute("cx")) * ph / pw)))
    pkg.touch(part)
    return {"label": "図の変更"}


# ---------------------------------------------------------------- アニメーション
def op_anim_add(pkg, args):
    part, doc = _slide(pkg, args)
    for i in _ids(args):
        find_shape(doc, i)
        anim.add_effect(doc, i, args.get("effect", "appear"), args.get("trigger", "clickEffect"))
    pkg.touch(part)
    return {"label": "アニメーションの追加"}


def op_anim_delete(pkg, args):
    part, doc = _slide(pkg, args)
    anim.delete_effect(doc, int(args["n"]))
    pkg.touch(part)
    return {"label": "アニメーションの削除"}


def op_anim_move(pkg, args):
    part, doc = _slide(pkg, args)
    anim.move_step(doc, int(args["group"]), int(args["delta"]))
    pkg.touch(part)
    return {"label": "アニメーションの順番"}


def op_anim_trigger(pkg, args):
    part, doc = _slide(pkg, args)
    anim.set_trigger(doc, int(args["n"]), args["trigger"])
    pkg.touch(part)
    return {"label": "アニメーションの開始"}


def op_anim_timing(pkg, args):
    part, doc = _slide(pkg, args)
    anim.set_timing(doc, int(args["n"]), args.get("dur"), args.get("delay"))
    pkg.touch(part)
    return {"label": "アニメーションの長さ"}


def op_anim_change(pkg, args):
    part, doc = _slide(pkg, args)
    anim.change_effect(doc, int(args["n"]), args["effect"])
    pkg.touch(part)
    return {"label": "アニメーションの種類"}


OPS = {"frame": op_frame, "rotate": op_rotate, "flip": op_flip, "visibility": op_visibility, "rename": op_rename,
       "text": op_text, "style": op_style, "text_style": op_text_style,
       "delete": op_delete, "duplicate": op_duplicate, "zorder": op_zorder, "add": op_add,
       "anim_add": op_anim_add, "anim_delete": op_anim_delete, "anim_move": op_anim_move,
       "anim_trigger": op_anim_trigger, "anim_timing": op_anim_timing, "anim_change": op_anim_change,
       **slides.OPS, **clip.OPS, **tables.OPS, **search.OPS}
DATA_OPS = {"image": op_image, "replace_image": op_replace_image}   # 画像のバイト列を受け取る操作


def apply(pkg: Package, op: str, args: dict, data: bytes | None = None) -> dict:
    if op in DATA_OPS:
        return DATA_OPS[op](pkg, args, data or b"")
    if op not in OPS:
        raise PptxError(f"操作 {op} は無い")
    return OPS[op](pkg, args)
