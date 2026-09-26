"""操作（edit・slides・clip・tables・search）で共通に使う XML の補助。"""
from __future__ import annotations

import re
from xml.dom import minidom

from .pptx import (A, NS_A, NS_P, NS_R, P, SHAPE_TAGS, Package, PptxError, find_shape, insert_ordered, kid, kids,
                   prefix_for, read_xfrm, shape_pr, xfrm_el)

SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
              "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"]
RPR_ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst", "effectDag",
             "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick",
             "hlinkMouseOver", "rtl", "extLst"]
LN_ORDER = ["noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash", "round", "bevel", "miter",
            "headEnd", "tailEnd", "extLst"]
SP_ORDER = ["nvSpPr", "spPr", "style", "txBody", "extLst"]
FILLS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]
TEXT_PH = {"title", "ctrTitle", "subTitle", "body", "obj"}
# 入れられる図形（prstGeom の名前 → PowerPoint の日本語の名前）
SHAPES = {"textbox": "テキスト ボックス", "line": "直線コネクタ", "arrow": "直線矢印コネクタ",
          "rect": "正方形/長方形", "roundRect": "四角形: 角を丸くする", "snip1Rect": "四角形: 1 つの角を切り取る",
          "ellipse": "楕円", "triangle": "二等辺三角形", "rtTriangle": "直角三角形", "diamond": "ひし形",
          "pentagon": "五角形", "hexagon": "六角形", "octagon": "八角形", "parallelogram": "平行四辺形",
          "trapezoid": "台形", "can": "円柱", "cube": "直方体", "donut": "円: 塗りつぶしなし", "plus": "十字形",
          "rightArrow": "矢印: 右", "leftArrow": "矢印: 左", "upArrow": "矢印: 上", "downArrow": "矢印: 下",
          "leftRightArrow": "矢印: 左右", "chevron": "矢印: 山形", "homePlate": "矢印: 五方向", "curvedRightArrow": "矢印: 右カーブ",
          "wedgeRectCallout": "吹き出し: 四角形", "wedgeRoundRectCallout": "吹き出し: 角を丸めた四角形",
          "wedgeEllipseCallout": "吹き出し: 円形", "cloudCallout": "思考の吹き出し: 雲形",
          "star5": "星: 5 pt", "star8": "星: 8 pt", "heart": "ハート", "sun": "太陽", "moon": "月", "cloud": "雲",
          "smileyFace": "スマイル", "noSmoking": "禁止", "flowChartProcess": "フローチャート: 処理",
          "flowChartDecision": "フローチャート: 判断", "flowChartTerminator": "フローチャート: 端子",
          "flowChartDocument": "フローチャート: 書類", "flowChartMagneticDisk": "フローチャート: 磁気ディスク",
          "leftBrace": "中かっこ（左）", "rightBrace": "中かっこ（右）", "leftBracket": "大かっこ（左）",
          "rightBracket": "大かっこ（右）"}


def frag(doc, xml: str) -> list:
    """p: a: r: を使って書いた XML の断片を、その文書の要素にする（文書の接頭辞に合わせる）。"""
    pre = {"p": prefix_for(doc, NS_P, "p"), "a": prefix_for(doc, NS_A, "a"), "r": prefix_for(doc, NS_R, "r")}
    for k, v in pre.items():
        if v != k:
            xml = re.sub(rf"(</?){k}:", rf"\g<1>{v}:" if v else r"\1", xml)
            xml = re.sub(rf"(\s){k}:(\w+=)", rf"\g<1>{v}:\2" if v else r"\1\2", xml)
    decl = " ".join(f'xmlns:{v}="{ns}"' if v else f'xmlns="{ns}"'
                    for v, ns in ((pre["p"], NS_P), (pre["a"], NS_A), (pre["r"], NS_R)))
    tmp = minidom.parseString(f"<wrap {decl}>{xml}</wrap>")
    return [doc.importNode(n, True) for n in tmp.documentElement.childNodes if n.nodeType == n.ELEMENT_NODE]


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


# ---------------------------------------------------------------- 共通
def _slide(pkg: Package, args: dict):
    part = pkg.slide_part(int(args["slide"]))
    return part, pkg.xml(part)


def _ids(args) -> list[int]:
    ids = args.get("ids")
    if ids is None and "id" in args:
        ids = [args["id"]]
    if not ids:
        raise PptxError("図形が選ばれていない")
    return [int(i) for i in ids]


def _set_xfrm(sh, local: dict, keep_size: bool = False) -> None:
    """図形の位置と大きさ（その図形の座標系での値）を書く。xfrm が無ければ作る。"""
    doc = sh.ownerDocument
    x = xfrm_el(sh)
    if x is None:
        if sh.localName == "graphicFrame":
            x = P(doc, "xfrm")
            sh.insertBefore(x, kid(sh, "graphic"))
        else:
            pr = shape_pr(sh)
            if pr is None:
                raise PptxError("この図形は動かせない")
            x = insert_ordered(pr, A(doc, "xfrm"), SPPR_ORDER)
    off, ext = kid(x, "off"), kid(x, "ext")
    if off is None:
        off = A(doc, "off")
        x.insertBefore(off, x.firstChild)
    if ext is None:
        ext = A(doc, "ext")
        x.insertBefore(ext, off.nextSibling)
    off.setAttribute("x", str(int(local["x"])))
    off.setAttribute("y", str(int(local["y"])))
    if not keep_size:
        ext.setAttribute("cx", str(max(0, int(local["w"]))))
        ext.setAttribute("cy", str(max(0, int(local["h"]))))
    if sh.localName == "grpSp" and kid(x, "chOff") is None:   # グループは子の座標系も要る
        co, ce = A(doc, "chOff"), A(doc, "chExt")
        co.setAttribute("x", off.getAttribute("x"))
        co.setAttribute("y", off.getAttribute("y"))
        ce.setAttribute("cx", ext.getAttribute("cx") or "0")
        ce.setAttribute("cy", ext.getAttribute("cy") or "0")
        x.appendChild(co)
        x.appendChild(ce)


def _refit_group(g) -> None:
    """グループの子を動かしたあと、グループの枠を子の範囲に合わせる（拡大率は保つ）。"""
    gx = read_xfrm(xfrm_el(g))
    if not gx or "ch" not in gx:
        return
    boxes = []
    for k in kids(g):
        if k.localName in SHAPE_TAGS:
            f = read_xfrm(xfrm_el(k))
            if not f:
                return
            boxes.append(f)
    if not boxes:
        return
    ch = gx["ch"]
    x0 = min(b["x"] for b in boxes)
    y0 = min(b["y"] for b in boxes)
    x1 = max(b["x"] + b["w"] for b in boxes)
    y1 = max(b["y"] + b["h"] for b in boxes)
    sx = gx["w"] / ch["w"] if ch["w"] else 1
    sy = gx["h"] / ch["h"] if ch["h"] else 1
    x = xfrm_el(g)
    kid(x, "off").setAttribute("x", str(round(gx["x"] + (x0 - ch["x"]) * sx)))
    kid(x, "off").setAttribute("y", str(round(gx["y"] + (y0 - ch["y"]) * sy)))
    kid(x, "ext").setAttribute("cx", str(round((x1 - x0) * sx)))
    kid(x, "ext").setAttribute("cy", str(round((y1 - y0) * sy)))
    kid(x, "chOff").setAttribute("x", str(x0))
    kid(x, "chOff").setAttribute("y", str(y0))
    kid(x, "chExt").setAttribute("cx", str(x1 - x0))
    kid(x, "chExt").setAttribute("cy", str(y1 - y0))


def _expand(doc, ids: list[int]) -> list:
    """グループを、中の図形（文字を持ちうるもの）に展開する。"""
    out = []
    for i in ids:
        sh, _ = find_shape(doc, i)
        stack = [sh]
        while stack:
            s = stack.pop()
            if s.localName == "grpSp":
                stack.extend(k for k in kids(s) if k.localName in SHAPE_TAGS)
            else:
                out.append(s)
    return out




def append_shape(doc, el):
    """図形をスライドの一番上（最前面）に足す。"""
    from .pptx import sp_tree
    t = sp_tree(doc)
    ext = kid(t, "extLst")
    t.insertBefore(el, ext) if ext is not None else t.appendChild(el)


def ensure_xfrm(pkg: Package, part: str, sh):
    """図形の xfrm（無ければ、レイアウトから継承している位置を書き込んで作る）。回転・反転の前に使う。"""
    from .pptx import inherited_xfrm
    x = xfrm_el(sh)
    if x is not None and kid(x, "off") is not None:
        return x
    f = inherited_xfrm(pkg, part, sh)
    if f is None:
        raise PptxError("この図形は位置を持たない（一度動かしてから）")
    _set_xfrm(sh, f)
    return xfrm_el(sh)
