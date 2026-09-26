"""コピー・貼り付けと、グループ化・グループ解除。

コピーは、図形の XML と、それが指す部品（画像など）の中身を覚えておく。
別のスライド・別の資料にも貼り付けられる（画像は貼り付け先に部品として足す）。
グラフ・SmartArt・埋め込みオブジェクトは、関係する部品が多いので、別のスライドへはコピーしない。
"""
from __future__ import annotations

from xml.dom import minidom

from . import anim
from .oputil import _ids, _refit_group, _set_xfrm, _slide, append_shape, frag
from .pptx import (NS_R, NS_XMLNS, RT, SHAPE_TAGS, PptxError, all_ids, cnvpr, find_shape, kid, kids, placeholder,
                   read_xfrm, to_slide, xfrm_el)

# 中身ごと写せる関係（画像・動画・音声）
COPYABLE = {RT + "image", RT + "video", RT + "audio", RT + "media", "http://schemas.microsoft.com/office/2007/relationships/media",
            "http://schemas.microsoft.com/office/2007/relationships/hdphoto"}
LINKS = {RT + "hyperlink", RT + "slide"}


def renumber(el, next_id: int) -> int:
    """写した図形（グループなら中の図形も）に新しい id を振る。名前は PowerPoint と同じく元のまま。"""
    for c in el.getElementsByTagNameNS("*", "cNvPr"):
        c.setAttribute("id", str(next_id))
        next_id += 1
    return next_id


def _r_attrs(el):
    """要素（とその中）の r:embed・r:id などの属性。"""
    out = []
    stack = [el]
    while stack:
        n = stack.pop()
        if n.nodeType != n.ELEMENT_NODE:
            continue
        for i in range(n.attributes.length):
            a = n.attributes.item(i)
            if a.namespaceURI == NS_R:
                out.append((n, a))
        stack.extend(n.childNodes)
    return out


def copy_shapes(pkg, slide: int, ids: list[int]) -> dict:
    """図形を覚える（画面の Ctrl+C）。戻り値は Deck が持っておき、貼り付けに渡す。"""
    part = pkg.slide_part(slide)
    doc = pkg.xml(part)
    shapes = []
    for i in ids:
        sh, chain = find_shape(doc, i)
        order = list(sh.parentNode.childNodes).index(sh)
        shapes.append((len(chain), order, sh, chain))
    shapes.sort(key=lambda t: (t[0], t[1]))             # 重なりの順番を保つ
    rels = {r.id: r for r in pkg.rels(part)}
    rinfo, xmls, frames = {}, [], []
    for _, _, sh, chain in shapes:
        for n, a in _r_attrs(sh):
            r = rels.get(a.value)
            if r is None or a.value in rinfo:
                continue
            if r.external:
                rinfo[a.value] = {"type": r.type, "target": r.target, "external": True}
            elif r.type in COPYABLE:
                rinfo[a.value] = {"type": r.type, "target": r.target, "data": pkg.parts.get(r.target, b""),
                                  "ctype": pkg.content_type(r.target)}
            elif r.type in LINKS:
                rinfo[a.value] = {"type": r.type, "target": r.target}
            else:
                raise PptxError("グラフ・SmartArt・埋め込みオブジェクトはコピーできない（同じスライドなら複製は使える）")
        # グループの中の図形は、スライドの座標にしてから写す
        c = sh.cloneNode(True)
        if chain:
            f = read_xfrm(xfrm_el(sh))
            if f:
                g = to_slide([read_xfrm(xfrm_el(x)) for x in chain], f)
                _set_xfrm(c, g, keep_size=False)
        xmls.append(c.toxml())
        frames.append(read_xfrm(xfrm_el(c)))
    root = doc.documentElement
    ns = {}
    for k in range(root.attributes.length):
        a = root.attributes.item(k)
        if a.name.startswith("xmlns:"):
            ns[a.name[6:]] = a.value
    return {"xml": xmls, "rels": rinfo, "ns": ns, "src": pkg.path.name, "src_part": part, "pastes": 0}


def paste_shapes(pkg, slide: int, clip: dict) -> list[int]:
    """覚えた図形を貼り付ける（画面の Ctrl+V）。同じスライドなら少しずらす。新しい id の列を返す。"""
    part = pkg.slide_part(slide)
    doc = pkg.xml(part)
    root = doc.documentElement
    for pre, uri in clip["ns"].items():     # 貼り付け先に無い接頭辞（a14 など）を宣言する
        cur = root.getAttribute("xmlns:" + pre)
        if not cur:
            root.setAttributeNS(NS_XMLNS, "xmlns:" + pre, uri)
        elif cur != uri:
            raise PptxError(f"名前空間の接頭辞 {pre} がこの資料では別の意味で使われているので、貼り付けられない")
    remap = {}
    same_pkg = clip["src"] == pkg.path.name
    for rid, info in clip["rels"].items():
        if info.get("external"):
            remap[rid] = pkg.add_rel(part, info["type"], info["target"], external=True)
        elif "data" in info:
            target = info["target"]
            if not (same_pkg and pkg.parts.get(target) == info["data"]):
                ext = target.rsplit(".", 1)[-1]
                target = pkg.new_part_name("ppt/media/pc_media{}." + ext)
                pkg.put(target, info["data"])
                if info.get("ctype"):
                    pkg.ensure_default(ext, info["ctype"])
            remap[rid] = pkg.add_rel(part, info["type"], target)
        elif same_pkg and pkg.has(info["target"]):
            remap[rid] = pkg.add_rel(part, info["type"], info["target"])
    decl = " ".join(f'xmlns:{p}="{u}"' for p, u in clip["ns"].items())
    same_slide = clip["src_part"] == part and same_pkg
    clip["pastes"] = clip.get("pastes", 0) + (1 if same_slide else 0)
    d = 180000 * clip["pastes"] if same_slide else 0
    nid = max(all_ids(doc) | {1}) + 1
    new = []
    for xml in clip["xml"]:
        tmp = minidom.parseString(f"<wrap {decl}>{xml}</wrap>")
        el = doc.importNode(tmp.documentElement.firstChild, True)
        for n, a in _r_attrs(el):
            if a.value in remap:
                a.value = remap[a.value]
            elif n.localName in ("hlinkClick", "hlinkMouseOver"):
                n.parentNode.removeChild(n)       # 貼り付け先に無いスライドへのリンクは外す
        new.append(nid)
        nid = renumber(el, nid)
        x = xfrm_el(el)
        if d and x is not None and kid(x, "off") is not None:
            off = kid(x, "off")
            off.setAttribute("x", str(int(off.getAttribute("x")) + d))
            off.setAttribute("y", str(int(off.getAttribute("y")) + d))
        append_shape(doc, el)
    pkg.touch(part)
    return new


# ---------------------------------------------------------------- グループ
def op_group(pkg, args):
    """選んだ図形をグループにする（Ctrl+G）。アニメーションの付いた図形は、PowerPoint と同じく効果を外す。"""
    part, doc = _slide(pkg, args)
    ids = _ids(args)
    if len(ids) < 2:
        raise PptxError("グループにするには図形を2つ以上選ぶ")
    items = [find_shape(doc, i) for i in ids]
    parent = items[0][0].parentNode
    if any(sh.parentNode is not parent for sh, _ in items):
        raise PptxError("グループの中と外の図形は、まとめてグループにできない")
    if any(placeholder(sh) for sh, _ in items):
        raise PptxError("プレースホルダーはグループにできない（PowerPoint と同じ）")
    frames = [read_xfrm(xfrm_el(sh)) for sh, _ in items]
    if any(f is None for f in frames):
        raise PptxError("位置を持たない図形はグループにできない")
    order = [n for n in kids(parent) if n.localName in SHAPE_TAGS]
    items.sort(key=lambda t: order.index(t[0]))
    x0 = min(f["x"] for f in frames)
    y0 = min(f["y"] for f in frames)
    x1 = max(f["x"] + f["w"] for f in frames)
    y1 = max(f["y"] + f["h"] for f in frames)
    nid = max(all_ids(doc) | {1}) + 1
    g = frag(doc, f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{nid}" name="グループ化 {nid}"/><p:cNvGrpSpPr/><p:nvPr/>'
                  f'</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="{x0}" y="{y0}"/><a:ext cx="{x1 - x0}" cy="{y1 - y0}"/>'
                  f'<a:chOff x="{x0}" y="{y0}"/><a:chExt cx="{x1 - x0}" cy="{y1 - y0}"/></a:xfrm></p:grpSpPr>'
                  f'</p:grpSp>')[0]
    parent.insertBefore(g, items[-1][0].nextSibling)
    for sh, _ in items:
        g.appendChild(sh)
    had = [e for e in anim.effects(doc) if e["spid"] in ids]
    anim.delete_for_shapes(doc, set(ids))
    for c in reversed(items[0][1]):
        _refit_group(c)
    pkg.touch(part)
    return {"label": "グループ化" + ("（中の図形のアニメーションは外した）" if had else ""), "select": [nid]}


def op_ungroup(pkg, args):
    """グループを解除する（Ctrl+Shift+G）。中の図形の位置・大きさ・回転をスライドの座標に直す。"""
    import math
    part, doc = _slide(pkg, args)
    selected = []
    for i in _ids(args):
        g, chain = find_shape(doc, i)
        if g.localName != "grpSp":
            raise PptxError("グループではない")
        gx = read_xfrm(xfrm_el(g))
        if gx is None:
            raise PptxError("位置を持たないグループは解除できない")
        if gx["flipH"] or gx["flipV"]:
            raise PptxError("反転したグループは解除できない（PowerPoint で解除する）")
        rot = gx.get("rot", 0)
        gcx, gcy = gx["x"] + gx["w"] / 2, gx["y"] + gx["h"] / 2
        parent = g.parentNode
        for c in [k for k in kids(g) if k.localName in SHAPE_TAGS]:
            cx = read_xfrm(xfrm_el(c))
            if cx:
                f = to_slide([gx], cx)
                if rot:
                    a = math.radians(rot)
                    mx, my = f["x"] + f["w"] / 2 - gcx, f["y"] + f["h"] / 2 - gcy
                    nx, ny = gcx + mx * math.cos(a) - my * math.sin(a), gcy + mx * math.sin(a) + my * math.cos(a)
                    f = {**f, "x": round(nx - f["w"] / 2), "y": round(ny - f["h"] / 2)}
                _set_xfrm(c, f)
                if rot:
                    x = xfrm_el(c)
                    x.setAttribute("rot", str(round(((cx.get("rot", 0) + rot) % 360) * 60000)))
            parent.insertBefore(c, g)
            selected.append(int(cnvpr(c).getAttribute("id")))
        parent.removeChild(g)
        anim.delete_for_shapes(doc, {i})
        for x in reversed(chain):
            _refit_group(x)
    pkg.touch(part)
    return {"label": "グループ解除", "select": selected}


def op_paste(pkg, args):
    """Deck が覚えているコピーを貼る（args["clip"] は Deck が入れる）。"""
    clip = args.get("clip")
    if not clip:
        raise PptxError("コピーした図形が無い")
    new = paste_shapes(pkg, int(args["slide"]), clip)
    return {"label": "貼り付け", "select": new}


OPS = {"group": op_group, "ungroup": op_ungroup, "paste": op_paste}
