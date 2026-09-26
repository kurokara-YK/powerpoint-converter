"""スライドの操作（新規・複製・削除・並べ替え・非表示）、ノート、画面切り替え、背景。"""
from __future__ import annotations

from .oputil import TEXT_PH, _esc, _slide, frag
from .pptx import (CT_NOTES, CT_SLIDE, NS_A, NS_P, NS_R, P, RT_LAYOUT, RT_NOTES, RT_NOTES_MASTER, RT_SLIDE,
                   PptxError, cnvpr, edit_body, kid, kids, notes_body_shape, path, placeholder, prefix_for,
                   text_body, top_shapes)

# ---------------------------------------------------------------- ノート
_NOTES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    f'<p:notes xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr>'
    '<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
    '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="スライド イメージ プレースホルダー 1"/><p:cNvSpPr><a:spLocks noGrp="1" '
    'noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="ノート プレースホルダー 2"/><p:cNvSpPr><a:spLocks noGrp="1"/>'
    '</p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>'
    '<a:lstStyle/><a:p><a:endParaRPr lang="ja-JP" altLang="en-US"/></a:p></p:txBody></p:sp>'
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>')


def op_notes(pkg, args):
    part, _ = _slide(pkg, args)
    notes = pkg.notes_of(part)
    if not notes:
        master = pkg.rel_of_type(pkg.main, RT_NOTES_MASTER)
        if not master:
            raise PptxError("この文書にはノートの型（ノート マスター）が無いので、ノートを作れない")
        notes = pkg.new_part_name("ppt/notesSlides/notesSlide{}.xml")
        pkg.put(notes, _NOTES_XML.encode())
        pkg.ensure_override(notes, CT_NOTES)
        pkg.add_rel(notes, RT_NOTES_MASTER, master)
        pkg.add_rel(notes, RT_SLIDE, part)
        pkg.add_rel(part, RT_NOTES, notes)
    ndoc = pkg.xml(notes)
    sh = notes_body_shape(ndoc)
    if sh is None:
        raise PptxError("ノートの本文の枠が無い")
    edit_body(text_body(sh), [str(t) for t in args["text"]])
    pkg.touch(notes)
    return {"label": "ノートの編集"}


# ---------------------------------------------------------------- スライドの操作
def _sections(pkg):
    """セクション（p14:sectionLst）の、各セクションのスライド id の並び（要素のまま）。"""
    out = []
    for s in pkg.pres.documentElement.getElementsByTagNameNS("*", "section"):
        lst = kid(s, "sldIdLst")
        if lst is not None:
            out.append(lst)
    return out


def _section_insert_after(pkg, after_id: str, new_id: str) -> None:
    for lst in _sections(pkg):
        for e in kids(lst, "sldId"):
            if e.getAttribute("id") == after_id:
                n = e.cloneNode(False)
                n.setAttribute("id", new_id)
                lst.insertBefore(n, e.nextSibling)
                return


def _section_remove(pkg, sid: str) -> None:
    for lst in _sections(pkg):
        for e in kids(lst, "sldId"):
            if e.getAttribute("id") == sid:
                lst.removeChild(e)


def _section_reorder(pkg) -> None:
    """スライドの並びを変えたあと、各セクションの枚数を保ったまま、新しい並びで詰め直す。"""
    secs = _sections(pkg)
    if not secs:
        return
    order = [s.getAttribute("id") for s in pkg.slide_ids()]
    k = 0
    for lst in secs:
        es = kids(lst, "sldId")
        for e in es:
            if k < len(order):
                e.setAttribute("id", order[k])
            k += 1


def _new_sld_id(pkg) -> str:
    ids = [int(s.getAttribute("id")) for s in pkg.slide_ids()]
    return str(max(ids + [255]) + 1)


def _add_slide_entry(pkg, new_part: str, after_index: int) -> int:
    """presentation.xml の sldIdLst に新しいスライドを after_index の後ろに足す。新しい番号（0 始まり）を返す。"""
    pres = pkg.pres
    rid = pkg.add_rel(pkg.main, RT_SLIDE, new_part)
    pkg.ensure_override(new_part, CT_SLIDE)
    sid = _new_sld_id(pkg)
    ids = pkg.slide_ids()
    lst = kid(pres.documentElement, "sldIdLst")
    if lst is None:        # sldIdLst は sldSz の前（sldMasterIdLst などの後）
        lst = P(pres, "sldIdLst")
        pres.documentElement.insertBefore(lst, kid(pres.documentElement, "sldSz"))
    el = P(pres, "sldId", {"id": sid})
    rp = prefix_for(pres, NS_R, "r")
    el.setAttributeNS(NS_R, f"{rp}:id", rid)
    if ids and 0 <= after_index < len(ids):
        lst.insertBefore(el, ids[after_index].nextSibling)
        _section_insert_after(pkg, ids[after_index].getAttribute("id"), sid)
    else:
        lst.appendChild(el)
        secs = _sections(pkg)
        if secs:
            n = P(pres, "sldId", {"id": sid})
            secs[-1].appendChild(n)
    pkg.touch(pkg.main)
    return [s.getAttribute("id") for s in pkg.slide_ids()].index(sid)


def op_slide_dup(pkg, args):
    i = int(args["index"])
    src = pkg.slide_part(i)
    pkg.flush()
    new = pkg.new_part_name("ppt/slides/slide{}.xml")
    pkg.put(new, pkg.parts[src])
    for r in pkg.rels(src):
        if r.type == RT_NOTES:
            continue
        if r.external:
            _add_external_rel(pkg, new, src, r.id)
        else:
            _copy_rel(pkg, new, r)
    notes = pkg.notes_of(src)
    if notes and pkg.has(notes):
        nn = pkg.new_part_name("ppt/notesSlides/notesSlide{}.xml")
        pkg.put(nn, pkg.parts[notes])
        pkg.ensure_override(nn, CT_NOTES)
        for r in pkg.rels(notes):
            if r.type == RT_SLIDE:
                _copy_rel(pkg, nn, r, target=new)
            elif not r.external:
                _copy_rel(pkg, nn, r)
        pkg.add_rel(new, RT_NOTES, nn)
    k = _add_slide_entry(pkg, new, i)
    return {"label": "スライドの複製", "slide": k}


def _copy_rel(pkg, part: str, r, target: str | None = None) -> None:
    """関係を同じ Id のまま写す（スライドの XML の r:embed などが Id で指しているため）。"""
    from .pptx import NS_PKG_REL, rels_name, relative
    doc = pkg.ensure_rels(part)
    el = doc.createElementNS(NS_PKG_REL, "Relationship")
    el.setAttribute("Id", r.id)
    el.setAttribute("Type", r.type)
    el.setAttribute("Target", relative(part, target or r.target))
    doc.documentElement.appendChild(el)
    pkg.touch(rels_name(part))


def _add_external_rel(pkg, part: str, src: str, rid: str) -> None:
    from .pptx import rels_name
    sdoc = pkg.xml(rels_name(src))
    for r in sdoc.documentElement.getElementsByTagNameNS("*", "Relationship"):
        if r.getAttribute("Id") == rid:
            d = pkg.ensure_rels(part)
            d.documentElement.appendChild(d.importNode(r, True))
            pkg.touch(rels_name(part))


def op_slide_delete(pkg, args):
    i = int(args["index"])
    ids = pkg.slide_ids()
    if len(ids) <= 1:
        raise PptxError("最後の1枚は消せない")
    part = pkg.slide_part(i)
    before = pkg.reachable([part], skip_types={RT_SLIDE, RT_LAYOUT})
    el = ids[i]
    sid = el.getAttribute("id")
    el.parentNode.removeChild(el)
    _section_remove(pkg, sid)
    pkg.drop_rel(pkg.main, el.getAttributeNS(NS_R, "id"))
    pkg.touch(pkg.main)
    after = pkg.reachable()
    for p in before:
        if p not in after:
            pkg.remove(p)
            pkg.drop_override(p)
    return {"label": "スライドの削除", "slide": min(i, len(ids) - 2)}


def op_slide_move(pkg, args):
    i, j = int(args["from"]), int(args["to"])
    ids = pkg.slide_ids()
    if not (0 <= i < len(ids) and 0 <= j < len(ids)) or i == j:
        return {"label": "スライドの移動", "slide": i}
    el = ids[i]
    lst = el.parentNode
    lst.removeChild(el)
    rest = [e for e in ids if e is not el]
    lst.insertBefore(el, rest[j] if j < len(rest) else None)
    _section_reorder(pkg)
    pkg.touch(pkg.main)
    return {"label": "スライドの移動", "slide": j}


def op_slide_hide(pkg, args):
    part = pkg.slide_part(int(args["index"]))
    root = pkg.xml(part).documentElement
    if args.get("hidden", True):
        root.setAttribute("show", "0")
    elif root.hasAttribute("show"):
        root.removeAttribute("show")
    pkg.touch(part)
    return {"label": "非表示スライド" if args.get("hidden", True) else "スライドの再表示"}


def op_slide_new(pkg, args):
    after = int(args.get("after", len(pkg.slide_ids()) - 1))
    layout = args.get("layout")
    if not layout:
        ids = pkg.slide_ids()
        layout = pkg.layout_of(pkg.slide_part(after)) if ids else pkg.layouts()[0]["part"]
    if not pkg.has(layout):
        raise PptxError(f"レイアウト {layout} が無い")
    shapes, nid = [], 2
    for sh in top_shapes(pkg.xml(layout)):
        ph = placeholder(sh)
        if not ph or ph["type"] in ("dt", "ftr", "sldNum", "hdr"):
            continue
        phel = kid(path(sh, "nvSpPr", "nvPr") or sh, "ph")
        attrs = "".join(f' {k}="{_esc(phel.getAttribute(k))}"' for k in ("type", "orient", "sz", "idx")
                        if phel is not None and phel.getAttribute(k))
        c = cnvpr(sh)
        name = _esc(c.getAttribute("name") if c is not None else f"プレースホルダー {nid}")
        body = ('<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP" altLang="en-US"/></a:p></p:txBody>'
                if ph["type"] in TEXT_PH else "")
        shapes.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{nid}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/>'
                      f'</p:cNvSpPr><p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr><p:spPr/>{body}</p:sp>')
        nid += 1
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           f'<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr>'
           '<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
           '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
           + "".join(shapes) +
           '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')
    new = pkg.new_part_name("ppt/slides/slide{}.xml")
    pkg.put(new, xml.encode())
    pkg.add_rel(new, RT_LAYOUT, layout)
    k = _add_slide_entry(pkg, new, after)
    return {"label": "新しいスライド", "slide": k}




# ---------------------------------------------------------------- 画面切り替え
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
# 画面の選択肢（LibreOffice のスライドショーでも動くもの）: 名前 → (要素, 日本語の名前)
TRANSITIONS = {"none": ("", "なし"), "fade": ('<p:fade/>', "フェード"), "push": ('<p:push dir="u"/>', "プッシュ"),
               "wipe": ('<p:wipe dir="r"/>', "ワイプ"), "cover": ('<p:cover dir="l"/>', "カバー"),
               "pull": ('<p:pull dir="l"/>', "アンカバー"), "split": ('<p:split orient="vert" dir="out"/>', "スプリット"),
               "randomBar": ('<p:randomBar dir="vert"/>', "ランダム ストライプ"), "circle": ('<p:circle/>', "図形（円）"),
               "diamond": ('<p:diamond/>', "図形（ひし形）"), "plus": ('<p:plus/>', "図形（プラス）"),
               "wedge": ('<p:wedge/>', "くさび形"), "wheel": ('<p:wheel spokes="1"/>', "ホイール"),
               "blinds": ('<p:blinds dir="vert"/>', "ブラインド"), "checker": ('<p:checker dir="horz"/>', "チェッカーボード"),
               "dissolve": ('<p:dissolve/>', "ディゾルブ"), "zoom": ('<p:zoom/>', "ズーム"), "cut": ('<p:cut/>', "カット")}
SPEEDS = {"slow": "遅く", "med": "普通", "fast": "速く"}


def _transition_elements(root) -> list:
    """スライドの画面切り替えの要素（p:transition と、それを包む mc:AlternateContent）。"""
    out = []
    for n in kids(root):
        if n.localName == "transition":
            out.append(n)
        elif n.localName == "AlternateContent" and n.getElementsByTagNameNS(NS_P, "transition"):
            out.append(n)
    return out


def transition_info(doc) -> dict:
    """画面に出す画面切り替え: {effect, label, speed, advance_after（秒）, advance_click}"""
    root = doc.documentElement
    for n in _transition_elements(root):
        t = n if n.localName == "transition" else n.getElementsByTagNameNS(NS_P, "transition")[-1]
        effect = next((k.localName for k in kids(t) if k.localName not in ("sndAc", "extLst")), "none")
        adv = t.getAttribute("advTm")
        return {"effect": effect, "label": TRANSITIONS.get(effect, ("", effect))[1],
                "speed": t.getAttribute("spd") or "fast", "advance_after": int(adv) / 1000 if adv else None,
                "advance_click": t.getAttribute("advClick") not in ("0", "false")}
    return {"effect": "none", "label": "なし", "speed": "fast", "advance_after": None, "advance_click": True}


def set_transition(doc, effect: str, speed: str = "med", advance_after: float | None = None,
                   advance_click: bool = True) -> None:
    if effect not in TRANSITIONS:
        raise PptxError(f"画面切り替え {effect} は使えない")
    root = doc.documentElement
    for n in _transition_elements(root):
        root.removeChild(n)
    if effect == "none" and advance_after is None and advance_click:
        return
    attrs = f' spd="{speed if speed in SPEEDS else "med"}"'
    if advance_after is not None:
        attrs += f' advTm="{round(float(advance_after) * 1000)}"'
    if not advance_click:
        attrs += ' advClick="0"'
    el = frag(doc, f"<p:transition{attrs}>{TRANSITIONS[effect][0]}</p:transition>")[0]
    # p:sld の子の順番: cSld, clrMapOvr, transition, timing, extLst
    ref = kid(root, "timing") or kid(root, "extLst")
    root.insertBefore(el, ref) if ref is not None else root.appendChild(el)


def op_transition(pkg, args):
    """画面切り替えを設定する。slides: 対象（0 始まり）の列。無ければ slide の1枚。all: true で全部。"""
    n = len(pkg.slide_parts())
    targets = range(n) if args.get("all") else [int(i) for i in args.get("slides", [args.get("slide", 0)])]
    for i in targets:
        part = pkg.slide_part(i)
        set_transition(pkg.xml(part), args.get("effect", "fade"), args.get("speed", "med"),
                       args.get("advance_after"), args.get("advance_click", True))
        pkg.touch(part)
    return {"label": "画面切り替え" + ("（すべてのスライド）" if args.get("all") else "")}


# ---------------------------------------------------------------- 背景
def background_info(doc) -> str | None:
    bg = path(doc.documentElement, "cSld", "bg")
    if bg is None:
        return None
    c = path(bg, "bgPr", "solidFill", "srgbClr")
    return "#" + c.getAttribute("val") if c is not None else "custom"


def op_background(pkg, args):
    """背景を単色にする（color: "#RRGGBB"）。color が無ければ、スライド マスターの背景に戻す。"""
    n = len(pkg.slide_parts())
    targets = range(n) if args.get("all") else [int(args.get("slide", 0))]
    color = args.get("color")
    for i in targets:
        part = pkg.slide_part(i)
        doc = pkg.xml(part)
        csld = kid(doc.documentElement, "cSld")
        for b in kids(csld, "bg"):
            csld.removeChild(b)
        if color:
            el = frag(doc, f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{color.lstrip("#").upper()}"/></a:solidFill>'
                           f'<a:effectLst/></p:bgPr></p:bg>')[0]
            csld.insertBefore(el, csld.firstChild)
        pkg.touch(part)
    return {"label": "背景" + ("（すべてのスライド）" if args.get("all") else "")}


OPS = {"notes": op_notes, "slide_dup": op_slide_dup, "slide_delete": op_slide_delete, "slide_move": op_slide_move,
       "slide_hide": op_slide_hide, "slide_new": op_slide_new, "transition": op_transition,
       "background": op_background}
