"""描画用の一時ファイルと、描いた結果の置き場（キャッシュ）。

スライドは1枚ずつ、「そのスライドだけを含む PPTX」を作って LibreOffice で描く。
  - 変わったスライドだけを描き直せる（全体を描くと 30 枚で 10 秒以上かかる）
  - 描く用の写しでは、図形の説明（descr）に "pc:<id>" を書き込む。LibreOffice の図形と
    PPTX の図形（cNvPr の id）を、これで1対1に対応づける
  - 非表示のスライドも描く（show="0" を外す）
  - スライド番号が正しく出るよう、開始番号（firstSlideNum）をそのスライドの番号にする

結果はスライドの指紋（そのスライドの描画に効く部品の中身から作る）ごとに保存する。
同じ中身なら、別のファイルでも、閉じて開き直しても、描き直さない。
"""
from __future__ import annotations

import hashlib
import re
import shutil
import subprocess
import time
import zipfile
import zlib
from pathlib import Path

from .office import CACHE
from .pptx import NS_R, RT_LAYOUT, RT_NOTES, RT_SLIDE, Package, rels_name

RENDER_VERSION = b"pc-render-4"
RENDERS = CACHE / "renders"

_CNVPR = re.compile(rb"""<((?:\w+:)?cNvPr)\b((?:[^>"']|"[^"]*"|'[^']*')*?)(/?)>""")
_DESCR = re.compile(rb"""\s+descr\s*=\s*(?:"[^"]*"|'[^']*')""")
_ID = re.compile(rb"""\sid\s*=\s*["'](\d+)["']""")
_SLD_OPEN = re.compile(rb"""<((?:\w+:)?sld)\b((?:[^>"']|"[^"]*"|'[^']*')*)>""")
_SHOW = re.compile(rb"""\s+show\s*=\s*["'](?:0|false)["']""")
_SLDIDLST = re.compile(rb"<(\w+:)?sldIdLst\b.*?</(\w+:)?sldIdLst>|<(\w+:)?sldIdLst\s*/>", re.S)
_SECTIONS = re.compile(rb"<(\w+:)?sectionLst\b.*?</(\w+:)?sectionLst>", re.S)


def tag_shapes(slide_xml: bytes) -> bytes:
    """全部の図形の説明を "pc:<id>" にする（描く用の写しだけ）。"""
    def f(m):
        attrs = _DESCR.sub(b"", m.group(2))
        i = _ID.search(attrs)
        if not i:
            return m.group(0)
        return b"<" + m.group(1) + b' descr="pc:' + i.group(1) + b'"' + attrs + m.group(3) + b">"
    return _CNVPR.sub(f, slide_xml)


def unhide(slide_xml: bytes) -> bytes:
    m = _SLD_OPEN.search(slide_xml)
    if not m:
        return slide_xml
    return slide_xml[:m.start()] + _SHOW.sub(b"", m.group(0)) + slide_xml[m.end():]


def _crc(pkg: Package, part: str) -> int:
    info = pkg.infos.get(part)
    if info is not None and part not in pkg.dirty:
        return info.CRC
    return zlib.crc32(pkg.parts.get(part, b""))


def deck_base(pkg: Package) -> bytes:
    """スライドの並びに関係しない、文書全体の設定（既定の文字の書式など）の指紋。

    XML の書き方（空要素の書き方・改行・引用符）の違いで変わらないよう、正規形（C14N）にしてから取る。
    """
    from xml.etree.ElementTree import canonicalize
    try:
        c = canonicalize(pkg.parts[pkg.main].decode("utf-8"), strip_text=True).encode()
    except Exception:  # noqa: BLE001  壊れた XML でも指紋は作る
        c = pkg.parts[pkg.main]
    pres = _SECTIONS.sub(b"", _SLDIDLST.sub(b"", c))
    return hashlib.sha1(pres).digest()


def slide_fp(pkg: Package, index: int, base: bytes | None = None) -> str:
    part = pkg.slide_part(index)
    h = hashlib.sha1(RENDER_VERSION)
    h.update(base or deck_base(pkg))
    if needs_number(pkg, index):      # 番号の欄が無ければ、並びが変わっても見た目は同じ
        h.update(str(pkg.first_number() + index).encode())
    h.update(repr(pkg.slide_size()).encode())
    for p in sorted(pkg.reachable([part], skip_types={RT_SLIDE, RT_NOTES})):
        h.update(p.encode())
        h.update(_crc(pkg, p).to_bytes(4, "little"))
    return h.hexdigest()[:24]


def deck_fp(pkg: Package) -> str:
    h = hashlib.sha1(RENDER_VERSION)
    for p in sorted(pkg.parts):
        h.update(p.encode())
        h.update(_crc(pkg, p).to_bytes(4, "little"))
    return h.hexdigest()[:24]


_BLANK = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld xmlns:a="http://schemas.openxmlformats.org/'
          b'drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          b'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr>'
          b'<p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>'
          b'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def needs_number(pkg: Package, index: int) -> bool:
    """スライドにスライド番号の欄があるか（あれば、番号が合うように前に空のスライドを並べて描く）。"""
    return b'"slidenum"' in pkg.parts[pkg.slide_part(index)]


def mini_deck(pkg: Package, index: int, dest: Path, tag: bool = True) -> int:
    """そのスライドだけを含む PPTX を書く（圧縮しない。速さのため）。描く対象のページ番号（0 始まり）を返す。

    LibreOffice は開始番号（firstSlideNum）を読まないので、スライド番号の欄があるスライドは、
    前に空のスライドを index 枚並べて番号を合わせる。
    """
    from .pptx import dump, kid, kids, parse
    target = pkg.slide_part(index)
    others = set(pkg.slide_parts()) - {target}
    pad = index + pkg.first_number() - 1 if needs_number(pkg, index) else 0
    pres = parse(pkg.parts[pkg.main])
    root = pres.documentElement
    lst = kid(root, "sldIdLst")
    keep_rid = None
    for i, s in enumerate(kids(lst, "sldId") if lst is not None else []):
        if i == index:
            keep_rid = s.getAttributeNS(NS_R, "id")
            keep_el = s
        else:
            lst.removeChild(s)
    for sec in list(root.getElementsByTagNameNS("*", "sectionLst")):   # セクションは枚数と合わなくなるので外す
        ext = sec.parentNode
        if ext is not None and ext.localName == "ext" and ext.parentNode is not None:
            ext.parentNode.removeChild(ext)
    rels = parse(pkg.parts[rels_name(pkg.main)])
    for r in list(rels.documentElement.getElementsByTagNameNS("*", "Relationship")):
        if r.getAttribute("Type") == RT_SLIDE and r.getAttribute("Id") != keep_rid:
            r.parentNode.removeChild(r)
    extra: dict[str, bytes] = {}
    if pad:
        layout = pkg.layout_of(target)
        rp = next((a.name[6:] for a in (root.attributes.item(i) for i in range(root.attributes.length))
                   if a.name.startswith("xmlns:") and a.value == NS_R), "r")
        for k in range(pad):
            name = f"ppt/slides/pc_pad{k + 1}.xml"
            rid = f"rIdPcPad{k + 1}"
            extra[name] = _BLANK
            extra[rels_name(name)] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.'
                'openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + RT_LAYOUT +
                '" Target="../slideLayouts/' + layout.rsplit("/", 1)[1] + '"/></Relationships>').encode()
            el = rels.createElementNS(rels.documentElement.namespaceURI, "Relationship")
            el.setAttribute("Id", rid)
            el.setAttribute("Type", RT_SLIDE)
            el.setAttribute("Target", "slides/" + name.rsplit("/", 1)[1])
            rels.documentElement.appendChild(el)
            sid = pres.createElementNS(keep_el.namespaceURI, keep_el.tagName)
            sid.setAttribute("id", str(2147483000 - pad + k))
            sid.setAttributeNS(NS_R, f"{rp}:id", rid)
            lst.insertBefore(sid, keep_el)
    keep = pkg.reachable(skip_parts=others)
    keep |= {"[Content_Types].xml", "_rels/.rels", pkg.main, rels_name(pkg.main)}
    ct = parse(pkg.parts["[Content_Types].xml"])
    for o in list(ct.documentElement.getElementsByTagNameNS("*", "Override")):
        if o.getAttribute("PartName").lstrip("/") not in keep:
            o.parentNode.removeChild(o)
    for name in extra:
        if name.endswith(".xml"):
            o = ct.createElementNS(ct.documentElement.namespaceURI, "Override")
            o.setAttribute("PartName", "/" + name)
            o.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
            ct.documentElement.appendChild(o)
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_STORED) as z:
        z.writestr("[Content_Types].xml", dump(ct, pkg.parts["[Content_Types].xml"]))
        for name in pkg.order:
            if name not in keep or name == "[Content_Types].xml" or name not in pkg.parts:
                continue
            data = pkg.parts[name]
            if name == pkg.main:
                data = dump(pres, data)
            elif name == rels_name(pkg.main):
                data = dump(rels, data)
            elif name == target:
                data = unhide(tag_shapes(data) if tag else data)
            z.writestr(name, data)
        for name, data in extra.items():
            z.writestr(name, data)
    return pad


def full_copy(pkg: Package, dest: Path, show_hidden: bool = False, tag: bool = False) -> None:
    """文書全体の写し（スライドショー・check・まとめて描く用）。
    show_hidden なら非表示のスライドも出す。tag なら図形の説明に "pc:<id>" を書き込む。"""
    slides = set(pkg.slide_parts())
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_STORED) as z:
        for name in pkg.order:
            if name not in pkg.parts:
                continue
            data = pkg.parts[name]
            if name in slides:
                if tag:
                    data = tag_shapes(data)
                if show_hidden:
                    data = unhide(data)
            z.writestr(name, data)


def render_dir(fp: str) -> Path:
    return RENDERS / fp


def is_rendered(fp: str) -> bool:
    return (render_dir(fp) / "meta.json").is_file()


def prune(days: float = 14) -> None:
    """長く使っていない描画結果を消す。"""
    if not RENDERS.is_dir():
        return
    limit = time.time() - days * 86400
    for d in RENDERS.iterdir():
        try:
            if d.stat().st_mtime < limit:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass


# ---------------------------------------------------------------- フォント
_FONT_CACHE: dict[str, str] = {}


def font_match(name: str) -> str:
    """fontconfig がそのフォント名で実際に使うフォント（ファミリー名の一覧の先頭）。"""
    if name not in _FONT_CACHE:
        try:
            r = subprocess.run(["fc-match", "-f", "%{family}", name], capture_output=True, text=True, timeout=5)
            _FONT_CACHE[name] = r.stdout.strip()
        except (OSError, subprocess.TimeoutExpired):
            _FONT_CACHE[name] = ""
    return _FONT_CACHE[name]


def _norm(s: str) -> str:
    return re.sub(r"[\s_-]", "", s).lower()


def font_report(used: dict[str, int]) -> list[dict]:
    """置き換わるフォントの一覧。文字幅が同じ代替（Carlito など）は、改行位置が変わらないので印を付ける。"""
    metric_compatible = {"calibri": "carlito", "cambria": "caladea", "arial": "liberationsans",
                         "helvetica": "liberationsans", "timesnewroman": "liberationserif",
                         "couriernew": "liberationmono", "arialnarrow": "liberationsansnarrow"}
    out = []
    for name, n in sorted(used.items(), key=lambda kv: -kv[1]):
        fam = font_match(name)
        fams = [_norm(f) for f in fam.split(",")]
        if _norm(name) in fams:
            continue
        sub = fam.split(",")[0] if fam else "?"
        out.append({"font": name, "uses": n, "replaced_by": sub,
                    "same_width": metric_compatible.get(_norm(name)) in fams})
    return out
