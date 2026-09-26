"""PPTX（zip の中の XML）の読み書き。

書き換えるのは、操作した部品（スライドの XML など）だけ。ほかの部品はバイト列のまま書き戻す。
LibreOffice で保存し直すと PowerPoint で開いたときに崩れることがあるので、保存には使わない。

XML は minidom で扱う。ElementTree は使われていない名前空間の宣言を落とし、
mc:Ignorable が指す接頭辞が消えて PowerPoint が「修復」を求めるため。
"""
from __future__ import annotations

import os
import posixpath
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.dom import minidom

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XMLNS = "http://www.w3.org/2000/xmlns/"

RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
RT_SLIDE = RT + "slide"
RT_LAYOUT = RT + "slideLayout"
RT_MASTER = RT + "slideMaster"
RT_NOTES = RT + "notesSlide"
RT_NOTES_MASTER = RT + "notesMaster"
RT_IMAGE = RT + "image"
RT_THEME = RT + "theme"
RT_OFFICE_DOC = RT + "officeDocument"

CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
CT_NOTES = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
CT_REL = "application/vnd.openxmlformats-package.relationships+xml"

EMU_PER_CM = 360000
EMU_PER_PT = 12700
EDITABLE_SUFFIXES = {".pptx", ".pptm", ".potx"}
VIEW_SUFFIXES = {".ppt", ".odp", ".pps", ".ppsx", ".key", ".fodp"}

# 画面で「その他のテキスト」の意味で扱うプレースホルダーの種類
_TITLE_PH = {"title", "ctrTitle"}


class PptxError(Exception):
    pass


# ---------------------------------------------------------------- minidom の補助
def kids(el, local: str | None = None, ns: str | None = None) -> list:
    """直下の要素。local を指定すると、その名前（接頭辞を除いた名前）のものだけ。"""
    out = []
    for n in el.childNodes:
        if n.nodeType != n.ELEMENT_NODE:
            continue
        if local and n.localName != local:
            continue
        if ns and n.namespaceURI != ns:
            continue
        out.append(n)
    return out


def kid(el, local: str, ns: str | None = None):
    for n in kids(el, local, ns):
        return n
    return None


def path(el, *locals_: str):
    """kid を順にたどる。途中で無ければ None。"""
    for l in locals_:
        if el is None:
            return None
        el = kid(el, l)
    return el


def descendants(el, local: str, ns: str | None = None) -> list:
    out = []
    for n in el.getElementsByTagNameNS(ns or "*", local):
        out.append(n)
    return out


def prefix_for(doc, ns: str, default: str) -> str:
    """その文書で ns に使われている接頭辞。無ければ default を根に宣言する。"""
    root = doc.documentElement
    for i in range(root.attributes.length):
        a = root.attributes.item(i)
        if a.name.startswith("xmlns:") and a.value == ns:
            return a.name[6:]
    if root.getAttribute("xmlns") == ns:
        return ""
    root.setAttributeNS(NS_XMLNS, "xmlns:" + default, ns)
    return default


def make(doc, ns: str, local: str, default_prefix: str, attrs: dict | None = None):
    pre = prefix_for(doc, ns, default_prefix)
    el = doc.createElementNS(ns, f"{pre}:{local}" if pre else local)
    for k, v in (attrs or {}).items():
        el.setAttribute(k, str(v))
    return el


def P(doc, local, attrs=None):
    return make(doc, NS_P, local, "p", attrs)


def A(doc, local, attrs=None):
    return make(doc, NS_A, local, "a", attrs)


def insert_ordered(parent, new, order: list[str]):
    """スキーマで順番が決まっている子を、正しい位置に入れる（同じ名前のものがあれば置き換える）。"""
    name = new.localName
    for old in kids(parent, name):
        parent.replaceChild(new, old)
        return new
    rank = order.index(name) if name in order else len(order)
    for n in kids(parent):
        r = order.index(n.localName) if n.localName in order else len(order)
        if r > rank:
            parent.insertBefore(new, n)
            return new
    parent.appendChild(new)
    return new


def text_of(el) -> str:
    return "".join(n.data for n in el.childNodes if n.nodeType == n.TEXT_NODE)


def set_text(el, s: str):
    for n in list(el.childNodes):
        el.removeChild(n)
    el.appendChild(el.ownerDocument.createTextNode(s))


_DECL = re.compile(rb"^\s*<\?xml[^>]*\?>")


def parse(data: bytes):
    return minidom.parseString(data)


def dump(doc, original: bytes | None = None) -> bytes:
    """XML 宣言は元のもの（standalone="yes" など）を保つ。"""
    body = doc.documentElement.toxml().encode("utf-8")
    m = _DECL.match(original or b"")
    decl = m.group(0).strip() if m else b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    return decl + b"\r\n" + body


# ---------------------------------------------------------------- 部品と関係
@dataclass
class Rel:
    id: str
    type: str
    target: str          # 解決した部品名（外部リンクなら元の文字列）
    external: bool = False


def rels_name(part: str) -> str:
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def resolve(part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target)).lstrip("/")


def relative(src_part: str, dst_part: str) -> str:
    return posixpath.relpath(dst_part, posixpath.dirname(src_part) or ".")


class Package:
    """PPTX を部品（名前→バイト列）として持つ。保存すると、変えていない部品はそのまま書き戻す。"""

    def __init__(self, path: Path):
        self.path = Path(path)
        try:
            z = zipfile.ZipFile(self.path)
        except zipfile.BadZipFile as e:
            raise PptxError(f"{self.path.name} は PPTX（zip）として開けない: {e}") from e
        with z:
            self.infos = {i.filename: i for i in z.infolist() if not i.is_dir()}
            self.parts = {n: z.read(n) for n in self.infos}
        self.order = list(self.infos)
        self.dirty: set[str] = set()
        self._doms: dict[str, object] = {}
        if "ppt/presentation.xml" not in self.parts:
            main = self._main_part()
            if not main:
                raise PptxError(f"{self.path.name} に presentation.xml が無い（PowerPoint の文書ではない）")
        self.main = self._main_part() or "ppt/presentation.xml"

    def _main_part(self) -> str | None:
        if "_rels/.rels" not in self.parts:
            return None
        for r in self.rels(""):
            if r.type == RT_OFFICE_DOC:
                return r.target
        return None

    # ---- 読み書き ----
    def has(self, part: str) -> bool:
        return part in self.parts

    def xml(self, part: str):
        """部品の DOM。同じ部品は同じ DOM を返す（書き換えたら touch する）。"""
        if part not in self._doms:
            if part not in self.parts:
                raise PptxError(f"{part} が無い")
            self._doms[part] = parse(self.parts[part])
        return self._doms[part]

    def touch(self, part: str):
        self.dirty.add(part)

    def put(self, part: str, data: bytes):
        if part not in self.parts:
            self.order.append(part)
        self.parts[part] = data
        self._doms.pop(part, None)
        self.dirty.add(part)

    def remove(self, part: str):
        if part in self.parts:
            del self.parts[part]
            self.order.remove(part)
            self._doms.pop(part, None)
            self.dirty.add(part)

    def flush(self) -> dict[str, bytes | None]:
        """DOM の変更をバイト列に戻す。戻り値は変わった部品（消えたものは None）。"""
        for part in list(self.dirty):
            if part in self._doms and part in self.parts:
                self.parts[part] = dump(self._doms[part], self.parts[part])
        out = {p: self.parts.get(p) for p in self.dirty}
        return out

    def save(self, dest: Path | None = None) -> None:
        """同じフォルダの一時ファイルに書いてから置き換える（書きかけを読まれないように）。"""
        self.flush()
        dest = Path(dest or self.path)
        fd, tmp = tempfile.mkstemp(prefix="." + dest.name + ".", suffix=".tmp", dir=dest.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w") as z:
                order = ["[Content_Types].xml"] + [p for p in self.order if p != "[Content_Types].xml"]
                for name in order:
                    if name not in self.parts:
                        continue
                    info = self.infos.get(name)
                    if info is None:
                        info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                        info.compress_type = zipfile.ZIP_DEFLATED
                    z.writestr(info, self.parts[name])
            if dest.exists():
                os.chmod(tmp, dest.stat().st_mode & 0o7777)
            os.replace(tmp, dest)
        except BaseException:
            if os.path.exists(tmp):
                os.unlink(tmp)
            raise
        self.dirty.clear()

    # ---- 関係 ----
    def rels(self, part: str) -> list[Rel]:
        name = rels_name(part) if part else "_rels/.rels"
        if name not in self.parts:
            return []
        out = []
        for r in self.xml(name).documentElement.getElementsByTagNameNS("*", "Relationship"):
            ext = r.getAttribute("TargetMode") == "External"
            t = r.getAttribute("Target")
            out.append(Rel(r.getAttribute("Id"), r.getAttribute("Type"), t if ext else resolve(part or "/", t), ext))
        return out

    def rel_target(self, part: str, rid: str) -> str | None:
        for r in self.rels(part):
            if r.id == rid:
                return r.target
        return None

    def ensure_rels(self, part: str):
        """部品の関係ファイルの DOM（無ければ空のものを作る）。"""
        name = rels_name(part)
        if name not in self.parts:
            self.put(name, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                           f'<Relationships xmlns="{NS_PKG_REL}"></Relationships>'.encode())
            self.ensure_default("rels", CT_REL)
        return self.xml(name)

    def add_rel(self, part: str, type_: str, target_part: str, external: bool = False) -> str:
        """関係を足して、その Id を返す。external なら target_part は URL などをそのまま書く。"""
        name = rels_name(part)
        doc = self.ensure_rels(part)
        used = {r.getAttribute("Id") for r in doc.documentElement.getElementsByTagNameNS("*", "Relationship")}
        n = 1
        while f"rId{n}" in used:
            n += 1
        rid = f"rId{n}"
        el = doc.createElementNS(NS_PKG_REL, "Relationship")
        el.setAttribute("Id", rid)
        el.setAttribute("Type", type_)
        el.setAttribute("Target", target_part if external else relative(part, target_part))
        if external:
            el.setAttribute("TargetMode", "External")
        doc.documentElement.appendChild(el)
        self.touch(name)
        return rid

    def drop_rel(self, part: str, rid: str) -> None:
        name = rels_name(part)
        if name not in self.parts:
            return
        doc = self.xml(name)
        for r in doc.documentElement.getElementsByTagNameNS("*", "Relationship"):
            if r.getAttribute("Id") == rid:
                r.parentNode.removeChild(r)
                self.touch(name)

    # ---- 種類（[Content_Types].xml）----
    def ensure_override(self, part: str, ctype: str) -> None:
        doc = self.xml("[Content_Types].xml")
        for o in doc.documentElement.getElementsByTagNameNS("*", "Override"):
            if o.getAttribute("PartName") == "/" + part:
                return
        el = doc.createElementNS(NS_CT, "Override")
        el.setAttribute("PartName", "/" + part)
        el.setAttribute("ContentType", ctype)
        doc.documentElement.appendChild(el)
        self.touch("[Content_Types].xml")

    def drop_override(self, part: str) -> None:
        doc = self.xml("[Content_Types].xml")
        for o in doc.documentElement.getElementsByTagNameNS("*", "Override"):
            if o.getAttribute("PartName") == "/" + part:
                o.parentNode.removeChild(o)
                self.touch("[Content_Types].xml")

    def ensure_default(self, ext: str, ctype: str) -> None:
        doc = self.xml("[Content_Types].xml")
        for d in doc.documentElement.getElementsByTagNameNS("*", "Default"):
            if d.getAttribute("Extension").lower() == ext.lower():
                return
        el = doc.createElementNS(NS_CT, "Default")
        el.setAttribute("Extension", ext)
        el.setAttribute("ContentType", ctype)
        root = doc.documentElement
        first = kid(root, "Override")
        root.insertBefore(el, first) if first else root.appendChild(el)
        self.touch("[Content_Types].xml")

    def content_type(self, part: str) -> str | None:
        doc = self.xml("[Content_Types].xml")
        for o in doc.documentElement.getElementsByTagNameNS("*", "Override"):
            if o.getAttribute("PartName") == "/" + part:
                return o.getAttribute("ContentType")
        ext = part.rsplit(".", 1)[-1].lower()
        for d in doc.documentElement.getElementsByTagNameNS("*", "Default"):
            if d.getAttribute("Extension").lower() == ext:
                return d.getAttribute("ContentType")
        return None

    def new_part_name(self, pattern: str) -> str:
        """pattern は 'ppt/slides/slide{}.xml' の形。使われていない一番小さい番号を返す。"""
        n = 1
        while pattern.format(n) in self.parts:
            n += 1
        return pattern.format(n)

    def reachable(self, roots: list[str] | None = None, skip_types: set[str] = frozenset(),
                  skip_parts: set[str] = frozenset()) -> set[str]:
        """_rels/.rels からたどれる部品（関係ファイルも含む）。"""
        seen: set[str] = set()
        stack = list(roots) if roots is not None else [r.target for r in self.rels("") if not r.external]
        seen.add("_rels/.rels")
        while stack:
            p = stack.pop()
            if p in seen or p not in self.parts or p in skip_parts:
                continue
            seen.add(p)
            rn = rels_name(p)
            if rn in self.parts:
                seen.add(rn)
                for r in self.rels(p):
                    if not r.external and r.type not in skip_types:
                        stack.append(r.target)
        return seen

    def collect_garbage(self) -> list[str]:
        """どこからも参照されなくなった部品（消したスライドの画像など）を消す。"""
        keep = self.reachable()
        keep |= {"[Content_Types].xml", "_rels/.rels"}
        gone = []
        for p in list(self.parts):
            if p in keep or p.startswith("customXml/") or p.startswith("docProps/"):
                continue
            if p.endswith(".rels"):
                owner = p.replace("_rels/", "")[:-5]
                if owner in self.parts and owner in keep:
                    continue
            self.remove(p)
            self.drop_override(p)
            gone.append(p)
        return gone

    # ---- プレゼンテーション ----
    @property
    def pres(self):
        return self.xml(self.main)

    def slide_size(self) -> tuple[int, int]:
        sz = kid(self.pres.documentElement, "sldSz")
        if sz is None:
            return 12192000, 6858000
        return int(sz.getAttribute("cx")), int(sz.getAttribute("cy"))

    def slide_ids(self) -> list:
        lst = kid(self.pres.documentElement, "sldIdLst")
        return kids(lst, "sldId") if lst is not None else []

    def slide_parts(self) -> list[str]:
        out = []
        for s in self.slide_ids():
            t = self.rel_target(self.main, s.getAttributeNS(NS_R, "id"))
            if t:
                out.append(t)
        return out

    def slide_part(self, index: int) -> str:
        parts = self.slide_parts()
        if not 0 <= index < len(parts):
            raise PptxError(f"スライド {index + 1} は無い（全 {len(parts)} 枚）")
        return parts[index]

    def first_number(self) -> int:
        v = self.pres.documentElement.getAttribute("firstSlideNum")
        return int(v) if v else 1

    def rel_of_type(self, part: str, type_: str) -> str | None:
        for r in self.rels(part):
            if r.type == type_ and not r.external:
                return r.target
        return None

    def layout_of(self, slide: str) -> str | None:
        return self.rel_of_type(slide, RT_LAYOUT)

    def master_of(self, layout: str) -> str | None:
        return self.rel_of_type(layout, RT_MASTER)

    def notes_of(self, slide: str) -> str | None:
        return self.rel_of_type(slide, RT_NOTES)

    def layouts(self) -> list[dict]:
        out = []
        for m in [r.target for r in self.rels(self.main) if r.type == RT_MASTER]:
            for r in self.rels(m):
                if r.type == RT_LAYOUT and r.target in self.parts:
                    csld = kid(self.xml(r.target).documentElement, "cSld")
                    out.append({"part": r.target, "name": csld.getAttribute("name") if csld else r.target})
        return out


# ---------------------------------------------------------------- 図形
def cnvpr(sh):
    """図形の p:cNvPr（id と名前を持つ要素）。"""
    for n in kids(sh):
        if n.localName.startswith("nv") and n.localName.endswith("Pr"):
            return kid(n, "cNvPr")
    return None


def nvpr(sh):
    for n in kids(sh):
        if n.localName.startswith("nv") and n.localName.endswith("Pr"):
            return kid(n, "nvPr")
    return None


def shape_id(sh) -> int | None:
    c = cnvpr(sh)
    try:
        return int(c.getAttribute("id")) if c is not None else None
    except ValueError:
        return None


def placeholder(sh) -> dict | None:
    nv = nvpr(sh)
    ph = kid(nv, "ph") if nv is not None else None
    if ph is None:
        return None
    return {"type": ph.getAttribute("type") or "obj", "idx": ph.getAttribute("idx") or ""}


SHAPE_TAGS = {"sp", "pic", "grpSp", "graphicFrame", "cxnSp", "contentPart"}


def sp_tree(doc):
    return path(doc.documentElement, "cSld", "spTree")


def top_shapes(doc) -> list:
    t = sp_tree(doc)
    return [n for n in kids(t) if n.localName in SHAPE_TAGS] if t is not None else []


def shape_pr(sh):
    """xfrm を持つ要素（p:spPr / p:grpSpPr / p:xfrm 直下の graphicFrame）。"""
    for name in ("spPr", "grpSpPr"):
        e = kid(sh, name)
        if e is not None:
            return e
    return None


def xfrm_el(sh):
    if sh.localName == "graphicFrame":
        return kid(sh, "xfrm")
    pr = shape_pr(sh)
    return kid(pr, "xfrm") if pr is not None else None


def read_xfrm(x) -> dict | None:
    if x is None:
        return None
    off, ext = kid(x, "off"), kid(x, "ext")
    if off is None or ext is None:
        return None
    d = {"x": int(off.getAttribute("x")), "y": int(off.getAttribute("y")),
         "w": int(ext.getAttribute("cx")), "h": int(ext.getAttribute("cy")),
         "rot": int(x.getAttribute("rot") or 0) / 60000,
         "flipH": x.getAttribute("flipH") in ("1", "true"), "flipV": x.getAttribute("flipV") in ("1", "true")}
    co, ce = kid(x, "chOff"), kid(x, "chExt")
    if co is not None and ce is not None:
        d["ch"] = {"x": int(co.getAttribute("x")), "y": int(co.getAttribute("y")),
                   "w": int(ce.getAttribute("cx")), "h": int(ce.getAttribute("cy"))}
    return d


def find_shape(doc, sid: int):
    """spTree の中（グループの中も含む）で id が sid の図形と、その祖先のグループの列。"""
    def walk(parent, chain):
        for n in kids(parent):
            if n.localName not in SHAPE_TAGS:
                continue
            if shape_id(n) == sid:
                return n, chain
            if n.localName == "grpSp":
                r = walk(n, chain + [n])
                if r:
                    return r
        return None
    t = sp_tree(doc)
    r = walk(t, []) if t is not None else None
    if not r:
        raise PptxError(f"図形 id={sid} が無い（ほかで直されたかもしれない）")
    return r


def all_ids(doc) -> set[int]:
    out = set()
    for c in doc.documentElement.getElementsByTagNameNS("*", "cNvPr"):
        try:
            out.add(int(c.getAttribute("id")))
        except ValueError:
            pass
    return out


def group_to_parent(g: dict, x: float, y: float) -> tuple[float, float]:
    ch = g.get("ch") or {"x": g["x"], "y": g["y"], "w": g["w"], "h": g["h"]}
    sx = g["w"] / ch["w"] if ch["w"] else 1
    sy = g["h"] / ch["h"] if ch["h"] else 1
    return g["x"] + (x - ch["x"]) * sx, g["y"] + (y - ch["y"]) * sy


def group_scale(chain_xfrms: list[dict]) -> tuple[float, float]:
    sx = sy = 1.0
    for g in chain_xfrms:
        ch = g.get("ch") or g
        sx *= g["w"] / ch["w"] if ch["w"] else 1
        sy *= g["h"] / ch["h"] if ch["h"] else 1
    return sx, sy


def to_slide(chain_xfrms: list[dict], f: dict) -> dict:
    """グループの中の座標（子の座標系）を、スライドの座標にする。"""
    x, y = f["x"], f["y"]
    x2, y2 = f["x"] + f["w"], f["y"] + f["h"]
    for g in reversed(chain_xfrms):
        x, y = group_to_parent(g, x, y)
        x2, y2 = group_to_parent(g, x2, y2)
    return {**f, "x": round(x), "y": round(y), "w": round(x2 - x), "h": round(y2 - y)}


def from_slide(chain_xfrms: list[dict], f: dict) -> dict:
    """スライドの座標を、グループの中の座標（子の座標系）にする。"""
    x, y = f["x"], f["y"]
    x2, y2 = f["x"] + f["w"], f["y"] + f["h"]
    for g in chain_xfrms:
        ch = g.get("ch") or g
        sx = g["w"] / ch["w"] if ch["w"] else 1
        sy = g["h"] / ch["h"] if ch["h"] else 1
        x, y = ch["x"] + (x - g["x"]) / sx, ch["y"] + (y - g["y"]) / sy
        x2, y2 = ch["x"] + (x2 - g["x"]) / sx, ch["y"] + (y2 - g["y"]) / sy
    return {**f, "x": round(x), "y": round(y), "w": round(x2 - x), "h": round(y2 - y)}


def inherited_xfrm(pkg: Package, slide: str, sh) -> dict | None:
    """xfrm を持たないプレースホルダーは、レイアウト→マスターの同じプレースホルダーの位置を使う。"""
    ph = placeholder(sh)
    if not ph:
        return None
    layout = pkg.layout_of(slide)
    for part in (layout, pkg.master_of(layout) if layout else None):
        if not part or not pkg.has(part):
            continue
        cands = [s for s in top_shapes(pkg.xml(part)) if placeholder(s)]
        hit = None
        if ph["idx"]:
            hit = next((s for s in cands if placeholder(s)["idx"] == ph["idx"]), None)
        if hit is None:
            t = "title" if ph["type"] in _TITLE_PH else ph["type"]
            hit = next((s for s in cands
                        if ("title" if placeholder(s)["type"] in _TITLE_PH else placeholder(s)["type"]) == t), None)
        if hit is None and ph["type"] in ("obj", "body", "subTitle"):
            hit = next((s for s in cands if placeholder(s)["type"] in ("body", "obj")), None)
        if hit is not None:
            x = read_xfrm(xfrm_el(hit))
            if x:
                return x
    return None


# ---------------------------------------------------------------- テキスト
SOFT_BREAK = "\v"   # 段落の中の改行（a:br）。画面では ↵ で見せる


def paragraph_segments(p) -> list[tuple[str, str, object]]:
    """段落の中身を (種類, 文字, 要素) の列にする。種類は r（文字）/ br（段落内改行）/ fld（ページ番号など）。"""
    segs = []
    for n in kids(p):
        if n.localName == "r":
            t = kid(n, "t")
            segs.append(("r", text_of(t) if t is not None else "", n))
        elif n.localName == "br":
            segs.append(("br", SOFT_BREAK, n))
        elif n.localName == "fld":
            t = kid(n, "t")
            segs.append(("fld", text_of(t) if t is not None else "", n))
    return segs


def paragraph_text(p) -> str:
    return "".join(s[1] for s in paragraph_segments(p))


def body_paragraphs(tx) -> list:
    return kids(tx, "p") if tx is not None else []


def body_text(tx) -> list[str]:
    return [paragraph_text(p) for p in body_paragraphs(tx)]


def _clone_rpr(doc, p):
    """段落の文字の書式（最初の run の rPr、無ければ endParaRPr）を複製する。"""
    for n in kids(p):
        if n.localName in ("r", "fld"):
            r = kid(n, "rPr")
            if r is not None:
                return r.cloneNode(True)
    e = kid(p, "endParaRPr")
    if e is not None:
        c = A(doc, "rPr")
        for i in range(e.attributes.length):
            a = e.attributes.item(i)
            c.setAttribute(a.name, a.value)
        for ch in kids(e):
            c.appendChild(ch.cloneNode(True))
        return c
    return None


def _new_run(doc, text: str, rpr):
    r = A(doc, "r")
    if rpr is not None:
        r.appendChild(rpr.cloneNode(True))
    t = A(doc, "t")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    return r


def _pieces(doc, text: str, rpr) -> list:
    """\v を a:br にしながら run の列を作る。"""
    out = []
    for i, chunk in enumerate(text.split(SOFT_BREAK)):
        if i:
            br = A(doc, "br")
            if rpr is not None:
                br.appendChild(rpr.cloneNode(True))
            out.append(br)
        if chunk:
            out.append(_new_run(doc, chunk, rpr))
    return out


def edit_paragraph(p, new: str) -> bool:
    """段落の文字を new にする。変わった範囲だけを書き換え、前後の run の書式（色・太字など）を保つ。"""
    doc = p.ownerDocument
    segs = paragraph_segments(p)
    old = "".join(s[1] for s in segs)
    if old == new:
        return False
    pre = 0
    while pre < min(len(old), len(new)) and old[pre] == new[pre]:
        pre += 1
    suf = 0
    while suf < min(len(old), len(new)) - pre and old[len(old) - 1 - suf] == new[len(new) - 1 - suf]:
        suf += 1
    d0, d1 = pre, len(old) - suf          # old[d0:d1] を ins に置き換える
    ins = new[pre:len(new) - suf]
    spans, pos = [], 0
    for kind, t, el in segs:
        spans.append((kind, t, el, pos, pos + len(t)))
        pos += len(t)

    # 挿入先の run。置き換えなら置き換えられる最初の文字の run（選んで打ち替えた文字は元の書式を引き継ぐ）、
    # 挿入だけなら直前の文字の run（続けて打った文字は前の書式を引き継ぐ）
    host = None
    if d1 > d0:
        host = next((sp for sp in spans if sp[0] == "r" and sp[3] <= d0 < sp[4]), None)
    if host is None:
        host = next((sp for sp in spans if sp[0] == "r" and sp[3] < d0 <= sp[4]), None)
    if host is None:
        host = next((sp for sp in spans if sp[0] == "r" and sp[3] == d0 and sp[4] > sp[3]), None)
    # 挿入先が無いときに新しい run を入れる位置と、その書式
    anchor = next((sp[2] for sp in spans if sp[3] >= d1 and not (d0 < sp[4] and sp[3] < d1)), None)
    prev_r = [sp for sp in spans if sp[0] == "r" and sp[4] <= d0]
    rpr = kid(prev_r[-1][2], "rPr") if prev_r else None
    rpr = rpr.cloneNode(True) if rpr is not None else _clone_rpr(doc, p)

    for kind, t, el, s, e in spans:
        overlap = max(s, d0) < min(e, d1)
        if kind == "r":
            keep_a = t[:d0 - s] if d0 > s else ""
            keep_b = t[d1 - s:] if d1 < e else ""
            if host is not None and el is host[2]:
                parts = ins.split(SOFT_BREAK)
                set_text(kid(el, "t"), keep_a + parts[0] + ("" if len(parts) > 1 else keep_b))
                if len(parts) > 1:
                    nxt = el.nextSibling
                    for x in _pieces(doc, SOFT_BREAK + SOFT_BREAK.join(parts[1:]) + keep_b, kid(el, "rPr")):
                        p.insertBefore(x, nxt)
            elif overlap:
                if keep_a + keep_b:
                    set_text(kid(el, "t"), keep_a + keep_b)
                else:
                    p.removeChild(el)
        elif overlap:          # br / fld は分けられないので、変わる範囲にかかれば消す
            p.removeChild(el)
    if host is None and ins:
        if anchor is None or anchor.parentNode is not p:
            anchor = kid(p, "endParaRPr")
        for x in _pieces(doc, ins, rpr):
            p.insertBefore(x, anchor) if anchor is not None else p.appendChild(x)
    for kind, t, el in paragraph_segments(p):   # 空になった run を消す
        if kind == "r" and t == "":
            p.removeChild(el)
    return True


def _new_paragraph_like(tpl, text: str):
    doc = tpl.ownerDocument
    p = A(doc, "p")
    ppr = kid(tpl, "pPr")
    if ppr is not None:
        p.appendChild(ppr.cloneNode(True))
    rpr = _clone_rpr(doc, tpl)
    for x in _pieces(doc, text, rpr):
        p.appendChild(x)
    end = kid(tpl, "endParaRPr")
    if end is not None:
        p.appendChild(end.cloneNode(True))
    return p


def edit_body(tx, new_paras: list[str]) -> bool:
    """テキスト本体の段落を new_paras にする。同じ段落はそのまま、変わった段落だけ書き換える。"""
    ps = body_paragraphs(tx)
    old = [paragraph_text(p) for p in ps]
    if old == new_paras:
        return False
    if not new_paras:
        new_paras = [""]
    a = 0
    while a < min(len(old), len(new_paras)) and old[a] == new_paras[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new_paras)) - a and old[len(old) - 1 - b] == new_paras[len(new_paras) - 1 - b]:
        b += 1
    om, nm = ps[a:len(ps) - b], new_paras[a:len(new_paras) - b]
    anchor = ps[len(ps) - b] if b else None
    tpl = (om[-1] if om else (ps[a - 1] if a else (ps[0] if ps else None)))
    for i in range(max(len(om), len(nm))):
        if i < len(om) and i < len(nm):
            edit_paragraph(om[i], nm[i])
            tpl = om[i]
        elif i < len(nm):
            if tpl is None:
                tpl = A(tx.ownerDocument, "p")
            np_ = _new_paragraph_like(tpl, nm[i])
            tx.insertBefore(np_, anchor) if anchor is not None else tx.appendChild(np_)
        else:
            tx.removeChild(om[i])
    if not body_paragraphs(tx):
        tx.appendChild(A(tx.ownerDocument, "p"))
    return True


def text_body(sh):
    return kid(sh, "txBody")


def first_rpr_info(tx) -> dict:
    """テキストの代表の書式（最初の run）。画面の書式の欄に出す。"""
    info = {}
    for p in body_paragraphs(tx):
        ppr = kid(p, "pPr")
        if ppr is not None and "algn" not in info and ppr.getAttribute("algn"):
            info["algn"] = ppr.getAttribute("algn")
        for n in kids(p):
            rpr = kid(n, "rPr") if n.localName in ("r", "fld") else (n if n.localName == "endParaRPr" else None)
            if rpr is None:
                continue
            if rpr.getAttribute("sz"):
                info.setdefault("sz", int(rpr.getAttribute("sz")) / 100)
            for k in ("b", "i"):
                if rpr.getAttribute(k):
                    info.setdefault(k, rpr.getAttribute(k) in ("1", "true"))
            if rpr.getAttribute("u"):
                info.setdefault("u", rpr.getAttribute("u") != "none")
            f = kid(rpr, "solidFill")
            c = kid(f, "srgbClr") if f is not None else None
            if c is not None:
                info.setdefault("color", "#" + c.getAttribute("val"))
            for k in ("latin", "ea"):
                e = kid(rpr, k)
                if e is not None and e.getAttribute("typeface"):
                    info.setdefault("font_" + k, e.getAttribute("typeface"))
    return info


def fill_info(sh) -> str | None:
    pr = shape_pr(sh)
    f = kid(pr, "solidFill") if pr is not None else None
    c = kid(f, "srgbClr") if f is not None else None
    if c is not None:
        return "#" + c.getAttribute("val")
    if pr is not None and kid(pr, "noFill") is not None:
        return "none"
    return None


# ---------------------------------------------------------------- 画面に渡すスライドの中身
KIND_LABEL = {"sp": "図形", "pic": "画像", "grpSp": "グループ", "graphicFrame": "表・グラフ", "cxnSp": "線",
              "contentPart": "インク"}


def _is_txbox(sh) -> bool:
    c = path(sh, "nvSpPr", "cNvSpPr")
    return c is not None and c.getAttribute("txBox") in ("1", "true")


def shape_model(pkg: Package, slide: str, sh, chain: list[dict]) -> dict:
    c = cnvpr(sh)
    x = read_xfrm(xfrm_el(sh))
    src = "xfrm"
    if x is None and not chain:
        x = inherited_xfrm(pkg, slide, sh)
        src = "inherited" if x else "none"
    frame = to_slide(chain, x) if x else None
    tx = text_body(sh)
    ph = placeholder(sh)
    m = {"id": shape_id(sh), "name": c.getAttribute("name") if c is not None else "", "kind": sh.localName,
         "label": KIND_LABEL.get(sh.localName, sh.localName), "ph": ph, "frame": frame, "frame_src": src,
         "rot": (x or {}).get("rot", 0), "hidden": c is not None and c.getAttribute("hidden") in ("1", "true"),
         "txbox": _is_txbox(sh)}
    if tx is not None:
        m["text"] = body_text(tx)
        m["style"] = first_rpr_info(tx)
    fi = fill_info(sh)
    if fi:
        m["fill"] = fi
    if sh.localName == "graphicFrame":
        from .tables import table_model
        t = table_model(sh)
        if t:
            m["table"] = t
            m["label"] = "表"
    pr = shape_pr(sh)
    ln = kid(pr, "ln") if pr is not None else None
    if ln is not None:
        c = path(ln, "solidFill", "srgbClr")
        dash = kid(ln, "prstDash")
        m["line"] = {"w": int(ln.getAttribute("w")) / 12700 if ln.getAttribute("w") else None,
                     "color": "#" + c.getAttribute("val") if c is not None else None,
                     "none": kid(ln, "noFill") is not None, "dash": dash.getAttribute("val") if dash is not None else "solid",
                     "head": (kid(ln, "headEnd").getAttribute("type") if kid(ln, "headEnd") is not None else "none"),
                     "tail": (kid(ln, "tailEnd").getAttribute("type") if kid(ln, "tailEnd") is not None else "none")}
    if tx is not None:
        bp = kid(tx, "bodyPr")
        if bp is not None:
            m["anchor"] = bp.getAttribute("anchor") or "t"
            m["autofit"] = ("shape" if kid(bp, "spAutoFit") is not None else
                            "shrink" if kid(bp, "normAutofit") is not None else "none")
        ppr = kid(body_paragraphs(tx)[0], "pPr") if body_paragraphs(tx) else None
        if ppr is not None:
            m["bullet"] = ("char" if kid(ppr, "buChar") is not None else "number" if kid(ppr, "buAutoNum") is not None
                           else "none" if kid(ppr, "buNone") is not None else "")
            ls = path(ppr, "lnSpc", "spcPct")
            if ls is not None:
                m["line_spacing"] = int(ls.getAttribute("val")) / 100000
    if sh.localName == "pic":
        m["image"] = True
    if sh.localName == "grpSp":
        gx = x
        m["children"] = [shape_model(pkg, slide, k, chain + ([gx] if gx else []))
                         for k in kids(sh) if k.localName in SHAPE_TAGS]
    return m


def slide_title(pkg: Package, part: str) -> str:
    doc = pkg.xml(part)
    shapes = top_shapes(doc)
    for sh in shapes:
        ph = placeholder(sh)
        if ph and ph["type"] in _TITLE_PH:
            t = " ".join(x for x in body_text(text_body(sh)) if x).replace(SOFT_BREAK, " ")
            if t.strip():
                return t.strip()
    for sh in shapes:
        ph = placeholder(sh)
        tx = text_body(sh)
        if tx is not None and not (ph and ph["type"] in ("sldNum", "dt", "ftr", "hdr")):
            t = " ".join(x for x in body_text(tx) if x).replace(SOFT_BREAK, " ").strip()
            if t:
                return t
    return ""


def is_hidden(pkg: Package, part: str) -> bool:
    return pkg.xml(part).documentElement.getAttribute("show") in ("0", "false")


def notes_text(pkg: Package, slide: str) -> list[str] | None:
    notes = pkg.notes_of(slide)
    if not notes or not pkg.has(notes):
        return None
    sh = notes_body_shape(pkg.xml(notes))
    return body_text(text_body(sh)) if sh is not None else []


def notes_body_shape(doc):
    for sh in top_shapes(doc):
        ph = placeholder(sh)
        if ph and ph["type"] == "body" and text_body(sh) is not None:
            return sh
    return None


def slide_model(pkg: Package, index: int) -> dict:
    from . import anim
    part = pkg.slide_part(index)
    doc = pkg.xml(part)
    shapes = [shape_model(pkg, part, sh, []) for sh in top_shapes(doc)]
    layout = pkg.layout_of(part)
    lname = ""
    if layout and pkg.has(layout):
        cs = kid(pkg.xml(layout).documentElement, "cSld")
        lname = cs.getAttribute("name") if cs is not None else ""
    names = {}

    def walk(ms):
        for m in ms:
            # 一覧では、文字があれば文字で見せる（Google Slides 由来の「Google Shape;12;p3」などは見分けにくい）
            text = " ".join(t for t in m.get("text", []) if t).replace(SOFT_BREAK, " ").strip()
            names[m["id"]] = f"「{text[:18]}」" if text else f"{m['label']} {m['name']}"
            walk(m.get("children", []))
    walk(shapes)
    effects = anim.effects(doc)
    for e in effects:
        e["shape"] = names.get(e["spid"], f"id={e['spid']}")
    from .slides import background_info, transition_info
    return {"index": index, "part": part, "hidden": is_hidden(pkg, part), "layout": lname,
            "shapes": shapes, "anims": effects, "notes": notes_text(pkg, part),
            "has_notes_master": bool(pkg.rel_of_type(pkg.main, RT_NOTES_MASTER)),
            "transition": transition_info(doc), "background": background_info(doc)}


def fonts_used(pkg: Package) -> dict[str, int]:
    """文書で使われているフォント名と、使われている回数（テーマの既定のフォントを含む）。"""
    out: dict[str, int] = {}
    pat = re.compile(rb'<a:(?:latin|ea|cs)\b[^>]*?typeface="([^"]*)"')
    for name, data in pkg.parts.items():
        if not name.endswith(".xml") or not (name.startswith("ppt/slides/") or name.startswith("ppt/slideLayouts/")
                                             or name.startswith("ppt/slideMasters/") or name.startswith("ppt/theme/")):
            continue
        for m in pat.finditer(data):
            f = m.group(1).decode("utf-8", "replace")
            if f and not f.startswith("+"):
                out[f] = out.get(f, 0) + 1
    return out


def is_editable(p: Path) -> bool:
    return p.suffix.lower() in EDITABLE_SUFFIXES
