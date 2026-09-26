"""表（a:tbl）の読み取りと編集。セルの文字、行・列の追加と削除、表の挿入。

結合したセル（gridSpan・rowSpan・hMerge・vMerge）がある表は、文字は直せるが、行・列の追加と削除はしない
（結合の範囲を正しく付け替えるのは PowerPoint に任せる）。
"""
from __future__ import annotations

from .oputil import _slide, append_shape, frag
from .pptx import PptxError, all_ids, body_text, edit_body, find_shape, kid, kids, path

TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"   # PowerPoint の既定（中間スタイル 2 - アクセント 1）


def tbl_of(sh):
    if sh is None or sh.localName != "graphicFrame":
        return None
    return path(sh, "graphic", "graphicData", "tbl")


def _cells(tr) -> list:
    return kids(tr, "tc")


def _merged(tbl) -> bool:
    for tc in tbl.getElementsByTagNameNS("*", "tc"):
        if any(tc.getAttribute(a) for a in ("gridSpan", "rowSpan", "hMerge", "vMerge")):
            return True
    return False


def table_model(sh) -> dict | None:
    """画面の表の編集に渡す中身。rows[行][列] = {text, span, rspan, merged}"""
    tbl = tbl_of(sh)
    if tbl is None:
        return None
    cols = [int(g.getAttribute("w") or 0) for g in kids(kid(tbl, "tblGrid"), "gridCol")]
    rows = []
    for tr in kids(tbl, "tr"):
        row = []
        for tc in _cells(tr):
            tx = kid(tc, "txBody")
            row.append({"text": body_text(tx) if tx is not None else [""],
                        "span": int(tc.getAttribute("gridSpan") or 1), "rspan": int(tc.getAttribute("rowSpan") or 1),
                        "merged": tc.getAttribute("hMerge") in ("1", "true") or tc.getAttribute("vMerge") in ("1", "true")})
        rows.append({"h": int(tr.getAttribute("h") or 0), "cells": row})
    return {"cols": cols, "rows": rows, "has_merge": _merged(tbl)}


def _table(pkg, args):
    part, doc = _slide(pkg, args)
    sh, _ = find_shape(doc, int(args["id"]))
    tbl = tbl_of(sh)
    if tbl is None:
        raise PptxError("表ではない")
    return part, doc, sh, tbl


def _ext(sh):
    x = kid(sh, "xfrm")
    return kid(x, "ext") if x is not None else None


def _clear_ids(el) -> None:
    """行・列を写したとき、PowerPoint の行 id・列 id（a16:rowId など）が重ならないよう外す。"""
    for e in list(el.getElementsByTagNameNS("*", "extLst")):
        if e.parentNode is el:
            el.removeChild(e)


def op_table_text(pkg, args):
    """セルの文字を変える。row・col は 0 始まり。text は段落の列。"""
    part, doc, sh, tbl = _table(pkg, args)
    trs = kids(tbl, "tr")
    r, c = int(args["row"]), int(args["col"])
    if not (0 <= r < len(trs)) or not (0 <= c < len(_cells(trs[r]))):
        raise PptxError("そのセルは無い")
    tc = _cells(trs[r])[c]
    tx = kid(tc, "txBody")
    if tx is None:
        tx = frag(doc, '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP"/></a:p></a:txBody>')[0]
        tc.insertBefore(tx, tc.firstChild)
    edit_body(tx, [str(t) for t in args["text"]])
    pkg.touch(part)
    return {"label": "表の文字"}


def op_table_row(pkg, args):
    """行を足す（where: above・below）・消す（where: delete）。表の高さも合わせる。"""
    part, doc, sh, tbl = _table(pkg, args)
    if _merged(tbl):
        raise PptxError("結合したセルがある表の行は、PowerPoint で足し引きする")
    trs = kids(tbl, "tr")
    r = int(args["row"])
    if not 0 <= r < len(trs):
        raise PptxError("その行は無い")
    ref = trs[r]
    h = int(ref.getAttribute("h") or 370840)
    ext = _ext(sh)
    if args["where"] == "delete":
        if len(trs) == 1:
            raise PptxError("最後の1行は消せない（表ごと消す）")
        tbl.removeChild(ref)
        if ext is not None:
            ext.setAttribute("cy", str(max(0, int(ext.getAttribute("cy")) - h)))
        label = "行の削除"
    else:
        new = ref.cloneNode(True)
        _clear_ids(new)
        for tc in _cells(new):
            tx = kid(tc, "txBody")
            if tx is not None:
                edit_body(tx, [""])
        tbl.insertBefore(new, ref if args["where"] == "above" else ref.nextSibling)
        if ext is not None:
            ext.setAttribute("cy", str(int(ext.getAttribute("cy")) + h))
        label = "行の挿入"
    pkg.touch(part)
    return {"label": label}


def op_table_col(pkg, args):
    """列を足す（where: left・right）・消す（where: delete）。表の幅も合わせる。"""
    part, doc, sh, tbl = _table(pkg, args)
    if _merged(tbl):
        raise PptxError("結合したセルがある表の列は、PowerPoint で足し引きする")
    grid = kid(tbl, "tblGrid")
    gcs = kids(grid, "gridCol")
    c = int(args["col"])
    if not 0 <= c < len(gcs):
        raise PptxError("その列は無い")
    w = int(gcs[c].getAttribute("w") or 0)
    ext = _ext(sh)
    if args["where"] == "delete":
        if len(gcs) == 1:
            raise PptxError("最後の1列は消せない（表ごと消す）")
        grid.removeChild(gcs[c])
        for tr in kids(tbl, "tr"):
            tr.removeChild(_cells(tr)[c])
        if ext is not None:
            ext.setAttribute("cx", str(max(0, int(ext.getAttribute("cx")) - w)))
        label = "列の削除"
    else:
        after = args["where"] == "right"
        g = gcs[c].cloneNode(True)
        _clear_ids(g)
        grid.insertBefore(g, gcs[c].nextSibling if after else gcs[c])
        for tr in kids(tbl, "tr"):
            tc = _cells(tr)[c]
            n = tc.cloneNode(True)
            tx = kid(n, "txBody")
            if tx is not None:
                edit_body(tx, [""])
            tr.insertBefore(n, tc.nextSibling if after else tc)
        if ext is not None:
            ext.setAttribute("cx", str(int(ext.getAttribute("cx")) + w))
        label = "列の挿入"
    pkg.touch(part)
    return {"label": label}


def op_table_add(pkg, args):
    """表を入れる（rows × cols。既定は 3 × 3）。"""
    part, doc = _slide(pkg, args)
    rows = max(1, min(50, int(args.get("rows", 3))))
    cols = max(1, min(30, int(args.get("cols", 3))))
    cx, cy = pkg.slide_size()
    w = round(cx * 0.7)
    h = 370840
    x, y = (cx - w) // 2, max(0, (cy - h * rows) // 2)
    cw = w // cols
    nid = max(all_ids(doc) | {1}) + 1
    cell = ('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP" altLang="en-US"/></a:p>'
            '</a:txBody><a:tcPr/></a:tc>')
    xml = (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{nid}" name="表 {nid}"/><p:cNvGraphicFramePr>'
           f'<a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
           f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cw * cols}" cy="{h * rows}"/></p:xfrm>'
           f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>'
           f'<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{TABLE_STYLE}</a:tableStyleId></a:tblPr><a:tblGrid>'
           + "".join(f'<a:gridCol w="{cw}"/>' for _ in range(cols)) + '</a:tblGrid>'
           + "".join(f'<a:tr h="{h}">' + cell * cols + '</a:tr>' for _ in range(rows))
           + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')
    append_shape(doc, frag(doc, xml)[0])
    pkg.touch(part)
    return {"label": "表の挿入", "select": [nid]}


OPS = {"table_text": op_table_text, "table_row": op_table_row, "table_col": op_table_col}
