"""スライドのアニメーション（p:timing）の読み取りと編集。

PowerPoint の「アニメーション ウィンドウ」と同じ単位で扱う。
  メインシーケンス
    └ クリックのまとまり（p:par。クリック1回で始まる）
        └ 時間のまとまり（p:par。「直前の動作の後」で区切られる）
            └ 効果（p:par。presetClass と nodeType を持つ cTn）
編集したあとは cTn の id を振り直し、bldLst（文字を持つ図形の一覧）を合わせる。
"""
from __future__ import annotations

from .pptx import P, PptxError, kid, kids, path, NS_P

# PowerPoint の日本語の画面での名前（よく使うものだけ。無いものは番号で出す）
PRESETS = {
    "entr": {1: "アピール", 2: "スライドイン", 3: "ブラインド", 4: "ボックス", 5: "チェッカーボード", 6: "サークル",
             8: "ダイヤモンド", 9: "ディゾルブイン", 10: "フェード", 12: "ピーク", 13: "プラス", 14: "ランダムストライプ",
             16: "スプリット", 17: "ストリップ", 18: "ホイール", 19: "ホイール", 20: "くさび", 21: "ホイール",
             22: "ワイプ", 23: "ズーム", 26: "バウンド", 37: "ライズアップ", 42: "フロートイン", 47: "ターン",
             53: "ズーム", 55: "スウィベル"},
    "exit": {1: "クリア", 2: "スライドアウト", 10: "フェード", 16: "スプリット", 21: "ホイール", 22: "ワイプ",
             23: "ズーム", 53: "ズーム", 42: "フロートアウト"},
    "emph": {1: "カラーで強調", 6: "拡大/収縮", 8: "スピン", 9: "透過性", 14: "ブリンク", 26: "パルス",
             27: "カラーパルス", 32: "シーソー"},
}
CLASS_LABEL = {"entr": "開始", "exit": "終了", "emph": "強調", "path": "アニメーションの軌跡", "verb": "動作",
               "mediacall": "メディア"}
TRIGGER_LABEL = {"clickEffect": "クリック時", "withEffect": "直前の動作と同時", "afterEffect": "直前の動作の後"}

# 画面から追加できる効果（kind, 名前）
ADDABLE = [("entr", "appear", "アピール"), ("entr", "fade", "フェード"), ("entr", "fly", "スライドイン（下から）"),
           ("entr", "wipe", "ワイプ（下から）"), ("exit", "disappear", "クリア"), ("exit", "fadeout", "フェード（終了）")]


def _ctn(par):
    return kid(par, "cTn")


def _cond_delay(ctn) -> str:
    c = path(ctn, "stCondLst", "cond")
    return c.getAttribute("delay") if c is not None else ""


def _is_click(group_par) -> bool:
    """クリックで始まるまとまりか（自動で始まるものは onBegin の条件を持つ）。"""
    st = kid(_ctn(group_par), "stCondLst")
    conds = kids(st, "cond") if st is not None else []
    return any(c.getAttribute("delay") == "indefinite" for c in conds) and \
        not any(c.getAttribute("evt") == "onBegin" for c in conds)


def main_seq(doc):
    """メインシーケンスの cTn（無ければ None）。"""
    timing = kid(doc.documentElement, "timing")
    if timing is None:
        return None
    for c in timing.getElementsByTagNameNS("*", "cTn"):
        if c.getAttribute("nodeType") == "mainSeq":
            return c
    return None


def _groups(seq_ctn) -> list:
    lst = kid(seq_ctn, "childTnLst")
    return kids(lst, "par") if lst is not None else []


def _effect_pars(group_par) -> list[tuple]:
    """クリックのまとまりの中の (時間のまとまり, 効果の par) の列。"""
    out = []
    for tpar in kids(kid(_ctn(group_par), "childTnLst") or _ctn(group_par), "par"):
        for epar in kids(kid(_ctn(tpar), "childTnLst") or _ctn(tpar), "par"):
            out.append((tpar, epar))
    return out


def _target(epar) -> int | None:
    for t in epar.getElementsByTagNameNS("*", "spTgt"):
        try:
            return int(t.getAttribute("spid"))
        except ValueError:
            pass
    return None


def _dur(epar) -> float:
    """効果の長さ（秒）。子の cTn の delay + dur の最大。"""
    best = 0.0
    for c in epar.getElementsByTagNameNS("*", "cTn"):
        d = c.getAttribute("dur")
        try:
            dur = int(d) if d and d != "indefinite" else 0
            dl = int(_cond_delay(c) or 0)
        except ValueError:
            continue
        best = max(best, (dur + dl) / 1000)
    return best


def effects(doc) -> list[dict]:
    """画面の一覧に出す効果の列。step はクリックの番号（0 はスライドが出たときに自動で始まるもの）。"""
    seq = main_seq(doc)
    if seq is None:
        return []
    out = []
    step = 0
    for gi, g in enumerate(_groups(seq)):
        click = _is_click(g)
        if click:
            step += 1
        for ei, (_, e) in enumerate(_effect_pars(g)):
            c = _ctn(e)
            cls = c.getAttribute("presetClass") or ""
            pid = int(c.getAttribute("presetID") or 0)
            name = PRESETS.get(cls, {}).get(pid) or (f"{CLASS_LABEL.get(cls, cls)} {pid}" if cls else "効果")
            out.append({"n": len(out), "group": gi, "step": step, "click": click,
                        "spid": _target(e), "cls": cls, "cls_label": CLASS_LABEL.get(cls, cls),
                        "name": name, "trigger": c.getAttribute("nodeType") or "",
                        "trigger_label": TRIGGER_LABEL.get(c.getAttribute("nodeType"), c.getAttribute("nodeType")),
                        "dur": round(_dur(e), 2)})
    return out


def other_sequences(doc) -> int:
    """メイン以外のシーケンス（図形をクリックして動くトリガーなど）の数。画面では触らない。"""
    timing = kid(doc.documentElement, "timing")
    if timing is None:
        return 0
    return sum(1 for c in timing.getElementsByTagNameNS("*", "cTn") if c.getAttribute("nodeType") == "interactiveSeq")


# ---------------------------------------------------------------- 編集
def _renumber(doc) -> None:
    timing = kid(doc.documentElement, "timing")
    if timing is None:
        return
    for i, c in enumerate(timing.getElementsByTagNameNS("*", "cTn"), 1):
        c.setAttribute("id", str(i))


def _sync_bld(doc) -> None:
    """bldLst を、アニメーションの付いた「文字を持つ図形」に合わせる。"""
    from .pptx import find_shape, text_body
    timing = kid(doc.documentElement, "timing")
    if timing is None:
        return
    used: dict[int, set[str]] = {}
    for c in timing.getElementsByTagNameNS("*", "cTn"):
        if c.getAttribute("presetClass") in ("entr", "exit", "emph", "path"):
            par = c.parentNode
            sp = _target(par)
            if sp is not None:
                used.setdefault(sp, set()).add(c.getAttribute("grpId") or "0")
    bld = kid(timing, "bldLst")
    if bld is not None:
        for b in kids(bld):
            try:
                sp = int(b.getAttribute("spid"))
            except ValueError:
                continue
            if sp not in used or b.getAttribute("grpId") not in used[sp]:
                bld.removeChild(b)
    have = {(b.getAttribute("spid"), b.getAttribute("grpId")) for b in kids(bld)} if bld is not None else set()
    for sp, gs in used.items():
        try:
            sh, _ = find_shape(doc, sp)
        except PptxError:
            continue
        if sh.localName != "sp" or text_body(sh) is None:
            continue
        for g in gs:
            if (str(sp), g) in have:
                continue
            if bld is None:
                bld = P(doc, "bldLst")
                timing.appendChild(bld)
            bld.appendChild(P(doc, "bldP", {"spid": sp, "grpId": g, "animBg": "1"}))
    if bld is not None and not kids(bld):
        timing.removeChild(bld)


def _cleanup(doc) -> None:
    """空になったまとまりを消し、効果が1つも無くなれば p:timing ごと消す。"""
    seq = main_seq(doc)
    if seq is not None:
        for g in _groups(seq):
            glst = kid(_ctn(g), "childTnLst")
            for tpar in kids(glst, "par") if glst is not None else []:
                tl = kid(_ctn(tpar), "childTnLst")
                if tl is None or not kids(tl, "par"):
                    glst.removeChild(tpar)
            if glst is None or not kids(glst, "par"):
                g.parentNode.removeChild(g)
            else:
                # クリックで始まるまとまりの最初の効果は「クリック時」にする
                first = _ctn(_effect_pars(g)[0][1])
                if _is_click(g) and first.getAttribute("nodeType") != "clickEffect":
                    first.setAttribute("nodeType", "clickEffect")
    timing = kid(doc.documentElement, "timing")
    if timing is not None:
        has_effect = any(c.getAttribute("presetClass") for c in timing.getElementsByTagNameNS("*", "cTn"))
        if not has_effect:
            doc.documentElement.removeChild(timing)
            return
    _sync_bld(doc)
    _renumber(doc)


def delete_effect(doc, n: int) -> None:
    for e in effects(doc):
        if e["n"] == n:
            break
    else:
        raise PptxError("その効果は無い（ほかで直されたかもしれない）")
    seq = main_seq(doc)
    flat = [ep for g in _groups(seq) for _, ep in _effect_pars(g)]
    rest = [(ep, _ctn(ep).getAttribute("nodeType")) for i, ep in enumerate(flat) if i != n]
    flat[n].parentNode.removeChild(flat[n])
    if rest:
        _rebuild(doc, rest)      # 待ち時間（「直前の動作の後」の開始時刻）を計算し直す
    else:
        _cleanup(doc)


def delete_for_shapes(doc, ids: set[int]) -> None:
    """図形を消したときに、その図形の効果も消す（残すと PowerPoint が修復を求める）。"""
    timing = kid(doc.documentElement, "timing")
    if timing is None:
        return
    for c in list(timing.getElementsByTagNameNS("*", "cTn")):
        if c.getAttribute("presetClass") or c.getAttribute("nodeType") in ("clickEffect", "withEffect", "afterEffect"):
            par = c.parentNode
            if par.parentNode is not None and _target(par) in ids:
                par.parentNode.removeChild(par)
    # メイン以外（トリガー）で消した図形を指すものは、シーケンスごと消す
    for c in list(timing.getElementsByTagNameNS("*", "cTn")):
        if c.getAttribute("nodeType") == "interactiveSeq":
            seq = c.parentNode
            if any(_target_of(t) in ids for t in seq.getElementsByTagNameNS("*", "spTgt")) and seq.parentNode:
                seq.parentNode.removeChild(seq)
    _cleanup(doc)


def _target_of(t) -> int | None:
    try:
        return int(t.getAttribute("spid"))
    except ValueError:
        return None


def move_step(doc, group: int, delta: int) -> None:
    """クリックのまとまりを前後に入れ替える（PowerPoint の「順番を前にする／後にする」）。"""
    seq = main_seq(doc)
    gs = _groups(seq) if seq is not None else []
    j = group + delta
    if not (0 <= group < len(gs) and 0 <= j < len(gs)):
        raise PptxError("これ以上動かせない")
    a, b = gs[group], gs[j]
    if not _is_click(a) or not _is_click(b):
        raise PptxError("スライドが出たときに自動で始まる効果は動かせない")
    parent = a.parentNode
    if delta < 0:
        parent.insertBefore(a, b)
    else:
        parent.insertBefore(b, a)
    _cleanup(doc)


def set_trigger(doc, n: int, trigger: str) -> None:
    """効果の開始のしかたを変える。まとまりを組み直す。"""
    if trigger not in TRIGGER_LABEL:
        raise PptxError(f"開始のしかた {trigger} は使えない")
    seq = main_seq(doc)
    flat = [ep for g in _groups(seq) for _, ep in _effect_pars(g)] if seq is not None else []
    if not 0 <= n < len(flat):
        raise PptxError("その効果は無い")
    epar = flat[n]
    _ctn(epar).setAttribute("nodeType", trigger)
    _rebuild(doc, [(ep, _ctn(ep).getAttribute("nodeType")) for ep in flat])


def _rebuild(doc, items: list[tuple]) -> None:
    """効果の列（par, nodeType）から、クリック・時間のまとまりを作り直す。"""
    seq = main_seq(doc)
    lst = kid(seq, "childTnLst")
    auto_first = False
    old_groups = _groups(seq)
    if old_groups and not _is_click(old_groups[0]):
        auto_first = True
    for g in old_groups:
        lst.removeChild(g)
    for ep, _ in items:
        if ep.parentNode is not None:
            ep.parentNode.removeChild(ep)
    group = tpar = None
    t_off = 0.0     # まとまりの中での、次の「後」の開始時刻（ミリ秒）
    t_len = 0.0
    for i, (ep, trig) in enumerate(items):
        if group is None or trig == "clickEffect":
            auto = group is None and trig != "clickEffect" and (auto_first or i == 0)
            group = _par(doc, "indefinite" if not auto else "0", auto_from_start=auto)
            lst.appendChild(group)
            tpar, t_off, t_len = None, 0.0, 0.0
        if tpar is None or trig == "afterEffect" and kids(kid(_ctn(tpar), "childTnLst"), "par"):
            t_off += t_len if tpar is not None else 0
            t_len = 0.0
            tpar = _par(doc, str(int(t_off)))
            kid(_ctn(group), "childTnLst").appendChild(tpar)
        kid(_ctn(tpar), "childTnLst").appendChild(ep)
        t_len = max(t_len, _dur(ep) * 1000)
    _cleanup(doc)


def _par(doc, delay: str, auto_from_start: bool = False):
    par = P(doc, "par")
    c = P(doc, "cTn", {"id": "0", "fill": "hold"})
    st = P(doc, "stCondLst")
    cond = P(doc, "cond", {"delay": delay})
    st.appendChild(cond)
    if auto_from_start:
        # スライドが出たときに自動で始まる（PowerPoint と同じ書き方）
        cond.setAttribute("delay", "indefinite")
        c2 = P(doc, "cond", {"evt": "onBegin", "delay": "0"})
        c2.appendChild(P(doc, "tn", {"val": "2"}))
        st.appendChild(c2)
    c.appendChild(st)
    c.appendChild(P(doc, "childTnLst"))
    par.appendChild(c)
    return par


def _ensure_main(doc):
    seq = main_seq(doc)
    if seq is not None:
        return seq
    root = doc.documentElement
    timing = kid(root, "timing")
    if timing is None:
        timing = P(doc, "timing")
        # p:timing は p:transition の後、p:extLst の前
        ext = kid(root, "extLst")
        root.insertBefore(timing, ext) if ext is not None else root.appendChild(timing)
    tn = kid(timing, "tnLst")
    if tn is None:
        tn = P(doc, "tnLst")
        timing.insertBefore(tn, timing.firstChild)
    par = P(doc, "par")
    root_ctn = P(doc, "cTn", {"id": "1", "dur": "indefinite", "restart": "never", "nodeType": "tmRoot"})
    rl = P(doc, "childTnLst")
    seq = P(doc, "seq", {"concurrent": "1", "nextAc": "seek"})
    seq_ctn = P(doc, "cTn", {"id": "2", "dur": "indefinite", "nodeType": "mainSeq"})
    seq_ctn.appendChild(P(doc, "childTnLst"))
    seq.appendChild(seq_ctn)
    for tag, evt in (("prevCondLst", "onPrev"), ("nextCondLst", "onNext")):
        cl = P(doc, tag)
        cond = P(doc, "cond", {"evt": evt, "delay": "0"})
        tgt = P(doc, "tgtEl")
        tgt.appendChild(P(doc, "sldTgt"))
        cond.appendChild(tgt)
        cl.appendChild(cond)
        seq.appendChild(cl)
    rl.appendChild(seq)
    root_ctn.appendChild(rl)
    par.appendChild(root_ctn)
    tn.appendChild(par)
    return seq_ctn


def _set_vis(doc, spid: int, val: str, delay: int):
    s = P(doc, "set")
    b = P(doc, "cBhvr")
    c = P(doc, "cTn", {"id": "0", "dur": "1", "fill": "hold"})
    st = P(doc, "stCondLst")
    st.appendChild(P(doc, "cond", {"delay": str(delay)}))
    c.appendChild(st)
    b.appendChild(c)
    b.appendChild(_tgt(doc, spid))
    an = P(doc, "attrNameLst")
    a = P(doc, "attrName")
    a.appendChild(doc.createTextNode("style.visibility"))
    an.appendChild(a)
    b.appendChild(an)
    s.appendChild(b)
    to = P(doc, "to")
    to.appendChild(P(doc, "strVal", {"val": val}))
    s.appendChild(to)
    return s


def _tgt(doc, spid: int):
    t = P(doc, "tgtEl")
    t.appendChild(P(doc, "spTgt", {"spid": spid}))
    return t


def _anim_effect(doc, spid: int, transition: str, filt: str, dur: int = 500):
    e = P(doc, "animEffect", {"transition": transition, "filter": filt})
    b = P(doc, "cBhvr")
    b.appendChild(P(doc, "cTn", {"id": "0", "dur": str(dur)}))
    b.appendChild(_tgt(doc, spid))
    e.appendChild(b)
    return e


def _anim_prop(doc, spid: int, attr: str, frm: str, to: str, dur: int = 500):
    a = P(doc, "anim", {"calcmode": "lin", "valueType": "num"})
    b = P(doc, "cBhvr", {"additive": "base"})
    b.appendChild(P(doc, "cTn", {"id": "0", "dur": str(dur), "fill": "hold"}))
    b.appendChild(_tgt(doc, spid))
    an = P(doc, "attrNameLst")
    x = P(doc, "attrName")
    x.appendChild(doc.createTextNode(attr))
    an.appendChild(x)
    b.appendChild(an)
    a.appendChild(b)
    tl = P(doc, "tavLst")
    for tm, v in (("0", frm), ("100000", to)):
        tav = P(doc, "tav", {"tm": tm})
        val = P(doc, "val")
        val.appendChild(P(doc, "strVal", {"val": v}))
        tav.appendChild(val)
        tl.appendChild(tav)
    a.appendChild(tl)
    return a


def _effect_par(doc, spid: int, effect: str, trigger: str):
    presets = {"appear": ("entr", 1, 0), "fade": ("entr", 10, 0), "fly": ("entr", 2, 4), "wipe": ("entr", 22, 4),
               "disappear": ("exit", 1, 0), "fadeout": ("exit", 10, 0)}
    if effect not in presets:
        raise PptxError(f"効果 {effect} は追加できない")
    cls, pid, sub = presets[effect]
    par = P(doc, "par")
    c = P(doc, "cTn", {"id": "0", "presetID": pid, "presetClass": cls, "presetSubtype": sub, "fill": "hold",
                       "grpId": "0", "nodeType": trigger})
    st = P(doc, "stCondLst")
    st.appendChild(P(doc, "cond", {"delay": "0"}))
    c.appendChild(st)
    ch = P(doc, "childTnLst")
    if effect == "appear":
        ch.appendChild(_set_vis(doc, spid, "visible", 0))
    elif effect == "fade":
        ch.appendChild(_set_vis(doc, spid, "visible", 0))
        ch.appendChild(_anim_effect(doc, spid, "in", "fade"))
    elif effect == "wipe":
        ch.appendChild(_set_vis(doc, spid, "visible", 0))
        ch.appendChild(_anim_effect(doc, spid, "in", "wipe(up)"))
    elif effect == "fly":
        ch.appendChild(_set_vis(doc, spid, "visible", 0))
        ch.appendChild(_anim_prop(doc, spid, "ppt_x", "#ppt_x", "#ppt_x"))
        ch.appendChild(_anim_prop(doc, spid, "ppt_y", "1+#ppt_h/2", "#ppt_y"))
    elif effect == "disappear":
        ch.appendChild(_set_vis(doc, spid, "hidden", 0))
    elif effect == "fadeout":
        ch.appendChild(_anim_effect(doc, spid, "out", "fade"))
        ch.appendChild(_set_vis(doc, spid, "hidden", 499))
    c.appendChild(ch)
    par.appendChild(c)
    return par


def add_effect(doc, spid: int, effect: str, trigger: str = "clickEffect") -> None:
    """効果を最後に足す（PowerPoint で「アニメーションの追加」を押したときと同じ）。"""
    if trigger not in TRIGGER_LABEL:
        raise PptxError(f"開始のしかた {trigger} は使えない")
    seq = _ensure_main(doc)
    flat = [(ep, _ctn(ep).getAttribute("nodeType")) for g in _groups(seq) for _, ep in _effect_pars(g)]
    flat.append((_effect_par(doc, spid, effect, trigger), trigger))
    _rebuild(doc, flat)


def has_other_timing(doc) -> bool:
    return other_sequences(doc) > 0


__all__ = ["effects", "delete_effect", "delete_for_shapes", "move_step", "set_trigger", "add_effect", "ADDABLE",
           "NS_P"]
