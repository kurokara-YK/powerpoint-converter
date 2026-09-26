"""開いている資料1つぶんの状態。

- ファイルの監視: 約 0.5 秒ごとに更新時刻と大きさを見る。外（Claude Code・LibreOffice・PowerPoint）で
  直されたら読み直し、変わったスライドだけを描き直す。書きかけ（zip として開けない間）は待つ
- 描画: LibreOffice に1枚ずつ頼む。今見ているスライドを先に、その近くを次に描く
- 操作: 画面からの操作は、毎回ファイルを読み直してから当てる。画面が見ている版（version）と
  ファイルが違えば当てずに断る（外で直された内容を上書きしないため）
- 元に戻す: 操作で変わった部品の、前と後の中身を持つ。外で直されたら履歴は捨てる
"""
from __future__ import annotations

import os
import shutil
import tempfile
import threading
import time
from pathlib import Path

from . import edit, pptx, render
from .office import CACHE, Office, OfficeError
from .pptx import Package, PptxError

PRESENT = CACHE / "present"
PREVIEW = CACHE / "preview"
CONVERTED = CACHE / "converted"
HISTORY_LIMIT = 100
BULK_MIN = 6          # 描く枚数がこれ以上なら、資料をまるごと読み込んで複数の LibreOffice で分担する
BULK_PER_OFFICE = 5   # LibreOffice 1つあたり、これくらいの枚数は受け持たせる（読み込みに 2〜4 秒かかるため）


class Conflict(PptxError):
    pass


def stat_version(p: Path) -> str:
    try:
        st = p.stat()
    except FileNotFoundError:
        return "missing"
    return f"{st.st_mtime_ns}-{st.st_size}"


def lock_owner(p: Path) -> str | None:
    """LibreOffice が GUI で開いていれば、そのロックファイルの持ち主。"""
    lock = p.with_name(f".~lock.{p.name}#")
    if not lock.exists():
        return None
    try:
        parts = lock.read_text(errors="replace").split(",")
        return parts[0] or parts[1] or "LibreOffice"
    except OSError:
        return "LibreOffice"


class Deck:
    def __init__(self, path: Path, office: Office, office2: Office | None = None, bulk: list[Office] | None = None):
        self.path = Path(path).resolve()
        self.office = office                 # スライドを描く（直したスライドを1枚ずつ）
        self.office2 = office2 or office     # スライドショー・PDF を書き出す（時間がかかるので別にする）
        self.bulk = bulk or []               # 初めて開いたときなど、まとめて描くときに分担する LibreOffice
        self.clipboard: dict = {}            # コピーした図形（App が全部の資料で同じものを渡す）
        self._bulk_active = False
        self._fails: dict[str, int] = {}
        self.editable = pptx.is_editable(self.path)
        self.lock = threading.RLock()
        self.cond = threading.Condition(self.lock)
        self.rev = 0                   # 画面に見せるものが変わるたびに増える
        self.version = ""
        self.pkg: Package | None = None
        self.slides: list[dict] = []
        self.status: dict[str, str] = {}    # 指紋 → "ready" / "error: …"
        self.current = 0
        self.undo: list[dict] = []
        self.redo: list[dict] = []
        self.error: str | None = None
        self.notice: dict | None = None     # 外で直されたときの知らせ（画面に出す）
        self.changed: dict[str, float] = {}  # 指紋 → 外で直されて描き直した時刻（一覧で強調する）
        self.present_status = "none"
        self.present_file: Path | None = None
        self.present_wanted = False
        self.fonts: list[dict] = []
        self._stop = threading.Event()
        self._reload(initial=True)
        render.prune()
        self._threads = [threading.Thread(target=self._watch, daemon=True),
                         threading.Thread(target=self._render_loop, daemon=True),
                         threading.Thread(target=self._present_loop, daemon=True)]
        for t in self._threads:
            t.start()

    def close(self) -> None:
        self._stop.set()
        with self.cond:
            self.cond.notify_all()

    # ---- 読み込み ----
    def _source(self) -> Path:
        """描く元の PPTX。.ppt や .odp は LibreOffice で PPTX にしてから扱う（見るだけ）。"""
        if self.editable:
            return self.path
        CONVERTED.mkdir(parents=True, exist_ok=True)
        import hashlib
        key = hashlib.sha1(f"{self.path}|{stat_version(self.path)}".encode()).hexdigest()[:24]
        out = CONVERTED / f"{key}.pptx"
        if not out.exists():
            with tempfile.TemporaryDirectory() as d:
                tmp = Path(d) / ("src" + self.path.suffix)
                shutil.copy2(self.path, tmp)
                self.office.export(tmp, out, "Impress Office Open XML")
        return out

    def _reload(self, initial: bool = False, external: bool = False) -> None:
        with self.lock:
            v = stat_version(self.path)
            try:
                pkg = Package(self._source())
                base = render.deck_base(pkg)
                slides = []
                for i, part in enumerate(pkg.slide_parts()):
                    doc = pkg.xml(part)
                    from . import anim
                    from .slides import transition_info
                    slides.append({"index": i, "number": pkg.first_number() + i, "part": part,
                                   "fp": render.slide_fp(pkg, i, base), "hidden": pptx.is_hidden(pkg, part),
                                   "title": pptx.slide_title(pkg, part), "anims": len(anim.effects(doc)),
                                   "transition": transition_info(doc)["effect"]})
            except (PptxError, OfficeError, OSError, KeyError, ValueError) as e:
                self.error = str(e)
                self.version = v
                self.rev += 1
                self.cond.notify_all()
                return
            old = {s["fp"] for s in self.slides}
            self.pkg, self.slides, self.version, self.error = pkg, slides, v, None
            for s in slides:
                if s["fp"] not in self.status and render.is_rendered(s["fp"]):
                    self.status[s["fp"]] = "ready"
            if external and not initial:
                new = [s for s in slides if s["fp"] not in old]
                now = time.time()
                for s in new:
                    self.changed[s["fp"]] = now
                self.undo.clear()
                self.redo.clear()
                self.notice = {"t": now, "text": f"外で直された内容を読み込んだ（{len(new)} 枚が変わった）"
                               if new else "外で保存された（見た目の変化は無い）",
                               "slides": [s["index"] for s in new]}
            self.fonts = render.font_report(pptx.fonts_used(pkg))
            self.present_status = "none"
            self.present_file = None
            self.rev += 1
            self.cond.notify_all()

    def _watch(self) -> None:
        last = self.version
        while not self._stop.wait(0.5):
            with self.lock:           # 自分の保存と入れ違いにならないよう、ロックの中で見る
                v = stat_version(self.path)
                if v == self.version:
                    last = v
                    continue
            if v != last:        # まだ書いている途中かもしれない。変化が止まるまで待つ
                last = v
                continue
            if v == "missing":
                with self.lock:
                    self.error = f"{self.path.name} が無くなった（名前が変わったか、消された）"
                    self.version = v
                    self.rev += 1
                    self.cond.notify_all()
                continue
            if self.editable:
                try:
                    import zipfile
                    with zipfile.ZipFile(self.path) as z:
                        z.getinfo("[Content_Types].xml")
                except (zipfile.BadZipFile, KeyError, OSError):
                    continue     # 書きかけ
            with self.lock:
                if stat_version(self.path) == self.version:
                    continue
                self._reload(external=True)

    # ---- 描く ----
    def _next_job(self):
        if not self.slides or self.pkg is None:
            return None
        n = len(self.slides)
        cur = min(self.current, n - 1)
        order = sorted(range(n), key=lambda i: (abs(i - cur) > 2, abs(i - cur), i))
        for i in order:
            s = self.slides[i]
            if s["fp"] not in self.status:
                return i, s["fp"], self.pkg
        return None

    def _render_loop(self) -> None:
        while not self._stop.is_set():
            with self.cond:
                pending = sum(1 for x in self.slides if x["fp"] not in self.status)
                if self.bulk and not self._bulk_active and pending >= BULK_MIN and self.pkg is not None:
                    self._bulk_active = True
                    threading.Thread(target=self._bulk_run, args=(self.pkg, pending), daemon=True).start()
                job = self._next_job()
                if job is None:
                    self.cond.wait(1.0)
                    continue
                i, fp, pkg = job
                self.status[fp] = "running"
            try:
                self._render_slide(pkg, i, fp)
                err = None
            except (OfficeError, PptxError, OSError) as e:
                err = str(e)
            self._finish(fp, err)

    def _finish(self, fp: str, err: str | None) -> None:
        """1枚描き終えたとき。失敗は1回目なら描き直しの列に戻す（LibreOffice が落ちただけのことがある）。"""
        with self.cond:
            if err is None:
                self.status[fp] = "ready"
            elif self._fails.get(fp, 0) == 0:
                self._fails[fp] = 1
                self.status.pop(fp, None)
            else:
                self.status[fp] = f"error: {err}"
            self.rev += 1
            self.cond.notify_all()

    def _bulk_run(self, pkg: Package, pending: int) -> None:
        """資料をまるごと（非表示も表示にし、図形に id を書き込んで）読み込み、複数の LibreOffice で分担して描く。
        1枚ずつ描くより速い（35 枚で約 35 秒 → 10 秒前後）。スライド番号も正しく出る。"""
        n = max(1, min(len(self.bulk), pending // BULK_PER_OFFICE))
        offices = self.bulk[:n]
        fps = [x["fp"] for x in self.slides]
        claimed: set[str] = set()
        try:
            with tempfile.TemporaryDirectory(prefix="pc-bulk-") as d:
                src = Path(d) / "deck.pptx"
                render.full_copy(pkg, src, show_hidden=True, tag=True)

                def claim():
                    with self.cond:
                        if self._stop.is_set() or self.pkg is not pkg:
                            return None
                        job = self._next_job()
                        if job is None:
                            return None
                        i, fp, _ = job
                        self.status[fp] = "running"
                        claimed.add(fp)
                        return i, Path(d) / f"out-{fp}"

                def done(i, out, err):
                    fp = fps[i]
                    if err is None:
                        dest = render.render_dir(fp)
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        if not render.is_rendered(fp):
                            shutil.rmtree(dest, ignore_errors=True)
                            shutil.move(str(out), str(dest))
                    self._finish(fp, err)

                errors = []

                def work(o: Office):
                    try:
                        o.render_pages(src, claim, done)
                    except (OfficeError, OSError) as e:
                        errors.append(str(e))
                ts = [threading.Thread(target=work, args=(o,), daemon=True) for o in offices]
                for t in ts:
                    t.start()
                for t in ts:
                    t.join()
        finally:
            for o in offices:       # メモリを空ける（次にまとめて描くときに起動し直す。1 秒ほど）
                with o.lock:
                    o.close()
            with self.cond:
                # 取ったのに描き終えていないもの（running のまま）は、1枚ずつ描く列に戻す
                for fp in claimed:
                    if self.status.get(fp) == "running":
                        self.status.pop(fp, None)
                self._bulk_active = False
                self.cond.notify_all()

    def _present_loop(self) -> None:
        """スライドショー用の SVG を、頼まれたら（一度頼まれたあとは、直すたびに）作る。"""
        while not self._stop.is_set():
            with self.cond:
                if not (self.present_wanted and self.present_status == "none" and self.pkg is not None):
                    self.cond.wait(1.0)
                    continue
                pkg = self.pkg
                fp = render.deck_fp(pkg)
                self.present_status = "running"
            try:
                path = self._render_present(pkg, fp)
            except (OfficeError, PptxError, OSError):
                path = None
            with self.cond:
                if self.pkg is pkg:
                    self.present_status = "ready" if path else "error"
                    self.present_file = path
                    self.rev += 1
                    self.cond.notify_all()

    def _render_slide(self, pkg: Package, i: int, fp: str) -> None:
        out = render.render_dir(fp)
        if render.is_rendered(fp):
            return
        with tempfile.TemporaryDirectory(prefix="pc-") as d:
            mini = Path(d) / "slide.pptx"
            page = render.mini_deck(pkg, i, mini)
            tmp_out = Path(d) / "out"
            self.office.render_slide(mini, tmp_out, page_index=page)
            out.parent.mkdir(parents=True, exist_ok=True)
            if out.exists():
                shutil.rmtree(out, ignore_errors=True)
            shutil.move(str(tmp_out), str(out))

    def _render_present(self, pkg: Package, fp: str) -> Path:
        """スライドショー用の SVG（LibreOffice のスライドショーの仕組みごと書き出す。アニメーションが動く）。"""
        PRESENT.mkdir(parents=True, exist_ok=True)
        out = PRESENT / f"{fp}.svg"
        if out.exists():
            return out
        with tempfile.TemporaryDirectory(prefix="pc-") as d:
            src = Path(d) / "deck.pptx"
            render.full_copy(pkg, src)
            tmp = Path(d) / "deck.svg"
            self.office2.export(src, tmp, "impress_svg_Export")
            shutil.move(str(tmp), str(out))
        return out

    def preview(self, i: int) -> dict:
        """そのスライドだけのスライドショーの SVG（編集画面でアニメーションをプレビューする用）。
        非表示のスライドも出す。スライド番号を合わせるための空のスライドが前に並ぶので、その数（page）も返す。"""
        with self.lock:
            if self.pkg is None:
                raise PptxError(self.error or "資料が読めていない")
            if not 0 <= i < len(self.slides):
                raise PptxError(f"スライド {i + 1} は無い")
            pkg, fp = self.pkg, self.slides[i]["fp"]
        PREVIEW.mkdir(parents=True, exist_ok=True)
        out = PREVIEW / f"{fp}.svg"
        page = render.needs_number(pkg, i) and i + pkg.first_number() - 1 or 0
        if not out.exists():
            with tempfile.TemporaryDirectory(prefix="pc-") as d:
                mini = Path(d) / "slide.pptx"
                page = render.mini_deck(pkg, i, mini, tag=False)
                tmp = Path(d) / "slide.svg"
                self.office.export(mini, tmp, "impress_svg_Export")
                shutil.move(str(tmp), str(out))
        return {"url": f"/preview/{fp}.svg", "page": page}

    def want_present(self) -> dict:
        with self.cond:
            self.present_wanted = True
            if self.pkg is not None and self.present_status == "none":
                f = PRESENT / f"{render.deck_fp(self.pkg)}.svg"
                if f.exists():
                    self.present_status, self.present_file = "ready", f
            self.cond.notify_all()
            return {"status": self.present_status}

    # ---- 画面に渡すもの ----
    def model(self) -> dict:
        with self.lock:
            pkg = self.pkg
            cx, cy = pkg.slide_size() if pkg else (12192000, 6858000)
            now = time.time()
            return {
                "name": self.path.name, "version": self.version, "rev": self.rev, "editable": self.editable,
                "error": self.error, "size": {"cx": cx, "cy": cy},
                "slides": [{**s, "status": self.status.get(s["fp"], "pending"),
                            "changed": now - self.changed.get(s["fp"], 0) < 8} for s in self.slides],
                "layouts": pkg.layouts() if pkg and self.editable else [],
                "fonts": self.fonts, "lock_by": lock_owner(self.path),
                "can_undo": bool(self.undo), "can_redo": bool(self.redo),
                "undo_label": self.undo[-1]["label"] if self.undo else "",
                "redo_label": self.redo[-1]["label"] if self.redo else "",
                "notice": self.notice, "present": self.present_status,
            }

    def slide(self, i: int) -> dict:
        import json
        with self.lock:
            if self.pkg is None:
                raise PptxError(self.error or "資料が読めていない")
            if not 0 <= i < len(self.slides):
                raise PptxError(f"スライド {i + 1} は無い")
            m = pptx.slide_model(self.pkg, i)
            s = self.slides[i]
            m["fp"] = s["fp"]
            m["status"] = self.status.get(s["fp"], "pending")
            m["number"] = s["number"]
        meta_f = render.render_dir(s["fp"]) / "meta.json"
        m["render"] = json.loads(meta_f.read_text()) if meta_f.is_file() else None
        return m

    def poll(self, since: int, cur: int | None, timeout: float = 20) -> dict:
        """画面からの問い合わせ。何か変わるまで（最長 timeout 秒）待ってから返す。"""
        with self.cond:
            if cur is not None and cur != self.current:
                self.current = cur
                self.cond.notify_all()
            end = time.time() + timeout
            while self.rev <= since and not self._stop.is_set():
                left = end - time.time()
                if left <= 0:
                    break
                self.cond.wait(left)
            return self.model()

    # ---- 操作 ----
    def apply(self, op: str, args: dict, version: str | None, data: bytes | None = None) -> dict:
        if not self.editable and op != "copy":
            raise PptxError(f"{self.path.suffix} は見るだけ（直せるのは .pptx）")
        with self.lock:
            if version and version != stat_version(self.path):
                raise Conflict("ほかで保存された内容がある。最新を読み込んだので、もう一度操作する")
            if op in ("copy", "cut"):
                from . import clip
                self.clipboard["data"] = clip.copy_shapes(self.pkg, int(args["slide"]), [int(i) for i in args["ids"]])
                if op == "copy":
                    return {"label": "コピー", "version": self.version, "copied": len(args["ids"])}
                op = "delete"            # 切り取り = コピーしてから消す
            if op == "paste":
                args = {**args, "clip": self.clipboard.get("data")}
            pkg = Package(self.path)
            before = dict(pkg.parts)
            result = edit.apply(pkg, op, args, data)
            changed = pkg.flush()
            if not changed:
                return {**result, "version": self.version}
            entry = {"label": result.get("label", op), "slide": args.get("slide", args.get("index")),
                     "before": {p: before.get(p) for p in changed}, "after": changed}
            pkg.save()
            self.undo.append(entry)
            del self.undo[:-HISTORY_LIMIT]
            self.redo.clear()
            self.notice = None
            self._reload()
            return {**result, "version": self.version}

    def _restore(self, parts: dict[str, bytes | None], version: str | None) -> None:
        if version and version != stat_version(self.path):
            raise Conflict("ほかで保存された内容がある。最新を読み込んだ")
        pkg = Package(self.path)
        for p, data in parts.items():
            if data is None:
                pkg.remove(p)
            else:
                pkg.put(p, data)
        pkg.save()

    def do_undo(self, version: str | None) -> dict:
        with self.lock:
            if not self.undo:
                raise PptxError("元に戻す操作が無い")
            e = self.undo.pop()
            self._restore(e["before"], version)
            self.redo.append(e)
            self._reload()
            return {"label": e["label"], "slide": e["slide"], "version": self.version}

    def do_redo(self, version: str | None) -> dict:
        with self.lock:
            if not self.redo:
                raise PptxError("やり直す操作が無い")
            e = self.redo.pop()
            self._restore(e["after"], version)
            self.undo.append(e)
            self._reload()
            return {"label": e["label"], "slide": e["slide"], "version": self.version}

    def find(self, q: str, case: bool = False) -> dict:
        from . import search
        with self.lock:
            if self.pkg is None:
                raise PptxError(self.error or "資料が読めていない")
            return {"results": search.find_all(self.pkg, q, case)}

    # ---- 書き出し ----
    def export_pdf(self) -> bytes:
        with self.lock:
            pkg = self.pkg
        if pkg is None:
            raise PptxError(self.error or "資料が読めていない")
        with tempfile.TemporaryDirectory(prefix="pc-") as d:
            src = Path(d) / "deck.pptx"
            render.full_copy(pkg, src)
            out = Path(d) / "deck.pdf"
            self.office2.export(src, out, "impress_pdf_Export")
            return out.read_bytes()


def touch_cache_dirs() -> None:
    for d in (CACHE, PRESENT, CONVERTED, render.RENDERS):
        d.mkdir(parents=True, exist_ok=True)
    os.utime(CACHE)
