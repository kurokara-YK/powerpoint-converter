"""LibreOffice を裏で常駐させて、スライドを描く。

- 画面は出さない（--headless）。GUI で開いている LibreOffice とぶつからないよう、専用のプロファイルを使う
- 起動は数秒かかるので1回だけにし、あとは UNO（ソケット）で頼む。落ちたら起動し直す
- LibreOffice に保存はさせない。描くだけ（保存すると PowerPoint で開いたときに崩れることがある）
- 利用者のファイルを直接は開かない（横にロックファイルができるため）。一時ファイルに写してから開く
"""
from __future__ import annotations

import json
import math
import os
import shutil
import signal
import socket
import subprocess
import threading
import time
from pathlib import Path

CACHE = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "powerpoint-converter"


class OfficeError(Exception):
    pass


def soffice_path() -> str:
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return p
    for p in ("/usr/lib/libreoffice/program/soffice", "/opt/libreoffice/program/soffice"):
        if os.path.exists(p):
            return p
    raise OfficeError("LibreOffice が見つからない（sudo apt install libreoffice-impress）")


def _uno():
    try:
        import uno  # noqa: F401  （Ubuntu では python3-uno。LibreOffice と一緒に入る）
        return uno
    except ImportError as e:
        raise OfficeError("Python から LibreOffice を操作する python3-uno が無い（sudo apt install python3-uno）") from e


def _pv(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


_UNO_LOCK = threading.Lock()
_RESOLVER = None


def _resolver():
    """UNO の接続の窓口（1つだけ作って使い回す）。複数のスレッドが同時に初めて作ると失敗するのでロックする。"""
    global _RESOLVER
    with _UNO_LOCK:
        if _RESOLVER is None:
            uno = _uno()
            local = uno.getComponentContext()
            _RESOLVER = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        return _RESOLVER


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class Office:
    """LibreOffice 1つ。呼び出しは1本ずつ（lock）。"""

    def __init__(self, name: str = "serve"):
        self.name = name
        self.profile = CACHE / f"lo-profile-{name}"
        self._profile_lock = None
        self._released = False       # 終了したあとは起動し直さない（描画中のスレッドが起動し直して残るのを防ぐ）
        self.proc: subprocess.Popen | None = None
        self.desktop = None
        self.ctx = None
        self.lock = threading.RLock()

    # ---- 起動と終了 ----
    def _claim_profile(self) -> None:
        """使っていないプロファイルを押さえる。同じプロファイルで2つ目の LibreOffice を起動すると、
        1つ目に処理を渡して終了してしまうため（serve を2つ動かしたときなど）。"""
        import fcntl
        if self._profile_lock is not None:
            return
        CACHE.mkdir(parents=True, exist_ok=True)
        for n in range(1, 50):
            name = self.name if n == 1 else f"{self.name}-{n}"
            f = open(CACHE / f"lo-profile-{name}.lock", "w")
            try:
                fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError:
                f.close()
                continue
            self._profile_lock = f
            self.profile = CACHE / f"lo-profile-{name}"
            return
        raise OfficeError("LibreOffice のプロファイルが全部使用中")

    def _start(self) -> None:
        _uno()                       # python3-uno が無ければ、ここで分かる言葉で止める
        self.close()
        self._claim_profile()
        self.profile.mkdir(parents=True, exist_ok=True)
        port = _free_port()
        url = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        self.proc = subprocess.Popen(
            [soffice_path(), f"-env:UserInstallation={self.profile.as_uri()}", "--headless", "--invisible",
             "--norestore", "--nologo", "--nodefault", "--nolockcheck", f"--accept={url}"],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
        resolver = _resolver()
        deadline = time.time() + 60   # 初回はプロファイルを作るので時間がかかる
        while True:
            try:
                self.ctx = resolver.resolve(f"uno:{url}")
                break
            except Exception:
                if self.proc.poll() is not None:
                    raise OfficeError("LibreOffice が起動しなかった")
                if time.time() > deadline:
                    self.close()
                    raise OfficeError("LibreOffice に接続できない（60 秒待った）")
                time.sleep(0.25)
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def close(self) -> None:
        self.desktop = self.ctx = None
        if self.proc and self.proc.poll() is None:
            try:
                os.killpg(self.proc.pid, signal.SIGTERM)
                self.proc.wait(5)
            except (ProcessLookupError, subprocess.TimeoutExpired, PermissionError):
                try:
                    os.killpg(self.proc.pid, signal.SIGKILL)
                except (ProcessLookupError, PermissionError):
                    pass
        self.proc = None

    def release(self) -> None:
        """終了するときに呼ぶ（LibreOffice を止め、プロファイルを手放す）。"""
        self._released = True
        with self.lock:
            self.close()
        if self._profile_lock is not None:
            self._profile_lock.close()
            self._profile_lock = None

    def _run(self, fn):
        """fn(desktop, ctx) を実行する。LibreOffice が落ちていたら起動し直して1回だけやり直す。"""
        with self.lock:
            for attempt in (0, 1):
                if self._released:
                    raise OfficeError("終了中")
                if self.desktop is None or self.proc is None or self.proc.poll() is not None:
                    self._start()
                try:
                    return fn(self.desktop, self.ctx)
                except OfficeError:
                    raise
                except Exception as e:
                    name = type(e).__name__
                    if attempt == 0 and ("Disposed" in name or "Runtime" in name or "Connection" in name
                                         or isinstance(e, (OSError, EOFError))):
                        self.close()
                        continue
                    raise OfficeError(f"LibreOffice での処理に失敗: {name}: {e}") from e

    def _load(self, desktop, path: Path):
        uno = _uno()
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0,
                                           (_pv("Hidden", True), _pv("RepairPackage", True)))
        if doc is None:
            raise OfficeError(f"{path.name} を LibreOffice で開けない")
        return doc

    # ---- 描く ----
    def _export(self, ctx, src, dest: Path, w: int, h: int, translucent: bool = False) -> None:
        uno = _uno()
        ex = ctx.ServiceManager.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", ctx)
        ex.setSourceDocument(src)
        fd = [_pv("PixelWidth", max(1, w)), _pv("PixelHeight", max(1, h))]
        if translucent:
            fd.append(_pv("Translucent", True))
        # FilterData は PropertyValue のタプルのまま渡す（uno.Any で包むと Translucent が効かない）
        ex.filter((_pv("URL", uno.systemPathToFileUrl(str(dest))), _pv("MediaType", "image/png"),
                   _pv("FilterData", tuple(fd))))

    @staticmethod
    def _rect(r) -> dict:
        return {"x": r.X, "y": r.Y, "w": r.Width, "h": r.Height}

    def _shape_info(self, s, deep: bool = True) -> dict:
        def get(name, default=None):
            try:
                return s.getPropertyValue(name)
            except Exception:
                return default
        info = {"desc": get("Description", "") or "", "name": get("Name", "") or "",
                "type": s.ShapeType.rsplit(".", 1)[-1], "bound": self._rect(s.BoundRect),
                "pos": {"x": s.Position.X, "y": s.Position.Y}, "size": {"w": s.Size.Width, "h": s.Size.Height},
                "rot": (get("RotateAngle", 0) or 0) / 100, "visible": bool(get("Visible", True)),
                "empty": bool(get("IsEmptyPresentationObject", False))}
        if deep and info["type"] == "GroupShape":
            info["children"] = [self._shape_info(s.getByIndex(i)) for i in range(s.Count)]
        return info

    def _overflow(self, s) -> int:
        """文字が枠からあふれている量（1/100 mm）。自動で縮小・拡大する枠は 0。"""
        try:
            if not s.getPropertyValue("String"):
                return 0
            if s.getPropertyValue("TextAutoGrowHeight"):
                return 0
            fit = s.getPropertyValue("TextFitToSize")
            if getattr(fit, "value", fit) not in ("NONE", 0):
                return 0
            if (s.getPropertyValue("RotateAngle") or 0) != 0:
                return 0
            h0 = s.Size.Height
            s.setPropertyValue("TextAutoGrowHeight", True)
            h1 = s.Size.Height
            s.setPropertyValue("TextAutoGrowHeight", False)
            return max(0, h1 - h0)
        except Exception:
            return 0

    def _page_checks(self, page, pw: int, ph: int, shapes: list[dict]) -> None:
        """図形ごとに、文字のあふれとスライドの外へのはみ出しを調べて shapes に書き足す。"""
        for i in range(page.Count):
            s = page.getByIndex(i)
            info = shapes[i]
            of = self._overflow(s)
            if of > 100:      # 1 mm を超えるものだけ
                info["overflow"] = of
            b = info["bound"]
            out = max(0, -b["x"]) + max(0, -b["y"]) + max(0, b["x"] + b["w"] - pw) + max(0, b["y"] + b["h"] - ph)
            if out > 500 and info["visible"]:      # 5 mm を超えるものだけ
                info["outside"] = out

    def _render_page(self, ctx, page, out: Path, width: int) -> dict:
        """読み込んだ文書の1ページを描く。out に full.png（全体）、bg.png（背景とマスター）、s<N>.png（図形）、
        meta.json を書く。背景を描くために図形を取り除くので、このページは描いたあと使えない（保存はしない）。"""
        t0 = time.time()
        out.mkdir(parents=True, exist_ok=True)
        pw, ph = page.Width, page.Height
        k = width / pw
        H = round(ph * k)
        self._export(ctx, page, out / "full.png", width, H)
        shapes = []
        for i in range(page.Count):
            s = page.getByIndex(i)
            info = self._shape_info(s)
            b = info["bound"]
            if info["visible"] and not info["empty"] and b["w"] > 0 and b["h"] > 0:
                self._export(ctx, s, out / f"s{i}.png", math.ceil(b["w"] * k), math.ceil(b["h"] * k), True)
                info["png"] = f"s{i}.png"
            shapes.append(info)
        self._page_checks(page, pw, ph, shapes)
        # 図形を取り除いて、背景とマスターの図形だけを描く（表などは Visible を持たないので、隠すのではなく取り除く）
        while page.Count:
            page.remove(page.getByIndex(0))
        self._export(ctx, page, out / "bg.png", width, H)
        meta = {"px": {"w": width, "h": H}, "page": {"w": pw, "h": ph}, "shapes": shapes,
                "ms": round((time.time() - t0) * 1000)}
        (out / "meta.json").write_text(json.dumps(meta, ensure_ascii=False))
        return meta

    def render_slide(self, pptx: Path, out: Path, width: int = 1920, page_index: int = 0) -> dict:
        """PPTX の1ページを描く（1枚だけの PPTX を描くとき）。"""
        def job(desktop, ctx):
            doc = self._load(desktop, pptx)
            try:
                return self._render_page(ctx, doc.DrawPages.getByIndex(page_index), out, width)
            finally:
                doc.close(True)
        return self._run(job)

    def render_pages(self, pptx: Path, claim, done, width: int = 1920) -> int:
        """文書を1回だけ読み込み、claim() が返すページを順に描く（初めて開いた資料をまとめて描くとき）。

        claim() は (ページ番号, 出力先) か None を返す。描き終えるたびに done(ページ番号, 出力先, エラー) を呼ぶ。
        複数の LibreOffice で同じ claim を分け合えば、並列に描ける。描いた枚数を返す。
        """
        def job(desktop, ctx):
            doc = self._load(desktop, pptx)
            n = 0
            try:
                while True:
                    j = claim()
                    if j is None:
                        return n
                    idx, out = j
                    try:
                        self._render_page(ctx, doc.DrawPages.getByIndex(idx), out, width)
                        done(idx, out, None)
                        n += 1
                    except Exception as e:  # noqa: BLE001  1ページの失敗で全部を止めない
                        done(idx, out, f"{type(e).__name__}: {e}")
            finally:
                doc.close(True)
        return self._run(job)

    def render_deck(self, pptx: Path, out: Path, width: int = 1280, checks: bool = True,
                    pages: list[int] | None = None) -> list[dict]:
        """スライドを PNG にする（check 用）。slide-01.png … と、各ページの図形の検査結果を返す。
        pages を渡すと、そのページだけ（複数の LibreOffice で分担するとき）。"""
        out.mkdir(parents=True, exist_ok=True)

        def job(desktop, ctx):
            doc = self._load(desktop, pptx)
            result = []
            try:
                n = doc.DrawPages.Count
                digits = max(2, len(str(n)))
                for i in (range(n) if pages is None else pages):
                    page = doc.DrawPages.getByIndex(i)
                    pw, ph = page.Width, page.Height
                    H = round(ph * width / pw)
                    f = out / f"slide-{i + 1:0{digits}d}.png"
                    self._export(ctx, page, f, width, H)
                    shapes = [self._shape_info(page.getByIndex(j), deep=False) for j in range(page.Count)]
                    if checks:
                        self._page_checks(page, pw, ph, shapes)
                    result.append({"index": i, "png": str(f), "page": {"w": pw, "h": ph}, "shapes": shapes})
            finally:
                doc.close(True)
            return result
        return self._run(job)

    def page_count(self, pptx: Path) -> int:
        def job(desktop, ctx):
            doc = self._load(desktop, pptx)
            try:
                return doc.DrawPages.Count
            finally:
                doc.close(True)
        return self._run(job)

    def export(self, src: Path, dest: Path, filter_name: str, filter_data: dict | None = None) -> None:
        """文書全体を別の形式で書き出す（スライドショー用の SVG、PDF）。"""
        uno = _uno()

        def job(desktop, ctx):
            doc = self._load(desktop, src)
            try:
                props = [_pv("FilterName", filter_name)]
                if filter_data:
                    fd = tuple(_pv(k, v) for k, v in filter_data.items())
                    props.append(_pv("FilterData", uno.Any("[]com.sun.star.beans.PropertyValue", fd)))
                doc.storeToURL(uno.systemPathToFileUrl(str(dest)), tuple(props))
            finally:
                doc.close(True)
        self._run(job)
