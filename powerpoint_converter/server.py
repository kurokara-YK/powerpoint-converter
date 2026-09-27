"""ローカルのブラウザで資料を選び、PowerPoint と同じ並びの画面で見て直すためのサーバ。

127.0.0.1 でしか待ち受けない。画面とのやり取りでは、パスは data からの相対パスで渡す。
書き換えるのは、開いている資料のファイルと、data の中への取り込みだけ。
"""
from __future__ import annotations

import json
import mimetypes
import os
import re
import signal
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from . import pptx, render
from .deck import Conflict, Deck, touch_cache_dirs
from .office import Office, OfficeError
from .pptx import EDITABLE_SUFFIXES, VIEW_SUFFIXES, PptxError

STATIC = Path(__file__).parent / "static"
_BAD_NAME = re.compile(r'[\\/:*?"<>|]')
DECK_SUFFIXES = EDITABLE_SUFFIXES | VIEW_SUFFIXES
_SKIP = {"__pycache__", "node_modules"}


def _safe_name(name: str, fallback: str = "") -> str:
    name = _BAD_NAME.sub("_", name).strip().lstrip(".")
    return name or fallback


def _is_deck(p: Path) -> bool:
    return p.is_file() and p.suffix.lower() in DECK_SUFFIXES and not p.name.startswith((".", "~$"))


class App:
    def __init__(self, data: Path, start: str = ""):
        self.data = data.resolve()
        self.start = start
        self.office = Office("serve")
        self.office2 = Office("present")     # スライドショー・PDF 用（必要になったときに起動する）
        # まとめて描くときに分担する LibreOffice（CPU の数に合わせて最大 4 つ。使うときだけ起動する）
        self.bulk = [Office(f"bulk{k}") for k in range(1, min(4, max(1, (os.cpu_count() or 4) // 4)) + 1)]
        self.deck: Deck | None = None
        self.clipboard: dict = {}
        self.session = 0
        self._lock = threading.Lock()

    def _data_path(self, rel: str) -> Path:
        p = (self.data / rel).resolve() if rel else self.data
        if p != self.data and self.data not in p.parents:
            raise PptxError(f"{rel} は data の外")
        return p

    def _rel(self, p: Path) -> str:
        return str(p.relative_to(self.data)) if p.is_relative_to(self.data) else str(p)

    # ---- 一覧 ----
    def browse(self, q: dict) -> dict:
        d = self._data_path(q.get("path", ""))
        if not d.is_dir():
            raise PptxError(f"{q.get('path')} が無い")
        folders, decks = [], []
        for p in sorted(d.iterdir(), key=lambda p: p.name.lower()):
            if p.name.startswith((".", "_")) or p.name in _SKIP:
                continue
            if p.is_dir():
                inner = [f for f in p.rglob("*") if _is_deck(f) and ".git" not in f.parts]
                mt = max([f.stat().st_mtime for f in inner] + [p.stat().st_mtime])
                folders.append({"name": p.name, "path": self._rel(p), "count": len(inner), "mtime": mt})
            elif _is_deck(p):
                st = p.stat()
                decks.append({"name": p.name, "path": self._rel(p), "mtime": st.st_mtime, "size": st.st_size,
                              "editable": pptx.is_editable(p), "open": self.deck is not None and self.deck.path == p})
        return {"path": self._rel(d) if d != self.data else "", "folders": folders, "decks": decks,
                "start": self.start}

    def mkdir(self, body: dict) -> dict:
        name = _safe_name(body.get("name", ""))
        if not name:
            raise PptxError("名前が空")
        d = self._data_path(body.get("path", "")) / name
        if d.exists():
            raise PptxError(f"{name} は既にある")
        d.mkdir(parents=True)
        return {"path": self._rel(d)}

    def import_(self, q: dict, data: bytes) -> dict:
        """ブラウザに落とされた資料を、一覧で開いているフォルダに置く。同じ名前があれば _2, _3 … を付ける。"""
        name = _safe_name(Path(q.get("name", "")).name, "slides.pptx")
        if Path(name).suffix.lower() not in DECK_SUFFIXES:
            raise PptxError(f"{name} は PowerPoint の資料ではない（{', '.join(sorted(DECK_SUFFIXES))}）")
        base = self._data_path(q.get("path", ""))
        base.mkdir(parents=True, exist_ok=True)
        stem, suf = Path(name).stem, Path(name).suffix
        dest, n = base / name, 2
        while dest.exists():
            dest, n = base / f"{stem}_{n}{suf}", n + 1
        dest.write_bytes(data)
        return {"path": self._rel(dest)}

    def thumb(self, rel: str) -> bytes | None:
        """資料の1枚目の絵（描いたことがあれば）。一覧のカードに出す。"""
        p = self._data_path(rel)
        if not _is_deck(p) or not pptx.is_editable(p):
            return None
        try:
            pkg = pptx.Package(p)
            fp = render.slide_fp(pkg, 0)
        except (PptxError, KeyError, ValueError, OSError):
            return None
        f = render.render_dir(fp) / "full.png"
        return f.read_bytes() if f.is_file() else None

    # ---- 開く ----
    def open(self, body: dict) -> dict:
        p = self._data_path(body["path"])
        if not _is_deck(p):
            raise PptxError(f"{body['path']} は開けない（PowerPoint の資料ではない）")
        with self._lock:
            if self.deck is None or self.deck.path != p.resolve():
                if self.deck:
                    self.deck.close()
                self.deck = Deck(p, self.office, self.office2, self.bulk)
                self.deck.clipboard = self.clipboard     # 別の資料に貼り付けられるよう、コピーは全部で1つ
                self.session += 1
        return self.info()

    def close_deck(self, _=None) -> dict:
        with self._lock:
            if self.deck:
                self.deck.close()
            self.deck = None
            self.session += 1
        return {"ok": True}

    def need(self) -> Deck:
        if not self.deck:
            raise PptxError("資料が開かれていない")
        return self.deck

    def info(self, _=None) -> dict:
        d = self.deck
        return {"session": self.session, "data": str(self.data), "start": self.start,
                "deck": {"path": self._rel(d.path), **d.model()} if d else None}

    def shutdown(self) -> None:
        if self.deck:
            self.deck.close()
        for o in [self.office, self.office2, *self.bulk]:
            o.release()


def make_handler(app: App):
    class Handler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def log_message(self, *a):
            pass

        def _send(self, code: int, body: bytes, ctype: str, cache: str = "no-store", extra: dict | None = None):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", cache)
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            try:
                self.wfile.write(body)
            except (BrokenPipeError, ConnectionResetError):
                pass      # ブラウザがページを再読み込み・閉じた（返事を待っていた相手がもういない）。害は無い

        def _json(self, obj, code: int = 200):
            self._send(code, json.dumps(obj, ensure_ascii=False).encode(), "application/json; charset=utf-8")

        def _api(self, fn, arg=None):
            try:
                self._json(fn(arg))
            except Conflict as e:
                self._json({"error": str(e), "conflict": True}, HTTPStatus.CONFLICT)
            except (PptxError, OfficeError) as e:
                self._json({"error": str(e)}, HTTPStatus.CONFLICT)
            except (KeyError, ValueError, TypeError) as e:
                self._json({"error": f"不正な要求: {e}"}, HTTPStatus.BAD_REQUEST)

        def _body(self) -> bytes:
            return self.rfile.read(int(self.headers.get("Content-Length", 0)))

        def _download(self, name: str, data: bytes, ctype: str):
            self._send(200, data, ctype, extra={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"})

        def _trusted(self, post: bool) -> bool:
            """ほかのサイトから送らせた要求を断る。Host は 127.0.0.1 か localhost だけ（DNS rebinding 対策）。
            書き換える要求（POST）は、画面が付ける X-PC ヘッダーが要る（付けるとブラウザが事前確認をするので、
            ほかのサイトからは送れない）。"""
            host = (self.headers.get("Host") or "").rsplit(":", 1)[0]
            if host not in ("127.0.0.1", "localhost", "[::1]"):
                self._send(403, b"forbidden host", "text/plain")
                return False
            if post and self.headers.get("X-PC") != "1":
                self._send(403, b"missing X-PC header", "text/plain")
                return False
            return True

        def do_GET(self):
            if not self._trusted(False):
                return
            u = urlparse(self.path)
            q = {k: v[0] for k, v in parse_qs(u.query).items()}
            if u.path == "/":
                return self._send(200, (STATIC / "index.html").read_bytes(), "text/html; charset=utf-8")
            if u.path.startswith("/render/"):
                # /render/<指紋>/<ファイル>。指紋は中身から作るので、同じ URL の中身は変わらない
                m = re.fullmatch(r"/render/([0-9a-f]{8,40})/((?:full|bg|s\d+)\.png|meta\.json)", u.path)
                f = render.render_dir(m.group(1)) / m.group(2) if m else None
                if not f or not f.is_file():
                    return self._send(404, b"not found", "text/plain")
                ctype = "image/png" if f.suffix == ".png" else "application/json"
                return self._send(200, f.read_bytes(), ctype, cache="max-age=31536000, immutable")
            if u.path == "/present.svg":
                d = app.deck
                f = d.present_file if d else None
                if not f or not f.is_file():
                    return self._send(404, b"not ready", "text/plain")
                return self._send(200, f.read_bytes(), "image/svg+xml")
            if u.path.startswith("/preview/"):       # アニメーションのプレビュー用の SVG（中身は指紋で決まる）
                m = re.fullmatch(r"/preview/([0-9a-f]{8,40})\.svg", u.path)
                from .deck import PREVIEW
                f = PREVIEW / f"{m.group(1)}.svg" if m else None
                if not f or not f.is_file():
                    return self._send(404, b"not found", "text/plain")
                return self._send(200, f.read_bytes(), "image/svg+xml", cache="max-age=31536000, immutable")
            if u.path == "/thumb":
                data = app.thumb(q.get("path", ""))
                if data is None:
                    return self._send(404, b"", "text/plain")
                return self._send(200, data, "image/png", cache="max-age=60")
            if u.path == "/download":
                try:
                    d = app.need()
                    if q.get("format") == "pdf":
                        return self._download(d.path.with_suffix(".pdf").name, d.export_pdf(), "application/pdf")
                    return self._download(d.path.name, d.path.read_bytes(),
                                          mimetypes.guess_type(d.path.name)[0] or "application/octet-stream")
                except (PptxError, OfficeError) as e:
                    return self._json({"error": str(e)}, 409)
            if u.path == "/api/poll":
                return self._api(lambda _: app.need().poll(int(q.get("since", -1)),
                                                            int(q["cur"]) if q.get("cur") else None))
            routes = {"/api/info": app.info, "/api/browse": app.browse, "/api/fonts": lambda q: fonts(),
                      "/api/find": lambda q: app.need().find(q.get("q", ""), q.get("case") == "1"),
                      "/api/slide": lambda q: app.need().slide(int(q["i"])),
                      "/api/present": lambda q: app.need().want_present(),
                      "/api/preview": lambda q: app.need().preview(int(q["i"]))}
            if u.path in routes:
                return self._api(routes[u.path], q)
            if u.path.startswith("/static/"):
                f = (STATIC / u.path[len("/static/"):]).resolve()
                if f.is_file() and STATIC.resolve() in f.parents:
                    ctype = mimetypes.guess_type(f.name)[0] or "application/octet-stream"
                    return self._send(200, f.read_bytes(), ctype)
            self._send(404, b"not found", "text/plain")

        def do_POST(self):
            if not self._trusted(True):
                return
            u = urlparse(self.path)
            q = {k: v[0] for k, v in parse_qs(u.query).items()}
            if u.path == "/api/import":
                data = self._body()
                return self._api(lambda _: app.import_(q, data))
            if u.path == "/api/image":           # 画像の挿入・差し替え（本文は画像のバイト列）
                data = self._body()
                args = {"slide": int(q["slide"]), "name": q.get("name", "画像")}
                op = "image"
                if q.get("replace"):
                    op, args["id"] = "replace_image", int(q["replace"])
                return self._api(lambda _: app.need().apply(op, args, q.get("version"), data))
            try:
                body = json.loads(self._body() or b"{}")
            except json.JSONDecodeError:
                return self._json({"error": "JSON が壊れている"}, 400)
            routes = {"/api/open": app.open, "/api/close": app.close_deck, "/api/mkdir": app.mkdir,
                      "/api/op": lambda b: app.need().apply(b["op"], b.get("args", {}), b.get("version")),
                      "/api/undo": lambda b: app.need().do_undo(b.get("version")),
                      "/api/redo": lambda b: app.need().do_redo(b.get("version"))}
            if u.path in routes:
                return self._api(routes[u.path], body)
            self._send(404, b"not found", "text/plain")

    return Handler


def serve(data: Path, start: str, deck: Path | None, port: int, open_browser: bool) -> None:
    touch_cache_dirs()
    app = App(data, start=start)
    httpd = None
    for p in range(port, port + 20):   # 使用中なら次の番号
        try:
            httpd = ThreadingHTTPServer(("127.0.0.1", p), make_handler(app))
            break
        except OSError:
            continue
    if httpd is None:
        raise SystemExit(f"ポート {port}〜{port + 19} が全部使用中")
    httpd.daemon_threads = True
    # 接続が切れただけのエラー（ブラウザの再読み込みなど）は端末に出さない
    _orig_handle_error = httpd.handle_error

    def handle_error(request, client_address):
        import sys
        if isinstance(sys.exc_info()[1], (BrokenPipeError, ConnectionResetError)):
            return
        _orig_handle_error(request, client_address)
    httpd.handle_error = handle_error
    # LibreOffice の起動（数秒）を先に済ませておく
    threading.Thread(target=lambda: _warm(app), daemon=True).start()
    if deck:
        app.open({"path": app._rel(deck.resolve())})
    signal.signal(signal.SIGTERM, lambda *_: (_ for _ in ()).throw(KeyboardInterrupt))
    url = f"http://127.0.0.1:{httpd.server_port}/"
    if deck:
        url += "#deck=" + quote(app._rel(deck.resolve()))
    print(f"powerpoint-converter: {url}   （Ctrl+C で終了）")
    if open_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        app.shutdown()
        httpd.server_close()


_FONTS: list[str] = []


def fonts() -> dict:
    """この PC で使えるフォントの名前（画面のフォントの一覧に出す）。"""
    import subprocess
    if not _FONTS:
        try:
            r = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True, timeout=10)
            names = set()
            for line in r.stdout.splitlines():
                for n in line.split(","):
                    n = n.strip()
                    if n and not n.startswith("."):
                        names.add(n)
            _FONTS.extend(sorted(names, key=lambda s: (not any(ord(c) > 0x3000 for c in s), s.lower())))
        except (OSError, subprocess.TimeoutExpired):
            pass
    return {"fonts": _FONTS}


def _warm(app: App) -> None:
    try:
        app.office._run(lambda d, c: None)
    except OfficeError as e:
        print(f"注意: {e}")
