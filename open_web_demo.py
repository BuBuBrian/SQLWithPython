"""Serve / open index.html in default browser.

If Flask is available → run dev server, else open file URI.
"""
from __future__ import annotations
import importlib.util, threading, webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent
INDEX_HTML = ROOT / "index.html"

def _have_flask() -> bool:
    return importlib.util.find_spec("flask") is not None

def _serve_flask():
    from flask import Flask, send_from_directory  # type: ignore
    app = Flask(__name__, static_folder=str(ROOT))

    @app.route("/")
    def _index():  # type: ignore
        return send_from_directory(ROOT, "index.html")

    @app.route("/<path:path>")
    def _static(path):  # type: ignore
        return send_from_directory(ROOT, path)

    print("* Running on http://127.0.0.1:5000  (Ctrl+C to quit)")
    app.run(debug=False, use_reloader=False)

def main():
    if not INDEX_HTML.exists():
        raise SystemExit("index.html not found")
    if _have_flask():
        threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
        _serve_flask()
    else:
        print("Flask not installed – opening static file …")
        webbrowser.open(INDEX_HTML.as_uri())

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass