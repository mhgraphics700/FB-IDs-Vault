"""Rebuilds backend/Code.gs: the Apps Script backend with the current app page (app/src/main/assets/index.html) embedded.
Run after every change to index.html if you also use the web-page version:  python3 tools/build_backend.py
"""
import re, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
gs = (root / "backend" / "Code.gs").read_text(encoding="utf-8")
html = (root / "app/src/main/assets/index.html").read_text(encoding="utf-8")
assert "`" not in html and "${" not in html, "index.html must not contain backticks or ${ (it sits inside a template literal)"
a = gs.index("const HTML = String.raw`") + len("const HTML = String.raw`")
b = gs.index("</html>`;", a) + len("</html>")
gs = gs[:a] + html.rstrip() + gs[b:]
(root / "backend" / "Code.gs").write_text(gs, encoding="utf-8")
print("Code.gs rebuilt,", len(gs), "chars")
