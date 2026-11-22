#!/usr/bin/env python3
"""Build index.html from src/: inlines the engine and the sample document.

Usage: python3 build.py
"""
import base64
from pathlib import Path

root = Path(__file__).parent
src = root / "src"

page = (src / "template.html").read_text(encoding="utf-8")
engine = (src / "engine.js").read_text(encoding="utf-8").replace(
    "if (typeof module !== 'undefined') module.exports = Taglift;", ""
)
sample = base64.b64encode((src / "sample.docx").read_bytes()).decode()

page = page.replace("/*ENGINE*/", engine).replace("/*SAMPLE*/", sample)

(root / "index.html").write_text(page, encoding="utf-8")
print(f"Wrote index.html ({len(page):,} bytes)")
