#!/usr/bin/env python3
"""
Build power-query-explorer.html from source files in src/.

Usage:
    uv run scripts/build.py

Source files (all in src/):
    head_meta.html    - <head> metadata (charset, CSP, title, favicon)
    libraries.html    - Embedded JS libraries (JSZip, Cytoscape, XPress9, DataModel decoder)
    styles.css        - All CSS
    body.html         - HTML body content (drop zone, panels, tabs, footer)
    app.js            - Main application JavaScript

Output:
    power-query-explorer.html - Single self-contained HTML file
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, 'src')
OUT_PATH = os.path.join(ROOT, 'power-query-explorer.html')


def read_src(filename):
    path = os.path.join(SRC_DIR, filename)
    if not os.path.exists(path):
        print(f"ERROR: Missing source file: {path}", file=sys.stderr)
        sys.exit(1)
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


def main():
    head_meta = read_src('head_meta.html')
    libraries = read_src('libraries.html')
    styles = read_src('styles.css')
    body = read_src('body.html')
    app_js = read_src('app.js')

    html = (
        '<!DOCTYPE html>\n'
        '<html lang="en">\n'
        '<head>\n'
        f'{head_meta}\n'
        f'{libraries}\n'
        '    <style>\n'
        f'{styles}\n'
        '    </style>\n'
        '</head>\n'
        '<body>\n'
        f'{body}\n'
        '\n'
        '<script>\n'
        f'{app_js}\n'
        '</script>\n'
        '</body>\n'
        '</html>'
    )

    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(OUT_PATH) / 1024
    print(f"Built {OUT_PATH} ({size_kb:.1f} KB)")


if __name__ == '__main__':
    main()
