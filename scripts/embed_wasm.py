#!/usr/bin/env python3
"""Embeds the XPress9 WASM binary and decoder module into power-query-explorer.html"""
import base64
import re
import sys
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML_PATH = os.path.join(ROOT, 'power-query-explorer.html')
WASM_PATH = os.path.join(ROOT, 'wasm', 'xpress9', 'xpress9.wasm')
XPRESS9_JS_PATH = os.path.join(ROOT, 'wasm', 'xpress9', 'xpress9.js')
DECODER_JS_PATH = os.path.join(ROOT, 'wasm', 'xpress9', 'datamodel-decoder.js')

def main():
    # Read inputs
    with open(WASM_PATH, 'rb') as f:
        wasm_b64 = base64.b64encode(f.read()).decode('ascii')
    with open(XPRESS9_JS_PATH, 'r') as f:
        xpress9_js = f.read().strip()
    with open(DECODER_JS_PATH, 'r') as f:
        decoder_js = f.read()
    with open(HTML_PATH, 'r') as f:
        html = f.read()

    # Remove node.js module exports from xpress9.js (not needed in browser)
    xpress9_js = re.sub(r'if\(typeof exports===.*$', '', xpress9_js, flags=re.MULTILINE).strip()

    # Inject WASM base64 into decoder
    decoder_js = decoder_js.replace("'%%XPRESS9_WASM_B64%%'", f"'{wasm_b64}'")

    # Build the script block
    script_block = f"""<script>/* XPress9 WASM Decompressor | (c) Microsoft Corporation (XPress9 library) — MIT License */
{xpress9_js}
</script>
<script>/* DataModel Decoder — ABF parser + minimal SQLite reader for V3 .pbix files */
{decoder_js}
</script>"""

    # Check if already embedded (for re-runs)
    if '/* XPress9 WASM Decompressor' in html:
        # Remove old embedded blocks
        html = re.sub(
            r'<script>/\* XPress9 WASM Decompressor.*?</script>\s*<script>/\* DataModel Decoder.*?</script>\s*',
            '', html, flags=re.DOTALL)

    # Insert after the Cytoscape.js </script> tag (line 56 area)
    # Find the first </script> followed by whitespace then <style>
    match = re.search(r'</script>(\s*)<style>', html)
    if match:
        insert_pos = match.start() + len('</script>')
        html = html[:insert_pos] + '\n' + script_block + html[insert_pos:]
    else:
        print("WARNING: Could not find </script> before <style>", file=sys.stderr)

    # Update CSP to allow WASM
    old_csp = "script-src 'unsafe-inline'"
    new_csp = "script-src 'unsafe-inline' 'wasm-unsafe-eval'"
    if old_csp in html and new_csp not in html:
        html = html.replace(old_csp, new_csp)

    # Update Path 3 in parsePbixFile to use extractFromDataModel
    old_path3 = """    // Path 3: If only DataModel (compressed ABF) exists, inform user
    if(!queries.length&&!dmEntry&&!zip.files['DataModelSchema']&&zip.files['DataModel']){
        errors.push('This file uses a compressed DataModel (V3 format). Power Query code is embedded in the compressed model and cannot be extracted in the browser. Try saving as .pbit (File \\u2192 Export \\u2192 Power BI template) for full query extraction.');
    }"""

    new_path3 = """    // Path 3: DataModel (XPress9-compressed ABF) — decompress and extract M code
    if(!queries.length&&zip.files['DataModel']){
        if(typeof extractFromDataModel==='function'){
            try{
                const dmData=await zip.files['DataModel'].async('arraybuffer');
                const dmQueries=await extractFromDataModel(dmData,file.name);
                addUnique(dmQueries);
            }catch(e){errors.push('DataModel: '+e.message);}
        }else{
            errors.push('DataModel decoder not available. This file uses compressed V3 format.');
        }
    }"""

    if old_path3 in html:
        html = html.replace(old_path3, new_path3)
    else:
        print("WARNING: Could not find Path 3 placeholder in HTML. Manual update needed.", file=sys.stderr)

    # Update limitations text
    old_limitation = "Some newer <code>.pbix</code> files embed queries inside a compressed DataModel that cannot be read in the browser. If you see this error, export as <code>.pbit</code> (File &rarr; Export &rarr; Power BI template) for full extraction."
    new_limitation = "V3 <code>.pbix</code> files with compressed DataModel are supported via built-in XPress9 decompression (WASM). Very large files (&gt;100MB) may be slow to decompress."

    if old_limitation in html:
        html = html.replace(old_limitation, new_limitation)

    with open(HTML_PATH, 'w') as f:
        f.write(html)

    print(f"Embedded XPress9 WASM ({len(wasm_b64)} bytes base64) into {HTML_PATH}")
    print(f"Total HTML size: {len(html)} bytes ({len(html)/1024:.1f} KB)")

if __name__ == '__main__':
    main()
