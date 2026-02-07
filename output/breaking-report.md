# Power Query Explorer - Breaking Report (Adversarial Analysis)

**Date:** 2026-02-07
**File analyzed:** `power-query-explorer.html`
**Analyst role:** Breaker (adversarial tester)

---

## Summary

The Power Query Explorer is a single-page HTML application that extracts and visualizes M code from Excel (.xlsx/.xlsm) and Power BI (.pbix) files. The analysis found **6 Critical**, **7 High**, **8 Medium**, and **7 Low** severity issues across XSS, parser correctness, memory safety, error handling, UI behavior, and security domains.

---

## 1. XSS / HTML Injection Vulnerabilities

### 1.1 [CRITICAL] Graph Legend XSS via Filenames

**Location:** Lines 1146-1151 (`renderGraph` function)

```javascript
legend.innerHTML = appState.files.map((file, i) => `
    <div class="legend-item">
        <div class="legend-dot" style="background: ${FILE_COLORS[i % FILE_COLORS.length]}"></div>
        ${file}
    </div>
`).join('');
```

**Problem:** The `file` variable (which is `file.name` from user-dropped files) is interpolated directly into HTML without escaping. A file named `<img src=x onerror=alert(1)>.xlsx` would execute arbitrary JavaScript.

**Reproduction:** Create a file named `"><img src=x onerror=alert(document.cookie)>.xlsx` and drop it into the tool.

**Fix:** Use `escapeHtml(file)` in the template literal.

---

### 1.2 [CRITICAL] Code Panel File Filter XSS via Filenames

**Location:** Lines 1159-1164 (`renderCodePanel` function)

```javascript
filters.innerHTML = appState.files.map((file, i) => `
    <label class="filter-checkbox">
        <input type="checkbox" checked data-file="${file}">
        <span style="color: ${FILE_COLORS[i % FILE_COLORS.length]}">${file}</span>
    </label>
`).join('');
```

**Problem:** The filename is injected raw into both an HTML attribute (`data-file="${file}"`) and into visible HTML text (`${file}`). A filename containing `"` can break out of the attribute. A filename containing `<script>` will execute in the text context.

**Fix:** Use `escapeHtml(file)` for both locations.

---

### 1.3 [CRITICAL] Code Panel File Section Headers XSS

**Location:** Lines 1174-1199 (`renderCodePanel` function)

```javascript
content.innerHTML = appState.files.map((file, fileIndex) => {
    ...
    return `
        <div class="file-section" data-file="${file}">
            ...
            <span class="file-name" ...>${file}</span>
            ...
            <div class="query-block" data-query="${q.name}" data-file="${q.fileName}">
                ...
                <span class="query-name">${q.name}</span>
                ...
                `<span class="query-deps">→ ${q.dependencies.join(', ')}</span>`
```

**Problem:** The filename (`file`), query name (`q.name`), and dependency names (`q.dependencies`) are all injected raw into HTML. The `q.code` is properly escaped via `escapeHtml(q.code)`, but none of the metadata fields are.

**Reproduction:** An attacker could craft a `.xlsx` file where the M code contains `shared <script>alert(1)</script> = ...` -- the query name would be injected as HTML.

**Note:** The `q.name` values come from regex parsing of M code, which extracts `[\w_#]+`, so the character set is limited for names. However, `q.dependencies` values are extracted via a less restrictive regex and filenames are fully user-controlled.

**Fix:** Apply `escapeHtml()` to all interpolated values in HTML templates.

---

### 1.4 [HIGH] Cytoscape Node Labels Could Contain Special Characters

**Location:** Lines 1067-1073

```javascript
const nodes = appState.queries.map(q => ({
    data: {
        id: q.fileName + '::' + q.name,
        label: q.name,
        fileName: q.fileName,
        color: fileColorMap[q.fileName]
    }
}));
```

**Problem:** Cytoscape renders labels as canvas text (not HTML), so this is not directly exploitable for XSS. However, filenames with `::` in them could create node ID collisions, causing data corruption in the graph or edges pointing to wrong nodes.

**Fix:** Use a safer delimiter or hash-based IDs.

---

## 2. Parser Edge Cases

### 2.1 [CRITICAL] `parseMCodeFile` Splits on `shared` Inside Strings and Comments

**Location:** Lines 944-1000

```javascript
const cleanCode = mCode.replace(/section\s+\w+\s*;/gi, '').trim();
const sharedParts = cleanCode.split(/(?=shared\s+)/i);
```

**Problem:** The split on `shared` is a naive text split that does not account for:
- The word `shared` appearing inside M string literals: `"This is a shared resource"`
- The word `shared` appearing inside M comments: `// shared helper used below`
- The word `shared` appearing inside multiline comments: `/* shared */`
- The `section` removal regex also fires inside strings/comments

**Reproduction:** An M file containing:
```
shared MyQuery = let
    description = "This is shared among teams",
    Source = ...
in Source;
```
This would incorrectly split at the `shared` inside the string, creating a broken query parse.

**Impact:** Queries are silently dropped, corrupted, or merged. The user sees incorrect results with no error message.

**Fix:** Implement a proper tokenizer or at minimum strip string literals and comments before splitting.

---

### 2.2 [HIGH] `parseMCodeFile` Section Removal Breaks Multi-Section Files

**Location:** Line 951

```javascript
const cleanCode = mCode.replace(/section\s+\w+\s*;/gi, '').trim();
```

**Problem:** This removes ALL section declarations globally. If an M file has multiple sections (rare but valid), all section headers are stripped and queries from different sections get mixed together with potential name collisions.

---

### 2.3 [HIGH] `findDependencies` Produces Massive False Positives

**Location:** Lines 1002-1028

```javascript
const identifierPattern = /(?:^|[=,(\s])([A-Z_][\w_]*?)(?:\s*[,)\[\.]|\s*$)/gim;
```

**Problem:** This regex matches nearly any capitalized identifier in the M code. Common false positives include:
- Step names within the same query (e.g., `Source`, `Filtered`, `Renamed`)
- String literal contents (e.g., `"Column Name"` where `Column` matches)
- Type annotations (e.g., `Int64.Type`)
- Function names from `#` prefixed library calls

The keyword exclusion list is incomplete -- it misses `Value`, `Expression`, `Splitter`, `Comparer`, `Combiner`, `Lines`, `Replacer`, `Uri`, `Logical`, `Type`, `Action`, `Error`, among many others.

**Impact:** The dependency graph shows many false edges. With enough queries, this makes the graph unreadable.

**Fix:** Cross-reference detected identifiers against actual query names in the file, rather than guessing.

---

### 2.4 [HIGH] `findDependencies` False Negatives for Quoted Identifiers

**Location:** Lines 1002-1028

**Problem:** M code allows query references using `#"Query Name With Spaces"` syntax. The regex `[A-Z_][\w_]*` cannot match these. Any query with spaces, special characters, or starting with lowercase will be missed entirely.

**Reproduction:** M code like `Source = #"My Data Source"` will not detect `My Data Source` as a dependency.

**Impact:** Missing edges in the dependency graph.

---

### 2.5 [MEDIUM] `extractFromDataMashup` Only Finds First ZIP Signature

**Location:** Lines 850-865

```javascript
for (let i = 0; i < bytes.length - 4; i++) {
    ...
    zipStart = i;
    break;
}
```

**Problem:** The DataMashup binary format has a defined header structure before the embedded ZIP. The code does a linear scan for the ZIP magic bytes, which could match a false positive earlier in the binary data if the mashup header happens to contain `PK\x03\x04` bytes. It also only finds the first match.

**Impact:** Corrupted or empty query extraction with no error reported.

---

### 2.6 [MEDIUM] `extractFromCustomXml` False Positive Base64 Matching

**Location:** Lines 900-919

```javascript
const base64Match = xmlContent.match(/[A-Za-z0-9+/=]{100,}/g);
```

**Problem:** This regex matches any 100+ character run of base64-like characters. In XML documents, this could match:
- Long attribute values
- Encoded images or fonts
- Other binary data that is not a DataMashup

Each false match triggers an `atob()` decode and a full `extractFromDataMashup()` parse attempt. With large XML files this could be very slow.

**Impact:** Performance degradation; potential for garbage data being parsed as queries.

---

### 2.7 [MEDIUM] `parseMCodeFile` Name Regex Is Too Restrictive

**Location:** Line 961

```javascript
const nameMatch = trimmed.match(/shared\s+([\w_#]+)\s*=/i);
```

**Problem:** M query names can contain spaces and special characters when quoted with `#"..."` syntax, e.g., `shared #"My Query" = ...`. The regex `[\w_#]+` cannot match these. Only unquoted simple names are detected.

**Impact:** Queries with quoted names are silently dropped.

---

### 2.8 [LOW] Duplicate Queries from Multiple Extraction Methods

**Location:** Lines 786-843 (`parseXlsxFile`)

**Problem:** The code runs up to 4 extraction methods sequentially (DataMashup root, customXml .bin, customXml item*.xml, connections.xml). Deduplication only happens for Method 2 (Section1.m within `extractFromDataMashup`, lines 883-888) and Method 3 (connections.xml, only added if `queries.length === 0`). Methods 0 and 1 can both find the same DataMashup and produce duplicate queries.

**Reproduction:** A PBIX file where `DataMashup` at root and a `customXml/*.bin` both contain the same embedded ZIP.

**Impact:** Duplicate query entries in the UI and graph.

---

## 3. Memory / Performance Issues

### 3.1 [HIGH] No File Size Limits

**Location:** Lines 1349-1379 (`processFiles`)

**Problem:** There is no validation on file size. A 500MB PBIX file will be fully loaded into memory as an ArrayBuffer, then JSZip will decompress it (potentially expanding to gigabytes), then the inner DataMashup ZIP will also be decompressed. This creates multiple copies in memory.

**Reproduction:** Drop a 200MB+ Excel file onto the page.

**Impact:** Browser tab crashes with out-of-memory error. No recovery possible without page reload.

**Fix:** Add a file size limit (e.g., 100MB) with a user warning.

---

### 3.2 [HIGH] Cytoscape Performance with Many Nodes

**Location:** Lines 1098-1142 (`renderGraph`)

**Problem:** The cytoscape COSE layout is O(n^2) in the number of nodes. With 500+ queries, the layout computation can freeze the browser tab for 10-30+ seconds. The `animate: false` flag helps but the computation itself is still blocking the main thread.

Additionally, each node gets a text label which Cytoscape must render on canvas, further degrading performance.

**Reproduction:** Create or combine files with 500+ Power Query queries.

**Impact:** Browser appears frozen. User may kill the tab thinking the application crashed.

**Fix:** Add a node count threshold. Above it, switch to a simpler/faster layout (e.g., `grid` or `concentric`), or paginate, or warn the user.

---

### 3.3 [MEDIUM] Linear ZIP Signature Scan on Large Binaries

**Location:** Lines 853-861

```javascript
for (let i = 0; i < bytes.length - 4; i++) {
    if (bytes[i] === zipSignature[0] && ...
```

**Problem:** This byte-by-byte scan over the entire DataMashup binary is O(n). For a 50MB DataMashup, this is 50 million iterations in a tight loop on the main thread.

**Impact:** UI freeze during processing.

---

### 3.4 [LOW] No Limit on Number of Files

**Problem:** Dropping a folder with hundreds of Excel files will process them all sequentially. Each file involves ZIP decompression, multiple extraction methods, and DOM rendering.

**Fix:** Add a file count limit or process in batches with progress reporting.

---

## 4. Error Handling Gaps

### 4.1 [CRITICAL] Silent Failures with console.log Only

**Location:** Multiple catch blocks (lines 798, 810, 821, 838, 891, 1363)

**Problem:** Every catch block in the parsing pipeline only does `console.log(...)`. The user sees no indication that parsing failed. If all methods fail silently, the user gets "No Power Queries found" with no explanation.

**Reproduction:** Drop a corrupted .xlsx file -- the user sees "No Power Queries found" even though the real problem is a parse error.

**Fix:** Collect errors and display them in the UI. At minimum, show a warning toast listing which files had parse errors.

---

### 4.2 [CRITICAL] No CDN Failure Handling

**Location:** Lines 7-8

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/cytoscape/3.28.1/cytoscape.min.js"></script>
```

**Problem:** If the CDN is down, blocked by corporate firewall/proxy, or the user is offline:
- `JSZip` will be undefined -- any file drop will throw `JSZip is not defined` with no user-facing error
- `cytoscape` will be undefined -- graph rendering will throw with no user-facing error

There is no check for library availability and no fallback.

**Reproduction:** Open the HTML file while disconnected from the internet, or in a corporate network that blocks CDN domains.

**Impact:** Complete application failure with cryptic console errors only.

**Fix:** Add `<script>` existence checks on page load. Show a clear error message if libraries are missing. Consider bundling the libraries inline or using SRI hashes.

---

### 4.3 [MEDIUM] No Way to Reset/Retry After Failure

**Location:** Lines 1349-1379

**Problem:** After `processFiles` runs:
- The drop zone is hidden (`display: none`)
- If all files fail to parse, the drop zone is shown again -- but if some files parse with zero queries each, `appState.files` is empty and the user sees a blank main content area with no way to go back

Even when files load successfully, there is no "clear" or "load more files" button. The user must reload the page.

**Fix:** Add a reset/clear button. Ensure the drop zone is always accessible.

---

### 4.4 [LOW] `processFiles` Does Not Handle Empty File List Gracefully

**Location:** Line 1370

```javascript
if (appState.queries.length > 0) { ... } else {
    document.getElementById('dropZone').style.display = 'block';
    showToast('No Power Queries found in the selected files');
}
```

**Problem:** If files parse but produce `appState.files = ['file1.xlsx']` and `appState.queries = []`, the stats show "Files: 0" because `appState.files` only gets populated when queries are found (line 1359-1361). This is confusing.

---

## 5. UI Bugs

### 5.1 [HIGH] Copy Button Loses SVG Icon After First Click

**Location:** Lines 1227-1231

```javascript
function updateSelectedCount() {
    const checkboxes = document.querySelectorAll('.query-checkbox:checked');
    const btn = document.getElementById('copyBtn');
    btn.textContent = `Copy Selected (${checkboxes.length})`;
}
```

**Problem:** `btn.textContent = ...` replaces ALL child nodes of the button, including the SVG icon. After the first call to `updateSelectedCount()` (which happens immediately in `renderCodePanel`), the copy button permanently loses its icon.

**Reproduction:** Load any file -- the copy button will immediately show text-only without the clipboard icon.

**Fix:** Use a dedicated `<span>` for the text and only update that span's textContent.

---

### 5.2 [HIGH] `readDirectory` Only Reads First Batch of Entries

**Location:** Lines 1332-1347

```javascript
async function readDirectory(dirEntry, files) {
    const reader = dirEntry.createReader();
    return new Promise((resolve) => {
        reader.readEntries(async (entries) => {
            ...
            resolve();
        });
    });
}
```

**Problem:** The `FileSystemDirectoryReader.readEntries()` API does NOT guarantee returning all entries in a single call. Per the spec, it may return a partial batch (typically up to ~100 entries), and you must call `readEntries()` repeatedly until it returns an empty array.

**Reproduction:** Drop a folder containing 150+ files.

**Impact:** Files beyond the first batch are silently ignored.

**Fix:** Loop on `readEntries()` until an empty array is returned.

---

### 5.3 [MEDIUM] Drop Zone Click Handler Conflicts with Label Click

**Location:** Lines 1323, 637-641

```javascript
dropZone.addEventListener('click', () => fileInput.click());
```

```html
<label class="browse-btn">
    Browse Files
    <input type="file" id="fileInput" multiple accept=".xlsx,.xlsm,.pbix">
</label>
```

**Problem:** Clicking the "Browse Files" label already triggers the file input (because the `<input>` is a child of the `<label>`). But the click event also bubbles up to the drop zone, which calls `fileInput.click()` a second time. This can cause the file dialog to open and immediately close, or open twice.

**Browser-dependent behavior:** Some browsers show the dialog twice; others show it once then immediately cancel.

**Fix:** In the click handler, check if the click target is the label or input and skip the programmatic click.

---

### 5.4 [MEDIUM] File Filter Checkboxes Don't Toggle Query Checkboxes

**Location:** Lines 1204-1213

```javascript
cb.addEventListener('change', (e) => {
    const file = e.target.dataset.file;
    const section = content.querySelector(`.file-section[data-file="${file}"]`);
    if (section) {
        section.style.display = e.target.checked ? 'block' : 'none';
    }
    updateSelectedCount();
});
```

**Problem:** Unchecking a file filter hides the file section via CSS (`display: none`), but the individual query checkboxes inside remain checked. When the user clicks "Copy Selected", queries from hidden files are still included because `document.querySelectorAll('.query-checkbox:checked')` finds them.

**Impact:** User thinks they deselected a file's queries, but they are still copied.

**Fix:** Either uncheck the query checkboxes when hiding a file, or exclude hidden queries from the selection logic.

---

### 5.5 [MEDIUM] Select All Toggle State Is Fragile

**Location:** Lines 1402-1407

```javascript
document.getElementById('selectAllBtn').addEventListener('click', () => {
    const checkboxes = document.querySelectorAll('.query-checkbox');
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    checkboxes.forEach(cb => cb.checked = !allChecked);
    updateSelectedCount();
});
```

**Problem:** The toggle logic checks if ALL checkboxes are checked. If even one checkbox is unchecked, clicking "Select All" will check them all. But the button text always says "Select All" -- it never changes to "Deselect All". There is no visual indication of state.

Also, if the user has some queries selected and clicks "Select All", it selects all. Clicking again deselects all. There is no way to return to the previous partial selection.

---

### 5.6 [LOW] `.xlsm` Accepted by Drop but Not Shown in UI Text

**Location:** Line 636 vs Line 639

```html
<h2>Drop Excel or Power BI files here</h2>
<p>Supports .xlsx and .pbix files with Power Query</p>
...
<input type="file" id="fileInput" multiple accept=".xlsx,.xlsm,.pbix">
```

**Problem:** The `accept` attribute includes `.xlsm` but the descriptive text only mentions `.xlsx` and `.pbix`. Users with `.xlsm` files may not know they can use the tool.

---

### 5.7 [LOW] No File Type Validation Before Processing

**Location:** Lines 1349-1379

**Problem:** Files from the file input are filtered by extension (line 1326), but within `processFiles` there is no validation. If a non-ZIP file somehow gets passed (e.g., a renamed `.xlsx`), JSZip will throw a cryptic error caught only by console.log.

---

## 6. Logic Bugs

### 6.1 [MEDIUM] CSS Selector Injection via Filenames in querySelector

**Location:** Lines 1207, and similar

```javascript
const section = content.querySelector(`.file-section[data-file="${file}"]`);
```

**Problem:** The `file` variable is a filename that could contain characters that break CSS selectors: `"`, `]`, `\`. A filename like `test"].xlsx` would cause `querySelector` to throw an exception or match the wrong element.

**Fix:** Use `CSS.escape()` for dynamic selector values, or use `querySelectorAll` + filter by dataset property.

---

### 6.2 [MEDIUM] Duplicate File Names Cause Data Corruption

**Location:** Lines 1356-1365

**Problem:** If two different files have the same name (e.g., from different folders), `appState.files.push(file.name)` pushes duplicates. The `fileColorMap` (line 1063) uses filename as key, so both files get the same color. The `byFile` grouping (line 1168) merges their queries. The CSS selectors for file sections break because multiple elements share the same `data-file` attribute.

**Impact:** Queries from different files are merged together. File filter checkboxes may only affect one of the file sections.

**Fix:** Use full paths or add disambiguation suffixes for duplicate names.

---

### 6.3 [LOW] Token Estimate Is Very Rough

**Location:** Lines 1044-1045

```javascript
const totalChars = appState.queries.reduce((sum, q) => sum + q.code.length + q.name.length + 20, 0);
const estimatedTokens = Math.round(totalChars / 4);
```

**Problem:** The estimate uses chars/4 which is a rough average for English text. M code has different tokenization characteristics (many symbols, CamelCase identifiers). The estimate also doesn't account for the `// === filename ===` and `// --- Query: name ---` headers added by `getSelectedCode()`.

---

## 7. Security Concerns

### 7.1 [HIGH] "100% Local" Claim Is Misleading

**Location:** Line 623

```html
100% local — files never leave your browser
```

**Problem:** While file data does stay in the browser, the page loads 4 external resources:
1. `cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js`
2. `cdnjs.cloudflare.com/ajax/libs/cytoscape/3.28.1/cytoscape.min.js`
3. `fonts.googleapis.com` (CSS)
4. `fonts.gstatic.com` (font files)

These CDN requests:
- Leak the user's IP address and the fact that they're using this tool
- Could theoretically be compromised (supply chain attack) to exfiltrate data
- Are blocked by many corporate firewalls, breaking the app entirely

There are no Subresource Integrity (SRI) hashes on the script tags.

**Fix:** Add SRI hashes at minimum. Consider bundling libraries inline for true offline/local operation. Update the claim to be accurate, e.g., "Files stay in your browser."

---

### 7.2 [MEDIUM] No Content Security Policy

**Problem:** The HTML file has no CSP meta tag. If any XSS vulnerability is exploited (see Section 1), there are no restrictions on what the injected script can do -- it can load external scripts, exfiltrate data, etc.

**Fix:** Add a strict CSP meta tag, e.g.:
```html
<meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'self' cdnjs.cloudflare.com 'unsafe-inline'; style-src 'self' fonts.googleapis.com 'unsafe-inline'; font-src fonts.gstatic.com;">
```

---

## Severity Summary

| Severity | Count | Key Issues |
|----------|-------|------------|
| **Critical** | 6 | XSS via filenames (3), `shared` inside strings breaks parser, silent failures hide errors, CDN failure = total app failure |
| **High** | 7 | Copy button icon loss, directory batch limit, query dependency false positives/negatives, no file size limits, Cytoscape freeze, "100% local" misleading |
| **Medium** | 8 | Filter checkboxes don't uncheck queries, drop zone double-click, CSS selector injection, duplicate filenames, no CSP, ZIP scan performance, base64 false matches, quoted name support |
| **Low** | 7 | Token estimate rough, xlsm not in UI text, no file type validation, no file count limit, Select All text never changes, duplicate queries across methods, no-queries edge case |

---

## Recommended Priority for Fixes

1. **XSS in templates** -- Apply `escapeHtml()` to ALL user-derived values in HTML templates (filenames, query names, dependency names). This is a one-pass fix.
2. **CDN resilience** -- Add SRI hashes and a library-availability check on page load.
3. **Copy button icon loss** -- Restructure `updateSelectedCount` to only update a text span.
4. **Silent errors** -- Accumulate parse errors and display them in the UI.
5. **`readDirectory` batch bug** -- Loop until empty array.
6. **Parser `shared` splitting** -- At minimum, strip string literals before splitting.
7. **File size/count limits** -- Add guards to prevent memory exhaustion.
8. **File filter / query checkbox sync** -- Ensure unchecking a file excludes its queries from copy.
