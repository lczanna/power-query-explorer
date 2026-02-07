# Power Query Explorer -- Devil's Advocate Design Review

**Reviewer:** devils-advocate agent
**File under review:** `power-query-explorer.html` (1421 lines)
**Date:** 2026-02-07

---

## 1. Single HTML File Architecture

**Verdict: Acceptable now, but already straining.**

At 1421 lines the file is already at the upper boundary of comfortable single-file maintenance. The code mixes three distinct concerns (CSS ~600 lines, HTML ~100 lines, JS ~700 lines) with no separation. Every edit to the dependency algorithm forces the developer to scroll past 600 lines of CSS.

**Where it breaks:**
- Beyond ~2000 lines, editors start losing track of structure. No LSP language server will give you good IntelliSense inside an inline `<script>` tag because the context is ambiguous (HTML? JS? CSS?).
- No minification means users download every comment and whitespace character. Currently ~40KB raw; acceptable, but the CDN dependencies (see below) dwarf this.
- No code splitting means the entire Cytoscape styling block, all prompt templates, and the parsing code load even if the user never uses them.

**Counterargument the design gets right:** Zero build step is genuinely valuable for a tool meant to be shared as a single file. A Power BI developer receiving this file can double-click and go. That experience would be destroyed by requiring `npm install`.

**Recommendation:** Keep the single-file approach but consider a build step that *produces* the single file from separate source files (e.g., a simple `cat` script or `esbuild --bundle`). Develop in multiple files, ship as one.

---

## 2. CDN Dependencies -- The "100% Local" Lie

**Verdict: This is misleading and should be fixed.**

Line 623-624 prominently displays:
```
100% local -- files never leave your browser
```

But lines 7-11 load three external resources:
- `cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js` (~120KB)
- `cdnjs.cloudflare.com/ajax/libs/cytoscape/3.28.1/cytoscape.min.js` (~700KB)
- `fonts.googleapis.com` (Google Fonts: JetBrains Mono + Plus Jakarta Sans)

**Problems:**
1. **Privacy violation of stated promise.** Every time a user opens this file, their browser makes requests to Cloudflare and Google, leaking their IP address, user agent, and referrer. For a tool explicitly marketed as privacy-first, this is deceptive.
2. **Offline breakage.** If the user is on a restricted network (common in enterprises where Power BI is used), the CDN scripts will fail and the entire app becomes non-functional. No fallback is provided.
3. **Supply chain risk.** CDN-hosted scripts can theoretically be compromised. For a tool that processes potentially sensitive business data, this matters.

**The Google Fonts issue specifically:** Google Fonts requests send the user's IP to Google. In GDPR jurisdictions, courts have ruled that loading Google Fonts without consent violates privacy regulations (LG Munchen, Jan 2022). This is not hypothetical.

**Recommendation:**
- Inline JSZip and Cytoscape.js directly into the HTML file using a build step. Yes, this makes the file ~900KB, but it becomes *genuinely* self-contained.
- Replace Google Fonts with system font stacks: `-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif` for body text and `'Cascadia Code', 'Fira Code', 'Consolas', monospace` for code. These are fine. No user will notice the difference.
- Or: Use `@font-face` with base64-encoded WOFF2 subsets if the specific fonts are truly important.

---

## 3. Cytoscape.js for the Dependency Graph

**Verdict: Overkill. Significantly overkill.**

Cytoscape.js weighs ~700KB minified. It is a full-featured graph theory library designed for biological network visualization with hundreds of layout algorithms, graph analysis functions, and rendering modes. Power Query Explorer uses exactly one feature: draw circles with lines between them (lines 1098-1142).

**What PQ Explorer actually needs:**
- Nodes (circles with labels)
- Directed edges (lines with arrows)
- One layout algorithm (force-directed)
- Pan and zoom (nice to have)
- Click to select (nice to have)

**Alternatives and their weights:**
| Library | Size (min) | Sufficient? |
|---------|-----------|-------------|
| Cytoscape.js | ~700KB | Massive overkill |
| D3.js (force module only) | ~30KB | Yes, with more code |
| vis-network | ~200KB | Yes |
| Custom SVG | 0KB | Yes, for <50 nodes |
| Mermaid.js | ~1.5MB | Even worse |
| Plain Canvas | 0KB | Yes, with more code |

For the typical Power Query workbook (5-30 queries), a hand-rolled SVG solution in ~150 lines of JS would be entirely sufficient and add zero dependency weight. A simple force-directed layout algorithm is ~50 lines.

**Counterargument:** Cytoscape's `cose` layout (line 1136) handles edge cases well -- disconnected components, large graphs, overlapping labels. Reimplementing this correctly is non-trivial.

**Recommendation:** If keeping a library, switch to the D3 force module (~30KB). If the goal is truly self-contained, implement a basic SVG renderer. Most Power Query files have <30 queries; brute-force positioning works fine at that scale.

---

## 4. LLM Prompt Templates

**Verdict: Useful concept, poor execution.**

The prompt templates (lines 731-769) are static strings that get prepended to the copied code. The user clicks "Analyze dependencies", the app copies `[prompt text] + [all query code]` to clipboard, and the user pastes into ChatGPT/Claude.

**What's wrong:**
1. **The templates are generic.** "Identify potential errors" and "Find performance optimizations" could apply to any code in any language. There is nothing Power Query-specific about them. A Power Query expert would want prompts about query folding, gateway data source compatibility, incremental refresh patterns, or parameterized data sources.
2. **No context is injected.** The templates don't include the dependency map, file names, or stats. The LLM has to re-derive structure that the tool already computed.
3. **No feedback loop.** The user copies, pastes into an LLM, reads the response, and... then what? There is no way to incorporate LLM suggestions back into the tool.
4. **The UX is confusing.** On line 715, `data-prompt="analyze"` copies to clipboard. But there is no indication that this is what will happen. A user might expect a modal or dropdown, not a silent clipboard write.

**Recommendation:**
- Make templates Power Query-specific. Include prompts about query folding, M language best practices, data source isolation, etc.
- Pre-inject the dependency graph and stats into the prompt. The LLM should know "Query X depends on Y and Z" without having to infer it from code.
- Add a visible text area showing the assembled prompt before copying, so users can edit it.
- Consider: Is this even the right feature? If the primary workflow is "extract M code, paste into LLM", maybe the tool should just optimize the *copy* experience and drop the prompt pretense.

---

## 5. Dark Theme Only

**Verdict: Accessibility problem.**

The entire color scheme is hardcoded dark (lines 13-28). There is no theme toggle, no `prefers-color-scheme` media query, and no light mode fallback.

**Who is excluded:**
- Users with certain visual impairments who need high-contrast light modes
- Users working in bright environments (office fluorescent lighting)
- Users who simply prefer light themes (which is still the majority in enterprise tools)
- Anyone who wants to print the output (dark theme prints as a wall of dark ink, or browsers strip backgrounds leaving invisible text)

**The print issue is especially bad.** A common workflow would be: extract queries, print the code view for a code review meeting. With the current CSS, this produces garbage output.

**Recommendation:**
- Add a `@media (prefers-color-scheme: light)` block that inverts the palette. This is ~30 lines of CSS.
- Add a `@media print` block that forces light backgrounds and dark text.
- Optionally add a theme toggle button in the header.

---

## 6. No Persistence

**Verdict: Acceptable for v1, but the lack of export is not.**

The app has zero persistence. Reload the page and everything vanishes. There is no localStorage, no IndexedDB, no export.

**Why persistence might not matter:** The input is always a file. The user can re-drop the file. Re-parsing a 50KB Excel file takes <1 second. The app is a *viewer*, not an editor.

**Why export does matter:** The user might want to:
- Save the extracted M code as `.m` files (without having to copy/paste one at a time)
- Export the dependency graph as JSON or SVG for documentation
- Export query metadata as CSV for reporting
- Share results with a colleague who does not have the original file

The "Copy Selected" button (line 691-698) is the only export mechanism, and it only produces unstructured text.

**Recommendation:**
- Skip localStorage persistence (adds complexity, little value).
- Add "Download as .m files" (one file per query or one merged file).
- Add "Export dependency graph as JSON".
- Add "Export as SVG" for the graph panel.

---

## 7. Dependency Detection Algorithm

**Verdict: Fundamentally fragile. Good enough for 70% of cases, silently wrong for 30%.**

The `findDependencies` function (lines 1002-1028) uses a regex to find identifiers and filters out a hardcoded keyword list. This approach has multiple structural problems:

**Problem 1: The regex is too broad (line 1007).**
```js
const identifierPattern = /(?:^|[=,(\s])([A-Z_][\w_]*?)(?:\s*[,)\[\.]|\s*$)/gim;
```
This matches any uppercase-starting identifier, but M language has many standard library functions that start with uppercase letters. The keyword list (lines 1013-1018) covers ~40 common ones, but the M standard library has hundreds. `Expression.Evaluate`, `Comparer.OrdinalIgnoreCase`, `MissingField.UseNull`, `JoinKind.Inner` -- none of these are in the filter list, and they will all be detected as "dependencies."

**Problem 2: The keyword list is incomplete and ad hoc (lines 1013-1018).**
`Source` is listed as a keyword (line 1016), but `Source` is not a keyword -- it is a conventional step name. If a user names a query "Source", it will never be detected as a dependency. Conversely, `Int64`, `Percentage`, and `Currency` (line 1018) are type names, not keywords. This conflates M language keywords, standard library namespaces, type names, and conventional step names into one blocklist.

**Problem 3: String literals are not excluded.**
If M code contains `"Hello World"` or a connection string like `"Server=MyServer;Database=Sales"`, the regex will scan inside it and may pick up false positives. Any identifier-like string inside a literal will be flagged as a dependency.

**Problem 4: Comments are not excluded.**
M supports `//` and `/* */` comments. The regex scans these as code.

**Problem 5: Case sensitivity confusion.**
The regex uses the `i` flag (case-insensitive), but the keyword list comparison uses `Array.includes()` which is case-sensitive. This means `table.AddColumn` will match `Table` as a dependency, but `Table` is in the filter list and will be excluded -- yet `table` (lowercase) will NOT be excluded because `table` !== `Table`.

**Recommendation:**
- At minimum, strip string literals and comments before scanning.
- Replace the blocklist approach with an allowlist: only flag identifiers that match known query names from the same file set.
- For a proper solution, implement a basic M language tokenizer that distinguishes identifiers, keywords, string literals, comments, and step names. This is ~200 lines of code for an 80% solution.

---

## 8. Error Philosophy

**Verdict: Terrible. Silent failures are user-hostile.**

The app has 8 `try/catch` blocks (lines 793-799, 806-811, 819-824, 831-839, 867-892, 903-917, 1357-1365). Every single one catches the error, logs to console, and continues silently.

**What the user sees when parsing fails:** Nothing. If all parsing methods fail for a file, line 1370 checks `appState.queries.length > 0` and if it is zero, shows the drop zone again with a toast: "No Power Queries found in the selected files" (line 1377). This message is indistinguishable from "the file genuinely has no queries."

**Concrete failure scenarios and their (lack of) UX:**
1. User drops a password-protected PBIX file. JSZip fails silently. User sees "No Power Queries found." User thinks the file has no queries.
2. The DataMashup binary format has changed in a new Excel version. Inner ZIP parse fails. User sees nothing.
3. A corrupted file causes an exception in `parseMCodeFile`. Partial results are returned with no indication that some queries were lost.

**Recommendation:**
- Collect errors into an `appState.errors` array.
- After processing, if errors exist, display a warning banner: "Processed 3 files. 1 file had parsing errors. Click for details."
- In the details, show which file failed, which parsing method was attempted, and what the error was.
- For the "no queries found" case, distinguish between "we parsed the file successfully and it has no queries" vs. "we failed to parse the file."

---

## 9. Token Estimation

**Verdict: Misleading. The `chars/4` rule is wrong for code.**

Line 1045:
```js
const estimatedTokens = Math.round(totalChars / 4);
```

The "1 token ~= 4 characters" heuristic comes from English prose. For code -- especially Power Query M code with camelCase identifiers, special characters, and structured syntax -- the ratio is closer to 1 token per 2.5-3 characters. This means the displayed estimate consistently *underestimates* actual token usage by 25-40%.

**Why it matters:** The stated purpose of this estimate is to help users judge whether the code fits in an LLM context window. An underestimate could lead users to paste 150K tokens of code into a 128K context window, causing silent truncation and bad LLM results.

**Recommendation:**
- Use a more conservative ratio: `chars / 3` for code.
- Better: use a proper tokenizer (tiktoken-js or equivalent), but this adds dependency weight.
- Best: label the estimate honestly. Instead of "~12,000 tokens", show "~12,000 chars (~3K-4K tokens)" to communicate the uncertainty.
- Or drop it entirely and show character count, which is a fact rather than an estimate.

---

## 10. Missing Features Users Would Expect

**Search.** There is no way to search across queries. With 30+ queries, finding the one that references "SalesData" requires manual scrolling. A simple text search box filtering the code panel would take ~20 lines of JS.

**Query-level graph interaction.** Clicking a node in the dependency graph does nothing useful. It should highlight the node's dependencies, scroll to that query in the code panel, or show a tooltip with the code preview.

**Multiple file accumulation.** Dropping new files replaces the previous results (line 1353-1354 clears state). Users working with related Excel files would want to accumulate queries across multiple drop operations.

**No file type indication in the UI.** The stats bar shows "Files: 2" but does not show which are XLSX vs PBIX. Given these have different parsing paths and different reliability levels, this matters.

**Keyboard shortcuts.** No keyboard shortcuts for copy (Ctrl+Shift+C?), select all, or tab switching. Power users will want these.

**Responsive behavior on small screens.** The `@media (max-width: 768px)` block (lines 588-606) only adjusts padding and header layout. The code panel, tabs, and graph have no mobile adaptations. The graph is practically unusable on a phone.

**Drag-and-drop feedback.** When a file is being parsed, the loading spinner (line 643-646) shows "Processing locally..." but gives no progress indication. For a folder with 20 files, the user sees a spinner with no idea how long to wait.

---

## Summary: Priority Rankings

| Issue | Severity | Effort to Fix | Priority |
|-------|----------|---------------|----------|
| CDN deps / privacy lie | High | Medium (build step) | **P0** |
| Silent error handling | High | Low (~30 min) | **P0** |
| Dependency detection bugs | Medium | Medium (~2 hours) | **P1** |
| No print/light theme | Medium | Low (~1 hour) | **P1** |
| Search functionality | Medium | Low (~30 min) | **P1** |
| Token estimation accuracy | Low | Low (~10 min) | **P2** |
| Cytoscape.js weight | Low | High (rewrite) | **P2** |
| Export options | Medium | Medium | **P2** |
| LLM prompt improvements | Low | Low | **P3** |
| Single-file architecture | Low | Medium | **P3** |

The app has a solid concept and clean visual design. The critical fixes are: (1) make the privacy claim honest by inlining or removing CDN deps, (2) stop swallowing errors silently, and (3) improve dependency detection to exclude string literals and comments at minimum. Everything else is polish.
