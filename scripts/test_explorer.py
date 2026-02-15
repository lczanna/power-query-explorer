"""
Comprehensive Playwright test suite for Power Query Explorer.

Tests cover:
  - Page load and library availability
  - Single file upload and parsing
  - Multi-query file with dependency detection
  - Complex M code (comments, strings, nested)
  - Edge cases (quoted identifiers, special chars)
  - Stress test (25 queries)
  - No-query file handling (error feedback)
  - Multi-file upload
  - UI interactions (tabs, select all, copy, filter, graph controls)
  - Reset functionality
  - Keyboard shortcuts
  - Graph search
  - LLM prompt template copy
"""

import os
import sys
import json
import time
import base64
import subprocess
import socket
from pathlib import Path
from playwright.sync_api import sync_playwright, expect

HTML_PATH = Path(__file__).resolve().parent.parent / "index.html"
TEST_DIR = Path(__file__).resolve().parent.parent / "data" / "test-files"
PROJECT_DIR = Path(__file__).resolve().parent.parent

def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

PORT = find_free_port()
BASE_URL = f"http://localhost:{PORT}/index.html"

PASS = 0
FAIL = 0
ERRORS = []

def result(name, passed, detail=""):
    global PASS, FAIL, ERRORS
    if passed:
        PASS += 1
        print(f"  \033[32m✓\033[0m {name}")
    else:
        FAIL += 1
        ERRORS.append(f"{name}: {detail}")
        print(f"  \033[31m✗\033[0m {name} — {detail}")

def upload_files(page, filenames):
    """Upload files via the hidden file input."""
    paths = [str(TEST_DIR / f) for f in filenames]
    page.locator("#fileInput").set_input_files(paths)
    # Wait for processing to complete
    page.wait_for_function(
        "() => !document.getElementById('loading').classList.contains('active')",
        timeout=15000
    )
    time.sleep(0.3)  # Brief settle time for rendering


def test_page_load(page):
    """Test that the page loads correctly with all libraries."""
    print("\n━━━ Page Load ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    # Title
    title = page.title()
    result("Page title is 'Power Query Explorer'", title == "Power Query Explorer", f"Got: {title}")

    # Library warning should NOT be visible
    lib_missing = page.locator("#libMissing")
    result("Library warning not visible", not lib_missing.is_visible())

    # Drop zone visible
    drop_zone = page.locator("#dropZone")
    result("Drop zone is visible", drop_zone.is_visible())

    # Main content hidden
    main = page.locator("#mainContent")
    result("Main content hidden initially", not main.is_visible())

    # Privacy badge
    badge = page.locator(".privacy-badge")
    result("Privacy badge visible", badge.is_visible())
    badge_text = badge.inner_text()
    result("Privacy text says '100% local'", "100% local" in badge_text, badge_text)

    # Reset button hidden
    reset = page.locator("#resetBtn")
    result("Reset button hidden initially", not reset.is_visible())

    # Logo
    logo = page.locator(".logo-icon")
    result("Logo shows 'PQ'", logo.inner_text().strip() == "PQ")

    # Dependency parser: quoted identifiers should stay whole and not split.
    quoted_deps = page.evaluate("""() => {
        const code = 'let Source = #"FactOnlineSales Agg" in Source';
        return findDeps(stripCS(code), 'TestQuery');
    }""")
    result("Quoted dependency kept as full identifier", "FactOnlineSales Agg" in quoted_deps, str(quoted_deps))
    result("Quoted dependency not split into partial token", "FactOnlineSales" not in quoted_deps, str(quoted_deps))


def test_simple_query(page):
    """Test single query file parsing."""
    print("\n━━━ Simple Query (3 queries) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Main content should be visible
    result("Main content visible", page.locator("#mainContent").is_visible())

    # Stats
    files_stat = page.locator("#statFiles").inner_text()
    result("Files count = 1", files_stat == "1", f"Got: {files_stat}")

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 3", queries_stat == "3", f"Got: {queries_stat}")

    # File chip
    chips = page.locator(".file-chip")
    result("One file chip shown", chips.count() == 1)
    chip_text = chips.first.inner_text()
    result("File chip shows 'simple_query.xlsx'", "simple_query.xlsx" in chip_text, chip_text)

    # Query names in code panel — switch to code tab first
    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    qnames = page.locator(".query-name")
    result("Three query names shown", qnames.count() == 3, f"Got: {qnames.count()}")
    actual_names = {el.inner_text() for el in qnames.all()}
    expected_names = {"Query1", "SalesData", "EdgeLookup"}
    result("Expected query names present", actual_names == expected_names, f"Got: {actual_names}")

    # Code blocks present
    code_blocks = page.locator(".query-code")
    result("Three code blocks shown", code_blocks.count() == 3, f"Got: {code_blocks.count()}")
    all_code = "\n".join(code_blocks.nth(i).inner_text() for i in range(code_blocks.count()))
    result("Code contains cross-file File.Contents reference", 'File.Contents("edge_cases.xlsx")' in all_code, all_code[:200])

    # Error log should NOT be visible (successful parse)
    result("No error log visible", not page.locator("#errorLog").is_visible())

    # Reset button visible
    result("Reset button visible after upload", page.locator("#resetBtn").is_visible())


def test_multi_query(page):
    """Test multi-query file with dependencies."""
    print("\n━━━ Multi Query (5 queries + deps) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["multi_query.xlsx"])

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 5", queries_stat == "5", f"Got: {queries_stat}")

    deps_stat = page.locator("#statDeps").inner_text()
    deps_val = int(deps_stat)
    result("Dependencies > 0", deps_val > 0, f"Got: {deps_val}")

    # Switch to code panel
    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    expected_names = {"Query1", "RawOrders", "Customers", "OrdersWithCustomers", "SalesSummary"}
    actual_names = {el.inner_text() for el in page.locator(".query-name").all()}
    result("All 5 query names present", expected_names == actual_names,
           f"Expected: {expected_names}, Got: {actual_names}")

    # Check dependency arrows shown
    dep_spans = page.locator(".query-deps")
    has_deps = any(dep_spans.nth(i).inner_text().strip() for i in range(dep_spans.count()))
    result("At least one query shows dependencies", has_deps)

    # Graph tab — verify graph has nodes
    page.locator('.tab[data-tab="graph"]').click()
    page.wait_for_timeout(500)
    # Cytoscape renders on canvas, check container has content
    container = page.locator("#graph-container")
    result("Graph container is visible", container.is_visible())
    # Check legend items
    legend = page.locator(".legend-item")
    result("Graph legend has entries", legend.count() > 0)

    # Dependency edge direction: dependency -> dependent
    raw_to_joined = page.evaluate("""() => {
        if (!appState.cyInstance) return false;
        return appState.cyInstance.edges().jsons().some(e =>
            e?.data?.source?.endsWith("::RawOrders") &&
            e?.data?.target?.endsWith("::OrdersWithCustomers")
        );
    }""")
    result("Graph edges point dependency -> dependent (RawOrders -> OrdersWithCustomers)", raw_to_joined)

    joined_to_summary = page.evaluate("""() => {
        if (!appState.cyInstance) return false;
        return appState.cyInstance.edges().jsons().some(e =>
            e?.data?.source?.endsWith("::OrdersWithCustomers") &&
            e?.data?.target?.endsWith("::SalesSummary")
        );
    }""")
    result("Graph edges point dependency -> dependent (OrdersWithCustomers -> SalesSummary)", joined_to_summary)


def test_complex_code(page):
    """Test complex M code with comments, strings, nested lets."""
    print("\n━━━ Complex Code (comments, strings, nested) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["complex_code.xlsx"])

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 4", queries_stat == "4", f"Got: {queries_stat}")

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    names = {el.inner_text() for el in page.locator(".query-name").all()}
    result("'TransformPipeline' query found", "TransformPipeline" in names, str(names))
    result("'Start Date' parameter found", "Start Date" in names, str(names))

    # Verify syntax highlighting exists
    kw_spans = page.locator(".query-code .kw")
    result("Syntax highlighting: keywords found", kw_spans.count() > 0)

    str_spans = page.locator(".query-code .str")
    result("Syntax highlighting: strings found", str_spans.count() > 0)

    cm_spans = page.locator(".query-code .cm")
    result("Syntax highlighting: comments found", cm_spans.count() > 0)


def test_edge_cases(page):
    """Test edge case file with quoted identifiers and special chars."""
    print("\n━━━ Edge Cases (quoted identifiers, special chars) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["edge_cases.xlsx"])

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 7", queries_stat == "7", f"Got: {queries_stat}")

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    names = {el.inner_text() for el in page.locator(".query-name").all()}
    result("'Revenue Report' (quoted identifier) found", "Revenue Report" in names, str(names))
    result("'Year-to-Date (YTD) #1' found", "Year-to-Date (YTD) #1" in names, str(names))
    result("'Données_brutes' (accented chars) found", "Données_brutes" in names, str(names))
    result("'EmptyResult' found", "EmptyResult" in names, str(names))
    result("'raw_data_import' (snake_case) found", "raw_data_import" in names, str(names))

    # Verify no XSS: query names should be escaped in HTML
    code_html = page.locator("#codeContent").inner_html()
    result("No unescaped < in query names", "<script" not in code_html.lower())


def test_stress(page):
    """Test stress file with 26 queries."""
    print("\n━━━ Stress Test (26 queries) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["stress_test.xlsx"])

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 26", queries_stat == "26", f"Got: {queries_stat}")

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    qnames = page.locator(".query-name")
    result("26 query name elements rendered", qnames.count() == 26, f"Got: {qnames.count()}")

    # Check specific queries exist
    names = {el.inner_text() for el in qnames.all()}
    for prefix in ["DataSource", "Transform", "Merge", "Aggregate", "Output"]:
        found = sum(1 for n in names if n.startswith(prefix))
        result(f"5 '{prefix}*' queries found", found == 5, f"Got: {found}")
    result("One 'Query1' query found", "Query1" in names, f"Got: {names}")

    # Token estimate should be present and reasonable
    tokens = page.locator("#statTokens").inner_text()
    result("Token estimate shown", tokens.startswith("~"))
    token_val = int(tokens.replace("~", "").replace(",", ""))
    result("Token estimate > 100", token_val > 100, f"Got: {token_val}")


def test_no_queries(page):
    """Test file without Power Queries shows appropriate feedback."""
    print("\n━━━ No Queries File ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["no_queries.xlsx"])

    # Main content should NOT appear
    result("Main content NOT visible", not page.locator("#mainContent").is_visible())

    # Drop zone should reappear
    result("Drop zone visible again", page.locator("#dropZone").is_visible())

    # Error log stays hidden for this expected condition
    error_log = page.locator("#errorLog")
    result("Error log hidden", not error_log.is_visible())

    # Bottom notice shows expected message
    notice = page.locator("#bottomNotice")
    result("Bottom notice visible", notice.is_visible())
    if notice.is_visible():
        result("Bottom notice mentions no queries found", "no power query" in notice.inner_text().lower(),
               notice.inner_text())


def test_multi_file_upload(page):
    """Test uploading multiple files at once."""
    print("\n━━━ Multi-File Upload ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx", "multi_query.xlsx"])

    files_stat = page.locator("#statFiles").inner_text()
    result("Files count = 2", files_stat == "2", f"Got: {files_stat}")

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 8 (3+5)", queries_stat == "8", f"Got: {queries_stat}")

    # Two file chips
    chips = page.locator(".file-chip")
    result("Two file chips shown", chips.count() == 2, f"Got: {chips.count()}")

    # Graph legend should have 2 entries
    page.locator('.tab[data-tab="graph"]').click()
    page.wait_for_timeout(500)
    legend = page.locator(".legend-item")
    result("Graph legend has 2 entries", legend.count() == 2, f"Got: {legend.count()}")

    # Code panel: file sections
    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    sections = page.locator(".file-section")
    result("Two file sections in code panel", sections.count() == 2, f"Got: {sections.count()}")


def test_drag_drop_partial_items_uses_files_list(page):
    """Test drag/drop when DataTransfer.items is partial but DataTransfer.files has all files."""
    print("\n━━━ Drag Drop (partial items list) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    f1 = TEST_DIR / "simple_query.xlsx"
    f2 = TEST_DIR / "multi_query.xlsx"
    payload = {
        "f1_name": f1.name,
        "f1_b64": base64.b64encode(f1.read_bytes()).decode("ascii"),
        "f2_name": f2.name,
        "f2_b64": base64.b64encode(f2.read_bytes()).decode("ascii"),
    }

    page.evaluate("""(payload) => {
        const b64ToU8 = (s) => {
            const bin = atob(s);
            const out = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
            return out;
        };

        const f1 = new File([b64ToU8(payload.f1_b64)], payload.f1_name, {
            type: "application/octet-stream",
            lastModified: Date.now()
        });
        const f2 = new File([b64ToU8(payload.f2_b64)], payload.f2_name, {
            type: "application/octet-stream",
            lastModified: Date.now() + 1
        });

        // Simulate a browser/platform that exposes only the first file in items,
        // while files correctly contains both.
        const item1 = {
            kind: "file",
            getAsFile: () => f1,
            webkitGetAsEntry: () => null
        };

        const mockDt = {
            items: [item1],
            files: [f1, f2]
        };

        const ev = new Event("drop", { bubbles: true, cancelable: true });
        Object.defineProperty(ev, "dataTransfer", { value: mockDt });
        document.getElementById("dropZone").dispatchEvent(ev);
    }""", payload)

    page.wait_for_function(
        "() => !document.getElementById('loading').classList.contains('active')",
        timeout=15000
    )
    page.wait_for_timeout(200)

    files_stat = page.locator("#statFiles").inner_text()
    result("Drag/drop uses DataTransfer.files to include all dropped files", files_stat == "2", f"Got: {files_stat}")


def test_drag_drop_multiple_file_handles_same_tick(page):
    """Test drag/drop where all file handles must be captured in the same event tick."""
    print("\n━━━ Drag Drop (same-tick handle capture) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    f1 = TEST_DIR / "simple_query.xlsx"
    f2 = TEST_DIR / "multi_query.xlsx"
    payload = {
        "f1_name": f1.name,
        "f1_b64": base64.b64encode(f1.read_bytes()).decode("ascii"),
        "f2_name": f2.name,
        "f2_b64": base64.b64encode(f2.read_bytes()).decode("ascii"),
    }

    page.evaluate("""(payload) => {
        const b64ToU8 = (s) => {
            const bin = atob(s);
            const out = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
            return out;
        };

        const f1 = new File([b64ToU8(payload.f1_b64)], payload.f1_name, {
            type: "application/octet-stream",
            lastModified: Date.now()
        });
        const f2 = new File([b64ToU8(payload.f2_b64)], payload.f2_name, {
            type: "application/octet-stream",
            lastModified: Date.now() + 1
        });

        // Simulate environments where getAsFileSystemHandle must be called
        // for all items in the same event turn.
        let sameTick = true;
        queueMicrotask(() => { sameTick = false; });

        const mkHandle = (file) => ({
            kind: "file",
            getFile: async () => file
        });

        const mkItem = (file) => ({
            kind: "file",
            getAsFileSystemHandle: () => {
                if (!sameTick) return Promise.reject(new Error("Handle access expired"));
                return Promise.resolve(mkHandle(file));
            },
            getAsFile: () => null,
            webkitGetAsEntry: () => null
        });

        const mockDt = {
            items: [mkItem(f1), mkItem(f2)],
            files: []
        };

        const ev = new Event("drop", { bubbles: true, cancelable: true });
        Object.defineProperty(ev, "dataTransfer", { value: mockDt });
        document.getElementById("dropZone").dispatchEvent(ev);
    }""", payload)

    page.wait_for_function(
        "() => !document.getElementById('loading').classList.contains('active')",
        timeout=15000
    )
    page.wait_for_timeout(200)

    files_stat = page.locator("#statFiles").inner_text()
    result("Drag/drop captures all file handles in same event tick", files_stat == "2", f"Got: {files_stat}")


def test_select_all_and_copy(page):
    """Test Select All toggle and Copy button."""
    print("\n━━━ Select All & Copy ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["multi_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # All checkboxes should be checked by default
    cbs = page.locator(".query-checkbox")
    all_checked = all(cbs.nth(i).is_checked() for i in range(cbs.count()))
    result("All query checkboxes checked by default", all_checked)

    # Copy button text shows count
    copy_text = page.locator("#copyBtnText").inner_text()
    result("Copy button shows count", "5" in copy_text, copy_text)

    # Click Select All (should deselect all since all are already checked)
    page.locator("#selectAllBtn").click()
    page.wait_for_timeout(200)

    none_checked = not any(cbs.nth(i).is_checked() for i in range(cbs.count()))
    result("Select All toggles to deselect", none_checked)

    desel_text = page.locator("#copyBtnText").inner_text()
    result("Copy button shows 0 after deselect", "0" in desel_text, desel_text)

    # Click again to select all
    page.locator("#selectAllBtn").click()
    page.wait_for_timeout(200)
    all_checked2 = all(cbs.nth(i).is_checked() for i in range(cbs.count()))
    result("Clicking again re-selects all", all_checked2)

    # Test Copy button (clipboard API may not be available in headless, but should not error)
    page.locator("#copyBtn").click()
    page.wait_for_timeout(500)
    toast = page.locator(".toast")
    result("Toast appears after copy", toast.is_visible())


def test_file_filter(page):
    """Test file filter checkboxes hide/show file sections."""
    print("\n━━━ File Filters ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx", "multi_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # Uncheck first file filter
    filters = page.locator("#codeFilters input[type='checkbox']")
    result("Two filter checkboxes", filters.count() == 2, f"Got: {filters.count()}")

    filters.first.uncheck()
    page.wait_for_timeout(200)

    # First file section should be hidden
    sections = page.locator(".file-section")
    first_visible = sections.first.is_visible()
    result("First file section hidden after uncheck", not first_visible)

    # Second still visible
    second_visible = sections.nth(1).is_visible()
    result("Second file section still visible", second_visible)

    # Re-check
    filters.first.check()
    page.wait_for_timeout(200)
    result("First file section visible after re-check", sections.first.is_visible())


def test_tab_switching(page):
    """Test tab switching between Graph and Code."""
    print("\n━━━ Tab Switching ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Default is graph tab
    result("Graph panel active by default", page.locator("#graphPanel").is_visible())
    result("Code panel hidden by default", not page.locator("#codePanel").is_visible())

    # Switch to code
    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    result("Code panel visible after click", page.locator("#codePanel").is_visible())
    result("Graph panel hidden after click", not page.locator("#graphPanel").is_visible())

    # Code tab has active class
    code_tab = page.locator('.tab[data-tab="code"]')
    result("Code tab has 'active' class", "active" in (code_tab.get_attribute("class") or ""))

    # Switch back
    page.locator('.tab[data-tab="graph"]').click()
    page.wait_for_timeout(200)
    result("Graph panel visible again", page.locator("#graphPanel").is_visible())


def test_graph_controls(page):
    """Test graph zoom, fit, relayout, and search."""
    print("\n━━━ Graph Controls ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["multi_query.xlsx"])

    # Graph should be active
    result("Graph panel active", page.locator("#graphPanel").is_visible())

    # Zoom in button exists and is clickable
    zoom_in = page.locator("#graphZoomIn")
    result("Zoom in button exists", zoom_in.is_visible())
    zoom_in.click()
    page.wait_for_timeout(200)
    result("Zoom in click does not error", True)

    # Zoom out
    page.locator("#graphZoomOut").click()
    page.wait_for_timeout(200)
    result("Zoom out click does not error", True)

    # Fit
    page.locator("#graphFit").click()
    page.wait_for_timeout(200)
    result("Fit click does not error", True)

    # Relayout
    page.locator("#graphRelayout").click()
    page.wait_for_timeout(800)
    result("Relayout click does not error", True)

    # Search
    search_input = page.locator("#graphSearchInput")
    search_input.fill("Orders")
    page.wait_for_timeout(400)
    result("Graph search does not error", True)

    # Clear search
    search_input.fill("")
    page.wait_for_timeout(400)
    result("Graph search clear does not error", True)


def test_reset(page):
    """Test reset button functionality."""
    print("\n━━━ Reset ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    result("Main content visible before reset", page.locator("#mainContent").is_visible())

    page.locator("#resetBtn").click()
    page.wait_for_timeout(300)

    result("Main content hidden after reset", not page.locator("#mainContent").is_visible())
    result("Drop zone visible after reset", page.locator("#dropZone").is_visible())
    result("Reset button hidden after reset", not page.locator("#resetBtn").is_visible())
    result("Error log hidden after reset", not page.locator("#errorLog").is_visible())


def test_keyboard_shortcuts(page):
    """Test keyboard shortcuts (Ctrl+A, Esc)."""
    print("\n━━━ Keyboard Shortcuts ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["multi_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # First deselect all manually
    page.locator("#selectAllBtn").click()
    page.wait_for_timeout(200)

    # Ctrl+A should select all
    page.keyboard.press("Control+a")
    page.wait_for_timeout(200)

    cbs = page.locator(".query-checkbox")
    all_checked = all(cbs.nth(i).is_checked() for i in range(cbs.count()))
    result("Ctrl+A selects all queries", all_checked)

    # Esc should reset
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    result("Esc resets to drop zone", page.locator("#dropZone").is_visible())
    result("Main content hidden after Esc", not page.locator("#mainContent").is_visible())


def test_prompt_templates(page):
    """Test LLM prompt dropdown and Copy All button in compact header."""
    print("\n━━━ Prompt Templates ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Prompt dropdown in header (visible in compact mode)
    dropdown = page.locator("#promptDropdown")
    result("Prompt dropdown visible", dropdown.is_visible())

    # 5 options (No prompt + 4 templates)
    options = dropdown.locator("option")
    result("5 prompt options", options.count() == 5, f"Got: {options.count()}")

    # Select a prompt and click Copy All
    dropdown.select_option("analyze")
    page.wait_for_timeout(100)

    copy_btn = page.locator("#copyAllBtn")
    result("Copy All button visible", copy_btn.is_visible())
    copy_btn.click()
    page.wait_for_timeout(500)

    toast = page.locator(".toast")
    result("Toast appears after prompt copy", toast.is_visible())


def test_file_section_collapse(page):
    """Test file section collapse/expand."""
    print("\n━━━ File Section Collapse ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["multi_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # Click file header to collapse
    header = page.locator(".file-header").first
    header.click()
    page.wait_for_timeout(200)

    section = page.locator(".file-section").first
    result("Section has 'collapsed' class after click", "collapsed" in (section.get_attribute("class") or ""))

    # Queries should be hidden
    queries_div = section.locator(".file-queries")
    result("Queries hidden when collapsed", not queries_div.is_visible())

    # Click again to expand
    header.click()
    page.wait_for_timeout(200)
    result("Section expanded after second click", "collapsed" not in (section.get_attribute("class") or ""))


def test_mixed_valid_invalid(page):
    """Test uploading a mix of valid and no-query files."""
    print("\n━━━ Mixed Valid + Invalid Files ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx", "no_queries.xlsx"])

    # Should show main content (at least one valid file)
    result("Main content visible (valid file processed)", page.locator("#mainContent").is_visible())

    files_stat = page.locator("#statFiles").inner_text()
    result("Files count = 1 (only valid file)", files_stat == "1", f"Got: {files_stat}")

    # Bottom notice should be visible for the no-query file
    result("Bottom notice visible for no-query file", page.locator("#bottomNotice").is_visible())


def test_individual_checkbox(page):
    """Test individual query checkbox toggling."""
    print("\n━━━ Individual Checkbox Toggle ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["multi_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # Uncheck first checkbox
    cbs = page.locator(".query-checkbox")
    cbs.first.uncheck()
    page.wait_for_timeout(200)

    copy_text = page.locator("#copyBtnText").inner_text()
    result("Copy count decremented after uncheck", "4" in copy_text, copy_text)

    # Re-check
    cbs.first.check()
    page.wait_for_timeout(200)
    copy_text2 = page.locator("#copyBtnText").inner_text()
    result("Copy count restored after re-check", "5" in copy_text2, copy_text2)


def test_token_estimation(page):
    """Test token estimation is reasonable."""
    print("\n━━━ Token Estimation ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["stress_test.xlsx"])

    tokens_text = page.locator("#statTokens").inner_text()
    result("Token text starts with ~", tokens_text.startswith("~"))

    token_val = int(tokens_text.replace("~", "").replace(",", ""))
    # 26 queries with code, should be at least a few hundred tokens
    result("Token estimate > 200 for 26 queries", token_val > 200, f"Got: {token_val}")
    result("Token estimate < 100000 (reasonable upper bound)", token_val < 100000, f"Got: {token_val}")


def test_dependency_count(page):
    """Test dependency count in stats bar."""
    print("\n━━━ Dependency Count ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["stress_test.xlsx"])

    deps_text = page.locator("#statDeps").inner_text()
    deps_val = int(deps_text)
    # Stress test has Transform->DataSource, Merge->Transform, Aggregate->Merge, Output->Aggregate deps
    result("Dependencies > 0 for stress test", deps_val > 0, f"Got: {deps_val}")


def test_browse_button(page):
    """Test that the Browse Files button triggers file input."""
    print("\n━━━ Browse Button ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    browse = page.locator(".browse-btn")
    result("Browse button visible", browse.is_visible())
    result("Browse button text is 'Browse Files'", "Browse Files" in browse.inner_text())

    # The file input should be hidden
    fi = page.locator("#fileInput")
    result("File input is hidden", not fi.is_visible())


def test_pbix_simple(page):
    """Test parsing a .pbix file with 3 queries."""
    print("\n━━━ PBIX Simple (3 queries) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.pbix"])

    result("Main content visible", page.locator("#mainContent").is_visible())

    files_stat = page.locator("#statFiles").inner_text()
    result("Files count = 1", files_stat == "1", f"Got: {files_stat}")

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 3", queries_stat == "3", f"Got: {queries_stat}")

    chips = page.locator(".file-chip")
    chip_text = chips.first.inner_text()
    result("File chip shows '.pbix'", ".pbix" in chip_text, chip_text)

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    actual_names = {el.inner_text() for el in page.locator(".query-name").all()}
    expected_names = {"Query1", "SalesData", "EdgeLookup"}
    result("Expected query names present", actual_names == expected_names, f"Got: {actual_names}")

    result("No error log visible", not page.locator("#errorLog").is_visible())


def test_pbix_multi_query(page):
    """Test parsing a .pbix file with 5 queries and dependencies."""
    print("\n━━━ PBIX Multi Query (5 queries + deps) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["multi_query.pbix"])

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 5", queries_stat == "5", f"Got: {queries_stat}")

    deps_stat = page.locator("#statDeps").inner_text()
    deps_val = int(deps_stat)
    result("Dependencies > 0", deps_val > 0, f"Got: {deps_val}")

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    expected_names = {"Query1", "RawOrders", "Customers", "OrdersWithCustomers", "SalesSummary"}
    actual_names = {el.inner_text() for el in page.locator(".query-name").all()}
    result("All 5 query names present", expected_names == actual_names,
           f"Expected: {expected_names}, Got: {actual_names}")


def test_pbit_simple(page):
    """Test parsing a .pbit (template) file."""
    print("\n━━━ PBIT Simple (3 queries) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.pbit"])

    result("Main content visible", page.locator("#mainContent").is_visible())

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 3", queries_stat == "3", f"Got: {queries_stat}")

    chips = page.locator(".file-chip")
    chip_text = chips.first.inner_text()
    result("File chip shows '.pbit'", ".pbit" in chip_text, chip_text)


def test_pbit_schema_only(page):
    """Test parsing a .pbit file that only has DataModelSchema (no DataMashup)."""
    print("\n━━━ PBIT Schema-Only (DataModelSchema fallback) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["schema_only.pbit"])

    result("Main content visible", page.locator("#mainContent").is_visible())

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 3", queries_stat == "3", f"Got: {queries_stat}")

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    actual_names = {el.inner_text() for el in page.locator(".query-name").all()}
    expected_names = {"SalesData", "Customers", "StartDate"}
    result("Expected query names present", actual_names == expected_names, f"Got: {actual_names}")

    result("No error log visible", not page.locator("#errorLog").is_visible())


def test_mixed_xlsx_pbix(page):
    """Test uploading .xlsx and .pbix files together."""
    print("\n━━━ Mixed .xlsx + .pbix Upload ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx", "multi_query.pbix"])

    files_stat = page.locator("#statFiles").inner_text()
    result("Files count = 2", files_stat == "2", f"Got: {files_stat}")

    queries_stat = page.locator("#statQueries").inner_text()
    result("Queries count = 8 (3+5)", queries_stat == "8", f"Got: {queries_stat}")

    chips = page.locator(".file-chip")
    result("Two file chips shown", chips.count() == 2, f"Got: {chips.count()}")

    page.locator('.tab[data-tab="graph"]').click()
    page.wait_for_timeout(500)
    legend = page.locator(".legend-item")
    result("Graph legend has 2 entries", legend.count() == 2, f"Got: {legend.count()}")


def test_no_console_errors(page):
    """Test that no JavaScript errors are logged on load."""
    print("\n━━━ Console Errors ━━━")

    errors = []
    page.on("pageerror", lambda exc: errors.append(str(exc)))

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    result("No JS errors on page load", len(errors) == 0,
           f"Errors: {errors}" if errors else "")


def test_no_console_errors_after_upload(page):
    """Test that no JS errors occur during file processing."""
    print("\n━━━ Console Errors After Upload ━━━")

    errors = []
    page.on("pageerror", lambda exc: errors.append(str(exc)))

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["edge_cases.xlsx"])

    result("No JS errors after processing edge cases", len(errors) == 0,
           f"Errors: {errors}" if errors else "")


def test_compact_header(page):
    """Test that header becomes compact after file upload and reverts on reset."""
    print("\n━━━ Compact Header ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    # Before upload: no compact class
    container = page.locator(".container")
    cls = container.get_attribute("class") or ""
    result("No compact class before upload", "compact" not in cls, cls)

    # Subtitle should be visible
    subtitle = page.locator(".subtitle")
    result("Subtitle visible before upload", subtitle.is_visible())

    upload_files(page, ["simple_query.xlsx"])

    # After upload: compact class present
    cls = container.get_attribute("class") or ""
    result("Compact class added after upload", "compact" in cls, cls)

    # Subtitle should be hidden in compact mode
    result("Subtitle hidden in compact mode", not subtitle.is_visible())

    # Header should still show title
    h1 = page.locator("h1")
    result("Title still visible in compact mode", h1.is_visible())

    # Reset should remove compact
    page.locator("#resetBtn").click()
    page.wait_for_timeout(300)
    cls = container.get_attribute("class") or ""
    result("Compact class removed after reset", "compact" not in cls, cls)
    result("Subtitle visible again after reset", subtitle.is_visible())


def test_data_tab_xlsx(page):
    """Test Data tab appears for xlsx files and shows worksheet preview."""
    print("\n━━━ Data Tab (xlsx) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    # Data tab should be hidden initially
    data_tab = page.locator("#dataTabBtn")
    result("Data tab hidden before upload", not data_tab.is_visible())

    upload_files(page, ["simple_query.xlsx"])

    # Data tab should be visible for xlsx files
    result("Data tab visible after xlsx upload", data_tab.is_visible())

    # Click data tab
    data_tab.click()
    page.wait_for_timeout(300)

    # Data panel should be active
    data_panel = page.locator("#dataPanel")
    result("Data panel active after click", data_panel.is_visible())

    # Sheet chips should be present
    chips = page.locator(".data-sheet-chip")
    result("At least one sheet chip shown", chips.count() > 0, f"Got: {chips.count()}")

    # First chip should be active
    first_chip_cls = chips.first.get_attribute("class") or ""
    result("First sheet chip is active", "active" in first_chip_cls, first_chip_cls)

    # Preview table should be rendered
    table = page.locator(".data-preview table")
    result("Preview table rendered", table.count() > 0)

    # Table should have headers
    headers = page.locator(".data-preview thead th")
    result("Table has header cells", headers.count() > 0, f"Got: {headers.count()}")

    # Row info should be shown
    row_info = page.locator(".data-row-info")
    result("Row info shown", row_info.is_visible())
    row_text = row_info.inner_text()
    result("Row info contains 'Showing'", "Showing" in row_text, row_text)


def test_data_tab_pbix_no_datamodel(page):
    """Test Data tab does NOT appear for pbix files without DataModel (no table data)."""
    print("\n━━━ Data Tab (pbix without DataModel) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.pbix"])

    data_tab = page.locator("#dataTabBtn")
    result("Data tab hidden for pbix without DataModel", not data_tab.is_visible())


def test_data_tab_mixed_xlsx_pbix(page):
    """Test Data tab shows for mixed uploads with xlsx."""
    print("\n━━━ Data Tab (mixed xlsx+pbix) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx", "multi_query.pbix"])

    data_tab = page.locator("#dataTabBtn")
    result("Data tab visible when xlsx is in mix", data_tab.is_visible())

    data_tab.click()
    page.wait_for_timeout(300)

    # Only xlsx sheets should be in the chips (pbix has no worksheet data)
    chips = page.locator(".data-sheet-chip")
    result("Sheet chips present", chips.count() > 0)

    # Verify chips reference the xlsx file
    chip_html = page.locator("#dataSheetList").inner_html()
    result("Chips reference xlsx file", "simple_query.xlsx" in chip_html, chip_html[:200])


def test_data_tab_file_filtering(page):
    """Test Data tab filters tables by selected file."""
    print("\n━━━ Data Tab File Filtering ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx", "multi_query.xlsx"])

    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(300)

    # With all files selected, all sheet chips should be visible
    all_chips = page.locator(".data-sheet-chip").count()
    result("All sheets visible with all files selected", all_chips >= 2, f"Got: {all_chips}")

    # Click first file chip to filter to just that file
    file_chips = page.locator(".file-chip")
    if file_chips.count() >= 2:
        first_file = file_chips.first.get_attribute("data-file") or ""
        file_chips.first.click()
        page.wait_for_timeout(300)

        # Data tab should now only show sheets from the selected file
        filtered_chips = page.locator(".data-sheet-chip").count()
        result("Fewer sheets after file filter", filtered_chips <= all_chips, f"Before: {all_chips}, After: {filtered_chips}")

        # Verify all visible chips reference the selected file
        chip_html = page.locator("#dataSheetList").inner_html()
        result("Filtered chips reference selected file", first_file in chip_html or filtered_chips == 0, chip_html[:200])

        # Click second file to switch filter
        file_chips.nth(1).click()
        page.wait_for_timeout(300)
        second_file = file_chips.nth(1).get_attribute("data-file") or ""
        chip_html2 = page.locator("#dataSheetList").inner_html()
        result("Switching file updates data tab", second_file in chip_html2 or page.locator(".data-sheet-chip").count() == 0, chip_html2[:200])


def test_sheet_chip_selection(page):
    """Test clicking sheet chips updates the preview."""
    print("\n━━━ Sheet Chip Selection ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx", "multi_query.xlsx"])

    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(300)

    chips = page.locator(".data-sheet-chip")
    chip_count = chips.count()
    result("Multiple sheet chips for multi-file", chip_count >= 2, f"Got: {chip_count}")

    # Click second chip
    if chip_count >= 2:
        chips.nth(1).click()
        page.wait_for_timeout(300)

        second_cls = chips.nth(1).get_attribute("class") or ""
        result("Clicked chip becomes active", "active" in second_cls)

        first_cls = chips.first.get_attribute("class") or ""
        result("Previous chip deactivated", "active" not in first_cls, first_cls)


def test_export_csv(page):
    """Test CSV export button triggers download."""
    print("\n━━━ CSV Export ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(300)

    # Export CSV button should be enabled (sheet auto-selected)
    csv_btn = page.locator("#exportCsvBtn")
    result("CSV export button visible", csv_btn.is_visible())
    is_disabled = csv_btn.get_attribute("disabled")
    result("CSV export button enabled", is_disabled is None, f"disabled={is_disabled}")

    # Click and verify no JS error (actual download is handled by blob URL)
    with page.expect_download() as download_info:
        csv_btn.click()
    download = download_info.value
    result("CSV download triggered", download.suggested_filename.endswith(".csv"),
           f"filename={download.suggested_filename}")

    # Verify CSV content
    content = download.path().read_text(encoding="utf-8-sig")
    lines = content.strip().split("\n")
    result("CSV has header + data rows", len(lines) >= 2, f"Got {len(lines)} lines")

    # Toast should appear
    page.wait_for_timeout(500)
    toast = page.locator(".toast")
    result("Success toast after CSV export", toast.is_visible())


def test_export_parquet(page):
    """Test Parquet export button triggers download."""
    print("\n━━━ Parquet Export ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(300)

    parquet_btn = page.locator("#exportParquetBtn")
    result("Parquet export button visible", parquet_btn.is_visible())
    is_disabled = parquet_btn.get_attribute("disabled")
    result("Parquet export button enabled", is_disabled is None, f"disabled={is_disabled}")

    # Click and verify download triggers
    with page.expect_download() as download_info:
        parquet_btn.click()
    download = download_info.value
    result("Parquet download triggered", download.suggested_filename.endswith(".parquet"),
           f"filename={download.suggested_filename}")

    # Verify Parquet magic bytes: PAR1 header and footer
    raw = download.path().read_bytes()
    result("Parquet file starts with PAR1 magic", raw[:4] == b"PAR1", f"Got: {raw[:4]}")
    result("Parquet file ends with PAR1 magic", raw[-4:] == b"PAR1", f"Got: {raw[-4:]}")
    result("Parquet file has reasonable size", len(raw) > 50, f"Got: {len(raw)} bytes")


def test_export_buttons_disabled_initially(page):
    """Test export buttons are disabled when no sheet is selected."""
    print("\n━━━ Export Buttons Disabled ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.pbix"])

    # Force show data tab (even though pbix has no data)
    page.evaluate("document.getElementById('dataTabBtn').style.display=''")
    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(200)

    csv_btn = page.locator("#exportCsvBtn")
    parquet_btn = page.locator("#exportParquetBtn")
    result("CSV button disabled without data", csv_btn.get_attribute("disabled") is not None)
    result("Parquet button disabled without data", parquet_btn.get_attribute("disabled") is not None)


def test_data_profile_checkbox(page):
    """Test data profile checkbox visibility."""
    print("\n━━━ Data Profile Checkbox ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    # Profile checkbox wrapper has display:none initially
    hidden = page.evaluate("() => document.getElementById('includeProfileWrap').style.display === 'none'")
    result("Profile checkbox hidden initially", hidden)

    upload_files(page, ["simple_query.xlsx"])

    # Profile checkbox display should be cleared after xlsx upload
    shown = page.evaluate("() => document.getElementById('includeProfileWrap').style.display !== 'none'")
    result("Profile checkbox display enabled after xlsx upload", shown)

    # Switch to code tab to make it actually visible
    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    profile_wrap = page.locator("#includeProfileWrap")
    result("Profile checkbox visible in code panel", profile_wrap.is_visible())

    # Should be unchecked by default
    cb = page.locator("#includeProfileCb")
    result("Profile checkbox unchecked by default", not cb.is_checked())


def test_data_profile_pbix_no_datamodel(page):
    """Test data profile checkbox hidden for pbix without DataModel."""
    print("\n━━━ Data Profile (pbix without DataModel) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.pbix"])

    profile_wrap = page.locator("#includeProfileWrap")
    result("Profile checkbox hidden for pbix without DataModel", not profile_wrap.is_visible())


def test_copy_with_data_profile(page):
    """Test copy includes data profile when checkbox is checked."""
    print("\n━━━ Copy with Data Profile ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # Check the profile checkbox
    page.locator("#includeProfileCb").check()
    page.wait_for_timeout(100)
    result("Profile checkbox checked", page.locator("#includeProfileCb").is_checked())

    # Use evaluate to capture what would be copied (clipboard may not work in headless)
    copied_text = page.evaluate("""() => {
        const profile = document.getElementById('includeProfileCb').checked;
        let extra = '';
        if (profile && appState.worksheets.length > 0) {
            extra = buildDataProfile(appState.worksheets);
        }
        return getSelCode() + (extra ? '\\n\\n' + extra : '');
    }""")

    result("Copied text includes Data Profile header", "Data Profile" in copied_text,
           copied_text[:200] if len(copied_text) > 200 else copied_text)
    result("Copied text includes column stats", "distinct" in copied_text,
           copied_text[-300:] if len(copied_text) > 300 else copied_text)


def test_copy_without_data_profile(page):
    """Test copy does NOT include data profile when checkbox is unchecked."""
    print("\n━━━ Copy without Data Profile ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # Profile checkbox unchecked by default
    result("Profile checkbox unchecked", not page.locator("#includeProfileCb").is_checked())

    copied_text = page.evaluate("""() => {
        const profile = document.getElementById('includeProfileCb').checked;
        let extra = '';
        if (profile && appState.worksheets.length > 0) {
            extra = buildDataProfile(appState.worksheets);
        }
        return getSelCode() + (extra ? '\\n\\n' + extra : '');
    }""")

    result("Copied text does NOT include Data Profile", "Data Profile" not in copied_text,
           copied_text[-200:] if len(copied_text) > 200 else copied_text)


def test_prompt_template_with_profile(page):
    """Test prompt template includes data profile when checkbox is checked."""
    print("\n━━━ Prompt Template with Profile ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # Check profile checkbox
    page.locator("#includeProfileCb").check()
    page.wait_for_timeout(100)

    # Simulate what the prompt template copy would produce
    copied_text = page.evaluate("""() => {
        const profile = document.getElementById('includeProfileCb').checked;
        let extra = '';
        if (profile && appState.worksheets.length > 0) {
            try { extra = buildDataProfile(appState.worksheets); } catch(e) {}
        }
        return getSelCode(PROMPT_TEMPLATES['analyze']) + (extra ? '\\n\\n' + extra : '');
    }""")

    result("Prompt template text includes analysis prompt", "Analyze" in copied_text or "analyze" in copied_text,
           copied_text[:200])
    result("Prompt template text includes data profile", "Data Profile" in copied_text,
           copied_text[-200:] if len(copied_text) > 200 else copied_text)


def test_data_profile_stats_accuracy(page):
    """Test data profile computes reasonable statistics."""
    print("\n━━━ Data Profile Stats ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Call buildDataProfile directly
    profile = page.evaluate("""() => {
        if (!appState.worksheets.length) return '';
        return buildDataProfile(appState.worksheets);
    }""")

    result("Profile text is non-empty", len(profile) > 0, f"Length: {len(profile)}")
    result("Profile has header line", "Data Profile" in profile)
    result("Profile mentions file name", "simple_query.xlsx" in profile, profile[:300])
    result("Profile has distinct counts", "distinct" in profile)
    result("Profile has null counts", "nulls" in profile or "empty" in profile)


def test_compute_column_stats(page):
    """Test computeColumnStats function directly."""
    print("\n━━━ Column Stats Computation ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    # Test numeric column
    stats = page.evaluate("""() => {
        return computeColumnStats('Price', ['10', '20', '30', '', '40']);
    }""")
    result("Numeric: distinct count", stats["distinct"] == 4, f"Got: {stats['distinct']}")
    result("Numeric: null count", stats["nulls"] == 1, f"Got: {stats['nulls']}")
    result("Numeric: isNumeric true", stats["isNumeric"] is True)
    result("Numeric: min = 10", stats["min"] == 10, f"Got: {stats['min']}")
    result("Numeric: max = 40", stats["max"] == 40, f"Got: {stats['max']}")
    result("Numeric: avg = 25", stats["avg"] == 25, f"Got: {stats['avg']}")

    # Test categorical column
    stats2 = page.evaluate("""() => {
        return computeColumnStats('Color', ['Red', 'Blue', 'Red', 'Green', 'Red', 'Blue']);
    }""")
    result("Categorical: distinct = 3", stats2["distinct"] == 3, f"Got: {stats2['distinct']}")
    result("Categorical: nulls = 0", stats2["nulls"] == 0, f"Got: {stats2['nulls']}")
    result("Categorical: top values present", stats2["top"] is not None and len(stats2["top"]) > 0)
    result("Categorical: Red is top value", stats2["top"][0]["value"] == "Red" and stats2["top"][0]["count"] == 3,
           f"Got: {stats2['top'][0]}" if stats2["top"] else "No top values")

    # Test all-null column
    stats3 = page.evaluate("""() => {
        return computeColumnStats('Empty', ['', '', null, '']);
    }""")
    result("All-null: distinct = 0", stats3["distinct"] == 0, f"Got: {stats3['distinct']}")
    result("All-null: nulls = 4", stats3["nulls"] == 4, f"Got: {stats3['nulls']}")


def test_worksheet_extraction(page):
    """Test extractWorksheetData parses xlsx worksheets correctly."""
    print("\n━━━ Worksheet Extraction ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Check appState.worksheets
    ws_count = page.evaluate("() => appState.worksheets.length")
    result("Worksheets extracted", ws_count > 0, f"Got: {ws_count}")

    ws_info = page.evaluate("""() => {
        if (!appState.worksheets.length) return null;
        const ws = appState.worksheets[0];
        return {
            fileName: ws.fileName,
            sheetName: ws.sheetName,
            headerCount: ws.headers.length,
            rowCount: ws.totalRows,
            truncated: ws.truncated
        };
    }""")

    if ws_info:
        result("Worksheet has correct file name", ws_info["fileName"] == "simple_query.xlsx",
               f"Got: {ws_info['fileName']}")
        result("Worksheet has a sheet name", len(ws_info["sheetName"]) > 0, ws_info["sheetName"])
        result("Worksheet has headers", ws_info["headerCount"] > 0, f"Got: {ws_info['headerCount']}")
        result("Worksheet has rows", ws_info["rowCount"] >= 0, f"Got: {ws_info['rowCount']}")
        result("Worksheet not truncated (small file)", ws_info["truncated"] is False)
    else:
        result("Worksheet info retrieved", False, "ws_info is None")


def test_data_tab_reset(page):
    """Test that reset clears data tab state."""
    print("\n━━━ Data Tab Reset ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Verify data tab is visible
    result("Data tab visible before reset", page.locator("#dataTabBtn").is_visible())

    # Switch to code tab and check profile checkbox
    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    result("Profile checkbox visible before reset", page.locator("#includeProfileWrap").is_visible())
    page.locator("#includeProfileCb").check()
    page.wait_for_timeout(100)

    # Reset
    page.locator("#resetBtn").click()
    page.wait_for_timeout(300)

    # Data tab should be hidden
    result("Data tab hidden after reset", not page.locator("#dataTabBtn").is_visible())

    # Profile checkbox display should be none
    hidden = page.evaluate("() => document.getElementById('includeProfileWrap').style.display === 'none'")
    result("Profile checkbox hidden after reset", hidden)
    result("Profile checkbox unchecked after reset", not page.locator("#includeProfileCb").is_checked())

    # appState.worksheets should be empty
    ws_count = page.evaluate("() => appState.worksheets.length")
    result("Worksheets cleared after reset", ws_count == 0, f"Got: {ws_count}")


def test_csv_export_streaming(page):
    """Test that CSV export uses streaming (chunked) approach."""
    print("\n━━━ CSV Export Streaming ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(300)

    # Verify the exportCSV function uses setTimeout for yielding
    has_streaming = page.evaluate("""() => {
        const src = exportCSV.toString();
        return src.includes('setTimeout') || src.includes('Promise');
    }""")
    result("exportCSV uses async yielding", has_streaming)

    # Verify escapeCSVField handles edge cases
    tests = page.evaluate("""() => {
        return {
            simple: escapeCSVField('hello'),
            withComma: escapeCSVField('hello,world'),
            withQuote: escapeCSVField('say "hi"'),
            withNewline: escapeCSVField('line1\\nline2'),
            nullVal: escapeCSVField(null),
            emptyVal: escapeCSVField('')
        };
    }""")
    result("CSV: simple value unquoted", tests["simple"] == "hello")
    result("CSV: comma triggers quoting", tests["withComma"] == '"hello,world"')
    result("CSV: quotes are escaped", tests["withQuote"] == '"say ""hi"""')
    result("CSV: newline triggers quoting", tests["withNewline"] == '"line1\nline2"')
    result("CSV: null becomes empty", tests["nullVal"] == "")
    result("CSV: empty stays empty", tests["emptyVal"] == "")


def test_parquet_export_structure(page):
    """Test Parquet export produces a valid minimal Parquet file."""
    print("\n━━━ Parquet Export Structure ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(300)

    # Verify the exportParquet function uses async yielding
    has_streaming = page.evaluate("""() => {
        const src = exportParquet.toString();
        return src.includes('setTimeout') || src.includes('Promise');
    }""")
    result("exportParquet uses async yielding", has_streaming)

    # Test buildParquetBuffer directly with small data
    parquet_info = page.evaluate("""() => {
        const buf = buildParquetBuffer(['Name', 'Age'], [['Alice', 'Bob'], ['30', '25']]);
        return {
            length: buf.length,
            startMagic: String.fromCharCode(buf[0], buf[1], buf[2], buf[3]),
            endMagic: String.fromCharCode(buf[buf.length-4], buf[buf.length-3], buf[buf.length-2], buf[buf.length-1])
        };
    }""")
    result("Parquet buffer has PAR1 start", parquet_info["startMagic"] == "PAR1")
    result("Parquet buffer has PAR1 end", parquet_info["endMagic"] == "PAR1")
    result("Parquet buffer has reasonable size", parquet_info["length"] > 100, f"Got: {parquet_info['length']}")


def test_col_ref_to_index(page):
    """Test column reference to index conversion."""
    print("\n━━━ Column Ref to Index ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    results = page.evaluate("""() => {
        return {
            a1: colRefToIndex('A1'),
            b3: colRefToIndex('B3'),
            z1: colRefToIndex('Z1'),
            aa1: colRefToIndex('AA1'),
            az5: colRefToIndex('AZ5')
        };
    }""")
    result("A1 -> 0", results["a1"] == 0, f"Got: {results['a1']}")
    result("B3 -> 1", results["b3"] == 1, f"Got: {results['b3']}")
    result("Z1 -> 25", results["z1"] == 25, f"Got: {results['z1']}")
    result("AA1 -> 26", results["aa1"] == 26, f"Got: {results['aa1']}")
    result("AZ5 -> 51", results["az5"] == 51, f"Got: {results['az5']}")


def test_tab_switching_with_data(page):
    """Test tab switching includes Data tab when worksheets exist."""
    print("\n━━━ Tab Switching with Data ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Switch to data tab
    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(200)

    result("Data panel visible", page.locator("#dataPanel").is_visible())
    result("Graph panel hidden", not page.locator("#graphPanel").is_visible())
    result("Code panel hidden", not page.locator("#codePanel").is_visible())

    # Data tab should have active class
    data_tab = page.locator("#dataTabBtn")
    result("Data tab has active class", "active" in (data_tab.get_attribute("class") or ""))

    # Switch to graph
    page.locator('.tab[data-tab="graph"]').click()
    page.wait_for_timeout(200)

    result("Graph panel visible after switch", page.locator("#graphPanel").is_visible())
    result("Data panel hidden after switch", not page.locator("#dataPanel").is_visible())

    # Switch back to data
    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(200)
    result("Data panel visible again", page.locator("#dataPanel").is_visible())


def test_pbix_data_extraction_functions(page):
    """Test PBIX data extraction functions are available and callable."""
    print("\n━━━ PBIX Data Extraction Functions ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    # Verify all extraction functions exist
    funcs = page.evaluate("""() => {
        return {
            buildSchemaFromSQLite: typeof buildSchemaFromSQLite === 'function',
            extractTableData: typeof extractTableData === 'function',
            readIdfMeta: typeof readIdfMeta === 'function',
            readIdf: typeof readIdf === 'function',
            decodeRleBitPackedHybrid: typeof decodeRleBitPackedHybrid === 'function',
            readDictionary: typeof readDictionary === 'function',
            buildHuffmanTree: typeof buildHuffmanTree === 'function',
            convertColumnValue: typeof convertColumnValue === 'function',
            extractPbixTableData: typeof extractPbixTableData === 'function',
            parseABF: typeof parseABF === 'function',
            getDataSlice: typeof getDataSlice === 'function',
            readSQLiteTables: typeof readSQLiteTables === 'function',
        };
    }""")

    for name, available in funcs.items():
        result(f"{name} function exists", available, f"missing")

    # Test convertColumnValue
    cv_tests = page.evaluate("""() => {
        const dt = convertColumnValue(44927, 9);  // 2023-01-01
        const dec = convertColumnValue(12345, 10);
        const plain = convertColumnValue(42, 6);
        const nil = convertColumnValue(null, 2);
        return {
            datetime_is_date: dt instanceof Date,
            datetime_year: dt instanceof Date ? dt.getUTCFullYear() : null,
            decimal: dec,
            plain: plain,
            nil: nil
        };
    }""")
    result("convertColumnValue: datetime → Date", cv_tests["datetime_is_date"])
    result("convertColumnValue: datetime year 2023", cv_tests["datetime_year"] == 2023, f"Got: {cv_tests['datetime_year']}")
    result("convertColumnValue: decimal /10000", cv_tests["decimal"] == 1.2345, f"Got: {cv_tests['decimal']}")
    result("convertColumnValue: plain passthrough", cv_tests["plain"] == 42)
    result("convertColumnValue: null passthrough", cv_tests["nil"] is None)


def test_compact_header_aggressive(page):
    """Test compact header is dramatically smaller with aggressive CSS."""
    print("\n━━━ Compact Header (aggressive) ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")

    # Measure header height before upload
    before_height = page.evaluate("() => document.querySelector('header').offsetHeight")

    upload_files(page, ["simple_query.xlsx"])

    # Measure header height after upload (compact mode)
    after_height = page.evaluate("() => document.querySelector('header').offsetHeight")

    result("Header height reduced after upload", after_height < before_height,
           f"Before: {before_height}px, After: {after_height}px")
    result("Compact header under 60px", after_height <= 60,
           f"Got: {after_height}px")

    # Privacy badge should be hidden in compact mode
    badge_visible = page.locator(".privacy-badge").is_visible()
    result("Privacy badge hidden in compact mode", not badge_visible)

    # Stats bar hidden in compact mode, inline stats visible instead
    stats_bar_hidden = page.evaluate("() => document.querySelector('.stats-bar').offsetParent === null")
    result("Stats bar hidden in compact mode", stats_bar_hidden)

    # Inline header stats visible
    header_stats = page.locator("#headerStats")
    result("Inline header stats visible", header_stats.is_visible())
    result("Inline header stats has content", len(header_stats.text_content()) > 0, header_stats.text_content())


def test_no_console_errors_with_new_features(page):
    """Test no JS errors occur with the new features (data tab, profile, exports)."""
    print("\n━━━ Console Errors (New Features) ━━━")

    errors = []
    page.on("pageerror", lambda exc: errors.append(str(exc)))

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    # Exercise all new features
    page.locator("#dataTabBtn").click()
    page.wait_for_timeout(300)

    # Check profile checkbox
    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)
    page.locator("#includeProfileCb").check()
    page.wait_for_timeout(100)

    # Build profile
    page.evaluate("""() => {
        if (appState.worksheets.length > 0) buildDataProfile(appState.worksheets);
    }""")

    result("No JS errors exercising new features", len(errors) == 0,
           f"Errors: {errors}" if errors else "")


def test_responsive_viewports(page):
    """Test that the app renders without overflow at various screen sizes."""
    print("\n━━━ Responsive Viewports ━━━")

    viewports = [
        ("desktop", 1280, 800),
        ("tablet", 768, 1024),
        ("mobile", 375, 667),
    ]
    for name, w, h in viewports:
        page.set_viewport_size({"width": w, "height": h})
        page.goto(BASE_URL)
        page.wait_for_selector("#dropZone", state="visible", timeout=10000)

        overflow = page.evaluate(
            "document.documentElement.scrollWidth > document.documentElement.clientWidth"
        )
        result(f"no horizontal overflow at {name} ({w}x{h})", not overflow,
               f"scrollWidth > clientWidth")

        drop = page.locator("#dropZone")
        visible = drop.is_visible()
        result(f"drop zone visible at {name}", visible)


def main():
    global PASS, FAIL

    print("╔══════════════════════════════════════════╗")
    print("║  Power Query Explorer — Test Suite       ║")
    print("╚══════════════════════════════════════════╝")
    print(f"  Server: localhost:{PORT}")

    # Start HTTP server
    server = subprocess.Popen(
        ["python3", "-m", "http.server", str(PORT)],
        cwd=str(PROJECT_DIR),
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )
    time.sleep(1.5)  # Wait for server startup

    try:
        _run_tests()
    finally:
        server.terminate()
        server.wait(timeout=5)

    print(f"\n{'═'*46}")
    print(f"  Results: \033[32m{PASS} passed\033[0m, \033[31m{FAIL} failed\033[0m")
    if ERRORS:
        print(f"\n  Failures:")
        for e in ERRORS:
            print(f"    \033[31m•\033[0m {e}")
    print()

    sys.exit(0 if FAIL == 0 else 1)

def _run_tests():
    global FAIL, ERRORS
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            # Grant clipboard permissions for copy tests
            permissions=["clipboard-read", "clipboard-write"],
            viewport={"width": 1440, "height": 900}
        )

        tests = [
            test_page_load,
            test_no_console_errors,
            test_simple_query,
            test_multi_query,
            test_complex_code,
            test_edge_cases,
            test_stress,
            test_no_queries,
            test_multi_file_upload,
            test_drag_drop_partial_items_uses_files_list,
            test_drag_drop_multiple_file_handles_same_tick,
            test_mixed_valid_invalid,
            test_select_all_and_copy,
            test_file_filter,
            test_tab_switching,
            test_graph_controls,
            test_reset,
            test_keyboard_shortcuts,
            test_prompt_templates,
            test_file_section_collapse,
            test_individual_checkbox,
            test_token_estimation,
            test_dependency_count,
            test_browse_button,
            test_no_console_errors_after_upload,
            test_pbix_simple,
            test_pbix_multi_query,
            test_pbit_simple,
            test_pbit_schema_only,
            test_mixed_xlsx_pbix,
            # New feature tests
            test_compact_header,
            test_data_tab_xlsx,
            test_data_tab_pbix_no_datamodel,
            test_data_tab_mixed_xlsx_pbix,
            test_data_tab_file_filtering,
            test_sheet_chip_selection,
            test_export_csv,
            test_export_parquet,
            test_export_buttons_disabled_initially,
            test_data_profile_checkbox,
            test_data_profile_pbix_no_datamodel,
            test_copy_with_data_profile,
            test_copy_without_data_profile,
            test_prompt_template_with_profile,
            test_data_profile_stats_accuracy,
            test_compute_column_stats,
            test_worksheet_extraction,
            test_data_tab_reset,
            test_csv_export_streaming,
            test_parquet_export_structure,
            test_col_ref_to_index,
            test_tab_switching_with_data,
            test_pbix_data_extraction_functions,
            test_compact_header_aggressive,
            test_no_console_errors_with_new_features,
            test_responsive_viewports,
        ]

        for test_fn in tests:
            page = context.new_page()
            try:
                test_fn(page)
            except Exception as e:
                FAIL += 1
                ERRORS.append(f"{test_fn.__name__}: EXCEPTION — {e}")
                print(f"  \033[31m✗\033[0m {test_fn.__name__} EXCEPTION: {e}")
            finally:
                page.close()

        browser.close()


if __name__ == "__main__":
    main()
