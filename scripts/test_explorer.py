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
import subprocess
import socket
from pathlib import Path
from playwright.sync_api import sync_playwright, expect

HTML_PATH = Path(__file__).resolve().parent.parent / "power-query-explorer.html"
TEST_DIR = Path(__file__).resolve().parent.parent / "data" / "test-files"
PROJECT_DIR = Path(__file__).resolve().parent.parent

def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

PORT = find_free_port()
BASE_URL = f"http://localhost:{PORT}/power-query-explorer.html"

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
    """Test LLM prompt template buttons."""
    print("\n━━━ Prompt Templates ━━━")

    page.goto(BASE_URL)
    page.wait_for_load_state("networkidle")
    upload_files(page, ["simple_query.xlsx"])

    page.locator('.tab[data-tab="code"]').click()
    page.wait_for_timeout(200)

    # 4 prompt template buttons
    templates = page.locator(".prompt-template")
    result("4 prompt template buttons", templates.count() == 4, f"Got: {templates.count()}")

    # Click first template (analyze)
    templates.first.click()
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
