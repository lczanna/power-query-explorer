# Scripts

This folder contains maintainer utilities for test data generation and browser regression tests.

## `create_test_files.py`

Generates query-enabled `.xlsx` fixtures under `data/test-files/` using a template workbook (`scripts/templates/query-template.xlsx`).

Run:

```bash
uv run python scripts/create_test_files.py
```

## `test_explorer.py`

Runs the Playwright-based end-to-end test suite against `power-query-explorer.html`.

Prerequisites:

```bash
uv sync
uv run playwright install chromium
```

Run tests:

```bash
uv run python scripts/test_explorer.py
```
