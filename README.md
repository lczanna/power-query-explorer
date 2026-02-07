## Power Query Explorer

Power Query Explorer is a single-file web app that extracts and visualizes Power Query M code from Excel workbooks (`.xlsx`) directly in the browser.

Live site: https://lczanna.github.io/power-query-explorer/

### What it does

- Extracts M queries from uploaded `.xlsx` files.
- Builds a dependency graph between queries.
- Shows per-query code with syntax highlighting.
- Supports multi-file uploads and query selection for copy-to-clipboard workflows.
- Runs locally in the browser (no server required for normal use).

### Quick start

1. Open the live site above, or open `index.html` locally.
2. Drop one or more `.xlsx` files.
3. Explore:
   - `Dependency Graph` tab for relationships.
   - `Code` tab for query text and copy/export workflows.

### Local development

Serve the project locally (recommended instead of `file://`):

```bash
python3 -m http.server 8000
```

Then open:

- `http://localhost:8000/`
- or `http://localhost:8000/power-query-explorer.html`

### Test data and scripts

Repository includes helper scripts under `scripts/`:

- `scripts/create_test_files.py` regenerates example `.xlsx` files in `data/test-files/`.
- `scripts/test_explorer.py` runs Playwright browser tests against the local HTML app.

See `scripts/README.md` for commands.

### Repository structure

- `power-query-explorer.html`: Main self-contained app.
- `index.html`: GitHub Pages entrypoint (redirects to `power-query-explorer.html`).
- `data/test-files/`: Sample Excel files used for validation and demos.
- `scripts/`: Test-data generation and automated browser tests.
- `pyproject.toml` + `uv.lock`: Python environment metadata for project scripts.

### GitHub Pages

The repo is configured to publish from `master` branch root (`/`).

Production URL:

- https://lczanna.github.io/power-query-explorer/
