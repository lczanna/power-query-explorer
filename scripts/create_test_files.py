"""
Generate test Excel (.xlsx) files with embedded Power Query M code.

The Power Query M code is stored in a DataMashup binary blob inside the xlsx ZIP.
The DataMashup binary contains:
  - A header with version info and size
  - An inner ZIP containing Formulas/Section1.m with the M code
  - The M code uses the format: section Section1; shared QueryName = let ... in ...;

Files are saved to data/test-files/.
"""

import io
import os
import struct
import zipfile

import openpyxl

OUTPUT_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "data",
    "test-files",
)


def build_datamashup_blob(m_code: str) -> bytes:
    """Build a DataMashup binary blob wrapping M code.

    The DataMashup format used by Excel:
      - 4 bytes: version (0x00 0x00 0x00 0x00)
      - 4 bytes: size of the package part (little-endian uint32)
      - N bytes: the inner ZIP package (containing Formulas/Section1.m)
      - After the ZIP: permission metadata (we add minimal trailing bytes)
    """
    # Create the inner ZIP containing the M code
    inner_zip_buf = io.BytesIO()
    with zipfile.ZipFile(inner_zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Content types file (required for valid OPC package)
        content_types = (
            '<?xml version="1.0" encoding="utf-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="m" ContentType="application/x-ms-m"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            "</Types>"
        )
        zf.writestr("[Content_Types].xml", content_types)

        # The M code file
        zf.writestr("Formulas/Section1.m", m_code)

    inner_zip_bytes = inner_zip_buf.getvalue()

    # Build the DataMashup blob
    blob = io.BytesIO()

    # Version: 4 zero bytes
    blob.write(b"\x00\x00\x00\x00")

    # Size of the package part (little-endian uint32)
    blob.write(struct.pack("<I", len(inner_zip_bytes)))

    # The inner ZIP data
    blob.write(inner_zip_bytes)

    # Minimal trailing permission/metadata section
    # (4 bytes size = 0 for empty permissions block)
    blob.write(struct.pack("<I", 0))

    return blob.getvalue()


def inject_datamashup(xlsx_path: str, datamashup_blob: bytes) -> None:
    """Inject a DataMashup blob into an xlsx file as customXml/item1.bin."""
    # Read existing xlsx into memory
    with open(xlsx_path, "rb") as f:
        xlsx_bytes = f.read()

    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zin:
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            # Copy all existing entries
            for item in zin.infolist():
                data = zin.read(item.filename)
                zout.writestr(item, data)

            # Add the DataMashup blob
            zout.writestr("customXml/item1.bin", datamashup_blob)

            # Add a relationship file so Excel knows about customXml
            rels_content = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.bin"/>'
                "</Relationships>"
            )
            # Only add if not already present
            if "customXml/_rels/item1.bin.rels" not in [
                i.filename for i in zin.infolist()
            ]:
                zout.writestr("customXml/_rels/item1.bin.rels", rels_content)

    with open(xlsx_path, "wb") as f:
        f.write(out_buf.getvalue())


def create_base_xlsx(path: str) -> None:
    """Create a minimal valid xlsx file with openpyxl."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Test Data"
    wb.save(path)


# ============================================================
# Test File Definitions
# ============================================================


def create_simple_query():
    """1. simple_query.xlsx - One simple Power Query."""
    path = os.path.join(OUTPUT_DIR, "simple_query.xlsx")
    create_base_xlsx(path)

    m_code = (
        "section Section1;\n"
        "shared SalesData = let\n"
        '    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n'
        '    ChangedType = Table.TransformColumnTypes(Source, {{"Amount", type number}, {"Date", type date}}),\n'
        '    FilteredRows = Table.SelectRows(ChangedType, each [Amount] > 100)\n'
        "in\n"
        "    FilteredRows;"
    )

    blob = build_datamashup_blob(m_code)
    inject_datamashup(path, blob)
    print(f"  Created: {path}")


def create_multi_query():
    """2. multi_query.xlsx - Multiple queries with dependencies."""
    path = os.path.join(OUTPUT_DIR, "multi_query.xlsx")
    create_base_xlsx(path)

    m_code = (
        "section Section1;\n"
        "shared RawOrders = let\n"
        '    Source = Csv.Document(File.Contents("C:\\Data\\orders.csv")),\n'
        "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),\n"
        '    ChangedType = Table.TransformColumnTypes(PromotedHeaders, {{"OrderID", Int64.Type}, {"CustomerID", Int64.Type}, {"Amount", type number}})\n'
        "in\n"
        "    ChangedType;\n"
        "shared Customers = let\n"
        '    Source = Csv.Document(File.Contents("C:\\Data\\customers.csv")),\n'
        "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),\n"
        '    ChangedType = Table.TransformColumnTypes(PromotedHeaders, {{"CustomerID", Int64.Type}, {"Name", type text}, {"Region", type text}})\n'
        "in\n"
        "    ChangedType;\n"
        "shared OrdersWithCustomers = let\n"
        "    Source = RawOrders,\n"
        '    Merged = Table.NestedJoin(Source, {"CustomerID"}, Customers, {"CustomerID"}, "CustomerData", JoinKind.LeftOuter),\n'
        '    Expanded = Table.ExpandTableColumn(Merged, "CustomerData", {"Name", "Region"})\n'
        "in\n"
        "    Expanded;\n"
        "shared SalesSummary = let\n"
        "    Source = OrdersWithCustomers,\n"
        '    Grouped = Table.Group(Source, {"Region"}, {{"TotalSales", each List.Sum([Amount]), type number}, {"OrderCount", each Table.RowCount(_), Int64.Type}})\n'
        "in\n"
        "    Grouped;"
    )

    blob = build_datamashup_blob(m_code)
    inject_datamashup(path, blob)
    print(f"  Created: {path}")


def create_complex_code():
    """3. complex_code.xlsx - Complex M code with comments, nested lets, special chars."""
    path = os.path.join(OUTPUT_DIR, "complex_code.xlsx")
    create_base_xlsx(path)

    m_code = (
        "section Section1;\n"
        "shared TransformPipeline = let\n"
        "    // Load data from SQL Server\n"
        '    Source = Sql.Database("server.example.com", "AdventureWorks", [Query=\n'
        '        "SELECT * FROM Sales.SalesOrderHeader WHERE OrderDate >= \'2023-01-01\'"\n'
        "    ]),\n"
        "\n"
        "    /* Multi-line comment:\n"
        "       This step filters for active orders\n"
        "       and handles null values */\n"
        "    FilteredRows = Table.SelectRows(Source, each\n"
        "        [Status] <> null\n"
        '        and [Status] <> "Cancelled"\n'
        "        and [TotalDue] > 0\n"
        "    ),\n"
        "\n"
        "    // Nested let expression for custom transformation\n"
        "    AddedColumns = let\n"
        "        withTax = Table.AddColumn(FilteredRows, \"TaxRate\", each\n"
        "            if [TerritoryID] >= 1 and [TerritoryID] <= 5 then 0.08\n"
        "            else if [TerritoryID] >= 6 and [TerritoryID] <= 10 then 0.065\n"
        "            else 0.05\n"
        "        ),\n"
        '        withTotal = Table.AddColumn(withTax, "GrossTotal", each [TotalDue] * (1 + [TaxRate]), type number)\n'
        "    in\n"
        "        withTotal,\n"
        "\n"
        "    // Handle special characters in column renaming\n"
        '    RenamedColumns = Table.RenameColumns(AddedColumns, {{"SalesOrderID", "Order #"}, {"CustomerID", "Cust_ID"}}),\n'
        "\n"
        "    // Error handling with try/otherwise\n"
        "    SafeConversion = Table.TransformColumns(RenamedColumns, {\n"
        '        {"GrossTotal", each try Number.Round(_, 2) otherwise 0, type number}\n'
        "    }),\n"
        "\n"
        "    // Dynamic parameter usage\n"
        '    ParamFiltered = if #"Start Date" <> null\n'
        '        then Table.SelectRows(SafeConversion, each [OrderDate] >= #"Start Date")\n'
        "        else SafeConversion,\n"
        "\n"
        "    Result = Table.Buffer(ParamFiltered)\n"
        "in\n"
        "    Result;\n"
        'shared #"Start Date" = let\n'
        '    Source = #date(2023, 1, 1)\n'
        "in\n"
        "    Source;"
    )

    blob = build_datamashup_blob(m_code)
    inject_datamashup(path, blob)
    print(f"  Created: {path}")


def create_stress_test():
    """4. stress_test.xlsx - 20+ queries for stress testing."""
    path = os.path.join(OUTPUT_DIR, "stress_test.xlsx")
    create_base_xlsx(path)

    queries = []

    # Base data source queries (5)
    for i in range(1, 6):
        queries.append(
            f"shared DataSource{i} = let\n"
            f'    Source = Excel.CurrentWorkbook(){{[Name="Table{i}"]}}[Content],\n'
            f"    ChangedType = Table.TransformColumnTypes(Source, "
            f'{{{{\"ID\", Int64.Type}}, {{\"Value\", type number}}, {{\"Category\", type text}}}})\n'
            f"in\n"
            f"    ChangedType"
        )

    # Transform queries that reference base data (5)
    for i in range(1, 6):
        queries.append(
            f"shared Transform{i} = let\n"
            f"    Source = DataSource{i},\n"
            f"    Filtered = Table.SelectRows(Source, each [Value] > {i * 10}),\n"
            f'    Added = Table.AddColumn(Filtered, "Computed", each [Value] * {i})\n'
            f"in\n"
            f"    Added"
        )

    # Merge queries that combine transforms (5)
    for i in range(1, 6):
        j = (i % 5) + 1
        queries.append(
            f"shared Merge{i} = let\n"
            f"    Left = Transform{i},\n"
            f"    Right = Transform{j},\n"
            f'    Merged = Table.NestedJoin(Left, {{"ID"}}, Right, {{"ID"}}, "Joined", JoinKind.Inner),\n'
            f'    Expanded = Table.ExpandTableColumn(Merged, "Joined", {{"Value", "Computed"}})\n'
            f"in\n"
            f"    Expanded"
        )

    # Aggregation queries (5)
    for i in range(1, 6):
        queries.append(
            f"shared Aggregate{i} = let\n"
            f"    Source = Merge{i},\n"
            f'    Grouped = Table.Group(Source, {{"Category"}}, '
            f'{{{{\"Total\", each List.Sum([Value]), type number}}, '
            f'{{\"Count\", each Table.RowCount(_), Int64.Type}}}}),\n'
            f'    Sorted = Table.Sort(Grouped, {{{{\"Total\", Order.Descending}}}})\n'
            f"in\n"
            f"    Sorted"
        )

    # Final output queries (5)
    for i in range(1, 6):
        queries.append(
            f"shared Output{i} = let\n"
            f"    Source = Aggregate{i},\n"
            f"    TopN = Table.FirstN(Source, 10),\n"
            f'    Renamed = Table.RenameColumns(TopN, {{{{\"Total\", \"Grand Total {i}\"}}, {{\"Count\", \"Record Count {i}\"}}}})\n'
            f"in\n"
            f"    Renamed"
        )

    m_code = "section Section1;\n" + ";\n".join(queries) + ";"

    blob = build_datamashup_blob(m_code)
    inject_datamashup(path, blob)
    print(f"  Created: {path}")


def create_no_queries():
    """5. no_queries.xlsx - Regular Excel with no Power Query."""
    path = os.path.join(OUTPUT_DIR, "no_queries.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Name"
    ws["B1"] = "Value"
    for i in range(2, 12):
        ws[f"A{i}"] = f"Item {i - 1}"
        ws[f"B{i}"] = i * 10
    wb.save(path)
    print(f"  Created: {path}")


def create_edge_cases():
    """6. edge_cases.xlsx - Query names with #, underscores, Unicode, and tricky patterns."""
    path = os.path.join(OUTPUT_DIR, "edge_cases.xlsx")
    create_base_xlsx(path)

    m_code = (
        "section Section1;\n"
        # Query name with # (quoted identifier)
        'shared #"Revenue Report" = let\n'
        '    Source = Excel.CurrentWorkbook(){[Name="Revenue"]}[Content]\n'
        "in\n"
        "    Source;\n"
        # Query name with underscores
        "shared raw_data_import = let\n"
        '    Source = Csv.Document(File.Contents("data.csv")),\n'
        "    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true])\n"
        "in\n"
        "    Headers;\n"
        # Query name with multiple # and spaces
        'shared #"Year-to-Date (YTD) #1" = let\n'
        "    Source = raw_data_import,\n"
        "    Filtered = Table.SelectRows(Source, each [Year] = Date.Year(DateTime.LocalNow()))\n"
        "in\n"
        "    Filtered;\n"
        # Short / minimal query (just a value)
        "shared EmptyResult = let\n"
        "    Source = null\n"
        "in\n"
        "    Source;\n"
        # Query with very long code (repeated steps)
        "shared LongProcessing = let\n"
        '    Source = Excel.CurrentWorkbook(){[Name="BigTable"]}[Content],\n'
        + "".join(
            [
                f'    Step{i} = Table.AddColumn({"Source" if i == 1 else f"Step{i-1}"}, "Col{i}", each [Value] * {i}),\n'
                for i in range(1, 21)
            ]
        )
        + "    Final = Table.Buffer(Step20)\n"
        "in\n"
        "    Final;\n"
        # Query with Unicode in name
        'shared #"Donn\u00e9es_brutes" = let\n'
        '    Source = Excel.CurrentWorkbook(){[Name="Feuil1"]}[Content]\n'
        "in\n"
        "    Source;"
    )

    blob = build_datamashup_blob(m_code)
    inject_datamashup(path, blob)
    print(f"  Created: {path}")


# ============================================================
# Main
# ============================================================

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Creating test files...")
    create_simple_query()
    create_multi_query()
    create_complex_code()
    create_stress_test()
    create_no_queries()
    create_edge_cases()
    print(f"\nAll test files created in: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
