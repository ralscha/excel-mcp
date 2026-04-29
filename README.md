# Excel MCP Server

Excel MCP Server is a Go MCP server and CLI tool for creating and editing Excel workbooks with [`github.com/xuri/excelize/v2`](https://github.com/qax-os/excelize).

It exposes a series of tools to inspect and manipulate Excel workbooks.

## Tools

### Workbook inspection

- `get_workbook_metadata`: list workbook sheets and optional used-range metadata.
- `describe_workbook`: return richer workbook structure including optional ranges, tables, charts, pivot tables, named ranges, merged cells, and data validations.
- `list_charts`: list charts for a single host sheet or across the entire workbook, with optional source-sheet filtering.
- `get_sheet_schema`: infer column names, types, blank counts, and sample values for a worksheet range.
- `find_in_workbook`: search workbook text or formulas with exact, contains, or regex matching, optional context cells, sheet scoping, and result limits.

### Table-style data operations

- `write_data_to_excel`: write object rows to a worksheet starting at a given cell.
- `read_data_from_excel`: read tabular worksheet data into JSON rows.
- `filter_rows`: return rows from a range that match one or more filters.
- `sort_range`: sort a tabular range by one or more columns.
- `upsert_rows`: update matching rows by key or append new rows within a declared table range.
- `create_table`: create an Excel table from a range.
- `create_pivot_table`: create a pivot table from worksheet data.
- `create_chart`: create a chart from worksheet data.

### Worksheet and range operations

- `create_workbook`: create a new Excel workbook.
- `create_worksheet`: create a new worksheet in an existing workbook.
- `copy_worksheet`: copy a worksheet within a workbook.
- `delete_worksheet`: delete a worksheet from a workbook.
- `rename_worksheet`: rename a worksheet in a workbook.
- `copy_range`: copy a range of cells to another location.
- `delete_range`: delete a range of cells and shift remaining cells.
- `clear_range`: clear cell values in a range without shifting surrounding cells.
- `insert_rows`: insert one or more rows starting at a specified row.
- `insert_columns`: insert one or more columns starting at a specified column.
- `delete_sheet_rows`: delete one or more rows starting at a specified row.
- `delete_sheet_columns`: delete one or more columns starting at a specified column.
- `set_column_widths`: set explicit column widths or auto-fit columns to their content.
- `set_row_heights`: set worksheet row heights.
- `format_range`: apply formatting to a range of cells.
- `merge_cells`: merge a range of cells.
- `unmerge_cells`: unmerge a previously merged range of cells.
- `get_merged_cells`: list merged-cell ranges in a worksheet.
- `validate_excel_range`: validate that a worksheet range exists and is properly formatted.
- `get_data_validation_info`: return data validation rules and metadata for a worksheet.
- `apply_formula`: apply an Excel formula to a cell.
- `validate_formula_syntax`: validate Excel formula syntax without applying it.

## Download 

Download the latest release from the [releases page](https://github.com/ralscha/excel-mcp/releases).

## Run from source

From the repository:

CLI:

```bash
go run ./cmd/excel-mcp cli list-tools
go run ./cmd/excel-mcp cli list-tools --json
go run ./cmd/excel-mcp cli tool-info create_workbook
go run ./cmd/excel-mcp cli tool-info --json create_workbook
```

MCP:

```bash
go run ./cmd/excel-mcp stdio
go run ./cmd/excel-mcp streamable-http
```

Build an executable:

```bash
go build ./cmd/excel-mcp
```

## Environment

### `stdio`

Tool calls must use absolute file paths.

### `streamable-http`

- `EXCEL_MCP_SERVER_PORT`: HTTP port for the `/mcp` endpoint. Defaults to `8000`.
- `EXCEL_FILES_PATH`: Root directory for relative workbook paths. Defaults to `./excel_files`.

Example:

```bash
export EXCEL_MCP_SERVER_PORT=8017
export EXCEL_FILES_PATH=/excel-files
go run ./cmd/excel-mcp streamable-http
```

### `cli`

- Run a tool directly from the shell with JSON input that matches the MCP tool arguments.
- By default, workbook paths must be absolute.
- Use `--rooted` to resolve relative workbook paths under `EXCEL_FILES_PATH`.
- Tool names match the MCP tool names exactly and are intended to remain stable long-term.
- Use `list-tools --json` for machine-readable discovery and `tool-info [--json] <tool-name>` for per-tool descriptions, input schemas, and output examples.

Examples:

```bash
go run ./cmd/excel-mcp cli list-tools
go run ./cmd/excel-mcp cli list-tools --json
go run ./cmd/excel-mcp cli tool-info create_workbook
go run ./cmd/excel-mcp cli tool-info --json create_workbook

cat <<'EOF' | go run ./cmd/excel-mcp cli read_data_from_excel --input -
{"filepath":"/workbooks/sales.xlsx","sheet_name":"Sheet1","start_cell":"A1","preview_only":true}
EOF

cat <<'EOF' | go run ./cmd/excel-mcp cli create_workbook --rooted --input -
{"filepath":"reports/q1.xlsx"}
EOF
```

## Client Config Examples

### Claude Desktop

```json
{
  "mcpServers": {
    "excel": {
      "command": "excel-mcp",
      "args": ["stdio"]
    }
  }
}
```

### VS Code / Cursor using `go run`

```json
{
  "mcpServers": {
    "excel": {
      "command": "go",
      "args": ["run", "./cmd/excel-mcp", "stdio"]
    }
  }
}
```

### HTTP client

```json
{
  "mcpServers": {
    "excel": {
      "url": "http://localhost:8000/mcp"
    }
  }
}
```

## Notes

- `stdio` mode requires absolute paths.
- `cli` mode defaults to absolute paths and can opt into rooted relative paths with `--rooted`.
- HTTP mode rejects absolute paths and resolves relative paths under `EXCEL_FILES_PATH`.
- `cli` mode supports `list-tools --json` for machine-readable discovery, `tool-info [--json] <tool-name>` for per-tool schema and output-example inspection, and `--input -` for stdin-fed JSON payloads.
- `cli` tool names intentionally match MCP tool names exactly.
- `describe_workbook` can include table, chart, pivot, named-range, merged-range, and validation metadata.
- `list_charts` supports workbook-wide listing and optional `source_sheet` filtering.
- `get_sheet_schema` defaults `sample_size` to `3` when omitted.
- `find_in_workbook` supports `text` and `formula` search types, `contains`, `exact`, and `regex` match modes, optional context, and `max_results`.
- `filter_rows` is a non-mutating table query tool. Filters are combined with AND semantics.
- `filter_rows` supports `equals`, `contains`, `gt`, `gte`, `lt`, `lte`, and `regex` operators.
- `sort_range` accepts header-name or 1-based column-index sort keys, with optional header preservation via `has_header`.
- `upsert_rows` is table-scoped and bounded by the declared range. It updates rows matched by `key_columns` and can append within the same range when `insert_if_missing=true`.
- `upsert_rows` requires explicit `match` and `values` objects per row, rejects key-column mutation, and fails when multiple existing rows match the same key.
- `create_chart` supports `line`, `column`, `bar`, `pie`, `doughnut`, `scatter`, and `area`.
- `create_pivot_table` supports `sum`, `count`, `average`, `avg`, `mean`, `max`, and `min` aggregation inputs.
- `create_table` requires non-empty, unique header cells.
- `delete_range` supports `up` and `left` shift directions.
- `clear_range` clears cell values in-place without shifting surrounding cells.
- `format_range` accepts built-in or custom `number_format` values, optional `protection` settings (`locked`, `hidden`), and an optional single `conditional_format` rule with an optional nested style.

## Example Tool Inputs

Filter rows from a table-like range:

```json
{
  "filepath": "/workbooks/sales.xlsx",
  "sheet_name": "Sheet1",
  "range": "A1:D200",
  "has_header": true,
  "filters": [
    {"column": "Status", "operator": "equals", "value": "Open"},
    {"column": "Revenue", "operator": "gte", "value": "1000"}
  ]
}
```

Sort a tabular range:

```json
{
  "filepath": "/workbooks/sales.xlsx",
  "sheet_name": "Sheet1",
  "range": "A1:D200",
  "has_header": true,
  "sort_keys": [
    {"column": "Region"},
    {"column": "Revenue", "descending": true}
  ]
}
```

Upsert rows into a table range:

```json
{
  "filepath": "/workbooks/sales.xlsx",
  "sheet_name": "Orders",
  "range": "A1:D500",
  "key_columns": ["OrderID"],
  "insert_if_missing": true,
  "rows": [
    {
      "match": {"OrderID": "A-1001"},
      "values": {"Status": "Paid", "Amount": 125.5}
    },
    {
      "match": {"OrderID": "A-1003"},
      "values": {"Customer": "Northwind", "Amount": 88, "Status": "Open"}
    }
  ]
}
```

Use `upsert_rows` as an append by providing a new key and enabling inserts:

```json
{
  "filepath": "/workbooks/sales.xlsx",
  "sheet_name": "Orders",
  "range": "A1:D500",
  "key_columns": ["OrderID"],
  "insert_if_missing": true,
  "rows": [
    {
      "match": {"OrderID": "A-2001"},
      "values": {"Customer": "Contoso", "Amount": 42, "Status": "Open"}
    }
  ]
}
```

Typical `upsert_rows` result shape:

```json
{
  "filepath": "/workbooks/sales.xlsx",
  "sheet_name": "Orders",
  "range": "A1:D500",
  "key_columns": ["OrderID"],
  "updated_count": 1,
  "inserted_count": 1,
  "skipped_count": 1,
  "results": [
    {
      "key": {"OrderID": "A-1001"},
      "action": "updated",
      "row_number": 2
    },
    {
      "key": {"OrderID": "A-1003"},
      "action": "inserted",
      "row_number": 4
    },
    {
      "key": {"OrderID": "A-4040"},
      "action": "skipped",
      "reason": "no match and insert_if_missing=false"
    }
  ]
}
```

Describe workbook structure with richer metadata:

```json
{
  "filepath": "/workbooks/sales.xlsx",
  "include_ranges": true,
  "include_tables": true,
  "include_charts": true,
  "include_pivots": true,
  "include_names": true,
  "include_merged": true,
  "include_validation": true
}
```

## Validation

```bash
go test ./...
```