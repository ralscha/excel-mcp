package server

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"sort"
	"strings"

	"github.com/google/jsonschema-go/jsonschema"
	"github.com/modelcontextprotocol/go-sdk/mcp"
)

type CLITool struct {
	Name          string             `json:"name"`
	Description   string             `json:"description"`
	InputSchema   *jsonschema.Schema `json:"input_schema"`
	OutputExample any                `json:"output_example,omitempty"`
}

type cliToolDef struct {
	meta   CLITool
	invoke func(context.Context, []byte) (*mcp.CallToolResult, error)
}

func ListCLITools() []CLITool {
	tools := cliToolDefs(Config{})
	list := make([]CLITool, 0, len(tools))
	for _, def := range tools {
		list = append(list, def.meta)
	}
	sort.Slice(list, func(i, j int) bool {
		return list[i].Name < list[j].Name
	})
	return list
}

func GetCLITool(name string) (CLITool, bool) {
	def, ok := cliToolDefs(Config{})[strings.TrimSpace(name)]
	if !ok {
		return CLITool{}, false
	}
	return def.meta, true
}

func RunCLITool(ctx context.Context, cfg Config, toolName string, payload []byte) (string, error) {
	toolName = strings.TrimSpace(toolName)
	if toolName == "" {
		return "", fmt.Errorf("tool name is required")
	}

	def, ok := cliToolDefs(cfg)[toolName]
	if !ok {
		return "", fmt.Errorf("unknown tool %q", toolName)
	}

	result, err := def.invoke(ctx, payload)
	if err != nil {
		return "", err
	}
	if result == nil {
		return "", nil
	}

	text := cliResultText(result)
	if result.IsError {
		if text == "" {
			text = "tool execution failed"
		}
		return "", errors.New(text)
	}
	return text, nil
}

func cliToolDefs(cfg Config) map[string]cliToolDef {
	tc := &toolContext{pathMode: cfg.PathMode, logger: cfg.Logger}
	return map[string]cliToolDef{
		"apply_formula":            cliTool("apply_formula", "Apply an Excel formula to a cell", tc.applyFormula, schemaFor[formulaArgs], textExample("Applied formula =SUM(B2:B10) to Summary!B11")),
		"clear_range":              cliTool("clear_range", "Clear cell values in a range without shifting cells", tc.clearRange, schemaFor[optionalEndRangeArgs], textExample("Cleared values in Sheet1!A2:D20")),
		"copy_range":               cliTool("copy_range", "Copy a range of cells to another location", tc.copyRange, schemaFor[copyRangeArgs], textExample("Copied range Sheet1!A1:D10 to Archive!A1")),
		"copy_worksheet":           cliTool("copy_worksheet", "Copy a worksheet within a workbook", tc.copyWorksheet, schemaFor[copyWorksheetArgs], textExample("Copied worksheet Template to Template Copy")),
		"create_chart":             cliTool("create_chart", "Create a chart in a worksheet", tc.createChart, schemaFor[chartArgs], textExample("Created line chart on Sheet1 anchored at E2")),
		"create_pivot_table":       cliTool("create_pivot_table", "Create a pivot table in a worksheet", tc.createPivotTable, schemaFor[pivotArgs], textExample("Created pivot table on Summary anchored at H2")),
		"create_table":             cliTool("create_table", "Create a native Excel table from a specified data range", tc.createTable, schemaFor[tableArgs], textExample("Created table OrdersTable from range Orders!A1:D200")),
		"create_workbook":          cliTool("create_workbook", "Create a new Excel workbook", tc.createWorkbook, schemaFor[fileArgs], textExample("Workbook created at C:\\workbooks\\sales.xlsx")),
		"create_worksheet":         cliTool("create_worksheet", "Create a new worksheet in an existing workbook", tc.createWorksheet, schemaFor[createWorksheetArgs], textExample("Worksheet Summary created in C:\\workbooks\\sales.xlsx")),
		"delete_range":             cliTool("delete_range", "Delete a range of cells and shift remaining cells", tc.deleteRange, schemaFor[deleteRangeArgs], textExample("Deleted range Sheet1!B2:D4 and shifted cells up")),
		"delete_sheet_columns":     cliTool("delete_sheet_columns", "Delete one or more columns starting at the specified column", tc.deleteSheetColumns, schemaFor[columnArgs], textExample("Deleted 2 columns from Sheet1 starting at column 3")),
		"delete_sheet_rows":        cliTool("delete_sheet_rows", "Delete one or more rows starting at the specified row", tc.deleteSheetRows, schemaFor[rowArgs], textExample("Deleted 3 rows from Sheet1 starting at row 8")),
		"delete_worksheet":         cliTool("delete_worksheet", "Delete a worksheet from a workbook", tc.deleteWorksheet, schemaFor[createWorksheetArgs], textExample("Worksheet Archive deleted from C:\\workbooks\\sales.xlsx")),
		"describe_workbook":        cliTool("describe_workbook", "Describe the workbook structure and artifacts", tc.describeWorkbook, schemaFor[describeWorkbookArgs], workbookDescriptionExample()),
		"filter_rows":              cliTool("filter_rows", "Filter tabular rows from a worksheet range", tc.filterRows, schemaFor[filterRowsArgs], filterRowsResultExample()),
		"find_in_workbook":         cliTool("find_in_workbook", "Search workbook cell values or formulas", tc.findInWorkbook, schemaFor[findInWorkbookArgs], findResultExample()),
		"format_range":             cliTool("format_range", "Apply formatting to a range of cells", tc.formatRange, schemaFor[formatRangeArgs], textExample("Formatted range Sheet1!A1:D1")),
		"get_data_validation_info": cliTool("get_data_validation_info", "Get data validation rules and metadata for a worksheet", tc.getDataValidationInfo, schemaFor[validationInfoArgs], dataValidationInfoExample()),
		"get_merged_cells":         cliTool("get_merged_cells", "Get merged cells in a worksheet", tc.getMergedCells, schemaFor[validationInfoArgs], map[string]any{"sheet_name": "Summary", "merged_ranges": []string{"A1:C1", "D4:E4"}}),
		"get_sheet_schema":         cliTool("get_sheet_schema", "Infer a worksheet schema and sample values", tc.getSheetSchema, schemaFor[getSheetSchemaArgs], sheetSchemaExample()),
		"get_workbook_metadata":    cliTool("get_workbook_metadata", "Get workbook metadata, including sheets and ranges", tc.getWorkbookMetadata, schemaFor[metadataArgs], workbookMetadataExample()),
		"insert_columns":           cliTool("insert_columns", "Insert one or more columns starting at the specified column", tc.insertColumns, schemaFor[columnArgs], textExample("Inserted 2 columns into Sheet1 starting at column 4")),
		"insert_rows":              cliTool("insert_rows", "Insert one or more rows starting at the specified row", tc.insertRows, schemaFor[rowArgs], textExample("Inserted 1 row into Sheet1 starting at row 5")),
		"list_charts":              cliTool("list_charts", "List charts in a worksheet or across a workbook", tc.listCharts, schemaFor[listChartsArgs], listChartsResultExample()),
		"merge_cells":              cliTool("merge_cells", "Merge a range of cells", tc.mergeCells, schemaFor[rangeArgs], textExample("Merged cells Summary!A1:C1")),
		"read_data_from_excel":     cliTool("read_data_from_excel", "Read data from an Excel worksheet", tc.readDataFromExcel, schemaFor[readDataArgs], readResultExample()),
		"rename_worksheet":         cliTool("rename_worksheet", "Rename a worksheet in a workbook", tc.renameWorksheet, schemaFor[renameWorksheetArgs], textExample("Renamed worksheet Q1 to Q1 2026")),
		"set_column_widths":        cliTool("set_column_widths", "Set column widths or auto-fit columns to their content", tc.setColumnWidths, mustSetColumnWidthsSchema, textExample("Updated column widths on Sheet1")),
		"set_row_heights":          cliTool("set_row_heights", "Set worksheet row heights", tc.setRowHeights, schemaFor[setRowHeightsArgs], textExample("Updated 3 row heights on Sheet1")),
		"sort_range":               cliTool("sort_range", "Sort a tabular range by one or more columns", tc.sortRange, schemaFor[sortRangeArgs], sortRangeResultExample()),
		"unmerge_cells":            cliTool("unmerge_cells", "Unmerge a previously merged range of cells", tc.unmergeCells, schemaFor[rangeArgs], textExample("Unmerged cells Summary!A1:C1")),
		"upsert_rows":              cliTool("upsert_rows", "Update matching table rows or append new rows within a table range", tc.upsertRows, schemaFor[upsertRowsArgs], upsertRowsResultExample()),
		"validate_excel_range":     cliTool("validate_excel_range", "Validate that a range exists and is properly formatted", tc.validateExcelRange, schemaFor[optionalEndRangeArgs], textExample("Range Sheet1!A1:D20 is valid")),
		"validate_formula_syntax":  cliTool("validate_formula_syntax", "Validate Excel formula syntax without applying it", tc.validateFormulaSyntax, schemaFor[validateFormulaArgs], textExample("formula syntax looks valid for Summary!B11")),
		"write_data_to_excel":      cliTool("write_data_to_excel", "Write data to an Excel worksheet", tc.writeDataToExcel, schemaFor[writeDataArgs], textExample("Wrote 25 rows to Orders starting at A1")),
	}
}

func cliTool[T any](name, description string, handler func(context.Context, *mcp.CallToolRequest, T) (*mcp.CallToolResult, any, error), schemaFn func() *jsonschema.Schema, outputExample any) cliToolDef {
	return cliToolDef{
		meta: CLITool{Name: name, Description: description, InputSchema: schemaFn(), OutputExample: outputExample},
		invoke: func(ctx context.Context, payload []byte) (*mcp.CallToolResult, error) {
			args, err := decodeCLIArgs[T](payload)
			if err != nil {
				return nil, err
			}
			result, _, err := handler(ctx, nil, args)
			return result, err
		},
	}
}

func schemaFor[T any]() *jsonschema.Schema {
	schema, err := jsonschema.For[T](nil)
	if err != nil {
		panic(fmt.Errorf("build CLI schema: %w", err))
	}
	return schema
}

func textExample(message string) string {
	return message
}

func workbookMetadataExample() map[string]any {
	return map[string]any{
		"filepath": "C:\\workbooks\\sales.xlsx",
		"sheets": []map[string]any{
			{"name": "Orders", "range": "A1:D200", "rows": 200, "cols": 4},
			{"name": "Summary", "range": "A1:H20", "rows": 20, "cols": 8},
		},
	}
}

func workbookDescriptionExample() map[string]any {
	return map[string]any{
		"filepath": "C:\\workbooks\\sales.xlsx",
		"sheets": []map[string]any{
			{
				"name":          "Orders",
				"range":         "A1:D200",
				"tables":        []map[string]any{{"name": "OrdersTable", "range": "A1:D200", "style_name": "TableStyleMedium9"}},
				"charts":        []map[string]any{{"sheet_name": "Orders", "anchor_cell": "F2", "title": "Revenue by Region", "chart_type": "bar"}},
				"pivot_tables":  []map[string]any{{"name": "PivotTable1", "data_range": "A1:D200", "pivot_table_range": "H2:K10", "rows": []string{"Region"}, "values": []map[string]any{{"data": "Revenue", "subtotal": "sum"}}}},
				"merged_ranges": []string{"A1:D1"},
			},
		},
	}
}

func listChartsResultExample() map[string]any {
	return map[string]any{
		"filepath":   "C:\\workbooks\\sales.xlsx",
		"sheet_name": "Dashboard",
		"charts": []map[string]any{
			{"sheet_name": "Dashboard", "anchor_cell": "B2", "title": "Revenue by Quarter", "chart_type": "column"},
		},
	}
}

func sheetSchemaExample() map[string]any {
	return map[string]any{
		"filepath":   "C:\\workbooks\\sales.xlsx",
		"sheet_name": "Orders",
		"range":      "A1:D200",
		"header_row": 1,
		"row_count":  199,
		"columns": []map[string]any{
			{"name": "OrderID", "column": "A", "inferred_type": "string", "blank_count": 0, "sample_values": []string{"A-1001", "A-1002"}},
			{"name": "Revenue", "column": "D", "inferred_type": "number", "blank_count": 1, "sample_values": []string{"125.5", "88"}},
		},
	}
}

func findResultExample() map[string]any {
	return map[string]any{
		"filepath":    "C:\\workbooks\\sales.xlsx",
		"query":       "Northwind",
		"search_type": "text",
		"match_mode":  "contains",
		"matches": []map[string]any{
			{"sheet_name": "Orders", "cell": "B12", "value": "Northwind", "context": []map[string]any{{"cell": "A12", "value": "A-1001"}}},
		},
	}
}

func readResultExample() map[string]any {
	return map[string]any{
		"filepath":   "C:\\workbooks\\sales.xlsx",
		"sheet_name": "Orders",
		"range":      "A1:D5",
		"rows": []map[string]any{
			{"OrderID": "A-1001", "Customer": "Northwind", "Amount": 125.5, "Status": "Paid"},
			{"OrderID": "A-1002", "Customer": "Contoso", "Amount": 88, "Status": "Open"},
		},
	}
}

func filterRowsResultExample() map[string]any {
	return map[string]any{
		"filepath":   "C:\\workbooks\\sales.xlsx",
		"sheet_name": "Orders",
		"range":      "A1:D200",
		"has_header": true,
		"filters": []map[string]any{
			{"column": "Status", "operator": "equals", "value": "Open"},
		},
		"rows": []map[string]any{
			{"OrderID": "A-1002", "Customer": "Contoso", "Amount": 88, "Status": "Open"},
		},
	}
}

func sortRangeResultExample() map[string]any {
	return map[string]any{
		"filepath":   "C:\\workbooks\\sales.xlsx",
		"sheet_name": "Orders",
		"range":      "A1:D200",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "Region"},
			{"column": "Revenue", "descending": true},
		},
	}
}

func upsertRowsResultExample() map[string]any {
	return map[string]any{
		"filepath":       "C:\\workbooks\\sales.xlsx",
		"sheet_name":     "Orders",
		"range":          "A1:D500",
		"key_columns":    []string{"OrderID"},
		"updated_count":  1,
		"inserted_count": 1,
		"skipped_count":  0,
		"results": []map[string]any{
			{"key": map[string]any{"OrderID": "A-1001"}, "action": "updated", "row_number": 2},
			{"key": map[string]any{"OrderID": "A-1003"}, "action": "inserted", "row_number": 4},
		},
	}
}

func dataValidationInfoExample() []map[string]any {
	return []map[string]any{
		{"sqref": "C2:C100", "type": "whole", "operator": "between", "formula1": "1", "formula2": "10", "error_title": "Invalid quantity"},
	}
}

func decodeCLIArgs[T any](payload []byte) (T, error) {
	var args T
	trimmed := strings.TrimSpace(string(payload))
	if trimmed == "" {
		trimmed = "{}"
	}
	if err := json.Unmarshal([]byte(trimmed), &args); err != nil {
		return args, fmt.Errorf("decode tool input: %w", err)
	}
	return args, nil
}

func cliResultText(result *mcp.CallToolResult) string {
	parts := make([]string, 0, len(result.Content))
	for _, content := range result.Content {
		if text, ok := content.(*mcp.TextContent); ok {
			parts = append(parts, text.Text)
		}
	}
	return strings.Join(parts, "\n")
}
