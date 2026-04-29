package server

import (
	"context"
	"encoding/json"
	"fmt"
	"log/slog"
	"math"

	"excel-mcp/internal/excel"

	"github.com/google/jsonschema-go/jsonschema"
	"github.com/modelcontextprotocol/go-sdk/mcp"
)

type Config struct {
	PathMode PathMode
	Logger   *slog.Logger
}

type toolContext struct {
	pathMode PathMode
	logger   *slog.Logger
}

func New(cfg Config) *mcp.Server {
	server := mcp.NewServer(&mcp.Implementation{Name: "excel-mcp", Version: "0.1.0"}, nil)
	tc := &toolContext{pathMode: cfg.PathMode, logger: cfg.Logger}

	mcp.AddTool(server, &mcp.Tool{Name: "create_workbook", Description: "Create a new Excel workbook"}, tc.createWorkbook)
	mcp.AddTool(server, &mcp.Tool{Name: "create_worksheet", Description: "Create a new worksheet in an existing workbook"}, tc.createWorksheet)
	mcp.AddTool(server, &mcp.Tool{Name: "describe_workbook", Description: "Describe the workbook structure and artifacts"}, tc.describeWorkbook)
	mcp.AddTool(server, &mcp.Tool{Name: "list_charts", Description: "List charts in a worksheet or across a workbook"}, tc.listCharts)
	mcp.AddTool(server, &mcp.Tool{Name: "get_sheet_schema", Description: "Infer a worksheet schema and sample values"}, tc.getSheetSchema)
	mcp.AddTool(server, &mcp.Tool{Name: "find_in_workbook", Description: "Search workbook cell values or formulas"}, tc.findInWorkbook)
	mcp.AddTool(server, &mcp.Tool{Name: "get_workbook_metadata", Description: "Get workbook metadata, including sheets and ranges"}, tc.getWorkbookMetadata)
	mcp.AddTool(server, &mcp.Tool{Name: "write_data_to_excel", Description: "Write data to an Excel worksheet"}, tc.writeDataToExcel)
	mcp.AddTool(server, &mcp.Tool{Name: "read_data_from_excel", Description: "Read data from an Excel worksheet"}, tc.readDataFromExcel)
	mcp.AddTool(server, &mcp.Tool{Name: "filter_rows", Description: "Filter tabular rows from a worksheet range"}, tc.filterRows)
	mcp.AddTool(server, &mcp.Tool{Name: "format_range", Description: "Apply formatting to a range of cells"}, tc.formatRange)
	mcp.AddTool(server, &mcp.Tool{Name: "merge_cells", Description: "Merge a range of cells"}, tc.mergeCells)
	mcp.AddTool(server, &mcp.Tool{Name: "unmerge_cells", Description: "Unmerge a previously merged range of cells"}, tc.unmergeCells)
	mcp.AddTool(server, &mcp.Tool{Name: "get_merged_cells", Description: "Get merged cells in a worksheet"}, tc.getMergedCells)
	mcp.AddTool(server, &mcp.Tool{Name: "apply_formula", Description: "Apply an Excel formula to a cell"}, tc.applyFormula)
	mcp.AddTool(server, &mcp.Tool{Name: "validate_formula_syntax", Description: "Validate Excel formula syntax without applying it"}, tc.validateFormulaSyntax)
	mcp.AddTool(server, &mcp.Tool{Name: "create_chart", Description: "Create a chart in a worksheet"}, tc.createChart)
	mcp.AddTool(server, &mcp.Tool{Name: "create_pivot_table", Description: "Create a pivot table in a worksheet"}, tc.createPivotTable)
	mcp.AddTool(server, &mcp.Tool{Name: "create_table", Description: "Create a native Excel table from a specified data range"}, tc.createTable)
	mcp.AddTool(server, &mcp.Tool{Name: "copy_worksheet", Description: "Copy a worksheet within a workbook"}, tc.copyWorksheet)
	mcp.AddTool(server, &mcp.Tool{Name: "delete_worksheet", Description: "Delete a worksheet from a workbook"}, tc.deleteWorksheet)
	mcp.AddTool(server, &mcp.Tool{Name: "rename_worksheet", Description: "Rename a worksheet in a workbook"}, tc.renameWorksheet)
	mcp.AddTool(server, &mcp.Tool{Name: "copy_range", Description: "Copy a range of cells to another location"}, tc.copyRange)
	mcp.AddTool(server, &mcp.Tool{Name: "sort_range", Description: "Sort a tabular range by one or more columns"}, tc.sortRange)
	mcp.AddTool(server, &mcp.Tool{Name: "upsert_rows", Description: "Update matching table rows or append new rows within a table range"}, tc.upsertRows)
	mcp.AddTool(server, &mcp.Tool{Name: "delete_range", Description: "Delete a range of cells and shift remaining cells"}, tc.deleteRange)
	mcp.AddTool(server, &mcp.Tool{Name: "clear_range", Description: "Clear cell values in a range without shifting cells"}, tc.clearRange)
	mcp.AddTool(server, &mcp.Tool{Name: "validate_excel_range", Description: "Validate that a range exists and is properly formatted"}, tc.validateExcelRange)
	mcp.AddTool(server, &mcp.Tool{Name: "get_data_validation_info", Description: "Get data validation rules and metadata for a worksheet"}, tc.getDataValidationInfo)
	mcp.AddTool(server, &mcp.Tool{Name: "insert_rows", Description: "Insert one or more rows starting at the specified row"}, tc.insertRows)
	mcp.AddTool(server, &mcp.Tool{Name: "insert_columns", Description: "Insert one or more columns starting at the specified column"}, tc.insertColumns)
	mcp.AddTool(server, &mcp.Tool{Name: "delete_sheet_rows", Description: "Delete one or more rows starting at the specified row"}, tc.deleteSheetRows)
	mcp.AddTool(server, &mcp.Tool{Name: "delete_sheet_columns", Description: "Delete one or more columns starting at the specified column"}, tc.deleteSheetColumns)
	mcp.AddTool(server, &mcp.Tool{Name: "set_column_widths", Description: "Set column widths or auto-fit columns to their content", InputSchema: mustSetColumnWidthsSchema()}, tc.setColumnWidths)
	mcp.AddTool(server, &mcp.Tool{Name: "set_row_heights", Description: "Set worksheet row heights"}, tc.setRowHeights)

	return server
}

func (tc *toolContext) resolve(inputPath string) (string, error) {
	return ResolvePath(tc.pathMode, inputPath)
}

func textResult(message string) *mcp.CallToolResult {
	return &mcp.CallToolResult{Content: []mcp.Content{&mcp.TextContent{Text: message}}}
}

func jsonResult(value any) (*mcp.CallToolResult, error) {
	data, err := json.MarshalIndent(value, "", "  ")
	if err != nil {
		return nil, err
	}
	return textResult(string(data)), nil
}

func toolError(err error) *mcp.CallToolResult {
	return &mcp.CallToolResult{IsError: true, Content: []mcp.Content{&mcp.TextContent{Text: err.Error()}}}
}

func mustSetColumnWidthsSchema() *jsonschema.Schema {
	schema, err := jsonschema.For[setColumnWidthsArgs](nil)
	if err != nil {
		panic(fmt.Errorf("build set_column_widths schema: %w", err))
	}
	trueValue := any(true)
	schema.AnyOf = []*jsonschema.Schema{
		{Required: []string{"widths"}},
		{
			Properties: map[string]*jsonschema.Schema{
				"auto_fit": {Const: &trueValue},
			},
			Required: []string{"auto_fit"},
		},
	}
	return schema
}

type fileArgs struct {
	Filepath string `json:"filepath" jsonschema:"Path to the Excel workbook"`
}

type createWorksheetArgs struct {
	Filepath  string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string `json:"sheet_name" jsonschema:"Name of the worksheet"`
}

type metadataArgs struct {
	Filepath      string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	IncludeRanges bool   `json:"include_ranges,omitempty" jsonschema:"Include row and column range metadata"`
}

type describeWorkbookArgs struct {
	Filepath          string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	IncludeRanges     bool   `json:"include_ranges,omitempty" jsonschema:"Include row and column range metadata"`
	IncludeTables     bool   `json:"include_tables,omitempty" jsonschema:"Include table metadata"`
	IncludeCharts     bool   `json:"include_charts,omitempty" jsonschema:"Include chart metadata"`
	IncludePivots     bool   `json:"include_pivots,omitempty" jsonschema:"Include pivot table metadata"`
	IncludeNames      bool   `json:"include_names,omitempty" jsonschema:"Include named range metadata"`
	IncludeMerged     bool   `json:"include_merged,omitempty" jsonschema:"Include merged range metadata"`
	IncludeValidation bool   `json:"include_validation,omitempty" jsonschema:"Include data validation metadata"`
}

type getSheetSchemaArgs struct {
	Filepath   string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName  string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	StartCell  string `json:"start_cell,omitempty" jsonschema:"Top-left cell for schema inference"`
	EndCell    string `json:"end_cell,omitempty" jsonschema:"Optional bottom-right cell for schema inference"`
	HeaderRow  int    `json:"header_row,omitempty" jsonschema:"Worksheet row number containing headers"`
	SampleSize int    `json:"sample_size,omitempty" jsonschema:"Maximum sample values per column"`
}

type listChartsArgs struct {
	Filepath    string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName   string `json:"sheet_name,omitempty" jsonschema:"Optional worksheet containing the charts"`
	SourceSheet string `json:"source_sheet,omitempty" jsonschema:"Optional worksheet referenced by chart source ranges"`
}

type findInWorkbookArgs struct {
	Filepath      string   `json:"filepath" jsonschema:"Path to the Excel workbook"`
	Query         string   `json:"query" jsonschema:"Text or formula fragment to search for"`
	Sheets        []string `json:"sheets,omitempty" jsonschema:"Optional sheet names to limit the search to"`
	MatchMode     string   `json:"match_mode,omitempty" jsonschema:"contains or exact"`
	SearchType    string   `json:"search_type,omitempty" jsonschema:"text or formula"`
	CaseSensitive bool     `json:"case_sensitive,omitempty" jsonschema:"Enable case-sensitive matching"`
	ContextRows   int      `json:"context_rows,omitempty" jsonschema:"Number of surrounding rows to include in each match context"`
	ContextCols   int      `json:"context_cols,omitempty" jsonschema:"Number of surrounding columns to include in each match context"`
	MaxResults    int      `json:"max_results,omitempty" jsonschema:"Maximum number of matches to return"`
}

type writeDataArgs struct {
	Filepath  string           `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string           `json:"sheet_name" jsonschema:"Name of the worksheet"`
	Data      []map[string]any `json:"data" jsonschema:"Array of objects to write"`
	StartCell string           `json:"start_cell,omitempty" jsonschema:"Top-left cell for the header row"`
}

type readDataArgs struct {
	Filepath    string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName   string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	StartCell   string `json:"start_cell,omitempty" jsonschema:"Top-left cell to start reading from"`
	EndCell     string `json:"end_cell,omitempty" jsonschema:"Optional bottom-right cell to stop reading at"`
	PreviewOnly bool   `json:"preview_only,omitempty" jsonschema:"Return at most 10 rows"`
}

type filterArgs struct {
	Column   string `json:"column" jsonschema:"Header name or 1-based column index within the range"`
	Operator string `json:"operator" jsonschema:"Filter operator: equals, contains, gt, gte, lt, lte, regex"`
	Value    string `json:"value" jsonschema:"Filter comparison value"`
}

type filterRowsArgs struct {
	Filepath  string       `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string       `json:"sheet_name" jsonschema:"Name of the worksheet"`
	Range     string       `json:"range" jsonschema:"Range to filter, such as A1:D20"`
	Filters   []filterArgs `json:"filters" jsonschema:"Filter predicates applied with AND semantics"`
	HasHeader bool         `json:"has_header,omitempty" jsonschema:"Treat the first row of the range as a header row"`
}

type rangeArgs struct {
	Filepath  string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	StartCell string `json:"start_cell" jsonschema:"Start cell"`
	EndCell   string `json:"end_cell" jsonschema:"End cell"`
}

type formatRangeArgs struct {
	Filepath     string                 `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName    string                 `json:"sheet_name" jsonschema:"Name of the worksheet"`
	StartCell    string                 `json:"start_cell" jsonschema:"Start cell"`
	EndCell      string                 `json:"end_cell,omitempty" jsonschema:"Optional end cell; defaults to start_cell"`
	Bold         bool                   `json:"bold,omitempty"`
	Italic       bool                   `json:"italic,omitempty"`
	Underline    bool                   `json:"underline,omitempty"`
	FontSize     int                    `json:"font_size,omitempty"`
	FontColor    string                 `json:"font_color,omitempty"`
	BGColor      string                 `json:"bg_color,omitempty"`
	BorderStyle  string                 `json:"border_style,omitempty"`
	BorderColor  string                 `json:"border_color,omitempty"`
	NumberFormat any                    `json:"number_format,omitempty" jsonschema:"Built-in Excel number format id or custom format string"`
	Alignment    string                 `json:"alignment,omitempty"`
	WrapText     bool                   `json:"wrap_text,omitempty"`
	MergeCells   bool                   `json:"merge_cells,omitempty"`
	Protection   *formatProtectionArgs  `json:"protection,omitempty" jsonschema:"Optional cell protection settings"`
	Conditional  *conditionalFormatArgs `json:"conditional_format,omitempty" jsonschema:"Optional conditional formatting rule"`
}

type formatProtectionArgs struct {
	Locked bool `json:"locked,omitempty" jsonschema:"Lock cells when worksheet protection is enabled"`
	Hidden bool `json:"hidden,omitempty" jsonschema:"Hide formulas when worksheet protection is enabled"`
}

type conditionalFormatStyleArgs struct {
	Bold         bool                  `json:"bold,omitempty"`
	Italic       bool                  `json:"italic,omitempty"`
	Underline    bool                  `json:"underline,omitempty"`
	FontSize     int                   `json:"font_size,omitempty"`
	FontColor    string                `json:"font_color,omitempty"`
	BGColor      string                `json:"bg_color,omitempty"`
	BorderStyle  string                `json:"border_style,omitempty"`
	BorderColor  string                `json:"border_color,omitempty"`
	NumberFormat any                   `json:"number_format,omitempty"`
	Alignment    string                `json:"alignment,omitempty"`
	WrapText     bool                  `json:"wrap_text,omitempty"`
	Protection   *formatProtectionArgs `json:"protection,omitempty"`
}

type conditionalFormatArgs struct {
	Type           string                      `json:"type" jsonschema:"Conditional format type, such as cell, 2_color_scale, or data_bar"`
	AboveAverage   bool                        `json:"above_average,omitempty"`
	Percent        bool                        `json:"percent,omitempty"`
	Criteria       string                      `json:"criteria,omitempty" jsonschema:"Conditional format criteria, such as >, <, ="`
	Value          string                      `json:"value,omitempty" jsonschema:"Comparison value for the rule"`
	MinType        string                      `json:"min_type,omitempty"`
	MidType        string                      `json:"mid_type,omitempty"`
	MaxType        string                      `json:"max_type,omitempty"`
	MinValue       string                      `json:"min_value,omitempty"`
	MidValue       string                      `json:"mid_value,omitempty"`
	MaxValue       string                      `json:"max_value,omitempty"`
	MinColor       string                      `json:"min_color,omitempty"`
	MidColor       string                      `json:"mid_color,omitempty"`
	MaxColor       string                      `json:"max_color,omitempty"`
	BarColor       string                      `json:"bar_color,omitempty"`
	BarBorderColor string                      `json:"bar_border_color,omitempty"`
	BarDirection   string                      `json:"bar_direction,omitempty"`
	BarOnly        bool                        `json:"bar_only,omitempty"`
	BarSolid       bool                        `json:"bar_solid,omitempty"`
	IconStyle      string                      `json:"icon_style,omitempty"`
	ReverseIcons   bool                        `json:"reverse_icons,omitempty"`
	IconsOnly      bool                        `json:"icons_only,omitempty"`
	StopIfTrue     bool                        `json:"stop_if_true,omitempty"`
	Format         *conditionalFormatStyleArgs `json:"format,omitempty" jsonschema:"Optional style applied when the conditional rule matches"`
}

type formulaArgs struct {
	Filepath  string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	Cell      string `json:"cell" jsonschema:"Target cell"`
	Formula   string `json:"formula" jsonschema:"Excel formula"`
}

type validateFormulaArgs struct {
	Filepath  string `json:"filepath,omitempty" jsonschema:"Optional path to the Excel workbook"`
	SheetName string `json:"sheet_name,omitempty" jsonschema:"Optional worksheet name"`
	Cell      string `json:"cell,omitempty" jsonschema:"Optional target cell"`
	Formula   string `json:"formula" jsonschema:"Excel formula to validate"`
}

func parseNumberFormat(value any) (int, string, error) {
	switch typed := value.(type) {
	case nil:
		return 0, "", nil
	case string:
		return 0, typed, nil
	case float64:
		if typed != math.Trunc(typed) {
			return 0, "", fmt.Errorf("invalid number_format: numeric values must be integers")
		}
		return int(typed), "", nil
	case int:
		return typed, "", nil
	default:
		return 0, "", fmt.Errorf("invalid number_format: expected integer or string")
	}
}

func parseProtectionArgs(value *formatProtectionArgs) *excel.ProtectionOptions {
	if value == nil {
		return nil
	}
	return &excel.ProtectionOptions{Locked: value.Locked, Hidden: value.Hidden}
}

func parseConditionalFormatStyleArgs(value *conditionalFormatStyleArgs) (*excel.ConditionalFormatStyleOptions, error) {
	if value == nil {
		return nil, nil
	}
	numberFormatID, customNumFmt, err := parseNumberFormat(value.NumberFormat)
	if err != nil {
		return nil, err
	}
	return &excel.ConditionalFormatStyleOptions{
		Bold:         value.Bold,
		Italic:       value.Italic,
		Underline:    value.Underline,
		FontSize:     value.FontSize,
		FontColor:    value.FontColor,
		BGColor:      value.BGColor,
		BorderStyle:  value.BorderStyle,
		BorderColor:  value.BorderColor,
		NumberFormat: numberFormatID,
		CustomNumFmt: customNumFmt,
		Alignment:    value.Alignment,
		WrapText:     value.WrapText,
		Protection:   parseProtectionArgs(value.Protection),
	}, nil
}

func parseConditionalFormatArgs(value *conditionalFormatArgs) (*excel.ConditionalFormatOptions, error) {
	if value == nil {
		return nil, nil
	}
	format, err := parseConditionalFormatStyleArgs(value.Format)
	if err != nil {
		return nil, err
	}
	return &excel.ConditionalFormatOptions{
		Type:           value.Type,
		AboveAverage:   value.AboveAverage,
		Percent:        value.Percent,
		Criteria:       value.Criteria,
		Value:          value.Value,
		MinType:        value.MinType,
		MidType:        value.MidType,
		MaxType:        value.MaxType,
		MinValue:       value.MinValue,
		MidValue:       value.MidValue,
		MaxValue:       value.MaxValue,
		MinColor:       value.MinColor,
		MidColor:       value.MidColor,
		MaxColor:       value.MaxColor,
		BarColor:       value.BarColor,
		BarBorderColor: value.BarBorderColor,
		BarDirection:   value.BarDirection,
		BarOnly:        value.BarOnly,
		BarSolid:       value.BarSolid,
		IconStyle:      value.IconStyle,
		ReverseIcons:   value.ReverseIcons,
		IconsOnly:      value.IconsOnly,
		StopIfTrue:     value.StopIfTrue,
		Format:         format,
	}, nil
}

type chartArgs struct {
	Filepath   string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName  string `json:"sheet_name" jsonschema:"Name of the worksheet containing the source data"`
	DataRange  string `json:"data_range" jsonschema:"Range with headers in row 1, series names in column A, and values in remaining cells (e.g. A1:C5)"`
	ChartType  string `json:"chart_type" jsonschema:"Chart type: line, column, bar, pie, doughnut, scatter, area"`
	TargetCell string `json:"target_cell" jsonschema:"Top-left anchor cell for the chart (e.g. E2)"`
	Title      string `json:"title,omitempty" jsonschema:"Optional chart title"`
	XAxis      string `json:"x_axis,omitempty" jsonschema:"Optional X-axis label"`
	YAxis      string `json:"y_axis,omitempty" jsonschema:"Optional Y-axis label"`
}

type pivotArgs struct {
	Filepath   string   `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName  string   `json:"sheet_name" jsonschema:"Name of the worksheet containing the source data"`
	DataRange  string   `json:"data_range" jsonschema:"Source data range including the header row (e.g. A1:D100)"`
	TargetCell string   `json:"target_cell" jsonschema:"Top-left anchor cell for the pivot table output"`
	Rows       []string `json:"rows" jsonschema:"Column names to use as row labels"`
	Values     []string `json:"values" jsonschema:"Column names to aggregate"`
	Columns    []string `json:"columns,omitempty" jsonschema:"Optional column names to use as column labels"`
	AggFunc    string   `json:"agg_func,omitempty" jsonschema:"Aggregation function: sum (default), count, average, avg, mean, max, min"`
}

type tableArgs struct {
	Filepath   string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName  string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	DataRange  string `json:"data_range" jsonschema:"Range including header row and at least one data row (e.g. A1:D10)"`
	TableName  string `json:"table_name,omitempty" jsonschema:"Optional table name (defaults to an auto-generated name)"`
	TableStyle string `json:"table_style,omitempty" jsonschema:"Optional Excel table style name (e.g. TableStyleMedium9)"`
}

type copyWorksheetArgs struct {
	Filepath    string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SourceSheet string `json:"source_sheet" jsonschema:"Name of the worksheet to copy"`
	TargetSheet string `json:"target_sheet" jsonschema:"Name for the new copied worksheet"`
}

type copyRangeArgs struct {
	Filepath    string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName   string `json:"sheet_name" jsonschema:"Name of the source worksheet"`
	SourceStart string `json:"source_start" jsonschema:"Top-left cell of the source range"`
	SourceEnd   string `json:"source_end" jsonschema:"Bottom-right cell of the source range"`
	TargetStart string `json:"target_start" jsonschema:"Top-left cell of the destination range"`
	TargetSheet string `json:"target_sheet,omitempty" jsonschema:"Optional destination worksheet name (defaults to source sheet)"`
}

type sortKeyArgs struct {
	Column     string `json:"column" jsonschema:"Header name or 1-based column index within the range"`
	Descending bool   `json:"descending,omitempty" jsonschema:"Sort descending when true"`
}

type sortRangeArgs struct {
	Filepath  string        `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string        `json:"sheet_name" jsonschema:"Name of the worksheet"`
	Range     string        `json:"range" jsonschema:"Range to sort, such as A1:D20"`
	SortKeys  []sortKeyArgs `json:"sort_keys" jsonschema:"Sort keys applied in order"`
	HasHeader bool          `json:"has_header,omitempty" jsonschema:"Treat the first row of the range as a header row"`
}

type upsertRowArgs struct {
	Match  map[string]any `json:"match" jsonschema:"Key-column values used to find an existing row"`
	Values map[string]any `json:"values" jsonschema:"Non-key column values to update or insert"`
}

type upsertRowsArgs struct {
	Filepath          string          `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName         string          `json:"sheet_name" jsonschema:"Name of the worksheet"`
	Range             string          `json:"range" jsonschema:"Table range including the header row, such as A1:D200"`
	KeyColumns        []string        `json:"key_columns" jsonschema:"Header names that identify a row"`
	Rows              []upsertRowArgs `json:"rows" jsonschema:"Rows to update or insert"`
	InsertIfMissing   bool            `json:"insert_if_missing,omitempty" jsonschema:"Append a row when no existing row matches the key"`
	CaseSensitiveKeys bool            `json:"case_sensitive_keys,omitempty" jsonschema:"Use case-sensitive key matching"`
}

type deleteRangeArgs struct {
	Filepath       string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName      string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	StartCell      string `json:"start_cell" jsonschema:"Top-left cell of the range to delete"`
	EndCell        string `json:"end_cell" jsonschema:"Bottom-right cell of the range to delete"`
	ShiftDirection string `json:"shift_direction,omitempty" jsonschema:"Direction to shift remaining cells: up (default) or left"`
}

type columnWidthEntry struct {
	Column string  `json:"column" jsonschema:"Column letter (e.g. A, B, AA)"`
	Width  float64 `json:"width" jsonschema:"Column width in character units"`
}

type setColumnWidthsArgs struct {
	Filepath     string             `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName    string             `json:"sheet_name" jsonschema:"Name of the worksheet"`
	Widths       []columnWidthEntry `json:"widths,omitempty" jsonschema:"Explicit column widths to apply"`
	AutoFit      bool               `json:"auto_fit,omitempty" jsonschema:"Automatically size each column to fit its longest cell value"`
	AutoFitRange string             `json:"auto_fit_range,omitempty" jsonschema:"Optional range to limit auto-fit content scanning (e.g. A1:D100); defaults to the full sheet"`
}

type rowHeightEntry struct {
	Row    int     `json:"row" jsonschema:"1-based row number"`
	Height float64 `json:"height" jsonschema:"Row height in points (e.g. 15 is the Excel default)"`
}

type setRowHeightsArgs struct {
	Filepath  string           `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string           `json:"sheet_name" jsonschema:"Name of the worksheet"`
	Heights   []rowHeightEntry `json:"heights" jsonschema:"Row heights to set"`
}

type renameWorksheetArgs struct {
	Filepath string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	OldName  string `json:"old_name" jsonschema:"Current name of the worksheet"`
	NewName  string `json:"new_name" jsonschema:"New name for the worksheet"`
}

type validationInfoArgs struct {
	Filepath  string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string `json:"sheet_name" jsonschema:"Name of the worksheet"`
}

type rowArgs struct {
	Filepath  string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	StartRow  int    `json:"start_row" jsonschema:"1-based row number to start at"`
	Count     int    `json:"count,omitempty" jsonschema:"Number of rows to insert or delete (defaults to 1)"`
}

type columnArgs struct {
	Filepath  string `json:"filepath" jsonschema:"Path to the Excel workbook"`
	SheetName string `json:"sheet_name" jsonschema:"Name of the worksheet"`
	StartCol  int    `json:"start_col" jsonschema:"1-based column number to start at"`
	Count     int    `json:"count,omitempty" jsonschema:"Number of columns to insert or delete (defaults to 1)"`
}

func (tc *toolContext) createWorkbook(_ context.Context, _ *mcp.CallToolRequest, args fileArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.CreateWorkbook(path)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) createWorksheet(_ context.Context, _ *mcp.CallToolRequest, args createWorksheetArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.CreateWorksheet(path, args.SheetName)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) getWorkbookMetadata(_ context.Context, _ *mcp.CallToolRequest, args metadataArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	metadata, err := excel.GetWorkbookMetadata(path, args.IncludeRanges)
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(metadata)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) describeWorkbook(_ context.Context, _ *mcp.CallToolRequest, args describeWorkbookArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	description, err := excel.DescribeWorkbook(path, excel.DescribeWorkbookOptions{
		IncludeRanges:     args.IncludeRanges,
		IncludeTables:     args.IncludeTables,
		IncludeCharts:     args.IncludeCharts,
		IncludePivots:     args.IncludePivots,
		IncludeNames:      args.IncludeNames,
		IncludeMerged:     args.IncludeMerged,
		IncludeValidation: args.IncludeValidation,
	})
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(description)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) listCharts(_ context.Context, _ *mcp.CallToolRequest, args listChartsArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	charts, err := excel.ListCharts(path, args.SheetName, args.SourceSheet)
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(charts)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) getSheetSchema(_ context.Context, _ *mcp.CallToolRequest, args getSheetSchemaArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	startCell := args.StartCell
	if startCell == "" {
		startCell = "A1"
	}
	schema, err := excel.GetSheetSchema(path, args.SheetName, startCell, args.EndCell, excel.SheetSchemaOptions{
		HeaderRow:  args.HeaderRow,
		SampleSize: args.SampleSize,
	})
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(schema)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) findInWorkbook(_ context.Context, _ *mcp.CallToolRequest, args findInWorkbookArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	resultValue, err := excel.FindInWorkbook(path, excel.FindOptions{
		Query:         args.Query,
		Sheets:        args.Sheets,
		MatchMode:     args.MatchMode,
		SearchType:    args.SearchType,
		CaseSensitive: args.CaseSensitive,
		ContextRows:   args.ContextRows,
		ContextCols:   args.ContextCols,
		MaxResults:    args.MaxResults,
	})
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(resultValue)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) writeDataToExcel(_ context.Context, _ *mcp.CallToolRequest, args writeDataArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	startCell := args.StartCell
	if startCell == "" {
		startCell = "A1"
	}
	message, err := excel.WriteData(path, args.SheetName, args.Data, startCell)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) readDataFromExcel(_ context.Context, _ *mcp.CallToolRequest, args readDataArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	startCell := args.StartCell
	if startCell == "" {
		startCell = "A1"
	}
	data, err := excel.ReadData(path, args.SheetName, startCell, args.EndCell, args.PreviewOnly)
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(data)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) filterRows(_ context.Context, _ *mcp.CallToolRequest, args filterRowsArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	filters := make([]excel.Filter, 0, len(args.Filters))
	for _, filter := range args.Filters {
		filters = append(filters, excel.Filter{Column: filter.Column, Operator: filter.Operator, Value: filter.Value})
	}
	resultValue, err := excel.FilterRows(path, args.SheetName, args.Range, excel.FilterRowsOptions{
		HasHeader: args.HasHeader,
		Filters:   filters,
	})
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(resultValue)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) formatRange(_ context.Context, _ *mcp.CallToolRequest, args formatRangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	endCell := args.EndCell
	if endCell == "" {
		endCell = args.StartCell
	}
	numberFormatID, customNumFmt, err := parseNumberFormat(args.NumberFormat)
	if err != nil {
		return toolError(err), nil, nil
	}
	conditionalFormat, err := parseConditionalFormatArgs(args.Conditional)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.FormatRange(path, args.SheetName, args.StartCell, endCell, excel.FormatRangeOptions{
		Bold:         args.Bold,
		Italic:       args.Italic,
		Underline:    args.Underline,
		FontSize:     args.FontSize,
		FontColor:    args.FontColor,
		BGColor:      args.BGColor,
		BorderStyle:  args.BorderStyle,
		BorderColor:  args.BorderColor,
		NumberFormat: numberFormatID,
		CustomNumFmt: customNumFmt,
		Alignment:    args.Alignment,
		WrapText:     args.WrapText,
		MergeCells:   args.MergeCells,
		Protection:   parseProtectionArgs(args.Protection),
		Conditional:  conditionalFormat,
	})
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) mergeCells(_ context.Context, _ *mcp.CallToolRequest, args rangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.MergeCells(path, args.SheetName, args.StartCell, args.EndCell)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) unmergeCells(_ context.Context, _ *mcp.CallToolRequest, args rangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.UnmergeCells(path, args.SheetName, args.StartCell, args.EndCell)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) getMergedCells(_ context.Context, _ *mcp.CallToolRequest, args validationInfoArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	merged, err := excel.GetMergedCells(path, args.SheetName)
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(map[string]any{"sheet_name": args.SheetName, "merged_ranges": merged})
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) applyFormula(_ context.Context, _ *mcp.CallToolRequest, args formulaArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.ApplyFormula(path, args.SheetName, args.Cell, args.Formula)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) validateFormulaSyntax(_ context.Context, _ *mcp.CallToolRequest, args validateFormulaArgs) (*mcp.CallToolResult, any, error) {
	if err := excel.ValidateFormulaSyntax(args.Formula); err != nil {
		return toolError(err), nil, nil
	}
	if args.SheetName != "" && args.Cell != "" {
		return textResult("formula syntax looks valid for " + args.SheetName + "!" + args.Cell), nil, nil
	}
	return textResult("formula syntax looks valid"), nil, nil
}

func (tc *toolContext) createChart(_ context.Context, _ *mcp.CallToolRequest, args chartArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.CreateChart(path, args.SheetName, args.DataRange, args.ChartType, args.TargetCell, excel.ChartOptions{Title: args.Title, XAxis: args.XAxis, YAxis: args.YAxis})
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) createPivotTable(_ context.Context, _ *mcp.CallToolRequest, args pivotArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.CreatePivotTable(path, args.SheetName, args.DataRange, args.TargetCell, excel.PivotOptions{Rows: args.Rows, Values: args.Values, Columns: args.Columns, AggFunc: args.AggFunc})
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) createTable(_ context.Context, _ *mcp.CallToolRequest, args tableArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.CreateTable(path, args.SheetName, args.DataRange, args.TableName, args.TableStyle)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) copyWorksheet(_ context.Context, _ *mcp.CallToolRequest, args copyWorksheetArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.CopyWorksheet(path, args.SourceSheet, args.TargetSheet)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) deleteWorksheet(_ context.Context, _ *mcp.CallToolRequest, args createWorksheetArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.DeleteWorksheet(path, args.SheetName)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) renameWorksheet(_ context.Context, _ *mcp.CallToolRequest, args renameWorksheetArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.RenameWorksheet(path, args.OldName, args.NewName)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) copyRange(_ context.Context, _ *mcp.CallToolRequest, args copyRangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.CopyRange(path, args.SheetName, args.SourceStart, args.SourceEnd, args.TargetStart, args.TargetSheet)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) sortRange(_ context.Context, _ *mcp.CallToolRequest, args sortRangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	sortKeys := make([]excel.SortKey, 0, len(args.SortKeys))
	for _, key := range args.SortKeys {
		sortKeys = append(sortKeys, excel.SortKey{Column: key.Column, Descending: key.Descending})
	}
	resultValue, err := excel.SortRange(path, args.SheetName, args.Range, excel.SortRangeOptions{
		HasHeader: args.HasHeader,
		SortKeys:  sortKeys,
	})
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(resultValue)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) upsertRows(_ context.Context, _ *mcp.CallToolRequest, args upsertRowsArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	rows := make([]excel.UpsertRow, 0, len(args.Rows))
	for _, row := range args.Rows {
		rows = append(rows, excel.UpsertRow{Match: row.Match, Values: row.Values})
	}
	resultValue, err := excel.UpsertRows(path, args.SheetName, args.Range, excel.UpsertRowsOptions{
		KeyColumns:        args.KeyColumns,
		Rows:              rows,
		InsertIfMissing:   args.InsertIfMissing,
		CaseSensitiveKeys: args.CaseSensitiveKeys,
	})
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(resultValue)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) deleteRange(_ context.Context, _ *mcp.CallToolRequest, args deleteRangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.DeleteRange(path, args.SheetName, args.StartCell, args.EndCell, args.ShiftDirection)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) clearRange(_ context.Context, _ *mcp.CallToolRequest, args rangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.ClearRange(path, args.SheetName, args.StartCell, args.EndCell)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) validateExcelRange(_ context.Context, _ *mcp.CallToolRequest, args rangeArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	message, err := excel.ValidateRange(path, args.SheetName, args.StartCell, args.EndCell)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) getDataValidationInfo(_ context.Context, _ *mcp.CallToolRequest, args validationInfoArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	items, err := excel.GetDataValidationInfo(path, args.SheetName)
	if err != nil {
		return toolError(err), nil, nil
	}
	result, err := jsonResult(items)
	if err != nil {
		return nil, nil, err
	}
	return result, nil, nil
}

func (tc *toolContext) insertRows(_ context.Context, _ *mcp.CallToolRequest, args rowArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	count := args.Count
	if count == 0 {
		count = 1
	}
	message, err := excel.InsertRows(path, args.SheetName, args.StartRow, count)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) insertColumns(_ context.Context, _ *mcp.CallToolRequest, args columnArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	count := args.Count
	if count == 0 {
		count = 1
	}
	message, err := excel.InsertColumns(path, args.SheetName, args.StartCol, count)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) deleteSheetRows(_ context.Context, _ *mcp.CallToolRequest, args rowArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	count := args.Count
	if count == 0 {
		count = 1
	}
	message, err := excel.DeleteSheetRows(path, args.SheetName, args.StartRow, count)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) deleteSheetColumns(_ context.Context, _ *mcp.CallToolRequest, args columnArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	count := args.Count
	if count == 0 {
		count = 1
	}
	message, err := excel.DeleteSheetColumns(path, args.SheetName, args.StartCol, count)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) setColumnWidths(_ context.Context, _ *mcp.CallToolRequest, args setColumnWidthsArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	widths := make([]excel.ColumnWidthEntry, 0, len(args.Widths))
	for _, w := range args.Widths {
		widths = append(widths, excel.ColumnWidthEntry{Column: w.Column, Width: w.Width})
	}
	message, err := excel.SetColumnWidths(path, args.SheetName, widths, args.AutoFit, args.AutoFitRange)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}

func (tc *toolContext) setRowHeights(_ context.Context, _ *mcp.CallToolRequest, args setRowHeightsArgs) (*mcp.CallToolResult, any, error) {
	path, err := tc.resolve(args.Filepath)
	if err != nil {
		return toolError(err), nil, nil
	}
	heights := make([]excel.RowHeightEntry, 0, len(args.Heights))
	for _, h := range args.Heights {
		heights = append(heights, excel.RowHeightEntry{Row: h.Row, Height: h.Height})
	}
	message, err := excel.SetRowHeights(path, args.SheetName, heights)
	if err != nil {
		return toolError(err), nil, nil
	}
	return textResult(message), nil, nil
}
