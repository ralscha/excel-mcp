package excel

type WorkbookMetadata struct {
	FilePath string          `json:"filepath"`
	Sheets   []SheetMetadata `json:"sheets"`
}

type SheetMetadata struct {
	Name  string `json:"name"`
	Range string `json:"range,omitempty"`
	Rows  int    `json:"rows,omitempty"`
	Cols  int    `json:"cols,omitempty"`
}

type DescribeWorkbookOptions struct {
	IncludeRanges     bool
	IncludeTables     bool
	IncludeCharts     bool
	IncludePivots     bool
	IncludeNames      bool
	IncludeMerged     bool
	IncludeValidation bool
}

type WorkbookDescription struct {
	FilePath    string                     `json:"filepath"`
	NamedRanges []NamedRangeDescription    `json:"named_ranges,omitempty"`
	Sheets      []WorkbookSheetDescription `json:"sheets"`
}

type WorkbookSheetDescription struct {
	Name            string                  `json:"name"`
	Range           string                  `json:"range,omitempty"`
	Rows            int                     `json:"rows,omitempty"`
	Cols            int                     `json:"cols,omitempty"`
	MergedRanges    []string                `json:"merged_ranges,omitempty"`
	Tables          []TableDescription      `json:"tables,omitempty"`
	Charts          []ChartDescription      `json:"charts,omitempty"`
	PivotTables     []PivotTableDescription `json:"pivot_tables,omitempty"`
	DataValidations []ValidationInfo        `json:"data_validations,omitempty"`
}

type NamedRangeDescription struct {
	Name     string `json:"name"`
	RefersTo string `json:"refers_to"`
	Scope    string `json:"scope,omitempty"`
	Comment  string `json:"comment,omitempty"`
}

type TableDescription struct {
	Name      string `json:"name"`
	Range     string `json:"range"`
	StyleName string `json:"style_name,omitempty"`
}

type ChartDescription struct {
	SheetName  string                   `json:"sheet_name,omitempty"`
	AnchorCell string                   `json:"anchor_cell,omitempty"`
	Title      string                   `json:"title,omitempty"`
	ChartType  string                   `json:"chart_type,omitempty"`
	ChartPath  string                   `json:"chart_path,omitempty"`
	Series     []ChartSeriesDescription `json:"series,omitempty"`
}

type ChartSeriesDescription struct {
	DisplayName     string `json:"display_name,omitempty"`
	NameRef         string `json:"name_ref,omitempty"`
	SourceSheet     string `json:"source_sheet,omitempty"`
	SourceRange     string `json:"source_range,omitempty"`
	CategoriesRange string `json:"categories_range,omitempty"`
	ValuesRange     string `json:"values_range,omitempty"`
}

type PivotTableDescription struct {
	Name            string                  `json:"name,omitempty"`
	DataRange       string                  `json:"data_range,omitempty"`
	PivotTableRange string                  `json:"pivot_table_range,omitempty"`
	Rows            []string                `json:"rows,omitempty"`
	Columns         []string                `json:"columns,omitempty"`
	Values          []PivotValueDescription `json:"values,omitempty"`
	StyleName       string                  `json:"style_name,omitempty"`
}

type PivotValueDescription struct {
	Data     string `json:"data"`
	Name     string `json:"name,omitempty"`
	Subtotal string `json:"subtotal,omitempty"`
}

type SheetSchemaOptions struct {
	HeaderRow  int
	SampleSize int
}

type SheetSchema struct {
	FilePath  string              `json:"filepath"`
	SheetName string              `json:"sheet_name"`
	Range     string              `json:"range"`
	HeaderRow int                 `json:"header_row"`
	RowCount  int                 `json:"row_count"`
	Columns   []SheetColumnSchema `json:"columns"`
}

type SheetColumnSchema struct {
	Name         string   `json:"name"`
	Column       string   `json:"column"`
	InferredType string   `json:"inferred_type"`
	BlankCount   int      `json:"blank_count"`
	SampleValues []string `json:"sample_values"`
}

type FindOptions struct {
	Query         string
	Sheets        []string
	MatchMode     string
	SearchType    string
	CaseSensitive bool
	ContextRows   int
	ContextCols   int
	MaxResults    int
}

type FindResult struct {
	FilePath   string      `json:"filepath"`
	Query      string      `json:"query"`
	SearchType string      `json:"search_type"`
	MatchMode  string      `json:"match_mode"`
	Matches    []FindMatch `json:"matches"`
}

type FindMatch struct {
	SheetName string        `json:"sheet_name"`
	Cell      string        `json:"cell"`
	Value     string        `json:"value,omitempty"`
	Formula   string        `json:"formula,omitempty"`
	Context   []ContextCell `json:"context,omitempty"`
}

type ContextCell struct {
	Cell  string `json:"cell"`
	Value string `json:"value,omitempty"`
}

type ListChartsResult struct {
	FilePath    string             `json:"filepath"`
	SheetName   string             `json:"sheet_name,omitempty"`
	SourceSheet string             `json:"source_sheet,omitempty"`
	Charts      []ChartDescription `json:"charts"`
}

type ReadResult struct {
	FilePath  string           `json:"filepath"`
	SheetName string           `json:"sheet_name"`
	Range     string           `json:"range"`
	Rows      []map[string]any `json:"rows"`
}

type ValidationInfo struct {
	Sqref       string `json:"sqref"`
	Type        string `json:"type,omitempty"`
	Operator    string `json:"operator,omitempty"`
	Formula1    string `json:"formula1,omitempty"`
	Formula2    string `json:"formula2,omitempty"`
	ErrorTitle  string `json:"error_title,omitempty"`
	ErrorBody   string `json:"error_body,omitempty"`
	PromptTitle string `json:"prompt_title,omitempty"`
	PromptBody  string `json:"prompt_body,omitempty"`
}

type ChartOptions struct {
	Title string `json:"title,omitempty"`
	XAxis string `json:"x_axis,omitempty"`
	YAxis string `json:"y_axis,omitempty"`
}

type PivotOptions struct {
	Rows    []string `json:"rows"`
	Values  []string `json:"values"`
	Columns []string `json:"columns,omitempty"`
	AggFunc string   `json:"agg_func,omitempty"`
}

type ProtectionOptions struct {
	Locked bool `json:"locked,omitempty"`
	Hidden bool `json:"hidden,omitempty"`
}

type ConditionalFormatStyleOptions struct {
	Bold         bool               `json:"bold,omitempty"`
	Italic       bool               `json:"italic,omitempty"`
	Underline    bool               `json:"underline,omitempty"`
	FontSize     int                `json:"font_size,omitempty"`
	FontColor    string             `json:"font_color,omitempty"`
	BGColor      string             `json:"bg_color,omitempty"`
	BorderStyle  string             `json:"border_style,omitempty"`
	BorderColor  string             `json:"border_color,omitempty"`
	NumberFormat int                `json:"number_format,omitempty"`
	CustomNumFmt string             `json:"custom_number_format,omitempty"`
	Alignment    string             `json:"alignment,omitempty"`
	WrapText     bool               `json:"wrap_text,omitempty"`
	Protection   *ProtectionOptions `json:"protection,omitempty"`
}

type ConditionalFormatOptions struct {
	Type           string                         `json:"type,omitempty"`
	AboveAverage   bool                           `json:"above_average,omitempty"`
	Percent        bool                           `json:"percent,omitempty"`
	Criteria       string                         `json:"criteria,omitempty"`
	Value          string                         `json:"value,omitempty"`
	MinType        string                         `json:"min_type,omitempty"`
	MidType        string                         `json:"mid_type,omitempty"`
	MaxType        string                         `json:"max_type,omitempty"`
	MinValue       string                         `json:"min_value,omitempty"`
	MidValue       string                         `json:"mid_value,omitempty"`
	MaxValue       string                         `json:"max_value,omitempty"`
	MinColor       string                         `json:"min_color,omitempty"`
	MidColor       string                         `json:"mid_color,omitempty"`
	MaxColor       string                         `json:"max_color,omitempty"`
	BarColor       string                         `json:"bar_color,omitempty"`
	BarBorderColor string                         `json:"bar_border_color,omitempty"`
	BarDirection   string                         `json:"bar_direction,omitempty"`
	BarOnly        bool                           `json:"bar_only,omitempty"`
	BarSolid       bool                           `json:"bar_solid,omitempty"`
	IconStyle      string                         `json:"icon_style,omitempty"`
	ReverseIcons   bool                           `json:"reverse_icons,omitempty"`
	IconsOnly      bool                           `json:"icons_only,omitempty"`
	StopIfTrue     bool                           `json:"stop_if_true,omitempty"`
	Format         *ConditionalFormatStyleOptions `json:"format,omitempty"`
}

type FormatRangeOptions struct {
	Bold         bool                      `json:"bold,omitempty"`
	Italic       bool                      `json:"italic,omitempty"`
	Underline    bool                      `json:"underline,omitempty"`
	FontSize     int                       `json:"font_size,omitempty"`
	FontColor    string                    `json:"font_color,omitempty"`
	BGColor      string                    `json:"bg_color,omitempty"`
	BorderStyle  string                    `json:"border_style,omitempty"`
	BorderColor  string                    `json:"border_color,omitempty"`
	NumberFormat int                       `json:"number_format,omitempty"`
	CustomNumFmt string                    `json:"custom_number_format,omitempty"`
	Alignment    string                    `json:"alignment,omitempty"`
	WrapText     bool                      `json:"wrap_text,omitempty"`
	MergeCells   bool                      `json:"merge_cells,omitempty"`
	Protection   *ProtectionOptions        `json:"protection,omitempty"`
	Conditional  *ConditionalFormatOptions `json:"conditional_format,omitempty"`
}

type SortKey struct {
	Column     string `json:"column"`
	Descending bool   `json:"descending,omitempty"`
}

type SortRangeOptions struct {
	HasHeader bool
	SortKeys  []SortKey
}

type SortRangeResult struct {
	FilePath  string    `json:"filepath"`
	SheetName string    `json:"sheet_name"`
	Range     string    `json:"range"`
	HasHeader bool      `json:"has_header"`
	SortKeys  []SortKey `json:"sort_keys"`
}

type UpsertRow struct {
	Match  map[string]any `json:"match"`
	Values map[string]any `json:"values"`
}

type UpsertRowsOptions struct {
	KeyColumns        []string
	Rows              []UpsertRow
	InsertIfMissing   bool
	CaseSensitiveKeys bool
}

type UpsertRowsResult struct {
	FilePath      string            `json:"filepath"`
	SheetName     string            `json:"sheet_name"`
	Range         string            `json:"range"`
	KeyColumns    []string          `json:"key_columns"`
	UpdatedCount  int               `json:"updated_count"`
	InsertedCount int               `json:"inserted_count"`
	SkippedCount  int               `json:"skipped_count"`
	Results       []UpsertRowResult `json:"results"`
}

type UpsertRowResult struct {
	Key       map[string]any `json:"key"`
	Action    string         `json:"action"`
	RowNumber int            `json:"row_number,omitempty"`
	Reason    string         `json:"reason,omitempty"`
}

type Filter struct {
	Column   string `json:"column"`
	Operator string `json:"operator"`
	Value    string `json:"value"`
}

type FilterRowsOptions struct {
	HasHeader bool
	Filters   []Filter
}

type FilterRowsResult struct {
	FilePath  string           `json:"filepath"`
	SheetName string           `json:"sheet_name"`
	Range     string           `json:"range"`
	HasHeader bool             `json:"has_header"`
	Filters   []Filter         `json:"filters"`
	Rows      []map[string]any `json:"rows"`
}

type ColumnWidthEntry struct {
	Column string  `json:"column"`
	Width  float64 `json:"width"`
}

type RowHeightEntry struct {
	Row    int     `json:"row"`
	Height float64 `json:"height"`
}
