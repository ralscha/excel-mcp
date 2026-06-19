package server

import (
	"archive/zip"
	"context"
	"encoding/json"
	"io"
	"path/filepath"
	"slices"
	"strings"
	"testing"

	"github.com/google/jsonschema-go/jsonschema"
	"github.com/modelcontextprotocol/go-sdk/mcp"
	"github.com/xuri/excelize/v2"
)

func TestCoreWorkbookFlow(t *testing.T) {
	ctx, clientSession := newTestClient(t)

	workbook := filepath.Join(t.TempDir(), "book.xlsx")

	callTool(t, ctx, clientSession, "create_workbook", map[string]any{"filepath": workbook})
	callTool(t, ctx, clientSession, "write_data_to_excel", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "A1",
		"data": []map[string]any{
			{"name": "alpha", "value": 1},
			{"name": "beta", "value": 2},
		},
	})
	read := callTool(t, ctx, clientSession, "read_data_from_excel", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "A1",
	})

	var payload struct {
		Rows []map[string]any `json:"rows"`
	}
	if err := json.Unmarshal([]byte(read), &payload); err != nil {
		t.Fatalf("unmarshal read payload: %v\n%s", err, read)
	}
	if len(payload.Rows) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(payload.Rows))
	}
	if payload.Rows[0]["name"] != "alpha" {
		t.Fatalf("unexpected first row: %+v", payload.Rows[0])
	}
}

func TestAdvancedToolArtifacts(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "advanced.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Series", "Jan", "Feb"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10, 12}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 8, 9}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "E1", &[]any{"Region", "Revenue"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "E2", &[]any{"North", 10}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "E3", &[]any{"South", 8}); err != nil {
			return err
		}
		return nil
	})

	callTool(t, ctx, clientSession, "create_table", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "E1:F3",
		"table_name":  "SalesTable",
		"table_style": "TableStyleMedium2",
	})
	callTool(t, ctx, clientSession, "create_chart", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "A1:C3",
		"chart_type":  "line",
		"target_cell": "H2",
		"title":       "Regional Trend",
	})
	callTool(t, ctx, clientSession, "create_pivot_table", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "E1:F3",
		"target_cell": "H20",
		"rows":        []string{"Region"},
		"values":      []string{"Revenue"},
		"agg_func":    "sum",
	})

	f, err := excelize.OpenFile(workbook)
	if err != nil {
		t.Fatalf("open workbook: %v", err)
	}
	defer func() { _ = f.Close() }()

	tables, err := f.GetTables("Sheet1")
	if err != nil {
		t.Fatalf("get tables: %v", err)
	}
	if len(tables) != 1 || tables[0].Name != "SalesTable" {
		t.Fatalf("unexpected tables: %+v", tables)
	}

	pivots, err := f.GetPivotTables("Sheet1")
	if err != nil {
		t.Fatalf("get pivot tables: %v", err)
	}
	if len(pivots) != 1 {
		t.Fatalf("expected 1 pivot table, got %d", len(pivots))
	}
	if len(pivots[0].Data) != 1 || pivots[0].Data[0].Subtotal != "Sum" {
		t.Fatalf("unexpected pivot definition: %+v", pivots[0])
	}

	if !zipHasEntry(t, workbook, "xl/charts/chart1.xml") {
		t.Fatal("expected chart XML entry in workbook")
	}
}

func TestFormatRangeCompatibility(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "format-range.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetCellValue("Sheet1", "A1", 1234.5); err != nil {
			return err
		}
		if err := f.SetCellValue("Sheet1", "A2", 8); err != nil {
			return err
		}
		return f.SetCellValue("Sheet1", "A3", 4)
	})

	callTool(t, ctx, clientSession, "format_range", map[string]any{
		"filepath":      workbook,
		"sheet_name":    "Sheet1",
		"start_cell":    "A1",
		"end_cell":      "A3",
		"bold":          true,
		"number_format": "$#,##0.00",
		"protection": map[string]any{
			"locked": true,
			"hidden": true,
		},
		"conditional_format": map[string]any{
			"type":     "cell",
			"criteria": ">",
			"value":    "6",
			"format": map[string]any{
				"font_color": "#FF0000",
				"bg_color":   "#FFF2CC",
			},
		},
	})

	stylesXML := zipEntryContents(t, workbook, "xl/styles.xml")
	if !strings.Contains(stylesXML, "$#,##0.00") {
		t.Fatalf("expected custom number format in styles.xml, got %q", stylesXML)
	}

	f, err := excelize.OpenFile(workbook)
	if err != nil {
		t.Fatalf("open formatted workbook: %v", err)
	}
	defer func() { _ = f.Close() }()

	styleID, err := f.GetCellStyle("Sheet1", "A1")
	if err != nil {
		t.Fatalf("get cell style: %v", err)
	}
	if styleID == 0 {
		t.Fatal("expected style to be applied to A1")
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		t.Fatalf("get style details: %v", err)
	}
	if style.Protection == nil || !style.Protection.Locked || !style.Protection.Hidden {
		t.Fatalf("expected cell protection to be applied, got %+v", style.Protection)
	}
	formats, err := f.GetConditionalFormats("Sheet1")
	if err != nil {
		t.Fatalf("get conditional formats: %v", err)
	}
	rules, ok := formats["A1:A3"]
	if !ok || len(rules) != 1 {
		t.Fatalf("expected one conditional format rule on A1:A3, got %+v", formats)
	}
	if rules[0].Type != "cell" || rules[0].Criteria != "greater than" || rules[0].Value != "6" || rules[0].Format == nil {
		t.Fatalf("unexpected conditional format rule: %+v", rules[0])
	}
}

func TestValidateFormulaSyntaxCompatibility(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "formula.xlsx")

	callTool(t, ctx, clientSession, "create_workbook", map[string]any{"filepath": workbook})

	message := callTool(t, ctx, clientSession, "validate_formula_syntax", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"cell":       "B2",
		"formula":    "=SUM(A1:A3)",
	})
	if !strings.Contains(message, "Sheet1!B2") {
		t.Fatalf("expected validation response to mention target cell, got %q", message)
	}

	assertToolErrorContains(t, ctx, clientSession, "validate_formula_syntax", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"cell":       "B2",
		"formula":    "SUM(A1:A3)",
	}, "formula must start with '='")
}

func TestCopyRangeAndDeleteRangeBehavior(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	copyWorkbook := filepath.Join(t.TempDir(), "copy-ranges.xlsx")
	deleteWorkbook := filepath.Join(t.TempDir(), "delete-ranges.xlsx")

	fillWorkbook := func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Name", "Q1", "Q2"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10, 20}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 30, 40}); err != nil {
			return err
		}
		return nil
	}
	createWorkbookFixture(t, copyWorkbook, fillWorkbook)
	createWorkbookFixture(t, deleteWorkbook, fillWorkbook)

	callTool(t, ctx, clientSession, "copy_range", map[string]any{
		"filepath":     copyWorkbook,
		"sheet_name":   "Sheet1",
		"source_start": "A2",
		"source_end":   "C3",
		"target_start": "E2",
	})

	f, err := excelize.OpenFile(copyWorkbook)
	if err != nil {
		t.Fatalf("open copy workbook: %v", err)
	}
	defer func() { _ = f.Close() }()

	if got, _ := f.GetCellValue("Sheet1", "E2"); got != "North" { //nolint:goconst
		t.Fatalf("expected copied cell E2 to be North, got %q", got)
	}
	if got, _ := f.GetCellValue("Sheet1", "G3"); got != "40" {
		t.Fatalf("expected copied cell G3 to be 40, got %q", got)
	}

	callTool(t, ctx, clientSession, "delete_range", map[string]any{
		"filepath":        deleteWorkbook,
		"sheet_name":      "Sheet1",
		"start_cell":      "B2",
		"end_cell":        "B3",
		"shift_direction": "left",
	})

	f2, err := excelize.OpenFile(deleteWorkbook)
	if err != nil {
		t.Fatalf("open delete workbook: %v", err)
	}
	defer func() { _ = f2.Close() }()

	if got, _ := f2.GetCellValue("Sheet1", "B2"); got != "20" {
		t.Fatalf("expected B2 to shift left to 20, got %q", got)
	}
	if got, _ := f2.GetCellValue("Sheet1", "B3"); got != "40" {
		t.Fatalf("expected B3 to shift left to 40, got %q", got)
	}
}

func TestClearRangePreservesLayout(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "clear-range.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Name", "Q1", "Q2"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10, 20}); err != nil {
			return err
		}
		return f.SetSheetRow("Sheet1", "A3", &[]any{"South", 30, 40})
	})

	callTool(t, ctx, clientSession, "clear_range", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "B2",
		"end_cell":   "C2",
	})

	f, err := excelize.OpenFile(workbook)
	if err != nil {
		t.Fatalf("open clear workbook: %v", err)
	}
	defer func() { _ = f.Close() }()

	if got, _ := f.GetCellValue("Sheet1", "A2"); got != "North" {
		t.Fatalf("expected A2 to stay North, got %q", got)
	}
	if got, _ := f.GetCellValue("Sheet1", "B2"); got != "" {
		t.Fatalf("expected B2 to be cleared, got %q", got)
	}
	if got, _ := f.GetCellValue("Sheet1", "C2"); got != "" {
		t.Fatalf("expected C2 to be cleared, got %q", got)
	}
	if got, _ := f.GetCellValue("Sheet1", "B3"); got != "30" {
		t.Fatalf("expected B3 to remain 30, got %q", got)
	}
	if got, _ := f.GetCellValue("Sheet1", "C3"); got != "40" {
		t.Fatalf("expected C3 to remain 40, got %q", got)
	}
}

func TestSingleCellRangeDefaultsAndReadEmptySheet(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "single-cell.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetCellValue("Sheet1", "A1", "Header"); err != nil {
			return err
		}
		if err := f.SetCellValue("Sheet1", "A2", "Value"); err != nil {
			return err
		}
		if _, err := f.NewSheet("Empty"); err != nil {
			return err
		}
		return nil
	})

	callTool(t, ctx, clientSession, "clear_range", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "A2",
	})
	callTool(t, ctx, clientSession, "validate_excel_range", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "A1",
	})

	f, err := excelize.OpenFile(workbook)
	if err != nil {
		t.Fatalf("open workbook: %v", err)
	}
	if got, _ := f.GetCellValue("Sheet1", "A2"); got != "" {
		t.Fatalf("expected A2 to be cleared, got %q", got)
	}
	if err := f.Close(); err != nil {
		t.Fatalf("close workbook after clear check: %v", err)
	}

	callTool(t, ctx, clientSession, "delete_range", map[string]any{
		"filepath":        workbook,
		"sheet_name":      "Sheet1",
		"start_cell":      "A1",
		"shift_direction": "up",
	})
	f, err = excelize.OpenFile(workbook)
	if err != nil {
		t.Fatalf("reopen workbook after delete: %v", err)
	}
	defer func() { _ = f.Close() }()
	if got, _ := f.GetCellValue("Sheet1", "A1"); got != "" {
		t.Fatalf("expected A1 to shift up to empty, got %q", got)
	}

	read := callTool(t, ctx, clientSession, "read_data_from_excel", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Empty",
		"start_cell": "A1",
	})
	var payload struct {
		Range string           `json:"range"`
		Rows  []map[string]any `json:"rows"`
	}
	if err := json.Unmarshal([]byte(read), &payload); err != nil {
		t.Fatalf("unmarshal empty sheet payload: %v\n%s", err, read)
	}
	if payload.Range != "A1:A1" || len(payload.Rows) != 0 {
		t.Fatalf("unexpected empty sheet payload: %+v", payload)
	}

	assertToolErrorContains(t, ctx, clientSession, "read_data_from_excel", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "B2",
		"end_cell":   "A1",
	}, "range end must be below and to the right of start_cell")

	assertToolErrorContains(t, ctx, clientSession, "insert_rows", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_row":  1,
		"count":      -1,
	}, "count must be positive")
}

func TestSortRange(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	singleAscWorkbook := filepath.Join(t.TempDir(), "sort-single-asc.xlsx")
	singleDescWorkbook := filepath.Join(t.TempDir(), "sort-single-desc.xlsx")
	multiWorkbook := filepath.Join(t.TempDir(), "sort-multi.xlsx")
	partialWorkbook := filepath.Join(t.TempDir(), "sort-partial.xlsx")
	errorWorkbook := filepath.Join(t.TempDir(), "sort-errors.xlsx")

	createWorkbookFixture(t, singleAscWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Name", "Score"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"Beta", 2}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"Alpha", 3}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A4", &[]any{"Gamma", 1}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, singleDescWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Name", "Score"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"Beta", 2}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"Alpha", 3}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A4", &[]any{"Gamma", 1}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, multiWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Region", "Quarter", "Revenue"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"West", "Q2", 30}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"East", "Q1", 20}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A4", &[]any{"West", "Q1", 40}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A5", &[]any{"East", "Q3", 10}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, partialWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Name", "Score", "Status", "Note"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"Beta", 2, "open", "keep-1"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"Alpha", 3, "hold", "keep-2"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A4", &[]any{"Gamma", 1, "open", "keep-3"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A5", &[]any{"Tail", 99, "done", "outside"}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, errorWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Name", "Score"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"Beta", 2}); err != nil {
			return err
		}
		return nil
	})

	read := callTool(t, ctx, clientSession, "sort_range", map[string]any{
		"filepath":   singleAscWorkbook,
		"sheet_name": "Sheet1",
		"range":      "A1:B4",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "Name", "descending": false},
		},
	})

	var singleAscPayload struct {
		FilePath  string `json:"filepath"`
		SheetName string `json:"sheet_name"`
		Range     string `json:"range"`
		SortKeys  []struct {
			Column     string `json:"column"`
			Descending bool   `json:"descending"`
		} `json:"sort_keys"`
	}
	if err := json.Unmarshal([]byte(read), &singleAscPayload); err != nil {
		t.Fatalf("unmarshal sort asc payload: %v\n%s", err, read)
	}
	if singleAscPayload.FilePath != singleAscWorkbook || singleAscPayload.SheetName != "Sheet1" || singleAscPayload.Range != "A1:B4" { //nolint:goconst
		t.Fatalf("unexpected sort asc payload: %+v", singleAscPayload)
	}
	if len(singleAscPayload.SortKeys) != 1 || singleAscPayload.SortKeys[0].Column != "Name" || singleAscPayload.SortKeys[0].Descending {
		t.Fatalf("unexpected sort asc keys: %+v", singleAscPayload.SortKeys)
	}

	fAsc, err := excelize.OpenFile(singleAscWorkbook)
	if err != nil {
		t.Fatalf("open asc workbook: %v", err)
	}
	defer func() { _ = fAsc.Close() }()
	if got, _ := fAsc.GetCellValue("Sheet1", "A1"); got != "Name" {
		t.Fatalf("expected header row preserved in A1, got %q", got)
	}
	if got, _ := fAsc.GetCellValue("Sheet1", "A2"); got != "Alpha" {
		t.Fatalf("expected first sorted row Alpha, got %q", got)
	}
	if got, _ := fAsc.GetCellValue("Sheet1", "A3"); got != "Beta" {
		t.Fatalf("expected second sorted row Beta, got %q", got)
	}
	if got, _ := fAsc.GetCellValue("Sheet1", "A4"); got != "Gamma" {
		t.Fatalf("expected third sorted row Gamma, got %q", got)
	}

	callTool(t, ctx, clientSession, "sort_range", map[string]any{
		"filepath":   singleDescWorkbook,
		"sheet_name": "Sheet1",
		"range":      "A1:B4",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "2", "descending": true},
		},
	})

	fDesc, err := excelize.OpenFile(singleDescWorkbook)
	if err != nil {
		t.Fatalf("open desc workbook: %v", err)
	}
	defer func() { _ = fDesc.Close() }()
	if got, _ := fDesc.GetCellValue("Sheet1", "A2"); got != "Alpha" {
		t.Fatalf("expected highest score first, got %q", got)
	}
	if got, _ := fDesc.GetCellValue("Sheet1", "A3"); got != "Beta" {
		t.Fatalf("expected middle score second, got %q", got)
	}
	if got, _ := fDesc.GetCellValue("Sheet1", "A4"); got != "Gamma" {
		t.Fatalf("expected lowest score last, got %q", got)
	}

	callTool(t, ctx, clientSession, "sort_range", map[string]any{
		"filepath":   multiWorkbook,
		"sheet_name": "Sheet1",
		"range":      "A1:C5",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "Region"},
			{"column": "Quarter", "descending": true},
		},
	})

	fMulti, err := excelize.OpenFile(multiWorkbook)
	if err != nil {
		t.Fatalf("open multi workbook: %v", err)
	}
	defer func() { _ = fMulti.Close() }()
	if got, _ := fMulti.GetCellValue("Sheet1", "A2"); got != "East" {
		t.Fatalf("expected East row first after multi sort, got %q", got)
	}
	if got, _ := fMulti.GetCellValue("Sheet1", "B2"); got != "Q3" {
		t.Fatalf("expected East Q3 first within East group, got %q", got)
	}
	if got, _ := fMulti.GetCellValue("Sheet1", "B3"); got != "Q1" {
		t.Fatalf("expected East Q1 second within East group, got %q", got)
	}
	if got, _ := fMulti.GetCellValue("Sheet1", "B4"); got != "Q2" {
		t.Fatalf("expected West Q2 first within West group, got %q", got)
	}
	if got, _ := fMulti.GetCellValue("Sheet1", "B5"); got != "Q1" {
		t.Fatalf("expected West Q1 last within West group, got %q", got)
	}

	callTool(t, ctx, clientSession, "sort_range", map[string]any{
		"filepath":   partialWorkbook,
		"sheet_name": "Sheet1",
		"range":      "A1:C4",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "Score"},
		},
	})

	fPartial, err := excelize.OpenFile(partialWorkbook)
	if err != nil {
		t.Fatalf("open partial workbook: %v", err)
	}
	defer func() { _ = fPartial.Close() }()
	if got, _ := fPartial.GetCellValue("Sheet1", "A2"); got != "Gamma" {
		t.Fatalf("expected partial sort first row Gamma, got %q", got)
	}
	if got, _ := fPartial.GetCellValue("Sheet1", "A4"); got != "Alpha" {
		t.Fatalf("expected partial sort last row Alpha, got %q", got)
	}
	if got, _ := fPartial.GetCellValue("Sheet1", "D2"); got != "keep-1" {
		t.Fatalf("expected column outside sort range to stay untouched at D2, got %q", got)
	}
	if got, _ := fPartial.GetCellValue("Sheet1", "D4"); got != "keep-3" {
		t.Fatalf("expected column outside sort range to stay untouched at D4, got %q", got)
	}
	if got, _ := fPartial.GetCellValue("Sheet1", "A5"); got != "Tail" {
		t.Fatalf("expected row outside sort range to stay untouched at A5, got %q", got)
	}

	assertToolErrorContains(t, ctx, clientSession, "sort_range", map[string]any{
		"filepath":   errorWorkbook,
		"sheet_name": "Sheet1",
		"range":      "A1:B2",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "Missing"},
		},
	}, "column \"Missing\" not found in header row")

	assertToolErrorContains(t, ctx, clientSession, "sort_range", map[string]any{
		"filepath":   errorWorkbook,
		"sheet_name": "Sheet1",
		"range":      "A1:B2",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "9"},
		},
	}, "column index 9 is out of range")

	assertToolErrorContains(t, ctx, clientSession, "sort_range", map[string]any{
		"filepath":   errorWorkbook,
		"sheet_name": "Missing",
		"range":      "A1:B2",
		"has_header": true,
		"sort_keys": []map[string]any{
			{"column": "Name"},
		},
	}, "sheet Missing does not exist")
}

func TestFilterRows(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "filter.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Region", "Revenue", "Status", "Owner"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10, "Open", "Ava"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 25, "Closed", "Ben"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A4", &[]any{"West", 18, "Open", "Cara"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A5", &[]any{"East", 30, "Open", "Drew"}); err != nil {
			return err
		}
		return nil
	})

	read := callTool(t, ctx, clientSession, "filter_rows", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"range":      "A1:D5",
		"has_header": true,
		"filters": []map[string]any{
			{"column": "Status", "operator": "equals", "value": "Open"},
		},
	})

	var equalsPayload struct {
		FilePath  string `json:"filepath"`
		SheetName string `json:"sheet_name"`
		Range     string `json:"range"`
		Filters   []struct {
			Column   string `json:"column"`
			Operator string `json:"operator"`
			Value    string `json:"value"`
		} `json:"filters"`
		Rows []map[string]any `json:"rows"`
	}
	if err := json.Unmarshal([]byte(read), &equalsPayload); err != nil {
		t.Fatalf("unmarshal filter equals payload: %v\n%s", err, read)
	}
	if equalsPayload.FilePath != workbook || equalsPayload.SheetName != "Sheet1" || equalsPayload.Range != "A1:D5" {
		t.Fatalf("unexpected filter payload target: %+v", equalsPayload)
	}
	if len(equalsPayload.Filters) != 1 || equalsPayload.Filters[0].Column != "Status" || equalsPayload.Filters[0].Operator != "equals" || equalsPayload.Filters[0].Value != "Open" {
		t.Fatalf("unexpected filter payload filters: %+v", equalsPayload.Filters)
	}
	if len(equalsPayload.Rows) != 3 {
		t.Fatalf("expected 3 open rows, got %+v", equalsPayload.Rows)
	}
	if equalsPayload.Rows[0]["Region"] != "North" || equalsPayload.Rows[2]["Owner"] != "Drew" {
		t.Fatalf("unexpected filtered rows: %+v", equalsPayload.Rows)
	}

	read = callTool(t, ctx, clientSession, "filter_rows", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"range":      "A1:D5",
		"has_header": true,
		"filters": []map[string]any{
			{"column": "2", "operator": "gte", "value": "20"},
			{"column": "4", "operator": "contains", "value": "e"},
		},
	})

	var multiPayload struct {
		Rows []map[string]any `json:"rows"`
	}
	if err := json.Unmarshal([]byte(read), &multiPayload); err != nil {
		t.Fatalf("unmarshal filter multi payload: %v\n%s", err, read)
	}
	if len(multiPayload.Rows) != 2 {
		t.Fatalf("expected 2 rows after numeric and contains filters, got %+v", multiPayload.Rows)
	}
	if multiPayload.Rows[0]["Region"] != "South" || multiPayload.Rows[1]["Region"] != "East" { //nolint:goconst
		t.Fatalf("unexpected multi-filter rows: %+v", multiPayload.Rows)
	}

	read = callTool(t, ctx, clientSession, "filter_rows", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"range":      "A1:C4",
		"has_header": true,
		"filters": []map[string]any{
			{"column": "Revenue", "operator": "lt", "value": "20"},
		},
	})

	var partialPayload struct {
		Rows []map[string]any `json:"rows"`
	}
	if err := json.Unmarshal([]byte(read), &partialPayload); err != nil {
		t.Fatalf("unmarshal filter partial payload: %v\n%s", err, read)
	}
	if len(partialPayload.Rows) != 2 {
		t.Fatalf("expected 2 rows from partial range, got %+v", partialPayload.Rows)
	}
	if _, ok := partialPayload.Rows[0]["Owner"]; ok {
		t.Fatalf("expected partial range to exclude Owner column: %+v", partialPayload.Rows[0])
	}
	if partialPayload.Rows[0]["Region"] != "North" || partialPayload.Rows[1]["Region"] != "West" {
		t.Fatalf("unexpected partial filter rows: %+v", partialPayload.Rows)
	}

	assertToolErrorContains(t, ctx, clientSession, "filter_rows", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"range":      "A1:D5",
		"has_header": true,
		"filters": []map[string]any{
			{"column": "Missing", "operator": "equals", "value": "Open"},
		},
	}, "column \"Missing\" not found in header row")

	assertToolErrorContains(t, ctx, clientSession, "filter_rows", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"range":      "A1:D5",
		"has_header": true,
		"filters": []map[string]any{
			{"column": "Status", "operator": "between", "value": "Open"},
		},
	}, "unsupported operator \"between\"")

	assertToolErrorContains(t, ctx, clientSession, "filter_rows", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Missing",
		"range":      "A1:D5",
		"has_header": true,
		"filters": []map[string]any{
			{"column": "Status", "operator": "equals", "value": "Open"},
		},
	}, "sheet Missing does not exist")
}

func TestUpsertRows(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	mixedWorkbook := filepath.Join(t.TempDir(), "upsert-mixed.xlsx")
	compositeWorkbook := filepath.Join(t.TempDir(), "upsert-composite.xlsx")
	skipWorkbook := filepath.Join(t.TempDir(), "upsert-skip.xlsx")
	duplicateWorkbook := filepath.Join(t.TempDir(), "upsert-duplicate.xlsx")
	capacityWorkbook := filepath.Join(t.TempDir(), "upsert-capacity.xlsx")
	errorWorkbook := filepath.Join(t.TempDir(), "upsert-errors.xlsx")

	createWorkbookFixture(t, mixedWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"OrderID", "Customer", "Amount", "Status"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"A-1001", "Contoso", 100, "Open"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"A-1002", "Fabrikam", 250, "Open"}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, compositeWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"OrderID", "LineID", "Sku", "Qty"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"A-1", "1", "SKU-1", 2}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"A-1", "2", "SKU-2", 3}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, skipWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"OrderID", "Customer", "Amount", "Status"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"A-1001", "Contoso", 100, "Open"}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, duplicateWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"OrderID", "Customer", "Amount"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"A-1001", "Contoso", 100}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"A-1001", "Contoso-2", 125}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, capacityWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"OrderID", "Customer", "Amount"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"A-1001", "Contoso", 100}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"A-1002", "Fabrikam", 125}); err != nil {
			return err
		}
		return nil
	})

	createWorkbookFixture(t, errorWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"OrderID", "Customer", "Amount", "Status"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"A-1001", "Contoso", 100, "Open"}); err != nil {
			return err
		}
		return nil
	})

	read := callTool(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          mixedWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:D6",
		"key_columns":       []string{"OrderID"},
		"insert_if_missing": true,
		"rows": []map[string]any{
			{
				"match":  map[string]any{"OrderID": "A-1001"},
				"values": map[string]any{"Status": "Paid", "Amount": 125},
			},
			{
				"match":  map[string]any{"OrderID": "A-1003"},
				"values": map[string]any{"Customer": "Northwind", "Amount": 88, "Status": "Open"},
			},
		},
	})

	var mixedPayload struct {
		FilePath      string   `json:"filepath"`
		SheetName     string   `json:"sheet_name"`
		Range         string   `json:"range"`
		KeyColumns    []string `json:"key_columns"`
		UpdatedCount  int      `json:"updated_count"`
		InsertedCount int      `json:"inserted_count"`
		SkippedCount  int      `json:"skipped_count"`
		Results       []struct {
			Key       map[string]any `json:"key"`
			Action    string         `json:"action"`
			RowNumber int            `json:"row_number,omitempty"`
			Reason    string         `json:"reason,omitempty"`
		} `json:"results"`
	}
	if err := json.Unmarshal([]byte(read), &mixedPayload); err != nil {
		t.Fatalf("unmarshal upsert mixed payload: %v\n%s", err, read)
	}
	if mixedPayload.FilePath != mixedWorkbook || mixedPayload.SheetName != "Sheet1" || mixedPayload.Range != "A1:D6" {
		t.Fatalf("unexpected upsert mixed target: %+v", mixedPayload)
	}
	if len(mixedPayload.KeyColumns) != 1 || mixedPayload.KeyColumns[0] != "OrderID" || mixedPayload.UpdatedCount != 1 || mixedPayload.InsertedCount != 1 || mixedPayload.SkippedCount != 0 {
		t.Fatalf("unexpected upsert mixed summary: %+v", mixedPayload)
	}
	if len(mixedPayload.Results) != 2 || mixedPayload.Results[0].Action != "updated" || mixedPayload.Results[0].RowNumber != 2 || mixedPayload.Results[1].Action != "inserted" || mixedPayload.Results[1].RowNumber != 4 {
		t.Fatalf("unexpected upsert mixed results: %+v", mixedPayload.Results)
	}

	fMixed, err := excelize.OpenFile(mixedWorkbook)
	if err != nil {
		t.Fatalf("open mixed workbook: %v", err)
	}
	defer func() { _ = fMixed.Close() }()
	if got, _ := fMixed.GetCellValue("Sheet1", "D2"); got != "Paid" {
		t.Fatalf("expected updated status in D2, got %q", got)
	}
	if got, _ := fMixed.GetCellValue("Sheet1", "B2"); got != "Contoso" {
		t.Fatalf("expected untouched customer in B2, got %q", got)
	}
	if got, _ := fMixed.GetCellValue("Sheet1", "A4"); got != "A-1003" {
		t.Fatalf("expected inserted key in A4, got %q", got)
	}
	if got, _ := fMixed.GetCellValue("Sheet1", "B4"); got != "Northwind" {
		t.Fatalf("expected inserted customer in B4, got %q", got)
	}

	read = callTool(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          compositeWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:D5",
		"key_columns":       []string{"OrderID", "LineID"},
		"insert_if_missing": true,
		"rows": []map[string]any{
			{
				"match":  map[string]any{"OrderID": "A-1", "LineID": "2"},
				"values": map[string]any{"Qty": 5},
			},
			{
				"match":  map[string]any{"OrderID": "A-1", "LineID": "3"},
				"values": map[string]any{"Sku": "SKU-3", "Qty": 1},
			},
		},
	})

	var compositePayload struct {
		UpdatedCount  int `json:"updated_count"`
		InsertedCount int `json:"inserted_count"`
	}
	if err := json.Unmarshal([]byte(read), &compositePayload); err != nil {
		t.Fatalf("unmarshal upsert composite payload: %v\n%s", err, read)
	}
	if compositePayload.UpdatedCount != 1 || compositePayload.InsertedCount != 1 {
		t.Fatalf("unexpected composite summary: %+v", compositePayload)
	}

	fComposite, err := excelize.OpenFile(compositeWorkbook)
	if err != nil {
		t.Fatalf("open composite workbook: %v", err)
	}
	defer func() { _ = fComposite.Close() }()
	if got, _ := fComposite.GetCellValue("Sheet1", "D3"); got != "5" {
		t.Fatalf("expected updated composite qty in D3, got %q", got)
	}
	if got, _ := fComposite.GetCellValue("Sheet1", "B4"); got != "3" {
		t.Fatalf("expected inserted composite line in B4, got %q", got)
	}

	read = callTool(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          skipWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:D4",
		"key_columns":       []string{"OrderID"},
		"insert_if_missing": false,
		"rows": []map[string]any{
			{
				"match":  map[string]any{"OrderID": "A-4040"},
				"values": map[string]any{"Status": "Cancelled"},
			},
		},
	})

	var skipPayload struct {
		UpdatedCount  int `json:"updated_count"`
		InsertedCount int `json:"inserted_count"`
		SkippedCount  int `json:"skipped_count"`
		Results       []struct {
			Action string `json:"action"`
			Reason string `json:"reason"`
		} `json:"results"`
	}
	if err := json.Unmarshal([]byte(read), &skipPayload); err != nil {
		t.Fatalf("unmarshal upsert skip payload: %v\n%s", err, read)
	}
	if skipPayload.UpdatedCount != 0 || skipPayload.InsertedCount != 0 || skipPayload.SkippedCount != 1 || len(skipPayload.Results) != 1 || skipPayload.Results[0].Action != "skipped" {
		t.Fatalf("unexpected skip payload: %+v", skipPayload)
	}

	assertToolErrorContains(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          duplicateWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:C4",
		"key_columns":       []string{"OrderID"},
		"insert_if_missing": true,
		"rows": []map[string]any{
			{
				"match":  map[string]any{"OrderID": "A-1001"},
				"values": map[string]any{"Amount": 999},
			},
		},
	}, "multiple existing rows matched key")

	assertToolErrorContains(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          capacityWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:C3",
		"key_columns":       []string{"OrderID"},
		"insert_if_missing": true,
		"rows": []map[string]any{
			{
				"match":  map[string]any{"OrderID": "A-1003"},
				"values": map[string]any{"Customer": "Northwind", "Amount": 88},
			},
		},
	}, "no empty rows remain within range for insert")

	assertToolErrorContains(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          errorWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:D3",
		"key_columns":       []string{"MissingKey"},
		"insert_if_missing": true,
		"rows": []map[string]any{
			{
				"match":  map[string]any{"MissingKey": "A-1001"},
				"values": map[string]any{"Status": "Paid"},
			},
		},
	}, "key column \"MissingKey\" not found in header row")

	assertToolErrorContains(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          errorWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:D3",
		"key_columns":       []string{"OrderID"},
		"insert_if_missing": true,
		"rows": []map[string]any{
			{
				"match":  map[string]any{},
				"values": map[string]any{"Status": "Paid"},
			},
		},
	}, "row 1 match is missing key column \"OrderID\"")

	assertToolErrorContains(t, ctx, clientSession, "upsert_rows", map[string]any{
		"filepath":          errorWorkbook,
		"sheet_name":        "Sheet1",
		"range":             "A1:D3",
		"key_columns":       []string{"OrderID"},
		"insert_if_missing": true,
		"rows": []map[string]any{
			{
				"match":  map[string]any{"OrderID": "A-1001"},
				"values": map[string]any{"OrderID": "A-2000"},
			},
		},
	}, "values cannot modify key column \"OrderID\"")
}

func TestAdvancedToolValidationErrors(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "errors.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Header", "Header"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"x", 1}); err != nil {
			return err
		}
		return nil
	})

	assertToolErrorContains(t, ctx, clientSession, "create_chart", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "A1:B2",
		"chart_type":  "hexbin",
		"target_cell": "D2",
	}, "unsupported chart_type")

	assertToolErrorContains(t, ctx, clientSession, "create_pivot_table", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "A1:B2",
		"target_cell": "D10",
		"rows":        []string{"Header"},
		"values":      []string{"Header"},
		"agg_func":    "median",
	}, "unsupported agg_func")

	assertToolErrorContains(t, ctx, clientSession, "create_table", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"data_range": "A1:B2",
	}, "table header cells must be unique")

	assertToolErrorContains(t, ctx, clientSession, "delete_range", map[string]any{
		"filepath":        workbook,
		"sheet_name":      "Sheet1",
		"start_cell":      "A1",
		"end_cell":        "A1",
		"shift_direction": "down",
	}, "shift_direction must be 'up' or 'left'")
}

func TestDescribeWorkbook(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "describe.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if _, err := f.NewSheet("Summary"); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Region", "Revenue", "Active"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10, true}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 20, false}); err != nil {
			return err
		}
		if err := f.MergeCell("Sheet1", "D1", "E1"); err != nil {
			return err
		}
		if err := f.SetCellValue("Sheet1", "D1", "Merged Title"); err != nil {
			return err
		}
		dv := excelize.NewDataValidation(true)
		dv.Sqref = "C2:C3"
		if err := dv.SetDropList([]string{"true", "false"}); err != nil {
			return err
		}
		dv.SetError(excelize.DataValidationErrorStyleStop, "Active value required", "Choose true or false")
		if err := f.AddDataValidation("Sheet1", dv); err != nil {
			return err
		}
		if err := f.SetDefinedName(&excelize.DefinedName{Name: "Regions", RefersTo: "Sheet1!$A$2:$A$3"}); err != nil {
			return err
		}
		return nil
	})

	callTool(t, ctx, clientSession, "create_table", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "A1:C3",
		"table_name":  "RegionTable",
		"table_style": "TableStyleMedium2",
	})
	callTool(t, ctx, clientSession, "create_pivot_table", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "A1:C3",
		"target_cell": "H20",
		"rows":        []string{"Region"},
		"values":      []string{"Revenue"},
		"agg_func":    "sum",
	})
	callTool(t, ctx, clientSession, "create_chart", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "A1:C3",
		"chart_type":  "column",
		"target_cell": "E5",
		"title":       "Revenue Chart",
	})

	read := callTool(t, ctx, clientSession, "describe_workbook", map[string]any{
		"filepath":           workbook,
		"include_ranges":     true,
		"include_merged":     true,
		"include_tables":     true,
		"include_charts":     true,
		"include_pivots":     true,
		"include_names":      true,
		"include_validation": true,
	})

	var payload struct {
		FilePath    string `json:"filepath"`
		NamedRanges []struct {
			Name     string `json:"name"`
			RefersTo string `json:"refers_to"`
			Scope    string `json:"scope"`
		} `json:"named_ranges"`
		Sheets []struct {
			Name         string   `json:"name"`
			Range        string   `json:"range"`
			Rows         int      `json:"rows"`
			Cols         int      `json:"cols"`
			MergedRanges []string `json:"merged_ranges"`
			Tables       []struct {
				Name      string `json:"name"`
				Range     string `json:"range"`
				StyleName string `json:"style_name"`
			} `json:"tables"`
			Charts []struct {
				AnchorCell string `json:"anchor_cell"`
				Title      string `json:"title"`
				ChartType  string `json:"chart_type"`
				ChartPath  string `json:"chart_path"`
				Series     []struct {
					DisplayName     string `json:"display_name"`
					NameRef         string `json:"name_ref"`
					SourceSheet     string `json:"source_sheet"`
					SourceRange     string `json:"source_range"`
					CategoriesRange string `json:"categories_range"`
					ValuesRange     string `json:"values_range"`
				} `json:"series"`
			} `json:"charts"`
			PivotTables []struct {
				Name            string   `json:"name"`
				DataRange       string   `json:"data_range"`
				PivotTableRange string   `json:"pivot_table_range"`
				Rows            []string `json:"rows"`
				Values          []struct {
					Data     string `json:"data"`
					Subtotal string `json:"subtotal"`
				} `json:"values"`
			} `json:"pivot_tables"`
			DataValidations []struct {
				Sqref      string `json:"sqref"`
				Type       string `json:"type"`
				Formula1   string `json:"formula1"`
				ErrorTitle string `json:"error_title"`
				ErrorBody  string `json:"error_body"`
			} `json:"data_validations"`
		} `json:"sheets"`
	}
	if err := json.Unmarshal([]byte(read), &payload); err != nil {
		t.Fatalf("unmarshal describe payload: %v\n%s", err, read)
	}
	if payload.FilePath != workbook {
		t.Fatalf("expected filepath %q, got %q", workbook, payload.FilePath)
	}
	if len(payload.Sheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(payload.Sheets))
	}
	if payload.Sheets[0].Name != "Sheet1" {
		t.Fatalf("unexpected first sheet: %+v", payload.Sheets[0])
	}
	if payload.Sheets[0].Range != "A1:E3" {
		t.Fatalf("expected Sheet1 range A1:E3, got %q", payload.Sheets[0].Range)
	}
	if payload.Sheets[0].Rows != 3 || payload.Sheets[0].Cols != 5 {
		t.Fatalf("unexpected Sheet1 dimensions: %+v", payload.Sheets[0])
	}
	if len(payload.Sheets[0].MergedRanges) != 1 || payload.Sheets[0].MergedRanges[0] != "D1:E1" {
		t.Fatalf("unexpected merged ranges: %+v", payload.Sheets[0].MergedRanges)
	}
	if len(payload.Sheets[0].Tables) != 1 || payload.Sheets[0].Tables[0].Name != "RegionTable" || payload.Sheets[0].Tables[0].Range != "A1:C3" || payload.Sheets[0].Tables[0].StyleName != "TableStyleMedium2" {
		t.Fatalf("unexpected tables: %+v", payload.Sheets[0].Tables)
	}
	if len(payload.Sheets[0].Charts) != 1 || payload.Sheets[0].Charts[0].AnchorCell != "E5" || payload.Sheets[0].Charts[0].Title != "Revenue Chart" || payload.Sheets[0].Charts[0].ChartType != "column" || payload.Sheets[0].Charts[0].ChartPath != "xl/charts/chart1.xml" {
		t.Fatalf("unexpected charts: %+v", payload.Sheets[0].Charts)
	}
	if len(payload.Sheets[0].Charts[0].Series) != 2 {
		t.Fatalf("unexpected chart series: %+v", payload.Sheets[0].Charts[0].Series)
	}
	if payload.Sheets[0].Charts[0].Series[0].DisplayName != "North" || payload.Sheets[0].Charts[0].Series[0].NameRef != "Sheet1!$A$2" || payload.Sheets[0].Charts[0].Series[0].SourceSheet != "Sheet1" || payload.Sheets[0].Charts[0].Series[0].SourceRange != "Sheet1!A1:C2" || payload.Sheets[0].Charts[0].Series[0].CategoriesRange != "Sheet1!$B$1:$C$1" || payload.Sheets[0].Charts[0].Series[0].ValuesRange != "Sheet1!$B$2:$C$2" {
		t.Fatalf("unexpected first chart series: %+v", payload.Sheets[0].Charts[0].Series[0])
	}
	if len(payload.Sheets[0].PivotTables) != 1 || payload.Sheets[0].PivotTables[0].Name == "" {
		t.Fatalf("unexpected pivot tables: %+v", payload.Sheets[0].PivotTables)
	}
	if payload.Sheets[0].PivotTables[0].DataRange != "Sheet1!A1:C3" {
		t.Fatalf("unexpected pivot data range: %+v", payload.Sheets[0].PivotTables[0])
	}
	if len(payload.Sheets[0].PivotTables[0].Rows) != 1 || payload.Sheets[0].PivotTables[0].Rows[0] != "Region" {
		t.Fatalf("unexpected pivot rows: %+v", payload.Sheets[0].PivotTables[0].Rows)
	}
	if len(payload.Sheets[0].PivotTables[0].Values) != 1 || payload.Sheets[0].PivotTables[0].Values[0].Data != "Revenue" || payload.Sheets[0].PivotTables[0].Values[0].Subtotal != "Sum" {
		t.Fatalf("unexpected pivot values: %+v", payload.Sheets[0].PivotTables[0].Values)
	}
	if len(payload.Sheets[0].DataValidations) != 1 || payload.Sheets[0].DataValidations[0].Sqref != "C2:C3" || payload.Sheets[0].DataValidations[0].Type != "list" || payload.Sheets[0].DataValidations[0].ErrorTitle != "Active value required" {
		t.Fatalf("unexpected validations: %+v", payload.Sheets[0].DataValidations)
	}
	if len(payload.NamedRanges) != 1 || payload.NamedRanges[0].Name != "Regions" || payload.NamedRanges[0].RefersTo != "Sheet1!$A$2:$A$3" {
		t.Fatalf("unexpected named ranges: %+v", payload.NamedRanges)
	}
	if payload.Sheets[1].Name != "Summary" { //nolint:goconst
		t.Fatalf("unexpected second sheet: %+v", payload.Sheets[1])
	}
}

func TestGetSheetSchema(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "schema.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Region", "Revenue", "Active", "Notes"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10, true, "priority"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 15.5, false, ""}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A4", &[]any{"West", 0, true, "watch"}); err != nil {
			return err
		}
		return nil
	})

	read := callTool(t, ctx, clientSession, "get_sheet_schema", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"start_cell":  "A1",
		"end_cell":    "D4",
		"header_row":  1,
		"sample_size": 2,
	})

	var payload struct {
		FilePath  string `json:"filepath"`
		SheetName string `json:"sheet_name"`
		Range     string `json:"range"`
		HeaderRow int    `json:"header_row"`
		RowCount  int    `json:"row_count"`
		Columns   []struct {
			Name         string   `json:"name"`
			Column       string   `json:"column"`
			InferredType string   `json:"inferred_type"`
			BlankCount   int      `json:"blank_count"`
			SampleValues []string `json:"sample_values"`
		} `json:"columns"`
	}
	if err := json.Unmarshal([]byte(read), &payload); err != nil {
		t.Fatalf("unmarshal schema payload: %v\n%s", err, read)
	}
	if payload.FilePath != workbook || payload.SheetName != "Sheet1" {
		t.Fatalf("unexpected payload target: %+v", payload)
	}
	if payload.Range != "A1:D4" {
		t.Fatalf("expected range A1:D4, got %q", payload.Range)
	}
	if payload.HeaderRow != 1 || payload.RowCount != 3 {
		t.Fatalf("unexpected schema shape: %+v", payload)
	}
	if len(payload.Columns) != 4 {
		t.Fatalf("expected 4 columns, got %d", len(payload.Columns))
	}

	if payload.Columns[0].Name != "Region" || payload.Columns[0].Column != "A" || payload.Columns[0].InferredType != "string" {
		t.Fatalf("unexpected Region column: %+v", payload.Columns[0])
	}
	if len(payload.Columns[0].SampleValues) != 2 || payload.Columns[0].SampleValues[0] != "North" || payload.Columns[0].SampleValues[1] != "South" {
		t.Fatalf("unexpected Region samples: %+v", payload.Columns[0].SampleValues)
	}

	if payload.Columns[1].Name != "Revenue" || payload.Columns[1].Column != "B" || payload.Columns[1].InferredType != "number" {
		t.Fatalf("unexpected Revenue column: %+v", payload.Columns[1])
	}

	if payload.Columns[2].Name != "Active" || payload.Columns[2].Column != "C" || payload.Columns[2].InferredType != "boolean" {
		t.Fatalf("unexpected Active column: %+v", payload.Columns[2])
	}

	if payload.Columns[3].Name != "Notes" || payload.Columns[3].Column != "D" || payload.Columns[3].InferredType != "string" || payload.Columns[3].BlankCount != 1 {
		t.Fatalf("unexpected Notes column: %+v", payload.Columns[3])
	}
	if len(payload.Columns[3].SampleValues) != 2 || payload.Columns[3].SampleValues[0] != "priority" || payload.Columns[3].SampleValues[1] != "watch" {
		t.Fatalf("unexpected Notes samples: %+v", payload.Columns[3].SampleValues)
	}
}

func TestGetSheetSchemaDefaultsAndErrors(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "schema-errors.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Region", "Revenue"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 15}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A4", &[]any{"West", 20}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A5", &[]any{"East", 25}); err != nil {
			return err
		}
		return nil
	})

	read := callTool(t, ctx, clientSession, "get_sheet_schema", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "A1",
		"end_cell":   "B5",
	})

	var payload struct {
		Columns []struct {
			SampleValues []string `json:"sample_values"`
		} `json:"columns"`
	}
	if err := json.Unmarshal([]byte(read), &payload); err != nil {
		t.Fatalf("unmarshal default schema payload: %v\n%s", err, read)
	}
	if len(payload.Columns) != 2 || len(payload.Columns[0].SampleValues) != 3 {
		t.Fatalf("expected default sample size of 3, got %+v", payload.Columns)
	}
	if payload.Columns[0].SampleValues[0] != "North" || payload.Columns[0].SampleValues[1] != "South" || payload.Columns[0].SampleValues[2] != "West" {
		t.Fatalf("unexpected default samples: %+v", payload.Columns[0].SampleValues)
	}

	assertToolErrorContains(t, ctx, clientSession, "get_sheet_schema", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "1A",
	}, "invalid start_cell")

	assertToolErrorContains(t, ctx, clientSession, "get_sheet_schema", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "A1",
		"end_cell":   "Z",
	}, "invalid end_cell")

	assertToolErrorContains(t, ctx, clientSession, "get_sheet_schema", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Missing",
		"start_cell": "A1",
	}, "sheet Missing does not exist")

	assertToolErrorContains(t, ctx, clientSession, "get_sheet_schema", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"start_cell": "A1",
		"end_cell":   "B5",
		"header_row": 9,
	}, "header_row must be within the selected range")
}

func TestFindInWorkbook(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "find.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if _, err := f.NewSheet("Summary"); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Region", "Revenue"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 20}); err != nil {
			return err
		}
		if err := f.SetCellFormula("Summary", "B2", "=SUM(Sheet1!B2:B3)"); err != nil {
			return err
		}
		if err := f.SetCellValue("Summary", "A1", "South region total"); err != nil {
			return err
		}
		return nil
	})

	read := callTool(t, ctx, clientSession, "find_in_workbook", map[string]any{
		"filepath":   workbook,
		"query":      "south",
		"match_mode": "contains",
	})

	var textPayload struct {
		Query   string `json:"query"`
		Matches []struct {
			SheetName string `json:"sheet_name"`
			Cell      string `json:"cell"`
			Value     string `json:"value"`
		} `json:"matches"`
	}
	if err := json.Unmarshal([]byte(read), &textPayload); err != nil {
		t.Fatalf("unmarshal find payload: %v\n%s", err, read)
	}
	if textPayload.Query != "south" || len(textPayload.Matches) != 2 {
		t.Fatalf("unexpected text matches: %+v", textPayload)
	}
	if textPayload.Matches[0].SheetName != "Sheet1" || textPayload.Matches[0].Cell != "A3" || textPayload.Matches[0].Value != "South" {
		t.Fatalf("unexpected first text match: %+v", textPayload.Matches[0])
	}

	read = callTool(t, ctx, clientSession, "find_in_workbook", map[string]any{
		"filepath":    workbook,
		"query":       "sum",
		"search_type": "formula",
	})

	var formulaPayload struct {
		Matches []struct {
			SheetName string `json:"sheet_name"`
			Cell      string `json:"cell"`
			Formula   string `json:"formula"`
		} `json:"matches"`
	}
	if err := json.Unmarshal([]byte(read), &formulaPayload); err != nil {
		t.Fatalf("unmarshal formula find payload: %v\n%s", err, read)
	}
	if len(formulaPayload.Matches) != 1 || formulaPayload.Matches[0].SheetName != "Summary" || formulaPayload.Matches[0].Cell != "B2" || formulaPayload.Matches[0].Formula != "=SUM(Sheet1!B2:B3)" {
		t.Fatalf("unexpected formula matches: %+v", formulaPayload.Matches)
	}

	read = callTool(t, ctx, clientSession, "find_in_workbook", map[string]any{
		"filepath":     workbook,
		"query":        "^south",
		"match_mode":   "regex",
		"context_rows": 1,
		"context_cols": 1,
	})

	var regexPayload struct {
		Matches []struct {
			SheetName string `json:"sheet_name"`
			Cell      string `json:"cell"`
			Value     string `json:"value"`
			Context   []struct {
				Cell  string `json:"cell"`
				Value string `json:"value"`
			} `json:"context"`
		} `json:"matches"`
	}
	if err := json.Unmarshal([]byte(read), &regexPayload); err != nil {
		t.Fatalf("unmarshal regex find payload: %v\n%s", err, read)
	}
	if len(regexPayload.Matches) != 2 {
		t.Fatalf("unexpected regex matches: %+v", regexPayload.Matches)
	}
	if regexPayload.Matches[0].Cell != "A3" || regexPayload.Matches[0].Value != "South" {
		t.Fatalf("unexpected first regex match: %+v", regexPayload.Matches[0])
	}
	if len(regexPayload.Matches[0].Context) == 0 {
		t.Fatalf("expected context for regex match: %+v", regexPayload.Matches[0])
	}
	var foundNeighbor bool
	for _, item := range regexPayload.Matches[0].Context {
		if item.Cell == "B3" && item.Value == "20" {
			foundNeighbor = true
			break
		}
	}
	if !foundNeighbor {
		t.Fatalf("expected B3 neighbor in context: %+v", regexPayload.Matches[0].Context)
	}

	read = callTool(t, ctx, clientSession, "find_in_workbook", map[string]any{
		"filepath":    workbook,
		"query":       "^south",
		"match_mode":  "regex",
		"sheets":      []string{"Summary"},
		"max_results": 1,
	})

	var scopedPayload struct {
		Matches []struct {
			SheetName string `json:"sheet_name"`
			Cell      string `json:"cell"`
			Value     string `json:"value"`
		} `json:"matches"`
	}
	if err := json.Unmarshal([]byte(read), &scopedPayload); err != nil {
		t.Fatalf("unmarshal scoped find payload: %v\n%s", err, read)
	}
	if len(scopedPayload.Matches) != 1 || scopedPayload.Matches[0].SheetName != "Summary" || scopedPayload.Matches[0].Cell != "A1" || scopedPayload.Matches[0].Value != "South region total" {
		t.Fatalf("unexpected scoped regex matches: %+v", scopedPayload.Matches)
	}

	read = callTool(t, ctx, clientSession, "find_in_workbook", map[string]any{
		"filepath":    workbook,
		"query":       "south",
		"match_mode":  "contains",
		"max_results": 1,
	})

	var limitedPayload struct {
		Matches []struct {
			SheetName string `json:"sheet_name"`
			Cell      string `json:"cell"`
		} `json:"matches"`
	}
	if err := json.Unmarshal([]byte(read), &limitedPayload); err != nil {
		t.Fatalf("unmarshal limited find payload: %v\n%s", err, read)
	}
	if len(limitedPayload.Matches) != 1 || limitedPayload.Matches[0].SheetName != "Sheet1" || limitedPayload.Matches[0].Cell != "A3" {
		t.Fatalf("unexpected limited matches: %+v", limitedPayload.Matches)
	}

	assertToolErrorContains(t, ctx, clientSession, "find_in_workbook", map[string]any{
		"filepath": workbook,
		"query":    "",
	}, "query is required")

	assertToolErrorContains(t, ctx, clientSession, "find_in_workbook", map[string]any{
		"filepath":   workbook,
		"query":      "(",
		"match_mode": "regex",
	}, "invalid regex query")
}

func TestListCharts(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "charts.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Series", "Q1", "Q2"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"North", 10, 20}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"South", 15, 25}); err != nil {
			return err
		}
		if _, err := f.NewSheet("Summary"); err != nil {
			return err
		}
		if err := f.SetSheetRow("Summary", "A1", &[]any{"Series", "Q1", "Q2"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Summary", "A2", &[]any{"Forecast", 11, 22}); err != nil {
			return err
		}
		return nil
	})

	callTool(t, ctx, clientSession, "create_chart", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Sheet1",
		"data_range":  "A1:C3",
		"chart_type":  "line",
		"target_cell": "F4",
		"title":       "Quarterly Trend",
	})
	callTool(t, ctx, clientSession, "create_chart", map[string]any{
		"filepath":    workbook,
		"sheet_name":  "Summary",
		"data_range":  "A1:C2",
		"chart_type":  "column",
		"target_cell": "E2",
		"title":       "Forecast Trend",
	})

	read := callTool(t, ctx, clientSession, "list_charts", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
	})

	var payload struct {
		FilePath  string `json:"filepath"`
		SheetName string `json:"sheet_name"`
		Charts    []struct {
			SheetName  string `json:"sheet_name"`
			AnchorCell string `json:"anchor_cell"`
			Title      string `json:"title"`
			ChartType  string `json:"chart_type"`
			ChartPath  string `json:"chart_path"`
			Series     []struct {
				DisplayName     string `json:"display_name"`
				NameRef         string `json:"name_ref"`
				SourceSheet     string `json:"source_sheet"`
				SourceRange     string `json:"source_range"`
				CategoriesRange string `json:"categories_range"`
				ValuesRange     string `json:"values_range"`
			} `json:"series"`
		} `json:"charts"`
	}
	if err := json.Unmarshal([]byte(read), &payload); err != nil {
		t.Fatalf("unmarshal list_charts payload: %v\n%s", err, read)
	}
	if payload.FilePath != workbook || payload.SheetName != "Sheet1" {
		t.Fatalf("unexpected list_charts target: %+v", payload)
	}
	if len(payload.Charts) != 1 || payload.Charts[0].AnchorCell != "F4" || payload.Charts[0].Title != "Quarterly Trend" || payload.Charts[0].ChartType != "line" {
		t.Fatalf("unexpected list_charts charts: %+v", payload.Charts)
	}
	if len(payload.Charts[0].Series) != 2 || payload.Charts[0].Series[1].DisplayName != "South" || payload.Charts[0].Series[1].NameRef != "Sheet1!$A$3" || payload.Charts[0].Series[1].SourceSheet != "Sheet1" || payload.Charts[0].Series[1].SourceRange != "Sheet1!A1:C3" {
		t.Fatalf("unexpected list_charts series: %+v", payload.Charts[0].Series)
	}

	read = callTool(t, ctx, clientSession, "list_charts", map[string]any{
		"filepath": workbook,
	})

	var allChartsPayload struct {
		Charts []struct {
			SheetName string `json:"sheet_name"`
			Title     string `json:"title"`
		} `json:"charts"`
	}
	if err := json.Unmarshal([]byte(read), &allChartsPayload); err != nil {
		t.Fatalf("unmarshal all list_charts payload: %v\n%s", err, read)
	}
	if len(allChartsPayload.Charts) != 2 {
		t.Fatalf("expected 2 workbook-wide charts, got %+v", allChartsPayload.Charts)
	}
	if allChartsPayload.Charts[0].SheetName != "Sheet1" || allChartsPayload.Charts[1].SheetName != "Summary" {
		t.Fatalf("unexpected workbook-wide chart ordering: %+v", allChartsPayload.Charts)
	}

	read = callTool(t, ctx, clientSession, "list_charts", map[string]any{
		"filepath":     workbook,
		"source_sheet": "Summary",
	})

	var filteredChartsPayload struct {
		Charts []struct {
			SheetName string `json:"sheet_name"`
			Title     string `json:"title"`
		} `json:"charts"`
	}
	if err := json.Unmarshal([]byte(read), &filteredChartsPayload); err != nil {
		t.Fatalf("unmarshal filtered list_charts payload: %v\n%s", err, read)
	}
	if len(filteredChartsPayload.Charts) != 1 || filteredChartsPayload.Charts[0].SheetName != "Summary" || filteredChartsPayload.Charts[0].Title != "Forecast Trend" {
		t.Fatalf("unexpected source-sheet filtered charts: %+v", filteredChartsPayload.Charts)
	}

	assertToolErrorContains(t, ctx, clientSession, "list_charts", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Missing",
	}, "sheet Missing does not exist")
}

func TestToolInputSchemaRequiredFields(t *testing.T) {
	ctx, clientSession := newTestClient(t)

	res, err := clientSession.ListTools(ctx, nil)
	if err != nil {
		t.Fatalf("list tools: %v", err)
	}

	type nestedRequirement struct {
		path     []string
		required []string
	}

	expectedRequired := map[string][]string{
		"apply_formula":            {"filepath", "sheet_name", "cell", "formula"},
		"clear_range":              {"filepath", "sheet_name", "start_cell"},
		"copy_range":               {"filepath", "sheet_name", "source_start", "source_end", "target_start"},
		"copy_worksheet":           {"filepath", "source_sheet", "target_sheet"},
		"create_chart":             {"filepath", "sheet_name", "data_range", "chart_type", "target_cell"},
		"create_pivot_table":       {"filepath", "sheet_name", "data_range", "target_cell", "rows", "values"},
		"create_table":             {"filepath", "sheet_name", "data_range"},
		"create_workbook":          {"filepath"},
		"create_worksheet":         {"filepath", "sheet_name"},
		"delete_range":             {"filepath", "sheet_name", "start_cell"},
		"delete_sheet_columns":     {"filepath", "sheet_name", "start_col"},
		"delete_sheet_rows":        {"filepath", "sheet_name", "start_row"},
		"delete_worksheet":         {"filepath", "sheet_name"},
		"describe_workbook":        {"filepath"},
		"filter_rows":              {"filepath", "sheet_name", "range", "filters"},
		"find_in_workbook":         {"filepath", "query"},
		"format_range":             {"filepath", "sheet_name", "start_cell"},
		"get_data_validation_info": {"filepath", "sheet_name"},
		"get_merged_cells":         {"filepath", "sheet_name"},
		"get_sheet_schema":         {"filepath", "sheet_name"},
		"get_workbook_metadata":    {"filepath"},
		"insert_columns":           {"filepath", "sheet_name", "start_col"},
		"insert_rows":              {"filepath", "sheet_name", "start_row"},
		"list_charts":              {"filepath"},
		"merge_cells":              {"filepath", "sheet_name", "start_cell", "end_cell"},
		"read_data_from_excel":     {"filepath", "sheet_name"},
		"rename_worksheet":         {"filepath", "old_name", "new_name"},
		"set_column_widths":        {"filepath", "sheet_name"},
		"set_row_heights":          {"filepath", "sheet_name", "heights"},
		"sort_range":               {"filepath", "sheet_name", "range", "sort_keys"},
		"unmerge_cells":            {"filepath", "sheet_name", "start_cell", "end_cell"},
		"upsert_rows":              {"filepath", "sheet_name", "range", "key_columns", "rows"},
		"validate_excel_range":     {"filepath", "sheet_name", "start_cell"},
		"validate_formula_syntax":  {"formula"},
		"write_data_to_excel":      {"filepath", "sheet_name", "data"},
	}

	expectedNested := map[string][]nestedRequirement{
		"filter_rows": {
			{path: []string{"filters", "items"}, required: []string{"column", "operator", "value"}},
		},
		"format_range": {
			{path: []string{"conditional_format"}, required: []string{"type"}},
		},
		"set_column_widths": {
			{path: []string{"widths", "items"}, required: []string{"column", "width"}},
		},
		"set_row_heights": {
			{path: []string{"heights", "items"}, required: []string{"row", "height"}},
		},
		"sort_range": {
			{path: []string{"sort_keys", "items"}, required: []string{"column"}},
		},
		"upsert_rows": {
			{path: []string{"rows", "items"}, required: []string{"match", "values"}},
		},
	}

	toolSchemas := make(map[string]*jsonschema.Schema, len(res.Tools))
	for _, tool := range res.Tools {
		data, err := json.Marshal(tool.InputSchema)
		if err != nil {
			t.Fatalf("marshal schema for %s: %v", tool.Name, err)
		}
		var schema jsonschema.Schema
		if err := json.Unmarshal(data, &schema); err != nil {
			t.Fatalf("unmarshal schema for %s: %v", tool.Name, err)
		}
		toolSchemas[tool.Name] = &schema
	}

	for toolName, required := range expectedRequired {
		schema := toolSchemas[toolName]
		if schema == nil {
			t.Fatalf("missing tool schema for %s", toolName)
		}
		assertRequiredFields(t, toolName, schema.Required, required)
	}

	for toolName, requirements := range expectedNested {
		schema := toolSchemas[toolName]
		if schema == nil {
			t.Fatalf("missing tool schema for %s", toolName)
		}
		for _, requirement := range requirements {
			nested := schema
			for _, segment := range requirement.path {
				switch segment {
				case "items":
					nested = nested.Items
				default:
					nested = nested.Properties[segment]
				}
				if nested == nil {
					t.Fatalf("missing nested schema for %s at %s", toolName, strings.Join(requirement.path, "."))
				}
			}
			assertRequiredFields(t, toolName+":"+strings.Join(requirement.path, "."), nested.Required, requirement.required)
		}
	}

	setColumnWidths := toolSchemas["set_column_widths"]
	if setColumnWidths == nil {
		t.Fatal("missing tool schema for set_column_widths")
	}
	if len(setColumnWidths.AnyOf) != 2 {
		t.Fatalf("expected set_column_widths anyOf alternatives, got %+v", setColumnWidths.AnyOf)
	}
	assertRequiredFields(t, "set_column_widths:anyOf[0]", setColumnWidths.AnyOf[0].Required, []string{"widths"})
	assertRequiredFields(t, "set_column_widths:anyOf[1]", setColumnWidths.AnyOf[1].Required, []string{"auto_fit"})
	if setColumnWidths.AnyOf[1].Properties["auto_fit"] == nil || setColumnWidths.AnyOf[1].Properties["auto_fit"].Const == nil {
		t.Fatalf("expected set_column_widths auto_fit=true constraint, got %+v", setColumnWidths.AnyOf[1].Properties["auto_fit"])
	}
	if value, ok := (*setColumnWidths.AnyOf[1].Properties["auto_fit"].Const).(bool); !ok || !value {
		t.Fatalf("expected set_column_widths auto_fit const true, got %+v", setColumnWidths.AnyOf[1].Properties["auto_fit"].Const)
	}
}

func assertRequiredFields(t *testing.T, name string, got, want []string) {
	t.Helper()
	got = append([]string(nil), got...)
	want = append([]string(nil), want...)
	slices.Sort(got)
	slices.Sort(want)
	if !slices.Equal(got, want) {
		t.Fatalf("unexpected required fields for %s: got %v want %v", name, got, want)
	}
}

func newTestClient(t *testing.T) (context.Context, *mcp.ClientSession) {
	t.Helper()
	ctx := context.Background()
	clientTransport, serverTransport := mcp.NewInMemoryTransports()

	serverSession, err := New(Config{PathMode: PathModeDirect}).Connect(ctx, serverTransport, nil)
	if err != nil {
		t.Fatalf("connect server: %v", err)
	}
	t.Cleanup(func() { _ = serverSession.Close() })

	client := mcp.NewClient(&mcp.Implementation{Name: "test-client", Version: "0.1.0"}, nil)
	clientSession, err := client.Connect(ctx, clientTransport, nil)
	if err != nil {
		t.Fatalf("connect client: %v", err)
	}
	t.Cleanup(func() { _ = clientSession.Close() })

	return ctx, clientSession
}

func createWorkbookFixture(t *testing.T, path string, fill func(*excelize.File) error) {
	t.Helper()
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	if err := fill(f); err != nil {
		t.Fatalf("fill workbook: %v", err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save workbook: %v", err)
	}
}

func zipHasEntry(t *testing.T, workbook, name string) bool {
	t.Helper()
	reader, err := zip.OpenReader(workbook)
	if err != nil {
		t.Fatalf("open zip workbook: %v", err)
	}
	defer func() { _ = reader.Close() }()
	for _, file := range reader.File {
		if file.Name == name {
			return true
		}
	}
	return false
}

func zipEntryContents(t *testing.T, workbook, name string) string {
	t.Helper()
	reader, err := zip.OpenReader(workbook)
	if err != nil {
		t.Fatalf("open zip workbook: %v", err)
	}
	defer func() { _ = reader.Close() }()
	for _, file := range reader.File {
		if file.Name != name {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			t.Fatalf("open zip entry %s: %v", name, err)
		}
		defer func() { _ = rc.Close() }()
		data, err := io.ReadAll(rc)
		if err != nil {
			t.Fatalf("read zip entry %s: %v", name, err)
		}
		return string(data)
	}
	t.Fatalf("zip entry %s not found", name)
	return ""
}

func assertToolErrorContains(t *testing.T, ctx context.Context, session *mcp.ClientSession, name string, args map[string]any, want string) {
	t.Helper()
	result, err := session.CallTool(ctx, &mcp.CallToolParams{Name: name, Arguments: args})
	if err != nil {
		t.Fatalf("call %s: %v", name, err)
	}
	if !result.IsError {
		t.Fatalf("expected tool %s to fail", name)
	}
	if len(result.Content) == 0 {
		t.Fatalf("tool %s failed with empty content", name)
	}
	text := result.Content[0].(*mcp.TextContent).Text
	if !strings.Contains(text, want) {
		t.Fatalf("expected error containing %q, got %q", want, text)
	}
}

func callTool(t *testing.T, ctx context.Context, session *mcp.ClientSession, name string, args map[string]any) string {
	t.Helper()
	result, err := session.CallTool(ctx, &mcp.CallToolParams{Name: name, Arguments: args})
	if err != nil {
		t.Fatalf("call %s: %v", name, err)
	}
	if result.IsError {
		if len(result.Content) == 0 {
			t.Fatalf("tool %s failed with empty content", name)
		}
		t.Fatalf("tool %s failed: %s", name, result.Content[0].(*mcp.TextContent).Text)
	}
	if len(result.Content) == 0 {
		return ""
	}
	return result.Content[0].(*mcp.TextContent).Text
}

func TestSetColumnWidthsAndRowHeights(t *testing.T) {
	ctx, clientSession := newTestClient(t)
	workbook := filepath.Join(t.TempDir(), "widths.xlsx")

	createWorkbookFixture(t, workbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Name", "Description", "Value"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"alpha", "a longer description here", 1}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A3", &[]any{"beta", "short", 2}); err != nil {
			return err
		}
		return nil
	})

	// Explicit widths
	callTool(t, ctx, clientSession, "set_column_widths", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"widths": []map[string]any{
			{"column": "A", "width": 20},
			{"column": "B", "width": 35},
		},
	})

	f, err := excelize.OpenFile(workbook)
	if err != nil {
		t.Fatalf("open workbook: %v", err)
	}
	defer func() { _ = f.Close() }()

	widthA, err := f.GetColWidth("Sheet1", "A")
	if err != nil {
		t.Fatalf("get col width A: %v", err)
	}
	if widthA != 20 {
		t.Fatalf("expected column A width 20, got %v", widthA)
	}
	widthB, err := f.GetColWidth("Sheet1", "B")
	if err != nil {
		t.Fatalf("get col width B: %v", err)
	}
	if widthB != 35 {
		t.Fatalf("expected column B width 35, got %v", widthB)
	}
	_ = f.Close()

	// Auto-fit
	autoFitWorkbook := filepath.Join(t.TempDir(), "autofit.xlsx")
	createWorkbookFixture(t, autoFitWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"A very long header that exceeds default width"}); err != nil {
			return err
		}
		return nil
	})

	callTool(t, ctx, clientSession, "set_column_widths", map[string]any{
		"filepath":   autoFitWorkbook,
		"sheet_name": "Sheet1",
		"auto_fit":   true,
	})

	fAuto, err := excelize.OpenFile(autoFitWorkbook)
	if err != nil {
		t.Fatalf("open auto-fit workbook: %v", err)
	}
	defer func() { _ = fAuto.Close() }()

	autoWidth, err := fAuto.GetColWidth("Sheet1", "A")
	if err != nil {
		t.Fatalf("get auto-fit col width: %v", err)
	}
	// "A very long header that exceeds default width" is 46 chars, so width should be > 8
	if autoWidth <= 8 {
		t.Fatalf("expected auto-fit width > 8, got %v", autoWidth)
	}

	// Row heights
	heightWorkbook := filepath.Join(t.TempDir(), "heights.xlsx")
	createWorkbookFixture(t, heightWorkbook, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Header"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{"Row 2"}); err != nil {
			return err
		}
		return nil
	})

	callTool(t, ctx, clientSession, "set_row_heights", map[string]any{
		"filepath":   heightWorkbook,
		"sheet_name": "Sheet1",
		"heights": []map[string]any{
			{"row": 1, "height": 30},
			{"row": 2, "height": 20},
		},
	})

	fH, err := excelize.OpenFile(heightWorkbook)
	if err != nil {
		t.Fatalf("open height workbook: %v", err)
	}
	defer func() { _ = fH.Close() }()

	h1, err := fH.GetRowHeight("Sheet1", 1)
	if err != nil {
		t.Fatalf("get row 1 height: %v", err)
	}
	if h1 != 30 {
		t.Fatalf("expected row 1 height 30, got %v", h1)
	}
	h2, err := fH.GetRowHeight("Sheet1", 2)
	if err != nil {
		t.Fatalf("get row 2 height: %v", err)
	}
	if h2 != 20 {
		t.Fatalf("expected row 2 height 20, got %v", h2)
	}

	// Error cases
	assertToolErrorContains(t, ctx, clientSession, "set_column_widths", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
	}, "anyOf: did not validate against any of")

	assertToolErrorContains(t, ctx, clientSession, "set_column_widths", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
		"widths": []map[string]any{
			{"column": "123!", "width": 10},
		},
	}, "invalid column")

	assertToolErrorContains(t, ctx, clientSession, "set_column_widths", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Missing",
		"auto_fit":   true,
	}, "sheet Missing does not exist")

	assertToolErrorContains(t, ctx, clientSession, "set_row_heights", map[string]any{
		"filepath":   workbook,
		"sheet_name": "Sheet1",
	}, "heights")

	assertToolErrorContains(t, ctx, clientSession, "set_row_heights", map[string]any{
		"filepath":   heightWorkbook,
		"sheet_name": "Missing",
		"heights": []map[string]any{
			{"row": 1, "height": 30},
		},
	}, "sheet Missing does not exist")
}
