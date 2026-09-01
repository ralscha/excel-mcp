package excel

import (
	"fmt"
	"path/filepath"
	"sync"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestTranslateFormula(t *testing.T) {
	formula := `SUM(A1,$B2,C$3,$D$4,"A1",LOG10(A1),'Sales Data'!B2)`
	want := `SUM(C4,$B5,E$3,$D$4,"A1",LOG10(C4),'Sales Data'!D5)`
	if got := translateFormula(formula, 2, 3); got != want {
		t.Fatalf("unexpected translated formula: got %q want %q", got, want)
	}
}

func TestCopyRangeHandlesOverlapAndPreservesCellMetadata(t *testing.T) {
	path := filepath.Join(t.TempDir(), "copy.xlsx")
	f := excelize.NewFile()
	styleID, err := f.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
	if err != nil {
		t.Fatalf("create style: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", 1); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "A2", 2); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle("Sheet1", "A2", "A2", styleID); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellFormula("Sheet1", "A3", "A1+A2"); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save fixture: %v", err)
	}
	_ = f.Close()

	if _, err := CopyRange(path, "Sheet1", "A1", "A3", "A2", ""); err != nil {
		t.Fatalf("copy range: %v", err)
	}
	f, err = excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open result: %v", err)
	}
	defer func() { _ = f.Close() }()
	if got, _ := f.GetCellValue("Sheet1", "A2"); got != "1" {
		t.Fatalf("overlapping copy corrupted first value: %q", got)
	}
	if got, _ := f.GetCellValue("Sheet1", "A3"); got != "2" {
		t.Fatalf("overlapping copy corrupted second value: %q", got)
	}
	if got, _ := f.GetCellFormula("Sheet1", "A4"); got != "A2+A3" {
		t.Fatalf("formula was not copied and translated: %q", got)
	}
	if got, _ := f.GetCellStyle("Sheet1", "A3"); got != styleID {
		t.Fatalf("style was not copied: got %d want %d", got, styleID)
	}
}

func TestSortRangePreservesAndTranslatesFormulasAndStyles(t *testing.T) {
	path := filepath.Join(t.TempDir(), "sort.xlsx")
	f := excelize.NewFile()
	styleID, err := f.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"#FFFF00"}}})
	if err != nil {
		t.Fatalf("create style: %v", err)
	}
	for cell, value := range map[string]any{"A1": "Name", "B1": "Label", "A2": "Beta", "A3": "Alpha"} {
		if err := f.SetCellValue("Sheet1", cell, value); err != nil {
			t.Fatal(err)
		}
	}
	if err := f.SetCellFormula("Sheet1", "B2", `A2&"!"`); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellFormula("Sheet1", "B3", `A3&"!"`); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle("Sheet1", "A3", "B3", styleID); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save fixture: %v", err)
	}
	_ = f.Close()

	if _, err := SortRange(path, "Sheet1", "A1:B3", SortRangeOptions{HasHeader: true, SortKeys: []SortKey{{Column: "Name"}}}); err != nil {
		t.Fatalf("sort range: %v", err)
	}
	f, err = excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open result: %v", err)
	}
	defer func() { _ = f.Close() }()
	if got, _ := f.GetCellFormula("Sheet1", "B2"); got != `A2&"!"` {
		t.Fatalf("formula was not translated with sorted row: %q", got)
	}
	if got, _ := f.GetCellStyle("Sheet1", "A2"); got != styleID {
		t.Fatalf("style was not moved with sorted row: got %d want %d", got, styleID)
	}
}

func TestDeleteRangeBeyondUsedCells(t *testing.T) {
	path := filepath.Join(t.TempDir(), "delete.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatalf("create workbook: %v", err)
	}
	if err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetCellValue("Sheet1", "A1", "keep")
	}); err != nil {
		t.Fatalf("fill workbook: %v", err)
	}
	if _, err := DeleteRange(path, "Sheet1", "B2", "Z100", "left"); err != nil {
		t.Fatalf("delete range outside used cells: %v", err)
	}
}

func TestCreateWorkbookRequiresExplicitOverwrite(t *testing.T) {
	path := filepath.Join(t.TempDir(), "existing.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatalf("create workbook: %v", err)
	}
	if err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetCellValue("Sheet1", "A1", "keep")
	}); err != nil {
		t.Fatalf("write marker: %v", err)
	}
	if _, err := CreateWorkbook(path, false); err == nil {
		t.Fatal("expected existing workbook to be protected from overwrite")
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open protected workbook: %v", err)
	}
	if got, _ := f.GetCellValue("Sheet1", "A1"); got != "keep" {
		t.Fatalf("workbook was unexpectedly replaced: marker=%q", got)
	}
	_ = f.Close()
	if _, err := CreateWorkbook(path, true); err != nil {
		t.Fatalf("explicit overwrite: %v", err)
	}
}

func TestReadDataReturnsJSONTypesAndRejectsDuplicateHeaders(t *testing.T) {
	path := filepath.Join(t.TempDir(), "typed.xlsx")
	f := excelize.NewFile()
	if err := f.SetSheetRow("Sheet1", "A1", &[]any{"Count", "Enabled", "Name"}); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSheetRow("Sheet1", "A2", &[]any{42, true, "alpha"}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save fixture: %v", err)
	}
	_ = f.Close()
	result, err := ReadData(path, "Sheet1", "A1", "C2", false)
	if err != nil {
		t.Fatalf("read typed data: %v", err)
	}
	if len(result.Rows) != 1 || result.Rows[0]["Count"] != float64(42) || result.Rows[0]["Enabled"] != true || result.Rows[0]["Name"] != "alpha" {
		t.Fatalf("unexpected typed row: %+v", result.Rows)
	}
	if err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetCellValue("Sheet1", "B1", "Count")
	}); err != nil {
		t.Fatalf("write duplicate header: %v", err)
	}
	if _, err := ReadData(path, "Sheet1", "A1", "C2", false); err == nil {
		t.Fatal("expected duplicate headers to be rejected")
	}
}

func TestApplyFormulaStoresValidExcelFormula(t *testing.T) {
	path := filepath.Join(t.TempDir(), "formula.xlsx")
	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "A1", 1); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "A2", 2); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save fixture: %v", err)
	}
	_ = f.Close()
	if _, err := ApplyFormula(path, "Sheet1", "A3", "=SUM(A1:A2)"); err != nil {
		t.Fatalf("apply formula: %v", err)
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open result: %v", err)
	}
	defer func() { _ = f.Close() }()
	if got, _ := f.GetCellFormula("Sheet1", "A3"); got != "SUM(A1:A2)" {
		t.Fatalf("formula contains an invalid leading equals sign: %q", got)
	}
	if got, err := f.CalcCellValue("Sheet1", "A3"); err != nil || got != "3" {
		t.Fatalf("formula does not calculate: value=%q error=%v", got, err)
	}
}

func TestChartRangesSupportSpecialSheetNames(t *testing.T) {
	path := filepath.Join(t.TempDir(), "chart.xlsx")
	f := excelize.NewFile()
	const sheetName = "R&D"
	if _, err := f.NewSheet(sheetName); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSheetRow(sheetName, "A1", &[]any{"Series", "Q1", "Q2"}); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSheetRow(sheetName, "A2", &[]any{"Revenue", 10, 20}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save fixture: %v", err)
	}
	_ = f.Close()
	if _, err := CreateChart(path, sheetName, "A1:C2", "line", "E2", ChartOptions{Title: "Revenue"}); err != nil {
		t.Fatalf("create chart: %v", err)
	}
	charts, err := ListCharts(path, sheetName, sheetName)
	if err != nil {
		t.Fatalf("list charts: %v", err)
	}
	if len(charts.Charts) != 1 || len(charts.Charts[0].Series) != 1 || charts.Charts[0].Series[0].SourceSheet != sheetName {
		t.Fatalf("chart source sheet was not preserved: %+v", charts.Charts)
	}
}

func TestConcurrentWorkbookMutations(t *testing.T) {
	path := filepath.Join(t.TempDir(), "concurrent.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatalf("create workbook: %v", err)
	}
	const workers = 12
	var wg sync.WaitGroup
	errs := make(chan error, workers)
	for index := range workers {
		wg.Add(1)
		go func() {
			defer wg.Done()
			_, err := CreateWorksheet(path, fmt.Sprintf("Sheet-%d", index))
			errs <- err
		}()
	}
	wg.Wait()
	close(errs)
	for err := range errs {
		if err != nil {
			t.Fatalf("concurrent mutation failed: %v", err)
		}
	}
	metadata, err := GetWorkbookMetadata(path, false)
	if err != nil {
		t.Fatalf("read workbook: %v", err)
	}
	if len(metadata.Sheets) != workers+1 {
		t.Fatalf("lost a concurrent update: got %d sheets want %d", len(metadata.Sheets), workers+1)
	}
}

func TestDeleteWorksheetReportsNoOpCases(t *testing.T) {
	path := filepath.Join(t.TempDir(), "sheets.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatalf("create workbook: %v", err)
	}
	if _, err := DeleteWorksheet(path, "Missing"); err == nil {
		t.Fatal("expected deleting a missing sheet to fail")
	}
	if _, err := DeleteWorksheet(path, "Sheet1"); err == nil {
		t.Fatal("expected deleting the only sheet to fail")
	}
}

func TestGetSheetSchemaSupportsEmptySheets(t *testing.T) {
	path := filepath.Join(t.TempDir(), "empty.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatalf("create workbook: %v", err)
	}
	schema, err := GetSheetSchema(path, "Sheet1", "A1", "", SheetSchemaOptions{})
	if err != nil {
		t.Fatalf("infer empty schema: %v", err)
	}
	if schema.Range != "A1:A1" || schema.RowCount != 0 || len(schema.Columns) != 1 || schema.Columns[0].Name != "column_1" {
		t.Fatalf("unexpected empty schema: %+v", schema)
	}
}
