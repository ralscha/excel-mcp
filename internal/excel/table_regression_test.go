package excel

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func tableRegressionWorkbook(t *testing.T, setup func(*excelize.File) error) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "table.xlsx")
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	if err := setup(f); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestSortComparisonIsTransitive(t *testing.T) {
	values := []string{"", "2", "10", "1a", "1", "true", "false", "TRUE", "T", "f", "0", "Alpha", "alpha"}
	for _, a := range values {
		for _, b := range values {
			for _, c := range values {
				if compareSortableValues(a, b) < 0 && compareSortableValues(b, c) < 0 && compareSortableValues(a, c) >= 0 {
					t.Fatalf("inconsistent sort order: %q < %q < %q but first is not less than last", a, b, c)
				}
				if compareSortableValues(a, b) == 0 && compareSortableValues(b, c) == 0 && compareSortableValues(a, c) != 0 {
					t.Fatalf("inconsistent sort equality for %q, %q, %q", a, b, c)
				}
			}
		}
	}
}

func TestUpsertMatchesFormattedNumericKeys(t *testing.T) {
	path := tableRegressionWorkbook(t, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"ID", "Value"}); err != nil {
			return err
		}
		if err := f.SetSheetRow("Sheet1", "A2", &[]any{1000000, "old"}); err != nil {
			return err
		}
		style, err := f.NewStyle(&excelize.Style{NumFmt: 3})
		if err != nil {
			return err
		}
		return f.SetCellStyle("Sheet1", "A2", "A2", style)
	})
	result, err := UpsertRows(path, "Sheet1", "A1:B3", UpsertRowsOptions{
		KeyColumns: []string{"ID"}, InsertIfMissing: true,
		Rows: []UpsertRow{{Match: map[string]any{"ID": float64(1000000)}, Values: map[string]any{"Value": "new"}}},
	})
	if err != nil {
		t.Fatal(err)
	}
	if result.UpdatedCount != 1 || result.InsertedCount != 0 || result.Results[0].RowNumber != 2 {
		t.Fatalf("numeric key did not match existing formatted cell: %+v", result)
	}
}

func TestUpsertCompositeKeysCannotCollide(t *testing.T) {
	keys := []string{"First", "Second"}
	first := canonicalUpsertKey(map[string]any{"First": "a\x1fb", "Second": "c"}, keys, true)
	second := canonicalUpsertKey(map[string]any{"First": "a", "Second": "b\x1fc"}, keys, true)
	if first == second {
		t.Fatal("distinct composite keys encoded to the same index entry")
	}
}

func TestUpsertDoesNotOverwriteUncachedFormulaRows(t *testing.T) {
	path := tableRegressionWorkbook(t, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"ID", "Value"}); err != nil {
			return err
		}
		return f.SetCellFormula("Sheet1", "B2", `IF(TRUE,"","")`)
	})
	result, err := UpsertRows(path, "Sheet1", "A1:B3", UpsertRowsOptions{
		KeyColumns: []string{"ID"}, InsertIfMissing: true,
		Rows: []UpsertRow{{Match: map[string]any{"ID": "new"}, Values: map[string]any{"Value": "inserted"}}},
	})
	if err != nil {
		t.Fatal(err)
	}
	if result.InsertedCount != 1 || result.Results[0].RowNumber != 3 {
		t.Fatalf("insert would overwrite formula row: %+v", result)
	}
}

func TestUpsertFailureLeavesWorkbookUnchanged(t *testing.T) {
	path := tableRegressionWorkbook(t, func(f *excelize.File) error {
		if err := f.SetSheetRow("Sheet1", "A1", &[]any{"ID", "Value"}); err != nil {
			return err
		}
		return f.SetSheetRow("Sheet1", "A2", &[]any{"existing", "old"})
	})
	before, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	_, err = UpsertRows(path, "Sheet1", "A1:B2", UpsertRowsOptions{
		KeyColumns: []string{"ID"}, InsertIfMissing: true,
		Rows: []UpsertRow{
			{Match: map[string]any{"ID": "existing"}, Values: map[string]any{"Value": "changed"}},
			{Match: map[string]any{"ID": "new"}, Values: map[string]any{"Value": "inserted"}},
		},
	})
	if err == nil {
		t.Fatal("expected exhausted capacity to reject the batch")
	}
	after, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(before) != string(after) {
		t.Fatal("failed upsert persisted a partial update")
	}
}

func TestFormulaValidationHandlesQuotedSheetNames(t *testing.T) {
	for _, formula := range []string{
		`='Sales (draft'!A1`, `='Sales "draft"'!A1`, `='Bob''s (sales'!A1+SUM(B1:B2)`,
		`=SUM(Table1[Cost (USD])`, `=SUM(Table1[Item '#])`,
	} {
		if err := ValidateFormulaSyntax(formula); err != nil {
			t.Errorf("rejected valid formula %q: %v", formula, err)
		}
	}
	if err := ValidateFormulaSyntax(`='Unterminated!A1`); err == nil {
		t.Fatal("expected unbalanced quoted sheet name to be rejected")
	}
	for _, formula := range []string{`=[foo(`, `=Table1[)`, `=Table1[Cost`, `=Table1]`} {
		if err := ValidateFormulaSyntax(formula); err == nil {
			t.Errorf("accepted unbalanced structured reference %q", formula)
		}
	}
}

func TestReadAndFilterRetainNumbersHiddenByFormatting(t *testing.T) {
	path := tableRegressionWorkbook(t, func(f *excelize.File) error {
		if err := f.SetCellValue("Sheet1", "A1", "Value"); err != nil {
			return err
		}
		if err := f.SetCellValue("Sheet1", "A2", 42); err != nil {
			return err
		}
		format := ";;;"
		style, err := f.NewStyle(&excelize.Style{CustomNumFmt: &format})
		if err != nil {
			return err
		}
		return f.SetCellStyle("Sheet1", "A2", "A2", style)
	})
	for _, endCell := range []string{"", "A2"} {
		result, err := ReadData(path, "Sheet1", "A1", endCell, false)
		if err != nil {
			t.Fatal(err)
		}
		if len(result.Rows) != 1 || result.Rows[0]["Value"] != float64(42) {
			t.Fatalf("hidden numeric row was lost with end_cell=%q: %+v", endCell, result.Rows)
		}
	}
	result, err := FilterRows(path, "Sheet1", "A1:A2", FilterRowsOptions{
		HasHeader: true, Filters: []Filter{{Column: "Value", Operator: "gt", Value: "40"}},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Rows) != 1 || result.Rows[0]["Value"] != float64(42) {
		t.Fatalf("hidden numeric row was lost when filtering: %+v", result.Rows)
	}
}

func TestApplyFormulaClearsPreviousCachedValue(t *testing.T) {
	path := tableRegressionWorkbook(t, func(f *excelize.File) error {
		return f.SetCellValue("Sheet1", "A1", 99)
	})
	if _, err := ApplyFormula(path, "Sheet1", "A1", "=1+2"); err != nil {
		t.Fatal(err)
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = f.Close() }()
	if value, err := f.GetCellValue("Sheet1", "A1"); err != nil || value == "99" {
		t.Fatalf("replaced formula retained stale cached value: %q, %v", value, err)
	}
	if value, err := f.CalcCellValue("Sheet1", "A1"); err != nil || value != "3" {
		t.Fatalf("new formula failed to calculate: %q, %v", value, err)
	}
}
