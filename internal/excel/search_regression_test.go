package excel

import (
	"math"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestFindLimitsAndSheetSelection(t *testing.T) {
	path := filepath.Join(t.TempDir(), "find.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatal(err)
	}
	if err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetSheetRow("Sheet1", "A1", &[]any{"match", "match"})
	}); err != nil {
		t.Fatal(err)
	}
	for _, limit := range []int{1, 2, 3} {
		result, err := FindInWorkbook(path, FindOptions{Query: "MATCH", MaxResults: limit, Sheets: []string{"sheet1", "Sheet1"}})
		if err != nil {
			t.Fatal(err)
		}
		if len(result.Matches) != min(limit, 2) || result.Truncated != (limit < 2) {
			t.Fatalf("limit %d: unexpected result: %+v", limit, result)
		}
		if result.Matches[0].SheetName != "Sheet1" {
			t.Fatalf("expected canonical sheet name, got %q", result.Matches[0].SheetName)
		}
	}
	if _, err := FindInWorkbook(path, FindOptions{Query: "match", MaxResults: 1, Sheets: []string{"Sheet1", "Missing"}}); err == nil {
		t.Fatal("expected invalid sheet to be rejected even after a result limit")
	}
	result, err := FindInWorkbook(path, FindOptions{Query: "absent"})
	if err != nil {
		t.Fatal(err)
	}
	if result.Matches == nil || len(result.Matches) != 0 || result.Truncated {
		t.Fatalf("expected empty non-null matches: %+v", result)
	}
}

func TestFindLargeContextDoesNotOverflow(t *testing.T) {
	path := filepath.Join(t.TempDir(), "context.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatal(err)
	}
	if err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetSheetRow("Sheet1", "A2", &[]any{"neighbor", "target"})
	}); err != nil {
		t.Fatal(err)
	}
	result, err := FindInWorkbook(path, FindOptions{Query: "target", ContextRows: math.MaxInt, ContextCols: math.MaxInt})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Matches) != 1 || len(result.Matches[0].Context) != 1 || result.Matches[0].Context[0].Cell != "A2" {
		t.Fatalf("expected adjacent cell in context: %+v", result)
	}
}

func TestFindSparseCellsAndUncachedFormula(t *testing.T) {
	path := filepath.Join(t.TempDir(), "sparse.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatal(err)
	}
	if err := withWorkbook(path, func(f *excelize.File) error {
		if err := f.SetCellValue("Sheet1", "XFD100", "far away"); err != nil {
			return err
		}
		return f.SetCellFormula("Sheet1", "C200", "SUM(A1:A2)")
	}); err != nil {
		t.Fatal(err)
	}
	for _, tt := range []struct{ searchType, query, cell string }{
		{"text", "far away", "XFD100"},
		{"formula", "SUM", "C200"},
	} {
		result, err := FindInWorkbook(path, FindOptions{Query: tt.query, SearchType: tt.searchType})
		if err != nil {
			t.Fatal(err)
		}
		if len(result.Matches) != 1 || result.Matches[0].Cell != tt.cell {
			t.Fatalf("%s search lost sparse cell: %+v", tt.searchType, result)
		}
	}
}

func TestFindMergedFormulaReportsOnlyAnchor(t *testing.T) {
	for _, trailingCell := range []bool{false, true} {
		name := "without trailing cell"
		if trailingCell {
			name = "with trailing cell"
		}
		t.Run(name, func(t *testing.T) {
			path := tableRegressionWorkbook(t, func(f *excelize.File) error {
				if err := f.SetCellFormula("Sheet1", "A1", "SUM(A3:A4)"); err != nil {
					return err
				}
				if err := f.MergeCell("Sheet1", "A1", "B2"); err != nil {
					return err
				}
				if trailingCell {
					if err := f.SetCellValue("Sheet1", "C1", "unrelated"); err != nil {
						return err
					}
					return f.SetCellValue("Sheet1", "C2", "unrelated")
				}
				return nil
			})
			for _, limit := range []int{1, 10} {
				result, err := FindInWorkbook(path, FindOptions{Query: "SUM", SearchType: "formula", MaxResults: limit})
				if err != nil {
					t.Fatal(err)
				}
				if len(result.Matches) != 1 || result.Matches[0].Cell != "A1" || result.Truncated {
					t.Fatalf("limit %d: merged formula produced duplicate matches or false truncation: %+v", limit, result)
				}
			}
		})
	}
}

func BenchmarkFindInSparseWorkbook(b *testing.B) {
	path := filepath.Join(b.TempDir(), "sparse.xlsx")
	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "XFD1000", "target"); err != nil {
		b.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		b.Fatal(err)
	}
	if err := f.Close(); err != nil {
		b.Fatal(err)
	}
	b.ReportAllocs()
	b.ResetTimer()
	for b.Loop() {
		result, err := FindInWorkbook(path, FindOptions{Query: "target", MaxResults: 1})
		if err != nil || len(result.Matches) != 1 {
			b.Fatalf("unexpected search result: %+v, %v", result, err)
		}
	}
}
