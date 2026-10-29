package excel

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestTranslateFormulaWholeReferencesAndEscapedText(t *testing.T) {
	tests := []struct {
		name, formula, want string
		cols, rows          int
	}{
		{"columns", `SUM(A:A,$B:C,D:$E,$F:$G)`, `SUM(C:C,$B:E,F:$E,$F:$G)`, 2, 3},
		{"rows", `SUM(1:1,$2:3,4:$5,$6:$7)`, `SUM(4:4,$2:6,7:$5,$6:$7)`, 2, 3},
		{"qualified", `SUM('Sales Data'!A:B,'O''Brien'!1:2)`, `SUM('Sales Data'!C:D,'O''Brien'!4:5)`, 2, 3},
		{"escaped text", `IF(A1="say ""Hi""","A:A","1:1")`, `IF(B2="say ""Hi""","A:A","1:1")`, 1, 1},
		{"column underflow", `SUM(A:A)`, `SUM(#REF!)`, -1, 0},
		{"column overflow", `SUM(XFD:XFD)`, `SUM(#REF!)`, 1, 0},
		{"row underflow", `SUM(1:1)`, `SUM(#REF!)`, 0, -1},
		{"row overflow", `SUM(1048576:1048576)`, `SUM(#REF!)`, 0, 1},
		{"absolute bounds", `SUM($A:$A,$1:$1)`, `SUM($A:$A,$1:$1)`, -1, -1},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := translateFormula(test.formula, test.cols, test.rows); got != test.want {
				t.Fatalf("translated formula = %q, want %q", got, test.want)
			}
		})
	}
}

func TestCopyFormulaClearsDestinationCachedValue(t *testing.T) {
	path := filepath.Join(t.TempDir(), "copy-formula.xlsx")
	f := excelize.NewFile()
	if err := f.SetCellFormula("Sheet1", "B1", "SUM(A:A)"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "C1", 999); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	_ = f.Close()
	if _, err := CopyRange(path, "Sheet1", "B1", "B1", "C1", ""); err != nil {
		t.Fatal(err)
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = f.Close() }()
	if got, err := f.GetCellFormula("Sheet1", "C1"); err != nil || got != "SUM(B:B)" {
		t.Fatalf("copied formula = %q, error = %v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "C1"); err != nil || got != "" {
		t.Fatalf("formula retained stale destination value: %q, error = %v", got, err)
	}
}

func TestHiddenNumericValueAndWorksheetBounds(t *testing.T) {
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	format := ";;;"
	style, err := f.NewStyle(&excelize.Style{CustomNumFmt: &format})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "C5", 42); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellStyle("Sheet1", "C5", "C5", style); err != nil {
		t.Fatal(err)
	}
	value, display, err := readCellJSONValue(f, "Sheet1", "C5")
	if err != nil || value != float64(42) || display != "" {
		t.Fatalf("hidden value = %#v, display = %q, error = %v", value, display, err)
	}
	rows, cols, rangeRef, err := worksheetBounds(f, "Sheet1")
	if err != nil || rows != 5 || cols != 3 || rangeRef != "A1:C5" {
		t.Fatalf("hidden-value bounds = %d x %d (%s), error = %v", rows, cols, rangeRef, err)
	}
	if err := f.SetCellFormula("Sheet1", "E8", "1+1"); err != nil {
		t.Fatal(err)
	}
	rows, cols, rangeRef, err = worksheetBounds(f, "Sheet1")
	if err != nil || rows != 8 || cols != 5 || rangeRef != "A1:E8" {
		t.Fatalf("formula bounds = %d x %d (%s), error = %v", rows, cols, rangeRef, err)
	}
}

func TestInferCellKindDoesNotTreatNumbersAsBooleans(t *testing.T) {
	for _, value := range []string{"0", "1", "0.0", "1.0"} {
		if got := inferCellKind(value); got != "number" {
			t.Errorf("kind(%q) = %s, want number", value, got)
		}
	}
	for _, value := range []string{"TRUE", "false", "True"} {
		if got := inferCellKind(value); got != "boolean" {
			t.Errorf("kind(%q) = %s, want boolean", value, got)
		}
	}
}

func TestReadDataDefaultBoundsExcludeMergedExtents(t *testing.T) {
	path := filepath.Join(t.TempDir(), "merged-title.xlsx")
	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "Title"); err != nil {
		t.Fatal(err)
	}
	if err := f.MergeCell("Sheet1", "A1", "C1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	_ = f.Close()
	result, err := ReadData(path, "Sheet1", "A1", "", false)
	if err != nil {
		t.Fatalf("read merged title: %v", err)
	}
	if result.Range != "A1:A1" || len(result.Rows) != 0 {
		t.Fatalf("read bounds expanded to merged cells: %+v", result)
	}
	metadata, err := GetWorkbookMetadata(path, true)
	if err != nil {
		t.Fatal(err)
	}
	if len(metadata.Sheets) != 1 || metadata.Sheets[0].Range != "A1:C1" {
		t.Fatalf("metadata lost merged extents: %+v", metadata)
	}
}
