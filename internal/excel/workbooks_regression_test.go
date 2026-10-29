package excel

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestWorksheetNameCollisionsPreserveWorkbook(t *testing.T) {
	path := filepath.Join(t.TempDir(), "sheets.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatal(err)
	}
	if _, err := CreateWorksheet(path, "Keep"); err != nil {
		t.Fatal(err)
	}
	if err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetCellValue("Keep", "A1", "important data")
	}); err != nil {
		t.Fatal(err)
	}
	before, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	for _, action := range []struct {
		name string
		run  func() (string, error)
	}{
		{"create", func() (string, error) { return CreateWorksheet(path, "keep") }},
		{"copy", func() (string, error) { return CopyWorksheet(path, "Sheet1", "KEEP") }},
		{"copy self", func() (string, error) { return CopyWorksheet(path, "Keep", "keep") }},
		{"rename", func() (string, error) { return RenameWorksheet(path, "Sheet1", "Keep") }},
	} {
		t.Run(action.name, func(t *testing.T) {
			if _, err := action.run(); err == nil || !strings.Contains(err.Error(), "already exists") {
				t.Fatalf("expected name collision, got %v", err)
			}
			after, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(before, after) {
				t.Fatal("rejected operation modified workbook")
			}
		})
	}
}

func TestRenameWorksheetResolvesSourceCase(t *testing.T) {
	path := filepath.Join(t.TempDir(), "rename.xlsx")
	if _, err := CreateWorkbook(path, false); err != nil {
		t.Fatal(err)
	}
	if _, err := RenameWorksheet(path, "sheet1", "Renamed"); err != nil {
		t.Fatal(err)
	}
	if _, err := RenameWorksheet(path, "RENAMED", "renamed"); err != nil {
		t.Fatal(err)
	}
	metadata, err := GetWorkbookMetadata(path, false)
	if err != nil {
		t.Fatal(err)
	}
	if len(metadata.Sheets) != 1 || metadata.Sheets[0].Name != "renamed" {
		t.Fatalf("unexpected sheets: %+v", metadata.Sheets)
	}
}
