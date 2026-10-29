package server

import (
	"context"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCLISchemaValidationPreventsMissingHeightMutation(t *testing.T) {
	path := filepath.Join(t.TempDir(), "heights.xlsx")
	f := excelize.NewFile()
	if err := f.SetRowHeight("Sheet1", 1, 42); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}

	for _, entry := range []map[string]any{{"row": 1}, {"row": 1, "height": nil}} {
		payload, err := json.Marshal(map[string]any{
			"filepath": path, "sheet_name": "Sheet1", "heights": []map[string]any{entry},
		})
		if err != nil {
			t.Fatal(err)
		}
		if _, err := RunCLITool(t.Context(), Config{}, "set_row_heights", payload); err == nil || !strings.Contains(err.Error(), "validate tool input") {
			t.Fatalf("expected schema validation error for %s, got %v", payload, err)
		}
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()
	height, err := f.GetRowHeight("Sheet1", 1)
	if err != nil || height != 42 {
		t.Fatalf("invalid input changed row height: height=%v err=%v", height, err)
	}
}

func TestCLISchemaValidationMatchesToolSchema(t *testing.T) {
	for _, tt := range []struct {
		name    string
		tool    string
		payload string
	}{
		{"null input", "create_workbook", `null`},
		{"missing required field", "create_workbook", `{}`},
		{"null boolean", "create_workbook", `{"filepath":"book.xlsx","overwrite":null}`},
		{"column width alternatives", "set_column_widths", `{"filepath":"book.xlsx","sheet_name":"Sheet1","auto_fit":false}`},
	} {
		t.Run(tt.name, func(t *testing.T) {
			_, err := RunCLITool(t.Context(), Config{}, tt.tool, []byte(tt.payload))
			if err == nil || !strings.Contains(err.Error(), "validate tool input") {
				t.Fatalf("expected schema validation failure, got %v", err)
			}
		})
	}
}

func TestCLICancelledContextDoesNotCreateWorkbook(t *testing.T) {
	path := filepath.Join(t.TempDir(), "cancelled.xlsx")
	payload, err := json.Marshal(map[string]string{"filepath": path})
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithCancel(t.Context())
	cancel()
	if _, err := RunCLITool(ctx, Config{}, "create_workbook", payload); !errors.Is(err, context.Canceled) {
		t.Fatalf("expected cancellation, got %v", err)
	}
	if _, err := os.Stat(path); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("cancelled tool created workbook: %v", err)
	}
}
