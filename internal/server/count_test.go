package server

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"testing"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"github.com/xuri/excelize/v2"
)

func TestRowAndColumnCounts(t *testing.T) {
	for _, protocol := range []string{"MCP", "CLI"} {
		t.Run(protocol, func(t *testing.T) {
			invoke := func(tool string, args map[string]any) error {
				payload, err := json.Marshal(args)
				if err != nil {
					return err
				}
				_, err = RunCLITool(t.Context(), Config{}, tool, payload)
				return err
			}
			if protocol == "MCP" {
				ctx, session := newTestClient(t)
				invoke = func(tool string, args map[string]any) error {
					result, err := session.CallTool(ctx, &mcp.CallToolParams{Name: tool, Arguments: args})
					if err != nil {
						return err
					}
					if result.IsError {
						return fmt.Errorf("tool failed: %s", cliResultText(result))
					}
					return nil
				}
			}
			for _, operation := range []struct {
				tool     string
				startKey string
				marker   string
				movedOne string
				movedTwo string
			}{
				{"insert_rows", "start_row", "A1", "A2", "A3"},
				{"insert_columns", "start_col", "A1", "B1", "C1"},
				{"delete_sheet_rows", "start_row", "A3", "A2", "A1"},
				{"delete_sheet_columns", "start_col", "C1", "B1", "A1"},
			} {
				t.Run(operation.tool, func(t *testing.T) {
					path := filepath.Join(t.TempDir(), "counts.xlsx")
					reset := func() {
						createWorkbookFixture(t, path, func(f *excelize.File) error {
							return f.SetCellValue("Sheet1", operation.marker, "marker")
						})
					}
					reset()
					args := map[string]any{"filepath": path, "sheet_name": "Sheet1", operation.startKey: 1}
					before, err := os.ReadFile(path)
					if err != nil {
						t.Fatal(err)
					}
					for _, count := range []any{0, -1, nil} {
						args["count"] = count
						if err := invoke(operation.tool, args); err == nil {
							t.Fatalf("expected count=%v to be rejected", count)
						}
						after, err := os.ReadFile(path)
						if err != nil {
							t.Fatal(err)
						}
						if !bytes.Equal(before, after) {
							t.Fatalf("count=%v modified the workbook", count)
						}
					}
					for _, count := range []int{1, 2} {
						reset()
						want := operation.movedOne
						delete(args, "count")
						if count == 2 {
							args["count"] = count
							want = operation.movedTwo
						}
						if err := invoke(operation.tool, args); err != nil {
							t.Fatal(err)
						}
						f, err := excelize.OpenFile(path)
						if err != nil {
							t.Fatal(err)
						}
						value, valueErr := f.GetCellValue("Sheet1", want)
						if err := f.Close(); err != nil {
							t.Fatal(err)
						}
						if valueErr != nil || value != "marker" {
							t.Fatalf("count=%d: expected marker at %s, got %q (%v)", count, want, value, valueErr)
						}
					}
				})
			}
		})
	}
}
