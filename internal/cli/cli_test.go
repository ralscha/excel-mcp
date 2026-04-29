package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"io"
	"log/slog"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

func TestRunListTools(t *testing.T) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer

	exitCode := Run(context.Background(), []string{"list-tools"}, &stdout, &stderr, testLogger())
	if exitCode != 0 {
		t.Fatalf("expected exit code 0, got %d: %s", exitCode, stderr.String())
	}
	if !strings.Contains(stdout.String(), "create_workbook") {
		t.Fatalf("expected list-tools output to include create_workbook, got %q", stdout.String())
	}
	if stderr.Len() != 0 {
		t.Fatalf("expected no stderr output, got %q", stderr.String())
	}
}

func TestRunListToolsJSON(t *testing.T) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer

	exitCode := Run(context.Background(), []string{"list-tools", "--json"}, &stdout, &stderr, testLogger())
	if exitCode != 0 {
		t.Fatalf("expected exit code 0, got %d: %s", exitCode, stderr.String())
	}
	var tools []struct {
		Name          string          `json:"name"`
		Description   string          `json:"description"`
		InputSchema   json.RawMessage `json:"input_schema"`
		OutputExample json.RawMessage `json:"output_example"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &tools); err != nil {
		t.Fatalf("expected valid json output: %v\n%s", err, stdout.String())
	}
	if len(tools) == 0 {
		t.Fatal("expected at least one tool")
	}
	if tools[0].Name == "" || len(tools[0].InputSchema) == 0 || len(tools[0].OutputExample) == 0 {
		t.Fatalf("expected tool metadata and schema, got %+v", tools[0])
	}
	if stderr.Len() != 0 {
		t.Fatalf("expected no stderr output, got %q", stderr.String())
	}
}

func TestRunToolInfoJSON(t *testing.T) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer

	exitCode := Run(context.Background(), []string{"tool-info", "--json", "create_workbook"}, &stdout, &stderr, testLogger())
	if exitCode != 0 {
		t.Fatalf("expected exit code 0, got %d: %s", exitCode, stderr.String())
	}
	var tool struct {
		Name        string `json:"name"`
		Description string `json:"description"`
		InputSchema struct {
			Required []string `json:"required"`
		} `json:"input_schema"`
		OutputExample string `json:"output_example"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &tool); err != nil {
		t.Fatalf("expected valid json output: %v\n%s", err, stdout.String())
	}
	if tool.Name != "create_workbook" {
		t.Fatalf("expected create_workbook info, got %+v", tool)
	}
	if len(tool.InputSchema.Required) == 0 || tool.InputSchema.Required[0] != "filepath" {
		t.Fatalf("expected filepath requirement in schema, got %+v", tool.InputSchema)
	}
	if tool.OutputExample == "" {
		t.Fatalf("expected output example, got %+v", tool)
	}
}

func TestRunToolInfoText(t *testing.T) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer

	exitCode := Run(context.Background(), []string{"tool-info", "create_workbook"}, &stdout, &stderr, testLogger())
	if exitCode != 0 {
		t.Fatalf("expected exit code 0, got %d: %s", exitCode, stderr.String())
	}
	output := stdout.String()
	if !strings.Contains(output, "name: create_workbook") || !strings.Contains(output, "input_schema:") || !strings.Contains(output, "output_example:") {
		t.Fatalf("expected text tool info output, got %q", output)
	}
}

func TestRunCreateWorkbookDirect(t *testing.T) {
	workbookPath := filepath.Join(t.TempDir(), "direct.xlsx")
	inputPath := writeInputFile(t, `{"filepath":`+strconv.Quote(workbookPath)+`}`)

	var stdout bytes.Buffer
	var stderr bytes.Buffer
	exitCode := Run(context.Background(), []string{"create_workbook", "--input", inputPath}, &stdout, &stderr, testLogger())
	if exitCode != 0 {
		t.Fatalf("expected exit code 0, got %d: %s", exitCode, stderr.String())
	}
	if _, err := os.Stat(workbookPath); err != nil {
		t.Fatalf("expected workbook to exist: %v", err)
	}
	if !strings.Contains(strings.ToLower(stdout.String()), "workbook") {
		t.Fatalf("expected success message, got %q", stdout.String())
	}
}

func TestRunCreateWorkbookRooted(t *testing.T) {
	root := t.TempDir()
	t.Setenv("EXCEL_FILES_PATH", root)
	inputPath := writeInputFile(t, `{"filepath":"nested/rooted.xlsx"}`)

	var stdout bytes.Buffer
	var stderr bytes.Buffer
	exitCode := Run(context.Background(), []string{"create_workbook", "--rooted", "--input", inputPath}, &stdout, &stderr, testLogger())
	if exitCode != 0 {
		t.Fatalf("expected exit code 0, got %d: %s", exitCode, stderr.String())
	}
	workbookPath := filepath.Join(root, "nested", "rooted.xlsx")
	if _, err := os.Stat(workbookPath); err != nil {
		t.Fatalf("expected rooted workbook to exist: %v", err)
	}
}

func TestRunUnknownTool(t *testing.T) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer

	exitCode := Run(context.Background(), []string{"missing_tool"}, &stdout, &stderr, testLogger())
	if exitCode != 1 {
		t.Fatalf("expected exit code 1, got %d", exitCode)
	}
	if !strings.Contains(stderr.String(), `unknown tool "missing_tool"`) {
		t.Fatalf("expected unknown tool error, got %q", stderr.String())
	}
}

func TestRunToolInfoUnknownTool(t *testing.T) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer

	exitCode := Run(context.Background(), []string{"tool-info", "missing_tool"}, &stdout, &stderr, testLogger())
	if exitCode != 1 {
		t.Fatalf("expected exit code 1, got %d", exitCode)
	}
	if !strings.Contains(stderr.String(), `unknown tool "missing_tool"`) {
		t.Fatalf("expected unknown tool error, got %q", stderr.String())
	}
}

func writeInputFile(t *testing.T, contents string) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "input.json")
	if err := os.WriteFile(path, []byte(contents), 0o600); err != nil {
		t.Fatalf("write input file: %v", err)
	}
	return path
}

func testLogger() *slog.Logger {
	return slog.New(slog.NewTextHandler(io.Discard, nil))
}
