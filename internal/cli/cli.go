package cli

import (
	"context"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"log/slog"
	"os"
	"strings"

	"excel-mcp/internal/server"
)

func Run(ctx context.Context, args []string, stdout, stderr io.Writer, logger *slog.Logger) int {
	if len(args) == 0 {
		writeUsage(stderr)
		return 2
	}

	switch args[0] {
	case "-h", "--help", "help":
		writeUsage(stdout)
		return 0
	case "list-tools":
		return runListTools(args[1:], stdout, stderr)
	case "tool-info":
		return runToolInfo(args[1:], stdout, stderr)
	}

	toolName := args[0]
	fs := flag.NewFlagSet("excel-mcp cli", flag.ContinueOnError)
	fs.SetOutput(stderr)
	inputPath := fs.String("input", "", "Read tool JSON input from a file path or - for stdin")
	rooted := fs.Bool("rooted", false, "Resolve file paths relative to EXCEL_FILES_PATH")
	if err := fs.Parse(args[1:]); err != nil {
		return 2
	}
	if fs.NArg() != 0 {
		writeError(stderr, "unexpected extra arguments: %s\n", strings.Join(fs.Args(), " "))
		return 2
	}

	payload, err := readPayload(*inputPath)
	if err != nil {
		writeError(stderr, "%v\n", err)
		return 1
	}

	pathMode := server.PathModeDirect
	if *rooted {
		pathMode = server.PathModeRooted
	}

	output, err := server.RunCLITool(ctx, server.Config{PathMode: pathMode, Logger: logger}, toolName, payload)
	if err != nil {
		writeError(stderr, "%v\n", err)
		return 1
	}
	if strings.TrimSpace(output) != "" {
		if _, err := io.WriteString(stdout, ensureTrailingNewline(output)); err != nil {
			writeError(stderr, "write output: %v\n", err)
			return 1
		}
	}
	return 0
}

func runListTools(args []string, stdout, stderr io.Writer) int {
	fs := flag.NewFlagSet("excel-mcp cli list-tools", flag.ContinueOnError)
	fs.SetOutput(stderr)
	jsonOutput := fs.Bool("json", false, "Emit machine-readable JSON output")
	if err := fs.Parse(args); err != nil {
		return 2
	}
	if fs.NArg() != 0 {
		writeError(stderr, "unexpected extra arguments: %s\n", strings.Join(fs.Args(), " "))
		return 2
	}
	return listTools(stdout, *jsonOutput)
}

func listTools(stdout io.Writer, jsonOutput bool) int {
	tools := server.ListCLITools()
	if jsonOutput {
		data, err := json.MarshalIndent(tools, "", "  ")
		if err != nil {
			return 1
		}
		_, err = io.WriteString(stdout, ensureTrailingNewline(string(data)))
		if err != nil {
			return 1
		}
		return 0
	}
	for _, tool := range tools {
		if _, err := fmt.Fprintf(stdout, "%s\t%s\n", tool.Name, tool.Description); err != nil {
			return 1
		}
	}
	return 0
}

func runToolInfo(args []string, stdout, stderr io.Writer) int {
	fs := flag.NewFlagSet("excel-mcp cli tool-info", flag.ContinueOnError)
	fs.SetOutput(stderr)
	jsonOutput := fs.Bool("json", false, "Emit machine-readable JSON output")
	if err := fs.Parse(args); err != nil {
		return 2
	}
	if fs.NArg() != 1 {
		writeLine(stderr, "usage: excel-mcp cli tool-info [--json] <tool-name>")
		return 2
	}

	tool, ok := server.GetCLITool(fs.Arg(0))
	if !ok {
		writeError(stderr, "unknown tool %q\n", fs.Arg(0))
		return 1
	}

	if *jsonOutput {
		data, err := json.MarshalIndent(tool, "", "  ")
		if err != nil {
			writeError(stderr, "marshal tool info: %v\n", err)
			return 1
		}
		if _, err := io.WriteString(stdout, ensureTrailingNewline(string(data))); err != nil {
			writeError(stderr, "write output: %v\n", err)
			return 1
		}
		return 0
	}

	data, err := json.MarshalIndent(tool.InputSchema, "", "  ")
	if err != nil {
		writeError(stderr, "marshal tool schema: %v\n", err)
		return 1
	}
	example, err := json.MarshalIndent(tool.OutputExample, "", "  ")
	if err != nil {
		writeError(stderr, "marshal output example: %v\n", err)
		return 1
	}
	if _, err := fmt.Fprintf(stdout, "name: %s\ndescription: %s\ninput_schema:\n%s\noutput_example:\n%s\n", tool.Name, tool.Description, string(data), string(example)); err != nil {
		writeError(stderr, "write output: %v\n", err)
		return 1
	}
	return 0
}

func readPayload(inputPath string) ([]byte, error) {
	if strings.TrimSpace(inputPath) == "" {
		return []byte("{}"), nil
	}
	if inputPath == "-" {
		data, err := io.ReadAll(os.Stdin)
		if err != nil {
			return nil, fmt.Errorf("read stdin: %w", err)
		}
		return data, nil
	}
	data, err := os.ReadFile(inputPath)
	if err != nil {
		return nil, fmt.Errorf("read input %q: %w", inputPath, err)
	}
	return data, nil
}

func ensureTrailingNewline(value string) string {
	if strings.HasSuffix(value, "\n") {
		return value
	}
	return value + "\n"
}

func writeError(w io.Writer, format string, args ...any) {
	if _, err := fmt.Fprintf(w, format, args...); err != nil {
		return
	}
}

func writeLine(w io.Writer, value string) {
	if _, err := fmt.Fprintln(w, value); err != nil {
		return
	}
}

func writeUsage(w io.Writer) {
	writeLine(w, "usage: excel-mcp cli <tool-name> [--input <path|->] [--rooted]")
	writeLine(w, "       excel-mcp cli list-tools [--json]")
	writeLine(w, "       excel-mcp cli tool-info [--json] <tool-name>")
	writeLine(w, "")
	writeLine(w, "The CLI expects JSON input that matches the MCP tool arguments.")
	writeLine(w, "Tool names match the MCP tool names exactly and are intended to remain stable.")
	writeLine(w, "Use --input - to read JSON from stdin. Without --input, the tool runs with {}.")
	writeLine(w, "Use --rooted to resolve relative workbook paths under EXCEL_FILES_PATH.")
}
