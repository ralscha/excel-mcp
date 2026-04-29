package server

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

type PathMode int

const (
	PathModeDirect PathMode = iota
	PathModeRooted
)

func ResolvePath(mode PathMode, inputPath string) (string, error) {
	if strings.TrimSpace(inputPath) == "" {
		return "", fmt.Errorf("filepath is required")
	}

	if mode == PathModeDirect {
		if !filepath.IsAbs(inputPath) {
			return "", fmt.Errorf("stdio mode requires an absolute filepath")
		}
		return filepath.Clean(inputPath), nil
	}

	if filepath.IsAbs(inputPath) {
		return "", fmt.Errorf("absolute filepaths are not allowed for HTTP transport")
	}

	clean := filepath.Clean(inputPath)
	if clean == "." || clean == "" {
		return "", fmt.Errorf("filepath is required")
	}
	if clean == ".." || strings.HasPrefix(clean, ".."+string(filepath.Separator)) {
		return "", fmt.Errorf("directory traversal is not allowed")
	}

	root := os.Getenv("EXCEL_FILES_PATH")
	if strings.TrimSpace(root) == "" {
		root = "excel_files"
	}

	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return "", fmt.Errorf("resolve EXCEL_FILES_PATH: %w", err)
	}

	resolved := filepath.Join(rootAbs, clean)
	resolvedAbs, err := filepath.Abs(resolved)
	if err != nil {
		return "", fmt.Errorf("resolve filepath: %w", err)
	}

	rel, err := filepath.Rel(rootAbs, resolvedAbs)
	if err != nil {
		return "", fmt.Errorf("validate filepath: %w", err)
	}
	if rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return "", fmt.Errorf("directory traversal is not allowed")
	}

	return resolvedAbs, nil
}
