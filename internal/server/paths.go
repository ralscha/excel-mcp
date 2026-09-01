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
			return "", fmt.Errorf("direct path mode requires an absolute filepath")
		}
		return filepath.Clean(inputPath), nil
	}

	if filepath.IsAbs(inputPath) {
		return "", fmt.Errorf("absolute filepaths are not allowed in rooted path mode")
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

	resolvedAbs, err := filepath.Abs(filepath.Join(rootAbs, clean))
	if err != nil {
		return "", fmt.Errorf("resolve filepath: %w", err)
	}

	rootReal, err := resolveExistingPath(rootAbs)
	if err != nil {
		return "", fmt.Errorf("resolve EXCEL_FILES_PATH symlinks: %w", err)
	}
	resolvedReal, err := resolveExistingPath(resolvedAbs)
	if err != nil {
		return "", fmt.Errorf("resolve filepath symlinks: %w", err)
	}
	rel, err := filepath.Rel(rootReal, resolvedReal)
	if err != nil {
		return "", fmt.Errorf("validate filepath: %w", err)
	}
	if rel == ".." || strings.HasPrefix(rel, ".."+string(filepath.Separator)) {
		return "", fmt.Errorf("directory traversal is not allowed")
	}

	return resolvedReal, nil
}

// resolveExistingPath evaluates symlinks in the nearest existing ancestor and
// then rejoins any path components that do not exist yet. This also protects
// create operations whose destination file has not been created.
func resolveExistingPath(value string) (string, error) {
	current := filepath.Clean(value)
	missing := make([]string, 0)
	for {
		_, err := os.Lstat(current)
		if err == nil {
			resolved, err := filepath.EvalSymlinks(current)
			if err != nil {
				return "", err
			}
			for index := len(missing) - 1; index >= 0; index-- {
				resolved = filepath.Join(resolved, missing[index])
			}
			return filepath.Clean(resolved), nil
		}
		if !os.IsNotExist(err) {
			return "", err
		}
		parent := filepath.Dir(current)
		if parent == current {
			return "", err
		}
		missing = append(missing, filepath.Base(current))
		current = parent
	}
}
