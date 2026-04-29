package server

import (
	"path/filepath"
	"testing"
)

func TestResolvePathDirectRequiresAbsolute(t *testing.T) {
	if _, err := ResolvePath(PathModeDirect, "relative.xlsx"); err == nil {
		t.Fatal("expected direct mode to reject relative paths")
	}

	abs := filepath.Join(t.TempDir(), "book.xlsx")
	resolved, err := ResolvePath(PathModeDirect, abs)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if resolved != abs {
		t.Fatalf("expected %q, got %q", abs, resolved)
	}
}

func TestResolvePathRootedRejectsTraversal(t *testing.T) {
	root := t.TempDir()
	t.Setenv("EXCEL_FILES_PATH", root)

	if _, err := ResolvePath(PathModeRooted, filepath.Join("..", "outside.xlsx")); err == nil {
		t.Fatal("expected traversal rejection")
	}

	resolved, err := ResolvePath(PathModeRooted, filepath.Join("reports", "q1.xlsx"))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	want := filepath.Join(root, "reports", "q1.xlsx")
	if resolved != want {
		t.Fatalf("expected %q, got %q", want, resolved)
	}
}
