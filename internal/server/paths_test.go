package server

import (
	"os"
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

func TestResolvePathRootedRejectsSymlinkEscape(t *testing.T) {
	root := t.TempDir()
	outside := t.TempDir()
	link := filepath.Join(root, "outside-link")
	if err := os.Symlink(outside, link); err != nil {
		t.Skipf("symlinks are unavailable: %v", err)
	}
	t.Setenv("EXCEL_FILES_PATH", root)
	if _, err := ResolvePath(PathModeRooted, filepath.Join("outside-link", "book.xlsx")); err == nil {
		t.Fatal("expected symlink escape rejection")
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
	rootReal, err := filepath.EvalSymlinks(root)
	if err != nil {
		t.Fatalf("resolve temporary root: %v", err)
	}
	want := filepath.Join(rootReal, "reports", "q1.xlsx")
	if resolved != want {
		t.Fatalf("expected %q, got %q", want, resolved)
	}
}
