package server

import "testing"

func TestRootedPathRejectsWindowsSpecialPaths(t *testing.T) {
	t.Setenv("EXCEL_FILES_PATH", t.TempDir())
	for _, path := range []string{
		`C:book.xlsx`,
		`\book.xlsx`,
		`book.xlsx:stream`,
		`NUL`,
		`CON`,
		`nested\COM1\book.xlsx`,
	} {
		t.Run(path, func(t *testing.T) {
			if resolved, err := ResolvePath(PathModeRooted, path); err == nil {
				t.Fatalf("special path %q resolved to %q", path, resolved)
			}
		})
	}
}
