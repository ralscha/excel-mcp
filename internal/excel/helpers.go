package excel

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"unicode"

	"github.com/xuri/excelize/v2"
)

const errFmtOpenWorkbook = "open workbook: %w"

var workbookLocks sync.Map

func ensureParentDir(path string) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o750); err != nil {
		return fmt.Errorf("create parent directory: %w", err)
	}
	return nil
}

func workbookLock(path string) *sync.RWMutex {
	key := filepath.Clean(path)
	if runtime.GOOS == "windows" {
		key = strings.ToLower(key)
	}
	lock, _ := workbookLocks.LoadOrStore(key, &sync.RWMutex{})
	return lock.(*sync.RWMutex)
}

func openWorkbook(path string) (*excelize.File, func(), error) {
	lock := workbookLock(path)
	lock.RLock()
	f, err := excelize.OpenFile(path)
	if err != nil {
		lock.RUnlock()
		return nil, nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	return f, func() {
		_ = f.Close()
		lock.RUnlock()
	}, nil
}

func withWorkbook(path string, fn func(*excelize.File) error) error {
	lock := workbookLock(path)
	lock.Lock()
	defer lock.Unlock()

	f, err := excelize.OpenFile(path)
	if err != nil {
		return fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()

	if err := fn(f); err != nil {
		return err
	}
	if err := f.Save(); err != nil {
		return fmt.Errorf("save workbook: %w", err)
	}
	return nil
}

func ensureSheetExists(f *excelize.File, sheetName string) error {
	index, err := f.GetSheetIndex(sheetName)
	if err != nil {
		return err
	}
	if index < 0 {
		return fmt.Errorf("sheet %s does not exist", sheetName)
	}
	return nil
}

func worksheetBounds(f *excelize.File, sheetName string) (int, int, string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return 0, 0, "", err
	}
	maxRow := len(rows)
	maxCol := 0
	for _, row := range rows {
		if len(row) > maxCol {
			maxCol = len(row)
		}
	}
	merged, err := f.GetMergeCells(sheetName)
	if err != nil {
		return 0, 0, "", err
	}
	for _, cell := range merged {
		_, _, err := excelize.CellNameToCoordinates(cell.GetStartAxis())
		if err != nil {
			return 0, 0, "", err
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(cell.GetEndAxis())
		if err != nil {
			return 0, 0, "", err
		}
		if endRow > maxRow {
			maxRow = endRow
		}
		if endCol > maxCol {
			maxCol = endCol
		}
	}
	if maxRow == 0 || maxCol == 0 {
		return maxRow, maxCol, "", nil
	}
	endCell, err := excelize.CoordinatesToCellName(maxCol, maxRow)
	if err != nil {
		return 0, 0, "", err
	}
	return maxRow, maxCol, "A1:" + endCell, nil
}

func inferCellKind(value string) string {
	if _, err := strconv.ParseBool(strings.ToLower(value)); err == nil {
		return "boolean"
	}
	if _, ok := parseFiniteFloat(value); ok {
		return "number"
	}
	return "string"
}

func parseFiniteFloat(value string) (float64, bool) {
	number, err := strconv.ParseFloat(strings.TrimSpace(value), 64)
	return number, err == nil && !math.IsNaN(number) && !math.IsInf(number, 0)
}

func dominantKind(counts map[string]int) string {
	if len(counts) == 0 {
		return "string"
	}
	bestKind := "string"
	bestCount := -1
	for _, kind := range []string{"string", "number", "boolean"} {
		if counts[kind] > bestCount {
			bestKind = kind
			bestCount = counts[kind]
		}
	}
	return bestKind
}

func normalizeRangeRef(rangeRef string) (string, string, int, int, int, int, error) {
	parts := strings.Split(rangeRef, ":")
	if len(parts) != 2 {
		return "", "", 0, 0, 0, 0, fmt.Errorf("range must be like A1:D5")
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(strings.TrimSpace(parts[0]))
	if err != nil {
		return "", "", 0, 0, 0, 0, err
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(strings.TrimSpace(parts[1]))
	if err != nil {
		return "", "", 0, 0, 0, 0, err
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	startCell, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return "", "", 0, 0, 0, 0, err
	}
	endCell, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return "", "", 0, 0, 0, 0, err
	}
	return startCell, endCell, startCol, startRow, endCol, endRow, nil
}

func normalizeRangeOrCellRef(rangeRef string) (string, string, int, int, int, int, error) {
	if !strings.Contains(rangeRef, ":") {
		rangeRef += ":" + rangeRef
	}
	return normalizeRangeRef(rangeRef)
}

func validateTableHeaders(f *excelize.File, sheetName string, startCol, endCol, headerRow int) error {
	seen := map[string]struct{}{}
	for col := startCol; col <= endCol; col++ {
		cell, err := excelize.CoordinatesToCellName(col, headerRow)
		if err != nil {
			return err
		}
		value, err := f.GetCellValue(sheetName, cell)
		if err != nil {
			return err
		}
		value = strings.TrimSpace(value)
		if value == "" {
			return fmt.Errorf("table header cells must be non-empty strings")
		}
		if _, ok := seen[value]; ok {
			return fmt.Errorf("table header cells must be unique")
		}
		seen[value] = struct{}{}
	}
	return nil
}

func stableColumns(data []map[string]any) []string {
	seen := map[string]struct{}{}
	cols := make([]string, 0)
	for _, row := range data {
		keys := make([]string, 0, len(row))
		for key := range row {
			keys = append(keys, key)
		}
		sort.Strings(keys)
		for _, key := range keys {
			if _, ok := seen[key]; ok {
				continue
			}
			seen[key] = struct{}{}
			cols = append(cols, key)
		}
	}
	return cols
}

func stringValue(value *string) string {
	if value == nil {
		return ""
	}
	return *value
}

func absoluteCellRef(cell string) string {
	letters := strings.TrimRight(cell, "0123456789")
	numbers := strings.TrimPrefix(cell, letters)
	return "$" + letters + "$" + numbers
}

func quoteSheetName(sheetName string) string {
	quote := sheetName == ""
	if _, _, err := excelize.CellNameToCoordinates(sheetName); err == nil {
		quote = true
	}
	for index, current := range []rune(sheetName) {
		if !(unicode.IsLetter(current) || current == '_' || current == '.' || (index > 0 && unicode.IsDigit(current))) {
			quote = true
			break
		}
	}
	if !quote {
		return sheetName
	}
	return "'" + strings.ReplaceAll(sheetName, "'", "''") + "'"
}
