package excel

import (
	"fmt"
	"maps"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func InsertRows(path, sheetName string, startRow, count int) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		return f.InsertRows(sheetName, startRow, count)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("inserted %d row(s) at %s!%d", count, sheetName, startRow), nil
}

func InsertColumns(path, sheetName string, startCol, count int) (string, error) {
	axis, err := excelize.ColumnNumberToName(startCol)
	if err != nil {
		return "", err
	}
	err = withWorkbook(path, func(f *excelize.File) error {
		return f.InsertCols(sheetName, axis, count)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("inserted %d column(s) at %s!%s", count, sheetName, axis), nil
}

func DeleteSheetRows(path, sheetName string, startRow, count int) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		for range count {
			if err := f.RemoveRow(sheetName, startRow); err != nil {
				return err
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("deleted %d row(s) at %s!%d", count, sheetName, startRow), nil
}

func DeleteSheetColumns(path, sheetName string, startCol, count int) (string, error) {
	axis, err := excelize.ColumnNumberToName(startCol)
	if err != nil {
		return "", err
	}
	err = withWorkbook(path, func(f *excelize.File) error {
		for range count {
			if err := f.RemoveCol(sheetName, axis); err != nil {
				return err
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("deleted %d column(s) at %s!%s", count, sheetName, axis), nil
}

func CreateTable(path, sheetName, dataRange, tableName, tableStyle string) (string, error) {
	if tableStyle == "" {
		tableStyle = "TableStyleMedium9"
	}
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		startCell, endCell, startCol, startRow, endCol, endRow, err := normalizeRangeRef(dataRange)
		if err != nil {
			return fmt.Errorf("invalid data_range: %w", err)
		}
		if endRow-startRow < 1 {
			return fmt.Errorf("table range must include a header row and at least one data row")
		}
		if err := validateTableHeaders(f, sheetName, startCol, endCol, startRow); err != nil {
			return err
		}
		return f.AddTable(sheetName, &excelize.Table{Range: startCell + ":" + endCell, Name: tableName, StyleName: tableStyle})
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("created table on %s!%s", sheetName, dataRange), nil
}

func CreatePivotTable(path, sheetName, dataRange, targetCell string, opts PivotOptions) (string, error) {
	if len(opts.Rows) == 0 || len(opts.Values) == 0 {
		return "", fmt.Errorf("rows and values are required")
	}
	pivotRows := make([]excelize.PivotTableField, 0, len(opts.Rows))
	for _, value := range opts.Rows {
		pivotRows = append(pivotRows, excelize.PivotTableField{Data: value})
	}
	pivotCols := make([]excelize.PivotTableField, 0, len(opts.Columns))
	for _, value := range opts.Columns {
		pivotCols = append(pivotCols, excelize.PivotTableField{Data: value})
	}
	subtotal, err := pivotSubtotal(opts.AggFunc)
	if err != nil {
		return "", err
	}
	pivotData := make([]excelize.PivotTableField, 0, len(opts.Values))
	for _, value := range opts.Values {
		pivotData = append(pivotData, excelize.PivotTableField{Data: value, Name: subtotal + " of " + value, Subtotal: subtotal})
	}
	err = withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		startCol, startRow, err := excelize.CellNameToCoordinates(targetCell)
		if err != nil {
			return fmt.Errorf("invalid target_cell: %w", err)
		}
		startCell, endCell, startDataCol, startDataRow, endDataCol, endDataRow, err := normalizeRangeRef(dataRange)
		if err != nil {
			return fmt.Errorf("invalid data_range: %w", err)
		}
		if endDataRow-startDataRow < 1 || endDataCol-startDataCol < 1 {
			return fmt.Errorf("data_range must include headers and at least one data row")
		}
		endColName, err := excelize.ColumnNumberToName(startCol + max(4, len(opts.Rows)+len(opts.Columns)+len(opts.Values)))
		if err != nil {
			return err
		}
		pivotRange := fmt.Sprintf("%s!%s:%s%d", sheetName, targetCell, endColName, startRow+max(10, (endDataRow-startDataRow)+6))
		return f.AddPivotTable(&excelize.PivotTableOptions{
			DataRange:       fmt.Sprintf("%s!%s:%s", sheetName, startCell, endCell),
			PivotTableRange: pivotRange,
			Rows:            pivotRows,
			Columns:         pivotCols,
			Data:            pivotData,
			RowGrandTotals:  true,
			ColGrandTotals:  true,
			ShowDrill:       true,
			ShowRowHeaders:  true,
			ShowColHeaders:  true,
		})
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("created pivot table at %s!%s", sheetName, targetCell), nil
}

func pivotSubtotal(value string) (string, error) {
	switch strings.ToLower(strings.TrimSpace(value)) {
	case "count":
		return "Count", nil
	case "average", "avg", "mean":
		return "Average", nil
	case "max":
		return "Max", nil
	case "min":
		return "Min", nil
	case "sum", "":
		return "Sum", nil
	default:
		return "", fmt.Errorf("unsupported agg_func %q; supported values are sum, count, average, avg, mean, max, min", value)
	}
}

func CopyRange(path, sheetName, sourceStart, sourceEnd, targetStart, targetSheet string) (string, error) {
	if targetSheet == "" {
		targetSheet = sheetName
	}
	_, _, startCol, startRow, endCol, endRow, err := normalizeRangeRef(sourceStart + ":" + sourceEnd)
	if err != nil {
		return "", fmt.Errorf("invalid source range: %w", err)
	}
	targetCol, targetRow, err := excelize.CellNameToCoordinates(targetStart)
	if err != nil {
		return "", err
	}
	err = withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		if err := ensureSheetExists(f, targetSheet); err != nil {
			return err
		}
		for rowOffset := 0; rowOffset <= endRow-startRow; rowOffset++ {
			for colOffset := 0; colOffset <= endCol-startCol; colOffset++ {
				sourceCell, _ := excelize.CoordinatesToCellName(startCol+colOffset, startRow+rowOffset)
				targetCell, _ := excelize.CoordinatesToCellName(targetCol+colOffset, targetRow+rowOffset)
				value, err := f.GetCellValue(sheetName, sourceCell)
				if err != nil {
					return err
				}
				if err := f.SetCellValue(targetSheet, targetCell, value); err != nil {
					return err
				}
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("copied range %s!%s:%s to %s!%s", sheetName, sourceStart, sourceEnd, targetSheet, targetStart), nil
}

func DeleteRange(path, sheetName, startCell, endCell, shiftDirection string) (string, error) {
	_, _, startCol, startRow, endCol, endRow, err := normalizeRangeRef(startCell + ":" + endCell)
	if err != nil {
		return "", fmt.Errorf("invalid range: %w", err)
	}
	shiftDirection = strings.ToLower(strings.TrimSpace(shiftDirection))
	if shiftDirection == "" {
		shiftDirection = "up"
	}
	if shiftDirection != "up" && shiftDirection != "left" {
		return "", fmt.Errorf("shift_direction must be 'up' or 'left'")
	}
	err = withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return err
		}
		maxRow := len(rows)
		maxCol := 0
		for _, row := range rows {
			if len(row) > maxCol {
				maxCol = len(row)
			}
		}
		rowSpan := endRow - startRow + 1
		colSpan := endCol - startCol + 1
		if shiftDirection == "left" {
			for row := startRow; row <= endRow; row++ {
				for col := startCol; col <= maxCol-colSpan; col++ {
					fromCell, _ := excelize.CoordinatesToCellName(col+colSpan, row)
					toCell, _ := excelize.CoordinatesToCellName(col, row)
					value, err := f.GetCellValue(sheetName, fromCell)
					if err != nil {
						return err
					}
					if err := f.SetCellValue(sheetName, toCell, value); err != nil {
						return err
					}
				}
				for col := maxCol - colSpan + 1; col <= maxCol; col++ {
					cell, _ := excelize.CoordinatesToCellName(col, row)
					if err := f.SetCellValue(sheetName, cell, ""); err != nil {
						return err
					}
				}
			}
			return nil
		}
		for col := startCol; col <= endCol; col++ {
			for row := startRow; row <= maxRow-rowSpan; row++ {
				fromCell, _ := excelize.CoordinatesToCellName(col, row+rowSpan)
				toCell, _ := excelize.CoordinatesToCellName(col, row)
				value, err := f.GetCellValue(sheetName, fromCell)
				if err != nil {
					return err
				}
				if err := f.SetCellValue(sheetName, toCell, value); err != nil {
					return err
				}
			}
			for row := maxRow - rowSpan + 1; row <= maxRow; row++ {
				cell, _ := excelize.CoordinatesToCellName(col, row)
				if err := f.SetCellValue(sheetName, cell, ""); err != nil {
					return err
				}
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("deleted range %s!%s:%s shifting %s", sheetName, startCell, endCell, shiftDirection), nil
}

func ClearRange(path, sheetName, startCell, endCell string) (string, error) {
	_, _, startCol, startRow, endCol, endRow, err := normalizeRangeRef(startCell + ":" + endCell)
	if err != nil {
		return "", fmt.Errorf("invalid range: %w", err)
	}
	err = withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		for row := startRow; row <= endRow; row++ {
			for col := startCol; col <= endCol; col++ {
				cell, err := excelize.CoordinatesToCellName(col, row)
				if err != nil {
					return err
				}
				if err := f.SetCellValue(sheetName, cell, ""); err != nil {
					return err
				}
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("cleared range %s!%s:%s", sheetName, startCell, endCell), nil
}

func SortRange(path, sheetName, rangeRef string, opts SortRangeOptions) (*SortRangeResult, error) {
	startCell, endCell, startCol, startRow, endCol, endRow, err := normalizeRangeRef(rangeRef)
	if err != nil {
		return nil, fmt.Errorf("invalid range: %w", err)
	}
	if len(opts.SortKeys) == 0 {
		return nil, fmt.Errorf("sort_keys is required")
	}
	if !opts.HasHeader {
		for _, key := range opts.SortKeys {
			if _, err := strconv.Atoi(strings.TrimSpace(key.Column)); err != nil {
				return nil, fmt.Errorf("column names require has_header=true")
			}
		}
	}

	err = withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}

		headers := make([]string, 0, endCol-startCol+1)
		for col := startCol; col <= endCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, startRow)
			if err != nil {
				return err
			}
			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return err
			}
			headers = append(headers, value)
		}

		resolvedKeys, err := resolveSortKeys(opts.SortKeys, headers, startCol, endCol, opts.HasHeader)
		if err != nil {
			return err
		}

		dataStartRow := startRow
		if opts.HasHeader {
			dataStartRow++
		}
		if dataStartRow > endRow {
			return nil
		}

		rows := make([]sortableRangeRow, 0, endRow-dataStartRow+1)
		for row := dataStartRow; row <= endRow; row++ {
			values := make([]string, 0, endCol-startCol+1)
			for col := startCol; col <= endCol; col++ {
				cell, err := excelize.CoordinatesToCellName(col, row)
				if err != nil {
					return err
				}
				value, err := f.GetCellValue(sheetName, cell)
				if err != nil {
					return err
				}
				values = append(values, value)
			}
			rows = append(rows, sortableRangeRow{Values: values})
		}

		sort.SliceStable(rows, func(leftIndex, rightIndex int) bool {
			left := rows[leftIndex]
			right := rows[rightIndex]
			for _, key := range resolvedKeys {
				comparison := compareSortableValues(left.Values[key.Index], right.Values[key.Index])
				if comparison == 0 {
					continue
				}
				if key.Descending {
					return comparison > 0
				}
				return comparison < 0
			}
			return false
		})

		for rowOffset, row := range rows {
			for colOffset, value := range row.Values {
				cell, err := excelize.CoordinatesToCellName(startCol+colOffset, dataStartRow+rowOffset)
				if err != nil {
					return err
				}
				if err := f.SetCellValue(sheetName, cell, coerceCellValue(value)); err != nil {
					return err
				}
			}
		}
		return nil
	})
	if err != nil {
		return nil, err
	}

	return &SortRangeResult{
		FilePath:  path,
		SheetName: sheetName,
		Range:     startCell + ":" + endCell,
		HasHeader: opts.HasHeader,
		SortKeys:  opts.SortKeys,
	}, nil
}

type sortableRangeRow struct {
	Values []string
}

type resolvedSortKey struct {
	Index      int
	Descending bool
}

func resolveSortKeys(keys []SortKey, headers []string, startCol, endCol int, hasHeader bool) ([]resolvedSortKey, error) {
	resolved := make([]resolvedSortKey, 0, len(keys))
	columnCount := endCol - startCol + 1
	for _, key := range keys {
		column := strings.TrimSpace(key.Column)
		if column == "" {
			return nil, fmt.Errorf("sort key column is required")
		}
		if columnIndex, err := strconv.Atoi(column); err == nil {
			if columnIndex < 1 || columnIndex > columnCount {
				return nil, fmt.Errorf("column index %d is out of range", columnIndex)
			}
			resolved = append(resolved, resolvedSortKey{Index: columnIndex - 1, Descending: key.Descending})
			continue
		}
		if !hasHeader {
			return nil, fmt.Errorf("column names require has_header=true")
		}
		matched := -1
		for index, header := range headers {
			if header == column {
				matched = index
				break
			}
		}
		if matched < 0 {
			return nil, fmt.Errorf("column %q not found in header row", column)
		}
		resolved = append(resolved, resolvedSortKey{Index: matched, Descending: key.Descending})
	}
	return resolved, nil
}

func compareSortableValues(left, right string) int {
	left = strings.TrimSpace(left)
	right = strings.TrimSpace(right)
	if left == right {
		return 0
	}
	if leftNumber, err := strconv.ParseFloat(left, 64); err == nil {
		if rightNumber, err := strconv.ParseFloat(right, 64); err == nil {
			switch {
			case leftNumber < rightNumber:
				return -1
			case leftNumber > rightNumber:
				return 1
			default:
				return 0
			}
		}
	}
	if leftBool, err := strconv.ParseBool(strings.ToLower(left)); err == nil {
		if rightBool, err := strconv.ParseBool(strings.ToLower(right)); err == nil {
			switch {
			case !leftBool && rightBool:
				return -1
			case leftBool && !rightBool:
				return 1
			default:
				return 0
			}
		}
	}
	leftLower := strings.ToLower(left)
	rightLower := strings.ToLower(right)
	if leftLower < rightLower {
		return -1
	}
	if leftLower > rightLower {
		return 1
	}
	if left < right {
		return -1
	}
	return 1
}

func coerceCellValue(value string) any {
	trimmed := strings.TrimSpace(value)
	if number, err := strconv.ParseFloat(trimmed, 64); err == nil {
		return number
	}
	if boolean, err := strconv.ParseBool(strings.ToLower(trimmed)); err == nil {
		return boolean
	}
	return value
}

func UpsertRows(path, sheetName, rangeRef string, opts UpsertRowsOptions) (*UpsertRowsResult, error) {
	startCell, endCell, startCol, startRow, endCol, endRow, err := normalizeRangeRef(rangeRef)
	if err != nil {
		return nil, fmt.Errorf("invalid range: %w", err)
	}
	if len(opts.KeyColumns) == 0 {
		return nil, fmt.Errorf("key_columns is required")
	}
	if len(opts.Rows) == 0 {
		return nil, fmt.Errorf("rows is required")
	}

	result := &UpsertRowsResult{
		FilePath:   path,
		SheetName:  sheetName,
		Range:      startCell + ":" + endCell,
		KeyColumns: append([]string(nil), opts.KeyColumns...),
	}

	err = withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}

		_, headerIndexes, err := readTableHeaders(f, sheetName, startCol, startRow, endCol)
		if err != nil {
			return err
		}
		keyIndexes, err := resolveUpsertKeyIndexes(opts.KeyColumns, headerIndexes)
		if err != nil {
			return err
		}
		if err := validateUpsertRows(opts.Rows, opts.KeyColumns); err != nil {
			return err
		}
		rowIndex, lastUsedRow, err := buildUpsertIndex(f, sheetName, startCol, startRow, endCol, endRow, opts.KeyColumns, keyIndexes, opts.CaseSensitiveKeys)
		if err != nil {
			return err
		}

		for _, item := range opts.Rows {
			if err := validateUpsertValues(item.Values, headerIndexes, opts.KeyColumns); err != nil {
				return err
			}
			keyPayload := upsertKeyPayload(item.Match, opts.KeyColumns)
			canonicalKey := canonicalUpsertKey(item.Match, opts.KeyColumns, opts.CaseSensitiveKeys)
			matches := rowIndex[canonicalKey]
			if len(matches) > 1 {
				return fmt.Errorf("multiple existing rows matched key")
			}
			if len(matches) == 1 {
				targetRow := matches[0]
				if err := writeUpsertValues(f, sheetName, targetRow, startCol, headerIndexes, item.Values); err != nil {
					return err
				}
				result.UpdatedCount++
				result.Results = append(result.Results, UpsertRowResult{Key: keyPayload, Action: "updated", RowNumber: targetRow})
				continue
			}
			if !opts.InsertIfMissing {
				result.SkippedCount++
				result.Results = append(result.Results, UpsertRowResult{Key: keyPayload, Action: "skipped", Reason: "no match and insert_if_missing=false"})
				continue
			}
			targetRow := max(lastUsedRow+1, startRow+1)
			if targetRow > endRow {
				return fmt.Errorf("no empty rows remain within range for insert")
			}
			insertValues := make(map[string]any, len(item.Values)+len(opts.KeyColumns))
			maps.Copy(insertValues, item.Values)
			for _, keyColumn := range opts.KeyColumns {
				insertValues[keyColumn] = item.Match[keyColumn]
			}
			if err := writeUpsertValues(f, sheetName, targetRow, startCol, headerIndexes, insertValues); err != nil {
				return err
			}
			lastUsedRow = targetRow
			rowIndex[canonicalKey] = []int{targetRow}
			result.InsertedCount++
			result.Results = append(result.Results, UpsertRowResult{Key: keyPayload, Action: "inserted", RowNumber: targetRow})
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	return result, nil
}

func readTableHeaders(f *excelize.File, sheetName string, startCol, startRow, endCol int) ([]string, map[string]int, error) {
	headers := make([]string, 0, endCol-startCol+1)
	indexes := make(map[string]int, endCol-startCol+1)
	for col := startCol; col <= endCol; col++ {
		cell, err := excelize.CoordinatesToCellName(col, startRow)
		if err != nil {
			return nil, nil, err
		}
		header, err := f.GetCellValue(sheetName, cell)
		if err != nil {
			return nil, nil, err
		}
		header = strings.TrimSpace(header)
		if header == "" {
			return nil, nil, fmt.Errorf("header row contains empty column names")
		}
		if _, exists := indexes[header]; exists {
			return nil, nil, fmt.Errorf("header row contains duplicate column %q", header)
		}
		indexes[header] = col - startCol
		headers = append(headers, header)
	}
	return headers, indexes, nil
}

func resolveUpsertKeyIndexes(keyColumns []string, headerIndexes map[string]int) (map[string]int, error) {
	resolved := make(map[string]int, len(keyColumns))
	for _, keyColumn := range keyColumns {
		index, ok := headerIndexes[keyColumn]
		if !ok {
			return nil, fmt.Errorf("key column %q not found in header row", keyColumn)
		}
		resolved[keyColumn] = index
	}
	return resolved, nil
}

func validateUpsertRows(rows []UpsertRow, keyColumns []string) error {
	allowedKeys := make(map[string]struct{}, len(keyColumns))
	for _, keyColumn := range keyColumns {
		allowedKeys[keyColumn] = struct{}{}
	}
	for index, row := range rows {
		for _, keyColumn := range keyColumns {
			if _, ok := row.Match[keyColumn]; !ok {
				return fmt.Errorf("row %d match is missing key column %q", index+1, keyColumn)
			}
		}
		for matchKey := range row.Match {
			if _, ok := allowedKeys[matchKey]; !ok {
				return fmt.Errorf("row %d match contains unsupported column %q", index+1, matchKey)
			}
		}
	}
	return nil
}

func validateUpsertValues(values map[string]any, headerIndexes map[string]int, keyColumns []string) error {
	keyColumnSet := make(map[string]struct{}, len(keyColumns))
	for _, keyColumn := range keyColumns {
		keyColumnSet[keyColumn] = struct{}{}
	}
	for column := range values {
		if _, ok := headerIndexes[column]; !ok {
			return fmt.Errorf("values contains unknown column %q", column)
		}
		if _, ok := keyColumnSet[column]; ok {
			return fmt.Errorf("values cannot modify key column %q", column)
		}
	}
	return nil
}

func buildUpsertIndex(f *excelize.File, sheetName string, startCol, startRow, endCol, endRow int, keyColumns []string, keyIndexes map[string]int, caseSensitive bool) (map[string][]int, int, error) {
	index := map[string][]int{}
	lastUsedRow := startRow
	for row := startRow + 1; row <= endRow; row++ {
		values := make([]string, 0, endCol-startCol+1)
		nonEmpty := false
		for col := startCol; col <= endCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return nil, 0, err
			}
			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return nil, 0, err
			}
			if strings.TrimSpace(value) != "" {
				nonEmpty = true
			}
			values = append(values, value)
		}
		if !nonEmpty {
			continue
		}
		lastUsedRow = row
		key := canonicalUpsertKeyFromValues(values, keyColumns, keyIndexes, caseSensitive)
		index[key] = append(index[key], row)
	}
	return index, lastUsedRow, nil
}

func canonicalUpsertKey(match map[string]any, keyColumns []string, caseSensitive bool) string {
	parts := make([]string, 0, len(keyColumns))
	for _, keyColumn := range keyColumns {
		parts = append(parts, normalizeUpsertKeyPart(match[keyColumn], caseSensitive))
	}
	return strings.Join(parts, "\x1f")
}

func canonicalUpsertKeyFromValues(values []string, keyColumns []string, keyIndexes map[string]int, caseSensitive bool) string {
	parts := make([]string, 0, len(keyColumns))
	for _, keyColumn := range keyColumns {
		parts = append(parts, normalizeUpsertKeyPart(values[keyIndexes[keyColumn]], caseSensitive))
	}
	return strings.Join(parts, "\x1f")
}

func normalizeUpsertKeyPart(value any, caseSensitive bool) string {
	text := strings.TrimSpace(fmt.Sprint(value))
	if !caseSensitive {
		text = strings.ToLower(text)
	}
	return text
}

func upsertKeyPayload(match map[string]any, keyColumns []string) map[string]any {
	result := make(map[string]any, len(keyColumns))
	for _, keyColumn := range keyColumns {
		result[keyColumn] = match[keyColumn]
	}
	return result
}

func writeUpsertValues(f *excelize.File, sheetName string, targetRow, startCol int, headerIndexes map[string]int, values map[string]any) error {
	for column, value := range values {
		cell, err := excelize.CoordinatesToCellName(startCol+headerIndexes[column], targetRow)
		if err != nil {
			return err
		}
		if err := f.SetCellValue(sheetName, cell, value); err != nil {
			return err
		}
	}
	return nil
}
