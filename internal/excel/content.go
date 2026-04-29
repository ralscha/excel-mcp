package excel

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func WriteData(path, sheetName string, data []map[string]any, startCell string) (string, error) {
	if len(data) == 0 {
		return "", fmt.Errorf("data is required")
	}
	columns := stableColumns(data)
	err := withWorkbook(path, func(f *excelize.File) error {
		startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
		if err != nil {
			return fmt.Errorf("invalid start_cell: %w", err)
		}
		for offset, column := range columns {
			cell, err := excelize.CoordinatesToCellName(startCol+offset, startRow)
			if err != nil {
				return err
			}
			if err := f.SetCellValue(sheetName, cell, column); err != nil {
				return err
			}
		}
		for rowIndex, row := range data {
			for colIndex, column := range columns {
				cell, err := excelize.CoordinatesToCellName(startCol+colIndex, startRow+rowIndex+1)
				if err != nil {
					return err
				}
				if err := f.SetCellValue(sheetName, cell, row[column]); err != nil {
					return err
				}
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("wrote %d row(s) to %s!%s", len(data), sheetName, startCell), nil
}

func ReadData(path, sheetName, startCell, endCell string, previewOnly bool) (*ReadResult, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()

	startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		return nil, fmt.Errorf("invalid start_cell: %w", err)
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	maxRow := len(rows)
	maxCol := 0
	for _, row := range rows {
		if len(row) > maxCol {
			maxCol = len(row)
		}
	}

	endCol := maxCol
	endRow := maxRow
	if endCell != "" {
		endCol, endRow, err = excelize.CellNameToCoordinates(endCell)
		if err != nil {
			return nil, fmt.Errorf("invalid end_cell: %w", err)
		}
	}
	if previewOnly && endRow-startRow > 10 {
		endRow = startRow + 10
	}

	result := &ReadResult{FilePath: path, SheetName: sheetName}
	rangeEnd, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return nil, err
	}
	result.Range = startCell + ":" + rangeEnd

	headers := make([]string, 0, endCol-startCol+1)
	for col := startCol; col <= endCol; col++ {
		cell, err := excelize.CoordinatesToCellName(col, startRow)
		if err != nil {
			return nil, err
		}
		value, err := f.GetCellValue(sheetName, cell)
		if err != nil {
			return nil, err
		}
		if value == "" {
			value = fmt.Sprintf("column_%d", col-startCol+1)
		}
		headers = append(headers, value)
	}

	for row := startRow + 1; row <= endRow; row++ {
		item := make(map[string]any, len(headers))
		empty := true
		for index, header := range headers {
			cell, err := excelize.CoordinatesToCellName(startCol+index, row)
			if err != nil {
				return nil, err
			}
			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return nil, err
			}
			if value != "" {
				empty = false
			}
			item[header] = value
		}
		if !empty {
			result.Rows = append(result.Rows, item)
		}
	}

	return result, nil
}

func ApplyFormula(path, sheetName, cell, formula string) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetCellFormula(sheetName, cell, formula)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("applied formula to %s!%s", sheetName, cell), nil
}

func ValidateFormulaSyntax(formula string) error {
	if !strings.HasPrefix(formula, "=") {
		return fmt.Errorf("formula must start with '='")
	}
	if strings.Count(formula, "(") != strings.Count(formula, ")") {
		return fmt.Errorf("formula has unbalanced parentheses")
	}
	if strings.Count(formula, `"`)%2 != 0 {
		return fmt.Errorf("formula has unbalanced quotes")
	}
	allowed := regexp.MustCompile(`^[=A-Za-z0-9_:$,+\-*/^&()<>.%"\s!]+$`)
	if !allowed.MatchString(formula) {
		return fmt.Errorf("formula contains unsupported characters")
	}
	return nil
}

func ValidateRange(path, sheetName, startCell, endCell string) (string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return "", fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()
	if _, _, err := excelize.CellNameToCoordinates(startCell); err != nil {
		return "", fmt.Errorf("invalid start_cell: %w", err)
	}
	if endCell != "" {
		if _, _, err := excelize.CellNameToCoordinates(endCell); err != nil {
			return "", fmt.Errorf("invalid end_cell: %w", err)
		}
	}
	if _, err := f.GetRows(sheetName); err != nil {
		return "", err
	}
	if endCell == "" {
		return fmt.Sprintf("validated range %s!%s", sheetName, startCell), nil
	}
	return fmt.Sprintf("validated range %s!%s:%s", sheetName, startCell, endCell), nil
}

func GetDataValidationInfo(path, sheetName string) ([]ValidationInfo, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()
	return getDataValidationInfoFromFile(f, sheetName)
}

func getDataValidationInfoFromFile(f *excelize.File, sheetName string) ([]ValidationInfo, error) {
	items, err := f.GetDataValidations(sheetName)
	if err != nil {
		return nil, err
	}
	result := make([]ValidationInfo, 0, len(items))
	for _, item := range items {
		result = append(result, ValidationInfo{
			Sqref:       item.Sqref,
			Type:        item.Type,
			Operator:    item.Operator,
			Formula1:    item.Formula1,
			Formula2:    item.Formula2,
			ErrorTitle:  stringValue(item.ErrorTitle),
			ErrorBody:   stringValue(item.Error),
			PromptTitle: stringValue(item.PromptTitle),
			PromptBody:  stringValue(item.Prompt),
		})
	}
	return result, nil
}

func FilterRows(path, sheetName, rangeRef string, opts FilterRowsOptions) (*FilterRowsResult, error) {
	startCell, endCell, startCol, startRow, endCol, endRow, err := normalizeRangeRef(rangeRef)
	if err != nil {
		return nil, fmt.Errorf("invalid range: %w", err)
	}
	if len(opts.Filters) == 0 {
		return nil, fmt.Errorf("filters is required")
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()
	if err := ensureSheetExists(f, sheetName); err != nil {
		return nil, err
	}

	headers := make([]string, 0, endCol-startCol+1)
	for col := startCol; col <= endCol; col++ {
		headerName := fmt.Sprintf("column_%d", col-startCol+1)
		if opts.HasHeader {
			cell, err := excelize.CoordinatesToCellName(col, startRow)
			if err != nil {
				return nil, err
			}
			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return nil, err
			}
			if strings.TrimSpace(value) != "" {
				headerName = value
			}
		}
		headers = append(headers, headerName)
	}

	resolvedFilters, err := resolveFilters(opts.Filters, headers, endCol-startCol+1, opts.HasHeader)
	if err != nil {
		return nil, err
	}

	dataStartRow := startRow
	if opts.HasHeader {
		dataStartRow++
	}
	result := &FilterRowsResult{
		FilePath:  path,
		SheetName: sheetName,
		Range:     startCell + ":" + endCell,
		HasHeader: opts.HasHeader,
		Filters:   opts.Filters,
	}
	for row := dataStartRow; row <= endRow; row++ {
		item := make(map[string]any, len(headers))
		values := make([]string, 0, len(headers))
		empty := true
		for index, header := range headers {
			cell, err := excelize.CoordinatesToCellName(startCol+index, row)
			if err != nil {
				return nil, err
			}
			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return nil, err
			}
			if value != "" {
				empty = false
			}
			item[header] = value
			values = append(values, value)
		}
		if empty {
			continue
		}
		matched, err := matchesAllFilters(values, resolvedFilters)
		if err != nil {
			return nil, err
		}
		if matched {
			result.Rows = append(result.Rows, item)
		}
	}

	return result, nil
}

type resolvedFilter struct {
	Index    int
	Operator string
	Value    string
}

func resolveFilters(filters []Filter, headers []string, columnCount int, hasHeader bool) ([]resolvedFilter, error) {
	resolved := make([]resolvedFilter, 0, len(filters))
	for _, filter := range filters {
		column := strings.TrimSpace(filter.Column)
		if column == "" {
			return nil, fmt.Errorf("filter column is required")
		}
		operator := strings.ToLower(strings.TrimSpace(filter.Operator))
		if !isSupportedFilterOperator(operator) {
			return nil, fmt.Errorf("unsupported operator %q", filter.Operator)
		}
		if columnIndex, err := strconv.Atoi(column); err == nil {
			if columnIndex < 1 || columnIndex > columnCount {
				return nil, fmt.Errorf("column index %d is out of range", columnIndex)
			}
			resolved = append(resolved, resolvedFilter{Index: columnIndex - 1, Operator: operator, Value: filter.Value})
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
		resolved = append(resolved, resolvedFilter{Index: matched, Operator: operator, Value: filter.Value})
	}
	return resolved, nil
}

func isSupportedFilterOperator(operator string) bool {
	switch operator {
	case "equals", "contains", "gt", "gte", "lt", "lte", "regex":
		return true
	default:
		return false
	}
}

func matchesAllFilters(values []string, filters []resolvedFilter) (bool, error) {
	for _, filter := range filters {
		matched, err := matchFilterValue(values[filter.Index], filter.Operator, filter.Value)
		if err != nil {
			return false, err
		}
		if !matched {
			return false, nil
		}
	}
	return true, nil
}

func matchFilterValue(actual, operator, expected string) (bool, error) {
	switch operator {
	case "equals":
		return strings.EqualFold(strings.TrimSpace(actual), strings.TrimSpace(expected)), nil
	case "contains":
		return strings.Contains(strings.ToLower(actual), strings.ToLower(expected)), nil
	case "regex":
		re, err := regexp.Compile(expected)
		if err != nil {
			return false, fmt.Errorf("invalid regex value: %w", err)
		}
		return re.MatchString(actual), nil
	case "gt", "gte", "lt", "lte":
		actualNumber, err := strconv.ParseFloat(strings.TrimSpace(actual), 64)
		if err != nil {
			return false, nil //nolint:nilerr // non-numeric cells don't match numeric comparisons
		}
		expectedNumber, err := strconv.ParseFloat(strings.TrimSpace(expected), 64)
		if err != nil {
			return false, fmt.Errorf("operator %q requires a numeric value", operator)
		}
		switch operator {
		case "gt":
			return actualNumber > expectedNumber, nil
		case "gte":
			return actualNumber >= expectedNumber, nil
		case "lt":
			return actualNumber < expectedNumber, nil
		default:
			return actualNumber <= expectedNumber, nil
		}
	default:
		return false, fmt.Errorf("unsupported operator %q", operator)
	}
}
