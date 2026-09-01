package excel

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"unicode"

	"github.com/xuri/excelize/v2"
)

func WriteData(path, sheetName string, data []map[string]any, startCell string) (string, error) {
	if len(data) == 0 {
		return "", fmt.Errorf("data is required")
	}
	columns := stableColumns(data)
	if len(columns) == 0 {
		return "", fmt.Errorf("data rows must contain at least one column")
	}
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
	f, closeWorkbook, err := openWorkbook(path)
	if err != nil {
		return nil, err
	}
	defer closeWorkbook()

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
		if endCol < startCol || endRow < startRow {
			return nil, fmt.Errorf("range end must be below and to the right of start_cell")
		}
	} else {
		endCol = max(endCol, startCol)
		endRow = max(endRow, startRow)
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
	seenHeaders := make(map[string]struct{}, endCol-startCol+1)
	for col := startCol; col <= endCol; col++ {
		cell, err := excelize.CoordinatesToCellName(col, startRow)
		if err != nil {
			return nil, err
		}
		value, err := f.GetCellValue(sheetName, cell)
		if err != nil {
			return nil, err
		}
		headers, err = appendHeader(headers, seenHeaders, value)
		if err != nil {
			return nil, err
		}
	}

	for row := startRow + 1; row <= endRow; row++ {
		item := make(map[string]any, len(headers))
		empty := true
		for index, header := range headers {
			cell, err := excelize.CoordinatesToCellName(startCol+index, row)
			if err != nil {
				return nil, err
			}
			value, displayValue, err := readCellJSONValue(f, sheetName, cell)
			if err != nil {
				return nil, err
			}
			if displayValue != "" {
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
	if err := ValidateFormulaSyntax(formula); err != nil {
		return "", err
	}
	err := withWorkbook(path, func(f *excelize.File) error {
		return f.SetCellFormula(sheetName, cell, strings.TrimPrefix(formula, "="))
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
	if strings.TrimSpace(strings.TrimPrefix(formula, "=")) == "" {
		return fmt.Errorf("formula expression is required after '='")
	}
	depth := 0
	inString := false
	runes := []rune(formula)
	for index := 1; index < len(runes); index++ {
		current := runes[index]
		if current == '"' {
			if inString && index+1 < len(runes) && runes[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
			continue
		}
		if unicode.IsControl(current) && current != '\t' && current != '\r' && current != '\n' {
			return fmt.Errorf("formula contains unsupported control characters")
		}
		if inString {
			continue
		}
		switch current {
		case '(':
			depth++
		case ')':
			depth--
			if depth < 0 {
				return fmt.Errorf("formula has unbalanced parentheses")
			}
		}
	}
	if inString {
		return fmt.Errorf("formula has unbalanced quotes")
	}
	if depth != 0 {
		return fmt.Errorf("formula has unbalanced parentheses")
	}
	return nil
}

func ValidateRange(path, sheetName, startCell, endCell string) (string, error) {
	f, closeWorkbook, err := openWorkbook(path)
	if err != nil {
		return "", err
	}
	defer closeWorkbook()
	startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		return "", fmt.Errorf("invalid start_cell: %w", err)
	}
	if endCell != "" {
		endCol, endRow, err := excelize.CellNameToCoordinates(endCell)
		if err != nil {
			return "", fmt.Errorf("invalid end_cell: %w", err)
		}
		if endCol < startCol || endRow < startRow {
			return "", fmt.Errorf("range end must be below and to the right of start_cell")
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
	f, closeWorkbook, err := openWorkbook(path)
	if err != nil {
		return nil, err
	}
	defer closeWorkbook()
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
			Sqref:            item.Sqref,
			Type:             item.Type,
			Operator:         item.Operator,
			Formula1:         item.Formula1,
			Formula2:         item.Formula2,
			AllowBlank:       item.AllowBlank,
			ShowDropDown:     item.ShowDropDown,
			ShowErrorMessage: item.ShowErrorMessage,
			ErrorStyle:       stringValue(item.ErrorStyle),
			ErrorTitle:       stringValue(item.ErrorTitle),
			ErrorBody:        stringValue(item.Error),
			ShowInputMessage: item.ShowInputMessage,
			PromptTitle:      stringValue(item.PromptTitle),
			PromptBody:       stringValue(item.Prompt),
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

	f, closeWorkbook, err := openWorkbook(path)
	if err != nil {
		return nil, err
	}
	defer closeWorkbook()
	if err := ensureSheetExists(f, sheetName); err != nil {
		return nil, err
	}

	headers := make([]string, 0, endCol-startCol+1)
	seenHeaders := make(map[string]struct{}, endCol-startCol+1)
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
				headerName = strings.TrimSpace(value)
			}
		}
		headers, err = appendHeader(headers, seenHeaders, headerName)
		if err != nil {
			return nil, err
		}
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
			jsonValue, value, err := readCellJSONValue(f, sheetName, cell)
			if err != nil {
				return nil, err
			}
			if value != "" {
				empty = false
			}
			item[header] = jsonValue
			values = append(values, jsonComparisonValue(jsonValue, value))
		}
		if empty {
			continue
		}
		if matchesAllFilters(values, resolvedFilters) {
			result.Rows = append(result.Rows, item)
		}
	}

	return result, nil
}

type resolvedFilter struct {
	Index    int
	Operator string
	Value    string
	Regex    *regexp.Regexp
	Number   float64
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
		item := resolvedFilter{Operator: operator, Value: filter.Value}
		if operator == "regex" {
			compiled, err := regexp.Compile(filter.Value)
			if err != nil {
				return nil, fmt.Errorf("invalid regex value: %w", err)
			}
			item.Regex = compiled
		}
		if operator == "gt" || operator == "gte" || operator == "lt" || operator == "lte" {
			number, ok := parseFiniteFloat(filter.Value)
			if !ok {
				return nil, fmt.Errorf("operator %q requires a numeric value", operator)
			}
			item.Number = number
		}
		if columnIndex, err := strconv.Atoi(column); err == nil {
			if columnIndex < 1 || columnIndex > columnCount {
				return nil, fmt.Errorf("column index %d is out of range", columnIndex)
			}
			item.Index = columnIndex - 1
			resolved = append(resolved, item)
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
		item.Index = matched
		resolved = append(resolved, item)
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

func matchesAllFilters(values []string, filters []resolvedFilter) bool {
	for _, filter := range filters {
		if !matchFilterValue(values[filter.Index], filter) {
			return false
		}
	}
	return true
}

func matchFilterValue(actual string, filter resolvedFilter) bool {
	switch filter.Operator {
	case "equals":
		return strings.EqualFold(strings.TrimSpace(actual), strings.TrimSpace(filter.Value))
	case "contains":
		return strings.Contains(strings.ToLower(actual), strings.ToLower(filter.Value))
	case "regex":
		return filter.Regex.MatchString(actual)
	case "gt", "gte", "lt", "lte":
		actualNumber, ok := parseFiniteFloat(actual)
		if !ok {
			return false
		}
		switch filter.Operator {
		case "gt":
			return actualNumber > filter.Number
		case "gte":
			return actualNumber >= filter.Number
		case "lt":
			return actualNumber < filter.Number
		default:
			return actualNumber <= filter.Number
		}
	default:
		return false
	}
}
