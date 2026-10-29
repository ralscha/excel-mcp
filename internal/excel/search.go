package excel

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

func FindInWorkbook(path string, opts FindOptions) (*FindResult, error) {
	if strings.TrimSpace(opts.Query) == "" {
		return nil, fmt.Errorf("query is required")
	}
	searchType := strings.ToLower(strings.TrimSpace(opts.SearchType))
	if searchType == "" {
		searchType = "text"
	}
	if searchType != "text" && searchType != "formula" {
		return nil, fmt.Errorf("search_type must be 'text' or 'formula'")
	}
	matchMode := strings.ToLower(strings.TrimSpace(opts.MatchMode))
	if matchMode == "" {
		matchMode = "contains"
	}
	if matchMode != "contains" && matchMode != "exact" && matchMode != "regex" {
		return nil, fmt.Errorf("match_mode must be 'contains', 'exact' or 'regex'")
	}
	if opts.ContextRows < 0 || opts.ContextCols < 0 {
		return nil, fmt.Errorf("context_rows and context_cols must be non-negative")
	}
	if opts.MaxResults < 0 {
		return nil, fmt.Errorf("max_results must be non-negative")
	}
	matcher, err := buildMatcher(opts.Query, matchMode, opts.CaseSensitive)
	if err != nil {
		return nil, err
	}
	maxResults := opts.MaxResults
	if maxResults <= 0 {
		maxResults = 50
	}

	f, closeWorkbook, err := openWorkbook(path)
	if err != nil {
		return nil, err
	}
	defer closeWorkbook()

	requestedSheets := opts.Sheets
	if len(requestedSheets) == 0 {
		requestedSheets = f.GetSheetList()
	}
	sheets := make([]string, 0, len(requestedSheets))
	seen := make(map[int]bool, len(requestedSheets))
	for _, sheetName := range requestedSheets {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return nil, err
		}
		index, err := f.GetSheetIndex(sheetName)
		if err != nil {
			return nil, err
		}
		if !seen[index] {
			seen[index] = true
			sheets = append(sheets, f.GetSheetName(index))
		}
	}
	result := &FindResult{FilePath: path, Query: opts.Query, SearchType: searchType, MatchMode: matchMode, Matches: []FindMatch{}}
	for _, sheetName := range sheets {
		if err := findInSheet(f, sheetName, opts, matcher, maxResults, result); err != nil {
			return nil, err
		}
		if result.Truncated {
			break
		}
	}
	return result, nil
}

func findInSheet(f *excelize.File, sheetName string, opts FindOptions, matcher func(string) bool, maxResults int, result *FindResult) error {
	var mergedRanges []mergedSearchRange
	if result.SearchType == "formula" {
		mergedCells, err := f.GetMergeCells(sheetName, true)
		if err != nil {
			return err
		}
		mergedRanges = make([]mergedSearchRange, 0, len(mergedCells))
		for _, merged := range mergedCells {
			_, _, startCol, startRow, endCol, endRow, err := normalizeRangeRef(merged.GetStartAxis() + ":" + merged.GetEndAxis())
			if err != nil {
				return err
			}
			mergedRanges = append(mergedRanges, mergedSearchRange{startCol, startRow, endCol, endRow})
		}
	}
	maxRow, maxCol := 0, 0
	includeContext := opts.ContextRows > 0 || opts.ContextCols > 0
	if includeContext {
		var err error
		maxRow, maxCol, _, err = worksheetBounds(f, sheetName)
		if err != nil {
			return err
		}
	}
	rows, err := f.Rows(sheetName)
	if err != nil {
		return err
	}
	defer func() { _ = rows.Close() }()
	row := 0
	for rows.Next() {
		row++
		values, err := rows.Columns()
		if err != nil {
			return err
		}
		for index, value := range values {
			col := index + 1
			if result.SearchType == "text" && (value == "" || !matcher(value)) {
				continue
			}
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return err
			}
			match := FindMatch{SheetName: sheetName, Cell: cell}
			if result.SearchType == "formula" {
				// GetCellFormula resolves every merged cell to its anchor. Only
				// report the anchor, even when a later value fills the row slice.
				if isMergedSearchContinuation(mergedRanges, row, col) {
					continue
				}
				formula, err := f.GetCellFormula(sheetName, cell)
				if err != nil {
					return err
				}
				if formula == "" {
					continue
				}
				formula = "=" + strings.TrimPrefix(formula, "=")
				if !matcher(formula) {
					continue
				}
				match.Formula = formula
			} else {
				match.Value = value
			}
			if len(result.Matches) == maxResults {
				result.Truncated = true
				return nil
			}
			if includeContext {
				match.Context, err = collectContext(f, sheetName, row, col, maxRow, maxCol, opts.ContextRows, opts.ContextCols)
				if err != nil {
					return err
				}
			}
			result.Matches = append(result.Matches, match)
		}
	}
	return rows.Error()
}

type mergedSearchRange struct {
	startCol, startRow, endCol, endRow int
}

func isMergedSearchContinuation(ranges []mergedSearchRange, row, col int) bool {
	for _, merged := range ranges {
		if row >= merged.startRow && row <= merged.endRow && col >= merged.startCol && col <= merged.endCol &&
			(row != merged.startRow || col != merged.startCol) {
			return true
		}
	}
	return false
}

func buildMatcher(query, matchMode string, caseSensitive bool) (func(string) bool, error) {
	if matchMode == "regex" {
		pattern := query
		if !caseSensitive {
			pattern = `(?i)` + pattern
		}
		re, err := regexp.Compile(pattern)
		if err != nil {
			return nil, fmt.Errorf("invalid regex query: %w", err)
		}
		return re.MatchString, nil
	}
	if !caseSensitive {
		query = strings.ToLower(query)
	}
	return func(haystack string) bool {
		if !caseSensitive {
			haystack = strings.ToLower(haystack)
		}
		if matchMode == "exact" {
			return haystack == query
		}
		return strings.Contains(haystack, query)
	}, nil
}

func collectContext(f *excelize.File, sheetName string, row, col, maxRow, maxCol, contextRows, contextCols int) ([]ContextCell, error) {
	startRow := max(1, row-contextRows)
	endRow := row + min(contextRows, maxRow-row)
	startCol := max(1, col-contextCols)
	endCol := col + min(contextCols, maxCol-col)
	context := make([]ContextCell, 0)
	for currentRow := startRow; currentRow <= endRow; currentRow++ {
		for currentCol := startCol; currentCol <= endCol; currentCol++ {
			if currentRow == row && currentCol == col {
				continue
			}
			cell, err := excelize.CoordinatesToCellName(currentCol, currentRow)
			if err != nil {
				return nil, err
			}
			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return nil, err
			}
			if value == "" {
				continue
			}
			context = append(context, ContextCell{Cell: cell, Value: value})
		}
	}
	return context, nil
}
