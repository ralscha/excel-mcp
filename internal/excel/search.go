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
	matcher, err := buildMatcher(opts.Query, matchMode, opts.CaseSensitive)
	if err != nil {
		return nil, err
	}
	maxResults := opts.MaxResults
	if maxResults <= 0 {
		maxResults = 50
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()

	sheets := opts.Sheets
	if len(sheets) == 0 {
		sheets = f.GetSheetList()
	}
	result := &FindResult{FilePath: path, Query: opts.Query, SearchType: searchType, MatchMode: matchMode}
	for _, sheetName := range sheets {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return nil, err
		}
		rows, cols, _, err := worksheetBounds(f, sheetName)
		if err != nil {
			return nil, err
		}
		for row := 1; row <= rows; row++ {
			for col := 1; col <= cols; col++ {
				cell, err := excelize.CoordinatesToCellName(col, row)
				if err != nil {
					return nil, err
				}
				match := FindMatch{SheetName: sheetName, Cell: cell}
				var haystack string
				if searchType == "formula" {
					haystack, err = f.GetCellFormula(sheetName, cell)
					if err != nil {
						return nil, err
					}
					match.Formula = haystack
				} else {
					haystack, err = f.GetCellValue(sheetName, cell)
					if err != nil {
						return nil, err
					}
					match.Value = haystack
				}
				if haystack == "" || !matcher(haystack) {
					continue
				}
				if opts.ContextRows > 0 || opts.ContextCols > 0 {
					match.Context, err = collectContext(f, sheetName, row, col, rows, cols, opts.ContextRows, opts.ContextCols)
					if err != nil {
						return nil, err
					}
				}
				result.Matches = append(result.Matches, match)
				if len(result.Matches) >= maxResults {
					return result, nil
				}
			}
		}
	}
	return result, nil
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
	return func(haystack string) bool {
		return matchesSearch(haystack, query, matchMode, caseSensitive)
	}, nil
}

func matchesSearch(haystack, needle, matchMode string, caseSensitive bool) bool {
	if !caseSensitive {
		haystack = strings.ToLower(haystack)
		needle = strings.ToLower(needle)
	}
	if matchMode == "exact" {
		return haystack == needle
	}
	return strings.Contains(haystack, needle)
}

func collectContext(f *excelize.File, sheetName string, row, col, maxRow, maxCol, contextRows, contextCols int) ([]ContextCell, error) {
	startRow := max(1, row-contextRows)
	endRow := min(maxRow, row+contextRows)
	startCol := max(1, col-contextCols)
	endCol := min(maxCol, col+contextCols)
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
