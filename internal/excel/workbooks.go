package excel

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func CreateWorkbook(path string) (string, error) {
	if err := ensureParentDir(path); err != nil {
		return "", err
	}
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	if err := f.SaveAs(path); err != nil {
		return "", fmt.Errorf("save workbook: %w", err)
	}
	return fmt.Sprintf("created workbook at %s", path), nil
}

func CreateWorksheet(path, sheetName string) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		_, err := f.NewSheet(sheetName)
		return err
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("created worksheet %q in %s", sheetName, path), nil
}

func CopyWorksheet(path, sourceSheet, targetSheet string) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sourceSheet); err != nil {
			return err
		}
		sourceIndex, err := f.GetSheetIndex(sourceSheet)
		if err != nil {
			return err
		}
		newIndex, err := f.NewSheet(targetSheet)
		if err != nil {
			return err
		}
		return f.CopySheet(sourceIndex, newIndex)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("copied worksheet %q to %q in %s", sourceSheet, targetSheet, path), nil
}

func DeleteWorksheet(path, sheetName string) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		return f.DeleteSheet(sheetName)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("deleted worksheet %q in %s", sheetName, path), nil
}

func RenameWorksheet(path, oldName, newName string) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, oldName); err != nil {
			return err
		}
		return f.SetSheetName(oldName, newName)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("renamed worksheet %q to %q in %s", oldName, newName, path), nil
}

func GetWorkbookMetadata(path string, includeRanges bool) (*WorkbookMetadata, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()

	metadata := &WorkbookMetadata{FilePath: path}
	for _, sheet := range f.GetSheetList() {
		item := SheetMetadata{Name: sheet}
		if includeRanges {
			rows, cols, sheetRange, err := worksheetBounds(f, sheet)
			if err != nil {
				return nil, err
			}
			item.Rows = rows
			item.Cols = cols
			item.Range = sheetRange
		}
		metadata.Sheets = append(metadata.Sheets, item)
	}
	return metadata, nil
}

func DescribeWorkbook(path string, opts DescribeWorkbookOptions) (*WorkbookDescription, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()

	description := &WorkbookDescription{FilePath: path}
	chartMap := map[string][]ChartDescription{}
	if opts.IncludeCharts {
		chartMap, err = describeCharts(path, f)
		if err != nil {
			return nil, err
		}
	}
	if opts.IncludeNames {
		for _, item := range f.GetDefinedName() {
			description.NamedRanges = append(description.NamedRanges, NamedRangeDescription{
				Name:     item.Name,
				RefersTo: strings.TrimPrefix(item.RefersTo, "="),
				Scope:    item.Scope,
				Comment:  item.Comment,
			})
		}
	}
	for _, sheetName := range f.GetSheetList() {
		sheet := WorkbookSheetDescription{Name: sheetName}
		if opts.IncludeRanges {
			rows, cols, sheetRange, err := worksheetBounds(f, sheetName)
			if err != nil {
				return nil, err
			}
			sheet.Rows = rows
			sheet.Cols = cols
			sheet.Range = sheetRange
		}
		if opts.IncludeMerged {
			merged, err := getMergedCellsFromFile(f, sheetName)
			if err != nil {
				return nil, err
			}
			sheet.MergedRanges = merged
		}
		if opts.IncludeTables {
			tables, err := f.GetTables(sheetName)
			if err != nil {
				return nil, err
			}
			for _, table := range tables {
				sheet.Tables = append(sheet.Tables, TableDescription{Name: table.Name, Range: table.Range, StyleName: table.StyleName})
			}
		}
		if opts.IncludeCharts {
			sheet.Charts = chartMap[sheetName]
		}
		if opts.IncludePivots {
			pivots, err := f.GetPivotTables(sheetName)
			if err != nil {
				return nil, err
			}
			for _, pivot := range pivots {
				sheet.PivotTables = append(sheet.PivotTables, pivotDescription(pivot))
			}
		}
		if opts.IncludeValidation {
			validations, err := getDataValidationInfoFromFile(f, sheetName)
			if err != nil {
				return nil, err
			}
			sheet.DataValidations = validations
		}
		description.Sheets = append(description.Sheets, sheet)
	}
	return description, nil
}

func ListCharts(path, sheetName, sourceSheet string) (*ListChartsResult, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()
	if sheetName != "" {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return nil, err
		}
	}
	if sourceSheet != "" {
		if err := ensureSheetExists(f, sourceSheet); err != nil {
			return nil, err
		}
	}
	charts, err := describeCharts(path, f)
	if err != nil {
		return nil, err
	}
	result := &ListChartsResult{FilePath: path, SheetName: sheetName, SourceSheet: sourceSheet}
	appendChart := func(hostSheet string, chart ChartDescription) {
		if sourceSheet != "" && !chartUsesSourceSheet(chart, sourceSheet) {
			return
		}
		chart.SheetName = hostSheet
		result.Charts = append(result.Charts, chart)
	}
	if sheetName != "" {
		for _, chart := range charts[sheetName] {
			appendChart(sheetName, chart)
		}
		return result, nil
	}
	for _, hostSheet := range f.GetSheetList() {
		for _, chart := range charts[hostSheet] {
			appendChart(hostSheet, chart)
		}
	}
	return result, nil
}

func chartUsesSourceSheet(chart ChartDescription, sheetName string) bool {
	for _, series := range chart.Series {
		if series.SourceSheet == sheetName {
			return true
		}
	}
	return false
}

func GetSheetSchema(path, sheetName, startCell, endCell string, opts SheetSchemaOptions) (*SheetSchema, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()

	if err := ensureSheetExists(f, sheetName); err != nil {
		return nil, err
	}

	startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		return nil, fmt.Errorf("invalid start_cell: %w", err)
	}
	var endCol, endRow int
	if endCell != "" {
		endCol, endRow, err = excelize.CellNameToCoordinates(endCell)
		if err != nil {
			return nil, fmt.Errorf("invalid end_cell: %w", err)
		}
	} else {
		rows, cols, _, err := worksheetBounds(f, sheetName)
		if err != nil {
			return nil, err
		}
		endCol = cols
		endRow = rows
	}
	if endCol < startCol || endRow < startRow {
		return nil, fmt.Errorf("range end must be below and to the right of start_cell")
	}

	headerRow := opts.HeaderRow
	if headerRow == 0 {
		headerRow = startRow
	}
	if headerRow < startRow || headerRow > endRow {
		return nil, fmt.Errorf("header_row must be within the selected range")
	}
	sampleSize := opts.SampleSize
	if sampleSize <= 0 {
		sampleSize = 3
	}

	rangeEnd, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return nil, err
	}
	schema := &SheetSchema{
		FilePath:  path,
		SheetName: sheetName,
		Range:     startCell + ":" + rangeEnd,
		HeaderRow: headerRow,
		RowCount:  endRow - headerRow,
	}

	for col := startCol; col <= endCol; col++ {
		headerCell, err := excelize.CoordinatesToCellName(col, headerRow)
		if err != nil {
			return nil, err
		}
		headerValue, err := f.GetCellValue(sheetName, headerCell)
		if err != nil {
			return nil, err
		}
		headerValue = strings.TrimSpace(headerValue)
		if headerValue == "" {
			headerValue = fmt.Sprintf("column_%d", col-startCol+1)
		}
		columnName, err := excelize.ColumnNumberToName(col)
		if err != nil {
			return nil, err
		}
		column := SheetColumnSchema{Name: headerValue, Column: columnName}
		observedKinds := map[string]int{}
		for row := headerRow + 1; row <= endRow; row++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return nil, err
			}
			value, err := f.GetCellValue(sheetName, cell)
			if err != nil {
				return nil, err
			}
			value = strings.TrimSpace(value)
			if value == "" {
				column.BlankCount++
				continue
			}
			kind := inferCellKind(value)
			observedKinds[kind]++
			if len(column.SampleValues) < sampleSize {
				column.SampleValues = append(column.SampleValues, value)
			}
		}
		column.InferredType = dominantKind(observedKinds)
		schema.Columns = append(schema.Columns, column)
	}

	return schema, nil
}

func pivotDescription(value excelize.PivotTableOptions) PivotTableDescription {
	description := PivotTableDescription{
		Name:            value.Name,
		DataRange:       value.DataRange,
		PivotTableRange: value.PivotTableRange,
		StyleName:       value.PivotTableStyleName,
	}
	for _, item := range value.Rows {
		description.Rows = append(description.Rows, item.Data)
	}
	for _, item := range value.Columns {
		description.Columns = append(description.Columns, item.Data)
	}
	for _, item := range value.Data {
		description.Values = append(description.Values, PivotValueDescription{Data: item.Data, Name: item.Name, Subtotal: item.Subtotal})
	}
	return description
}
