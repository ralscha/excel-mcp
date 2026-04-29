package excel

import (
	"archive/zip"
	"fmt"
	"io"
	"path"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func describeCharts(workbookPath string, workbook *excelize.File) (map[string][]ChartDescription, error) {
	reader, err := zip.OpenReader(workbookPath)
	if err != nil {
		return nil, fmt.Errorf("open workbook zip: %w", err)
	}
	defer func() { _ = reader.Close() }()
	entries := map[string]*zip.File{}
	for _, entry := range reader.File {
		entries[entry.Name] = entry
	}
	workbookXML, err := readZipText(entries, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}
	workbookRelsXML, err := readZipText(entries, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}
	sheetRelByName := map[string]string{}
	for _, match := range regexp.MustCompile(`<sheet[^>]*name="([^"]+)"[^>]*(?:r:id|relationships:id)="([^"]+)"`).FindAllStringSubmatch(workbookXML, -1) {
		sheetRelByName[match[1]] = match[2]
	}
	workbookTargets := relTargets(workbookRelsXML)
	result := map[string][]ChartDescription{}
	for sheetName, relID := range sheetRelByName {
		sheetPath := normalizeZipTarget("xl", workbookTargets[relID])
		if sheetPath == "" {
			continue
		}
		sheetRelsXML, err := readZipTextOptional(entries, relsPathFor(sheetPath))
		if err != nil {
			return nil, err
		}
		if sheetRelsXML == "" {
			continue
		}
		result[sheetName] = append(result[sheetName], drawingAnchorsXML(workbook, sheetPath, relTargets(sheetRelsXML), entries)...)
	}
	return result, nil
}

func drawingAnchorsXML(workbook *excelize.File, sheetPath string, sheetTargets map[string]string, entries map[string]*zip.File) []ChartDescription {
	var result []ChartDescription
	for _, drawingTarget := range sheetTargets {
		if !strings.Contains(drawingTarget, "drawings/") {
			continue
		}
		drawingPath := normalizeZipTarget(path.Dir(sheetPath), drawingTarget)
		drawingXML, err := readZipText(entries, drawingPath)
		if err != nil {
			continue
		}
		drawingTargets := relTargets(mustZipText(entries, relsPathFor(drawingPath)))
		for _, anchorMatch := range regexp.MustCompile(`(?s)<xdr:(?:oneCellAnchor|twoCellAnchor).*?<xdr:from>.*?<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>.*?</xdr:from>.*?<c:chart[^>]*r:id="([^"]+)"`).FindAllStringSubmatch(drawingXML, -1) {
			colZero, _ := strconv.Atoi(anchorMatch[1])
			rowZero, _ := strconv.Atoi(anchorMatch[2])
			chartPath := normalizeZipTarget(path.Dir(drawingPath), drawingTargets[anchorMatch[3]])
			chartXML, err := readZipText(entries, chartPath)
			if err != nil {
				continue
			}
			result = append(result, ChartDescription{
				AnchorCell: mustCellName(colZero+1, rowZero+1),
				Title:      chartTitle(chartXML),
				ChartType:  chartType(chartXML),
				ChartPath:  chartPath,
				Series:     chartSeries(workbook, chartXML),
			})
		}
	}
	return result
}

func relTargets(relsXML string) map[string]string {
	result := map[string]string{}
	for _, match := range regexp.MustCompile(`<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"`).FindAllStringSubmatch(relsXML, -1) {
		result[match[1]] = strings.ReplaceAll(match[2], `\`, "/")
	}
	return result
}

func relsPathFor(target string) string {
	dir := path.Dir(target)
	base := path.Base(target)
	if dir == "." {
		return "_rels/" + base + ".rels"
	}
	return path.Join(dir, "_rels", base+".rels")
}

func normalizeZipTarget(baseDir, target string) string {
	if target == "" {
		return ""
	}
	if after, ok := strings.CutPrefix(target, "/"); ok {
		return after
	}
	return path.Clean(path.Join(baseDir, target))
}

func readZipText(entries map[string]*zip.File, name string) (string, error) {
	s, err := readZipTextOptional(entries, name)
	if err != nil {
		return "", err
	}
	if s == "" {
		return "", fmt.Errorf("zip entry %s not found", name)
	}
	return s, nil
}

func readZipTextOptional(entries map[string]*zip.File, name string) (string, error) {
	entry := entries[name]
	if entry == nil {
		return "", nil
	}
	reader, err := entry.Open()
	if err != nil {
		return "", err
	}
	defer func() { _ = reader.Close() }()
	data, err := io.ReadAll(reader)
	if err != nil {
		return "", err
	}
	return string(data), nil
}

func mustZipText(entries map[string]*zip.File, name string) string {
	text, _ := readZipTextOptional(entries, name)
	return text
}

func chartTitle(chartXML string) string {
	matches := regexp.MustCompile(`<a:t>(.*?)</a:t>`).FindAllStringSubmatch(chartXML, -1)
	var builder strings.Builder
	for _, match := range matches {
		builder.WriteString(match[1])
	}
	return builder.String()
}

func chartSeries(workbook *excelize.File, chartXML string) []ChartSeriesDescription {
	matches := regexp.MustCompile(`(?s)<(?:c:)?ser>.*?<(?:c:)?tx>.*?<f>(.*?)</f>.*?</(?:c:)?tx>.*?<(?:c:)?cat>.*?<f>(.*?)</f>.*?</(?:c:)?cat>.*?<(?:c:)?val>.*?<f>(.*?)</f>.*?</(?:c:)?val>.*?</(?:c:)?ser>`).FindAllStringSubmatch(chartXML, -1)
	series := make([]ChartSeriesDescription, 0, len(matches))
	for _, match := range matches {
		item := ChartSeriesDescription{NameRef: match[1], CategoriesRange: match[2], ValuesRange: match[3]}
		item.DisplayName = resolveSeriesDisplayName(workbook, item.NameRef)
		item.SourceSheet, item.SourceRange = normalizeSeriesSource(item.NameRef, item.CategoriesRange, item.ValuesRange)
		series = append(series, item)
	}
	return series
}

type sheetRangeRef struct {
	Sheet     string
	StartCell string
	EndCell   string
}

func resolveSeriesDisplayName(workbook *excelize.File, nameRef string) string {
	ref, ok := parseSheetRangeRef(nameRef)
	if !ok || ref.StartCell != ref.EndCell {
		return ""
	}
	value, err := workbook.GetCellValue(ref.Sheet, ref.StartCell)
	if err != nil {
		return ""
	}
	return value
}

func normalizeSeriesSource(refs ...string) (string, string) {
	parsed := make([]sheetRangeRef, 0, len(refs))
	for _, ref := range refs {
		parsedRef, ok := parseSheetRangeRef(ref)
		if !ok {
			continue
		}
		parsed = append(parsed, parsedRef)
	}
	if len(parsed) == 0 {
		return "", ""
	}
	sheetName := parsed[0].Sheet
	minCol := 0
	minRow := 0
	maxCol := 0
	maxRow := 0
	for index, ref := range parsed {
		if ref.Sheet != sheetName {
			return "", ""
		}
		startCol, startRow, err := excelize.CellNameToCoordinates(ref.StartCell)
		if err != nil {
			return "", ""
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(ref.EndCell)
		if err != nil {
			return "", ""
		}
		if index == 0 || startCol < minCol {
			minCol = startCol
		}
		if index == 0 || startRow < minRow {
			minRow = startRow
		}
		if index == 0 || endCol > maxCol {
			maxCol = endCol
		}
		if index == 0 || endRow > maxRow {
			maxRow = endRow
		}
	}
	startCell, err := excelize.CoordinatesToCellName(minCol, minRow)
	if err != nil {
		return sheetName, ""
	}
	endCell, err := excelize.CoordinatesToCellName(maxCol, maxRow)
	if err != nil {
		return sheetName, ""
	}
	if startCell == endCell {
		return sheetName, formatSheetRangeRef(sheetName, startCell, "")
	}
	return sheetName, formatSheetRangeRef(sheetName, startCell, endCell)
}

func parseSheetRangeRef(ref string) (sheetRangeRef, bool) {
	ref = strings.TrimSpace(strings.TrimPrefix(ref, "="))
	separator := strings.LastIndex(ref, "!")
	if separator <= 0 || separator >= len(ref)-1 {
		return sheetRangeRef{}, false
	}
	sheetName := parseSheetName(ref[:separator])
	cellRef := ref[separator+1:]
	parts := strings.Split(cellRef, ":")
	if len(parts) > 2 || len(parts) == 0 {
		return sheetRangeRef{}, false
	}
	startCell := normalizeCellRef(parts[0])
	endCell := startCell
	if len(parts) == 2 {
		endCell = normalizeCellRef(parts[1])
	}
	if _, _, err := excelize.CellNameToCoordinates(startCell); err != nil {
		return sheetRangeRef{}, false
	}
	if _, _, err := excelize.CellNameToCoordinates(endCell); err != nil {
		return sheetRangeRef{}, false
	}
	return sheetRangeRef{Sheet: sheetName, StartCell: startCell, EndCell: endCell}, true
}

func parseSheetName(value string) string {
	if len(value) >= 2 && strings.HasPrefix(value, "'") && strings.HasSuffix(value, "'") {
		return strings.ReplaceAll(value[1:len(value)-1], "''", "'")
	}
	return value
}

func normalizeCellRef(value string) string {
	return strings.ReplaceAll(strings.TrimSpace(value), "$", "")
}

func formatSheetRangeRef(sheetName, startCell, endCell string) string {
	quotedSheet := sheetName
	if strings.ContainsAny(sheetName, " '!") {
		quotedSheet = "'" + strings.ReplaceAll(sheetName, "'", "''") + "'"
	}
	if endCell == "" || startCell == endCell {
		return quotedSheet + "!" + startCell
	}
	return quotedSheet + "!" + startCell + ":" + endCell
}

func chartType(chartXML string) string {
	if strings.Contains(chartXML, `<c:lineChart`) || strings.Contains(chartXML, `<lineChart`) {
		return "line"
	}
	if strings.Contains(chartXML, `<c:barChart`) || strings.Contains(chartXML, `<barChart`) {
		barDir := regexp.MustCompile(`<(?:c:)?barDir[^>]*val="([^"]+)"`).FindStringSubmatch(chartXML)
		if len(barDir) == 2 && barDir[1] == "col" {
			return "column"
		}
		return "bar"
	}
	if strings.Contains(chartXML, `<c:pieChart`) || strings.Contains(chartXML, `<pieChart`) {
		return "pie"
	}
	if strings.Contains(chartXML, `<c:doughnutChart`) || strings.Contains(chartXML, `<doughnutChart`) {
		return "doughnut"
	}
	if strings.Contains(chartXML, `<c:scatterChart`) || strings.Contains(chartXML, `<scatterChart`) {
		return "scatter"
	}
	if strings.Contains(chartXML, `<c:areaChart`) || strings.Contains(chartXML, `<areaChart`) {
		return "area"
	}
	return "unknown"
}

func mustCellName(col, row int) string {
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return ""
	}
	return cell
}

func CreateChart(path, sheetName, dataRange, chartType, targetCell string, opts ChartOptions) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		chart, err := chartFromRange(sheetName, dataRange, chartType, opts)
		if err != nil {
			return err
		}
		return f.AddChart(sheetName, targetCell, chart)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("created %s chart at %s!%s", chartType, sheetName, targetCell), nil
}

func chartFromRange(sheetName, dataRange, chartType string, opts ChartOptions) (*excelize.Chart, error) {
	_, _, startCol, startRow, endCol, endRow, err := normalizeRangeRef(dataRange)
	if err != nil {
		return nil, fmt.Errorf("invalid data_range: %w", err)
	}
	if endRow-startRow < 1 || endCol-startCol < 1 {
		return nil, fmt.Errorf("data_range must include headers and at least one series")
	}
	chartKind, err := chartTypeValue(chartType)
	if err != nil {
		return nil, err
	}
	series := make([]excelize.ChartSeries, 0, endRow-startRow)
	for row := startRow + 1; row <= endRow; row++ {
		nameCell, _ := excelize.CoordinatesToCellName(startCol, row)
		catStart, _ := excelize.CoordinatesToCellName(startCol+1, startRow)
		catEnd, _ := excelize.CoordinatesToCellName(endCol, startRow)
		valStart, _ := excelize.CoordinatesToCellName(startCol+1, row)
		valEnd, _ := excelize.CoordinatesToCellName(endCol, row)
		series = append(series, excelize.ChartSeries{
			Name:       fmt.Sprintf("%s!%s", sheetName, absoluteCellRef(nameCell)),
			Categories: fmt.Sprintf("%s!%s:%s", sheetName, absoluteCellRef(catStart), absoluteCellRef(catEnd)),
			Values:     fmt.Sprintf("%s!%s:%s", sheetName, absoluteCellRef(valStart), absoluteCellRef(valEnd)),
		})
	}
	chart := &excelize.Chart{Type: chartKind, Series: series}
	if opts.Title != "" {
		chart.Title = []excelize.RichTextRun{{Text: opts.Title}}
	}
	if opts.XAxis != "" {
		chart.XAxis.Title = []excelize.RichTextRun{{Text: opts.XAxis}}
	}
	if opts.YAxis != "" {
		chart.YAxis.Title = []excelize.RichTextRun{{Text: opts.YAxis}}
	}
	return chart, nil
}

func chartTypeValue(chartType string) (excelize.ChartType, error) {
	switch strings.ToLower(strings.TrimSpace(chartType)) {
	case "line":
		return excelize.Line, nil
	case "column", "col":
		return excelize.Col, nil
	case "bar":
		return excelize.Bar, nil
	case "pie":
		return excelize.Pie, nil
	case "doughnut", "donut":
		return excelize.Doughnut, nil
	case "scatter":
		return excelize.Scatter, nil
	case "area":
		return excelize.Area, nil
	default:
		return 0, fmt.Errorf("unsupported chart_type %q; supported values are line, column, bar, pie, doughnut, scatter, area", chartType)
	}
}
