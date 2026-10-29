package excel

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"sort"
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
	var workbookDocument struct {
		Sheets []struct {
			Name  string `xml:"name,attr"`
			RelID string `xml:"id,attr"`
		} `xml:"sheets>sheet"`
	}
	if err := xml.Unmarshal([]byte(workbookXML), &workbookDocument); err != nil {
		return nil, fmt.Errorf("parse workbook metadata: %w", err)
	}
	sheetRelByName := make(map[string]string, len(workbookDocument.Sheets))
	for _, sheet := range workbookDocument.Sheets {
		sheetRelByName[sheet.Name] = sheet.RelID
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
	drawingTargets := make([]string, 0, len(sheetTargets))
	for _, drawingTarget := range sheetTargets {
		if !strings.Contains(drawingTarget, "drawings/") {
			continue
		}
		drawingTargets = append(drawingTargets, drawingTarget)
	}
	sort.Strings(drawingTargets)
	for _, drawingTarget := range drawingTargets {
		drawingPath := normalizeZipTarget(path.Dir(sheetPath), drawingTarget)
		drawingXML, err := readZipText(entries, drawingPath)
		if err != nil {
			continue
		}
		drawingTargets := relTargets(mustZipText(entries, relsPathFor(drawingPath)))
		var drawing struct {
			Anchors []drawingAnchorXML `xml:",any"`
		}
		if err := xml.Unmarshal([]byte(drawingXML), &drawing); err != nil {
			continue
		}
		for _, anchor := range drawing.Anchors {
			if anchor.Chart.RelID == "" {
				continue
			}
			chartPath := normalizeZipTarget(path.Dir(drawingPath), drawingTargets[anchor.Chart.RelID])
			chartXML, err := readZipText(entries, chartPath)
			if err != nil {
				continue
			}
			chart, err := parseChartMetadata(workbook, chartXML)
			if err != nil {
				continue
			}
			if anchor.From != nil {
				chart.AnchorCell = mustCellName(anchor.From.Col+1, anchor.From.Row+1)
			}
			chart.ChartPath = chartPath
			result = append(result, chart)
		}
	}
	return result
}

type drawingAnchorXML struct {
	From *struct {
		Col int `xml:"col"`
		Row int `xml:"row"`
	} `xml:"from"`
	Chart struct {
		RelID string `xml:"id,attr"`
	} `xml:"graphicFrame>graphic>graphicData>chart"`
}

func relTargets(relsXML string) map[string]string {
	result := map[string]string{}
	var relationships struct {
		Items []struct {
			ID     string `xml:"Id,attr"`
			Target string `xml:"Target,attr"`
		} `xml:"Relationship"`
	}
	if err := xml.Unmarshal([]byte(relsXML), &relationships); err != nil {
		return result
	}
	for _, relationship := range relationships.Items {
		result[relationship.ID] = strings.ReplaceAll(relationship.Target, `\`, "/")
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

type chartReferenceXML struct {
	Formula string   `xml:"f"`
	Cached  []string `xml:"strCache>pt>v"`
}

type chartSourceXML struct {
	StringRef chartReferenceXML `xml:"strRef"`
	NumberRef chartReferenceXML `xml:"numRef"`
}

func (source chartSourceXML) formula() string {
	if source.StringRef.Formula != "" {
		return source.StringRef.Formula
	}
	return source.NumberRef.Formula
}

type chartSeriesXML struct {
	Name struct {
		Reference chartReferenceXML `xml:"strRef"`
		Value     string            `xml:"v"`
	} `xml:"tx"`
	Categories chartSourceXML `xml:"cat"`
	Values     chartSourceXML `xml:"val"`
	XValues    chartSourceXML `xml:"xVal"`
	YValues    chartSourceXML `xml:"yVal"`
}

func parseChartMetadata(workbook *excelize.File, chartXML string) (ChartDescription, error) {
	var document struct {
		Chart struct {
			Title struct {
				Runs      []string          `xml:"tx>rich>p>r>t"`
				Reference chartReferenceXML `xml:"tx>strRef"`
			} `xml:"title"`
			PlotArea struct {
				Groups []struct {
					XMLName xml.Name
					BarDir  struct {
						Value string `xml:"val,attr"`
					} `xml:"barDir"`
					Series []chartSeriesXML `xml:"ser"`
				} `xml:",any"`
			} `xml:"plotArea"`
		} `xml:"chart"`
	}
	if err := xml.Unmarshal([]byte(chartXML), &document); err != nil {
		return ChartDescription{}, err
	}
	result := ChartDescription{
		Title:     strings.Join(document.Chart.Title.Runs, ""),
		ChartType: "unknown",
		Series:    make([]ChartSeriesDescription, 0),
	}
	if result.Title == "" {
		result.Title = resolveSeriesDisplayName(workbook, document.Chart.Title.Reference.Formula)
		if result.Title == "" {
			result.Title = strings.Join(document.Chart.Title.Reference.Cached, "")
		}
	}
	for _, group := range document.Chart.PlotArea.Groups {
		kind := ""
		switch group.XMLName.Local {
		case "lineChart":
			kind = "line"
		case "barChart":
			kind = "bar"
			if group.BarDir.Value == "col" {
				kind = "column"
			}
		case "pieChart":
			kind = "pie"
		case "doughnutChart":
			kind = "doughnut"
		case "scatterChart":
			kind = "scatter"
		case "areaChart":
			kind = "area"
		}
		if kind != "" && result.ChartType == "unknown" {
			result.ChartType = kind
		}
		for _, series := range group.Series {
			item := ChartSeriesDescription{
				NameRef:         series.Name.Reference.Formula,
				CategoriesRange: series.Categories.formula(),
				ValuesRange:     series.Values.formula(),
			}
			if item.CategoriesRange == "" {
				item.CategoriesRange = series.XValues.formula()
			}
			if item.ValuesRange == "" {
				item.ValuesRange = series.YValues.formula()
			}
			item.DisplayName = resolveSeriesDisplayName(workbook, item.NameRef)
			if item.DisplayName == "" {
				item.DisplayName = series.Name.Value
				if item.DisplayName == "" && len(series.Name.Reference.Cached) > 0 {
					item.DisplayName = series.Name.Reference.Cached[0]
				}
			}
			item.SourceSheet, item.SourceRange = normalizeSeriesSource(item.NameRef, item.CategoriesRange, item.ValuesRange)
			result.Series = append(result.Series, item)
		}
	}
	return result, nil
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
	quotedSheet := quoteSheetName(sheetName)
	if endCell == "" || startCell == endCell {
		return quotedSheet + "!" + startCell
	}
	return quotedSheet + "!" + startCell + ":" + endCell
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
			Name:       formatSheetRangeRef(sheetName, absoluteCellRef(nameCell), ""),
			Categories: formatSheetRangeRef(sheetName, absoluteCellRef(catStart), absoluteCellRef(catEnd)),
			Values:     formatSheetRangeRef(sheetName, absoluteCellRef(valStart), absoluteCellRef(valEnd)),
		})
	}
	chart := &excelize.Chart{Type: chartKind, Series: series}
	if opts.Title != "" {
		chart.Title = excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: opts.Title}}}
	}
	if opts.XAxis != "" {
		chart.XAxis.Title = excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: opts.XAxis}}}
	}
	if opts.YAxis != "" {
		chart.YAxis.Title = excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: opts.YAxis}}}
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
