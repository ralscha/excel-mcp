package excel

import (
	"archive/zip"
	"bytes"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

const metadataChartFixture = `<q:chartSpace xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:text="http://schemas.openxmlformats.org/drawingml/2006/main">
<q:chart>
 <q:title><q:tx><q:rich><text:p><text:r><text:t>Sales &amp; </text:t></text:r><text:r><text:t>Revenue</text:t></text:r></text:p></q:rich></q:tx></q:title>
 <q:plotArea>
  <q:barChart><q:barDir val="col"/><q:ser>
   <q:tx><q:v>Literal &amp; series</q:v></q:tx>
   <q:cat><q:strRef><q:f>Sheet1!$B$1:$C$1</q:f></q:strRef></q:cat>
   <q:val><q:numRef><q:f>Sheet1!$B$2:$C$2</q:f></q:numRef></q:val>
  </q:ser></q:barChart>
  <q:catAx><q:title><q:tx><q:rich><text:p><text:r><text:t>Months</text:t></text:r></text:p></q:rich></q:tx></q:title></q:catAx>
  <q:valAx><q:title><q:tx><q:rich><text:p><text:r><text:t>USD</text:t></text:r></text:p></q:rich></q:tx></q:title></q:valAx>
 </q:plotArea>
</q:chart></q:chartSpace>`

func TestChartMetadataUsesStructureAndNamespaceIndependentNames(t *testing.T) {
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	chart, err := parseChartMetadata(f, metadataChartFixture)
	if err != nil {
		t.Fatal(err)
	}
	if chart.Title != "Sales & Revenue" {
		t.Fatalf("chart title contains axis titles: %q", chart.Title)
	}
	if chart.ChartType != "column" || len(chart.Series) != 1 {
		t.Fatalf("unexpected chart metadata: %+v", chart)
	}
	series := chart.Series[0]
	if series.DisplayName != "Literal & series" || series.NameRef != "" || series.CategoriesRange != "Sheet1!$B$1:$C$1" || series.ValuesRange != "Sheet1!$B$2:$C$2" || series.SourceRange != "Sheet1!B1:C2" {
		t.Fatalf("unexpected series: %+v", series)
	}
}

func TestDrawingChartDoesNotUsePrecedingPictureAnchor(t *testing.T) {
	drawing := `<d:wsDr xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <d:twoCellAnchor><d:from><d:col>0</d:col><d:row>0</d:row></d:from><d:pic/></d:twoCellAnchor>
 <d:oneCellAnchor><d:from><d:col>4</d:col><d:row>8</d:row></d:from><d:graphicFrame><a:graphic><a:graphicData><q:chart r:id="chart1"/></a:graphicData></a:graphic></d:graphicFrame></d:oneCellAnchor>
</d:wsDr>`
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	for name, content := range map[string]string{
		"xl/drawings/drawing1.xml":            drawing,
		"xl/drawings/_rels/drawing1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart1" Target="../charts/chart1.xml"/></Relationships>`,
		"xl/charts/chart1.xml":                metadataChartFixture,
	} {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write([]byte(content)); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	reader, err := zip.NewReader(bytes.NewReader(buffer.Bytes()), int64(buffer.Len()))
	if err != nil {
		t.Fatal(err)
	}
	entries := make(map[string]*zip.File, len(reader.File))
	for _, entry := range reader.File {
		entries[entry.Name] = entry
	}
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	charts := drawingAnchorsXML(f, "xl/worksheets/sheet1.xml", map[string]string{"drawing1": "../drawings/drawing1.xml"}, entries)
	if len(charts) != 1 || charts[0].AnchorCell != "E9" || charts[0].Title != "Sales & Revenue" {
		t.Fatalf("unexpected drawing chart metadata: %+v", charts)
	}
}

func TestChartMetadataScatterAndCachedSeriesNames(t *testing.T) {
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	chart, err := parseChartMetadata(f, `<chartSpace><chart><plotArea><scatterChart><ser>
<tx><strRef><f>Missing!A1</f><strCache><pt idx="0"><v>Cached label</v></pt></strCache></strRef></tx>
<xVal><numRef><f>Sheet1!A2:A3</f></numRef></xVal><yVal><numRef><f>Sheet1!B2:B3</f></numRef></yVal>
</ser></scatterChart></plotArea></chart></chartSpace>`)
	if err != nil {
		t.Fatal(err)
	}
	if chart.ChartType != "scatter" || len(chart.Series) != 1 || chart.Series[0].DisplayName != "Cached label" || chart.Series[0].CategoriesRange != "Sheet1!A2:A3" || chart.Series[0].ValuesRange != "Sheet1!B2:B3" {
		t.Fatalf("unexpected scatter metadata: %+v", chart)
	}
}

func TestListChartsFiltersMixedSourcesAndResolvesSheetCase(t *testing.T) {
	path := filepath.Join(t.TempDir(), "mixed-sources.xlsx")
	f := excelize.NewFile()
	if _, err := f.NewSheet("Labels"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Labels", "A1", "Revenue"); err != nil {
		t.Fatal(err)
	}
	if err := f.AddChart("Sheet1", "D1", &excelize.Chart{
		Type:   excelize.Line,
		Series: []excelize.ChartSeries{{Name: "Labels!$A$1", Categories: "Sheet1!$A$2:$A$3", Values: "Sheet1!$B$2:$B$3"}},
	}); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := f.Close(); err != nil {
		t.Fatal(err)
	}
	for _, source := range []string{"sheet1", "labels"} {
		result, err := ListCharts(path, "sheet1", source)
		if err != nil {
			t.Fatal(err)
		}
		if len(result.Charts) != 1 || result.Charts[0].SheetName != "Sheet1" || result.Charts[0].Series[0].DisplayName != "Revenue" {
			t.Fatalf("source %s: unexpected charts: %+v", source, result.Charts)
		}
	}
}
