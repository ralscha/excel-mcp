package excel

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func FormatRange(path, sheetName, startCell, endCell string, opts FormatRangeOptions) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		style := newCellStyle(
			opts.Bold,
			opts.Italic,
			opts.Underline,
			opts.FontSize,
			opts.FontColor,
			opts.BGColor,
			opts.BorderStyle,
			opts.BorderColor,
			opts.NumberFormat,
			opts.CustomNumFmt,
			opts.Alignment,
			opts.WrapText,
			opts.Protection,
		)
		styleID, err := f.NewStyle(style)
		if err != nil {
			return err
		}
		if err := f.SetCellStyle(sheetName, startCell, endCell, styleID); err != nil {
			return err
		}
		if err := setConditionalFormat(f, sheetName, startCell, endCell, opts.Conditional); err != nil {
			return err
		}
		if opts.MergeCells {
			if err := f.MergeCell(sheetName, startCell, endCell); err != nil {
				return err
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("formatted range %s!%s:%s", sheetName, startCell, endCell), nil
}

func newCellStyle(
	bold, italic, underline bool,
	fontSize int,
	fontColor, bgColor, borderStyleValue, borderColor string,
	numberFormat int,
	customNumFmt, alignment string,
	wrapText bool,
	protection *ProtectionOptions,
) *excelize.Style {
	style := &excelize.Style{}
	if bold || italic || underline || fontSize > 0 || fontColor != "" {
		style.Font = &excelize.Font{
			Bold:      bold,
			Italic:    italic,
			Underline: condUnderline(underline),
			Size:      float64(fontSize),
			Color:     fontColor,
		}
	}
	if bgColor != "" {
		style.Fill = excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{bgColor}}
	}
	if borderStyleValue != "" {
		style.Border = []excelize.Border{{Type: "left", Color: borderColor, Style: borderStyle(borderStyleValue)}, {Type: "right", Color: borderColor, Style: borderStyle(borderStyleValue)}, {Type: "top", Color: borderColor, Style: borderStyle(borderStyleValue)}, {Type: "bottom", Color: borderColor, Style: borderStyle(borderStyleValue)}}
	}
	if numberFormat > 0 {
		style.NumFmt = numberFormat
	}
	if customNumFmt != "" {
		style.CustomNumFmt = &customNumFmt
	}
	if alignment != "" || wrapText {
		style.Alignment = &excelize.Alignment{Horizontal: alignment, WrapText: wrapText}
	}
	if protection != nil {
		style.Protection = &excelize.Protection{Locked: protection.Locked, Hidden: protection.Hidden}
	}
	return style
}

func setConditionalFormat(f *excelize.File, sheetName, startCell, endCell string, opts *ConditionalFormatOptions) error {
	if opts == nil {
		return nil
	}
	if strings.TrimSpace(opts.Type) == "" {
		return fmt.Errorf("conditional_format.type is required")
	}
	rangeRef := startCell
	if endCell != "" && endCell != startCell {
		rangeRef = startCell + ":" + endCell
	}
	rule := excelize.ConditionalFormatOptions{
		Type:           opts.Type,
		AboveAverage:   opts.AboveAverage,
		Percent:        opts.Percent,
		Criteria:       opts.Criteria,
		Value:          opts.Value,
		MinType:        opts.MinType,
		MidType:        opts.MidType,
		MaxType:        opts.MaxType,
		MinValue:       opts.MinValue,
		MidValue:       opts.MidValue,
		MaxValue:       opts.MaxValue,
		MinColor:       opts.MinColor,
		MidColor:       opts.MidColor,
		MaxColor:       opts.MaxColor,
		BarColor:       opts.BarColor,
		BarBorderColor: opts.BarBorderColor,
		BarDirection:   opts.BarDirection,
		BarOnly:        opts.BarOnly,
		BarSolid:       opts.BarSolid,
		IconStyle:      opts.IconStyle,
		ReverseIcons:   opts.ReverseIcons,
		IconsOnly:      opts.IconsOnly,
		StopIfTrue:     opts.StopIfTrue,
	}
	if opts.Format != nil {
		styleID, err := f.NewConditionalStyle(newCellStyle(
			opts.Format.Bold,
			opts.Format.Italic,
			opts.Format.Underline,
			opts.Format.FontSize,
			opts.Format.FontColor,
			opts.Format.BGColor,
			opts.Format.BorderStyle,
			opts.Format.BorderColor,
			opts.Format.NumberFormat,
			opts.Format.CustomNumFmt,
			opts.Format.Alignment,
			opts.Format.WrapText,
			opts.Format.Protection,
		))
		if err != nil {
			return err
		}
		rule.Format = &styleID
	}
	return f.SetConditionalFormat(sheetName, rangeRef, []excelize.ConditionalFormatOptions{rule})
}

func condUnderline(enabled bool) string {
	if enabled {
		return "single"
	}
	return ""
}

func borderStyle(value string) int {
	switch strings.ToLower(value) {
	case "dashed":
		return 3
	case "dotted":
		return 4
	case "double":
		return 6
	default:
		return 1
	}
}

func MergeCells(path, sheetName, startCell, endCell string) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		return f.MergeCell(sheetName, startCell, endCell)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("merged %s!%s:%s", sheetName, startCell, endCell), nil
}

func UnmergeCells(path, sheetName, startCell, endCell string) (string, error) {
	err := withWorkbook(path, func(f *excelize.File) error {
		return f.UnmergeCell(sheetName, startCell, endCell)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("unmerged %s!%s:%s", sheetName, startCell, endCell), nil
}

func GetMergedCells(path, sheetName string) ([]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf(errFmtOpenWorkbook, err)
	}
	defer func() { _ = f.Close() }()
	return getMergedCellsFromFile(f, sheetName)
}

func getMergedCellsFromFile(f *excelize.File, sheetName string) ([]string, error) {
	merged, err := f.GetMergeCells(sheetName)
	if err != nil {
		return nil, err
	}
	out := make([]string, 0, len(merged))
	for _, cell := range merged {
		out = append(out, cell.GetStartAxis()+":"+cell.GetEndAxis())
	}
	return out, nil
}

func SetColumnWidths(path, sheetName string, widths []ColumnWidthEntry, autoFit bool, autoFitRange string) (string, error) {
	if !autoFit && len(widths) == 0 {
		return "", fmt.Errorf("at least one of widths or auto_fit must be provided")
	}
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		if autoFit {
			var startCol, startRow, endCol, endRow int
			if autoFitRange != "" {
				var err error
				_, _, startCol, startRow, endCol, endRow, err = normalizeRangeRef(autoFitRange)
				if err != nil {
					return fmt.Errorf("invalid auto_fit_range: %w", err)
				}
			} else {
				rows, cols, _, err := worksheetBounds(f, sheetName)
				if err != nil {
					return err
				}
				startCol, startRow, endCol, endRow = 1, 1, cols, rows
			}
			maxLens := make(map[int]float64, endCol-startCol+1)
			for row := startRow; row <= endRow; row++ {
				for col := startCol; col <= endCol; col++ {
					cell, err := excelize.CoordinatesToCellName(col, row)
					if err != nil {
						return err
					}
					value, err := f.GetCellValue(sheetName, cell)
					if err != nil {
						return err
					}
					if l := float64(len(value)); l > maxLens[col] {
						maxLens[col] = l
					}
				}
			}
			for col, maxLen := range maxLens {
				colName, err := excelize.ColumnNumberToName(col)
				if err != nil {
					return err
				}
				width := maxLen + 2
				if width < 8 {
					width = 8
				}
				if width > 60 {
					width = 60
				}
				if err := f.SetColWidth(sheetName, colName, colName, width); err != nil {
					return err
				}
			}
		}
		for _, entry := range widths {
			colName := strings.ToUpper(strings.TrimSpace(entry.Column))
			if colName == "" {
				return fmt.Errorf("column is required")
			}
			if _, _, err := excelize.CellNameToCoordinates(colName + "1"); err != nil {
				return fmt.Errorf("invalid column %q", entry.Column)
			}
			if entry.Width < 0 {
				return fmt.Errorf("width must be non-negative")
			}
			if err := f.SetColWidth(sheetName, colName, colName, entry.Width); err != nil {
				return err
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("set column widths in %s", sheetName), nil
}

func SetRowHeights(path, sheetName string, heights []RowHeightEntry) (string, error) {
	if len(heights) == 0 {
		return "", fmt.Errorf("heights is required")
	}
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		for _, entry := range heights {
			if entry.Row < 1 {
				return fmt.Errorf("row must be >= 1")
			}
			if entry.Height < 0 {
				return fmt.Errorf("height must be non-negative")
			}
			if err := f.SetRowHeight(sheetName, entry.Row, entry.Height); err != nil {
				return err
			}
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("set %d row height(s) in %s", len(heights), sheetName), nil
}
