package excel

import (
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type cellSnapshot struct {
	cellType excelize.CellType
	formula  string
	rawValue string
	styleID  int
}

func readCellSnapshot(f *excelize.File, sheetName, cell string) (cellSnapshot, error) {
	cellType, err := f.GetCellType(sheetName, cell)
	if err != nil {
		return cellSnapshot{}, err
	}
	formula, err := f.GetCellFormula(sheetName, cell)
	if err != nil {
		return cellSnapshot{}, err
	}
	rawValue, err := f.GetCellValue(sheetName, cell, excelize.Options{RawCellValue: true})
	if err != nil {
		return cellSnapshot{}, err
	}
	styleID, err := f.GetCellStyle(sheetName, cell)
	if err != nil {
		return cellSnapshot{}, err
	}
	return cellSnapshot{cellType: cellType, formula: formula, rawValue: rawValue, styleID: styleID}, nil
}

func writeCellSnapshot(f *excelize.File, sheetName, cell string, snapshot cellSnapshot, colDelta, rowDelta int) error {
	var err error
	if snapshot.formula != "" {
		formula := translateFormula(strings.TrimPrefix(snapshot.formula, "="), colDelta, rowDelta)
		err = f.SetCellFormula(sheetName, cell, formula)
	} else {
		switch snapshot.cellType {
		case excelize.CellTypeBool:
			err = f.SetCellBool(sheetName, cell, snapshot.rawValue == "1" || strings.EqualFold(snapshot.rawValue, "true"))
		case excelize.CellTypeNumber:
			var number float64
			number, err = strconv.ParseFloat(snapshot.rawValue, 64)
			if err == nil {
				err = f.SetCellFloat(sheetName, cell, number, -1, 64)
			}
		case excelize.CellTypeUnset:
			err = f.SetCellDefault(sheetName, cell, snapshot.rawValue)
		default:
			err = f.SetCellStr(sheetName, cell, snapshot.rawValue)
		}
	}
	if err != nil {
		return err
	}
	return f.SetCellStyle(sheetName, cell, cell, snapshot.styleID)
}

func readCellJSONValue(f *excelize.File, sheetName, cell string) (any, string, error) {
	displayValue, err := f.GetCellValue(sheetName, cell)
	if err != nil {
		return nil, "", err
	}
	if displayValue == "" {
		return "", "", nil
	}
	cellType, err := f.GetCellType(sheetName, cell)
	if err != nil {
		return nil, "", err
	}
	switch cellType {
	case excelize.CellTypeBool:
		rawValue, err := f.GetCellValue(sheetName, cell, excelize.Options{RawCellValue: true})
		if err != nil {
			return nil, "", err
		}
		return rawValue == "1" || strings.EqualFold(rawValue, "true"), displayValue, nil
	case excelize.CellTypeNumber, excelize.CellTypeUnset:
		rawValue, err := f.GetCellValue(sheetName, cell, excelize.Options{RawCellValue: true})
		if err != nil {
			return nil, "", err
		}
		number, err := strconv.ParseFloat(rawValue, 64)
		if err == nil && !math.IsNaN(number) && !math.IsInf(number, 0) {
			return number, displayValue, nil
		}
	}
	return displayValue, displayValue, nil
}

func appendHeader(headers []string, seen map[string]struct{}, value string) ([]string, error) {
	value = strings.TrimSpace(value)
	if value == "" {
		value = fmt.Sprintf("column_%d", len(headers)+1)
	}
	if _, exists := seen[value]; exists {
		return nil, fmt.Errorf("header row contains duplicate column %q", value)
	}
	seen[value] = struct{}{}
	return append(headers, value), nil
}

func jsonComparisonValue(value any, displayValue string) string {
	switch typed := value.(type) {
	case float64:
		return strconv.FormatFloat(typed, 'g', -1, 64)
	case bool:
		return strconv.FormatBool(typed)
	default:
		return displayValue
	}
}

func snapshotComparisonValue(snapshot cellSnapshot, displayValue string) string {
	switch snapshot.cellType {
	case excelize.CellTypeBool:
		if snapshot.rawValue == "1" || strings.EqualFold(snapshot.rawValue, "true") {
			return "true"
		}
		return "false"
	case excelize.CellTypeNumber, excelize.CellTypeUnset:
		if _, ok := parseFiniteFloat(snapshot.rawValue); ok {
			return snapshot.rawValue
		}
	}
	return displayValue
}
