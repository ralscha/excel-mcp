package excel

import (
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/efp"
	"github.com/xuri/excelize/v2"
)

var formulaRangePattern = regexp.MustCompile(`^((?:'[^']*(?:''[^']*)*'|[^!]+)!)?(\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(\$?[A-Za-z]{1,3}\$?[0-9]+))?$`)
var formulaCellPattern = regexp.MustCompile(`^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$`)
var formulaColumnRangePattern = regexp.MustCompile(`^((?:'[^']*(?:''[^']*)*'|[^!]+)!)?(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})$`)
var formulaRowRangePattern = regexp.MustCompile(`^((?:'[^']*(?:''[^']*)*'|[^!]+)!)?(\$?[0-9]+):(\$?[0-9]+)$`)

func translateFormula(formula string, colDelta, rowDelta int) string {
	if formula == "" || (colDelta == 0 && rowDelta == 0) {
		return formula
	}
	parser := efp.ExcelParser()
	parser.Parse(formula)
	for index := range parser.Tokens.Items {
		token := &parser.Tokens.Items[index]
		if token.TType == efp.TokenTypeOperand && token.TSubType == efp.TokenSubTypeText {
			token.TValue = strings.ReplaceAll(token.TValue, `"`, `""`)
		}
		if token.TType != efp.TokenTypeOperand || token.TSubType != efp.TokenSubTypeRange {
			continue
		}
		match := formulaRangePattern.FindStringSubmatch(token.TValue)
		if match == nil {
			if columns := formulaColumnRangePattern.FindStringSubmatch(token.TValue); columns != nil {
				token.TValue = translateWholeRange(columns, colDelta, true)
			} else if rows := formulaRowRangePattern.FindStringSubmatch(token.TValue); rows != nil {
				token.TValue = translateWholeRange(rows, rowDelta, false)
			}
			continue
		}
		start, ok := translateCellReference(match[2], colDelta, rowDelta)
		if !ok {
			token.TValue = "#REF!"
			continue
		}
		token.TValue = quoteFormulaSheetPrefix(match[1]) + start
		if match[3] != "" {
			end, ok := translateCellReference(match[3], colDelta, rowDelta)
			if !ok {
				token.TValue = "#REF!"
				continue
			}
			token.TValue += ":" + end
		}
	}
	return parser.Render()
}

func translateWholeRange(parts []string, delta int, columns bool) string {
	refs := [2]string{}
	for index, ref := range parts[2:] {
		absolute := strings.HasPrefix(ref, "$")
		value := strings.TrimPrefix(ref, "$")
		var number int
		var err error
		limit := excelize.TotalRows
		if columns {
			number, err = excelize.ColumnNameToNumber(value)
			limit = excelize.MaxColumns
		} else {
			number, err = strconv.Atoi(value)
		}
		if err != nil || number < 1 || number > limit {
			return "#REF!"
		}
		if !absolute {
			number += delta
		}
		if number < 1 || number > limit {
			return "#REF!"
		}
		if columns {
			value, _ = excelize.ColumnNumberToName(number)
		} else {
			value = strconv.Itoa(number)
		}
		if absolute {
			value = "$" + value
		}
		refs[index] = value
	}
	return quoteFormulaSheetPrefix(parts[1]) + refs[0] + ":" + refs[1]
}

func translateCellReference(ref string, colDelta, rowDelta int) (string, bool) {
	parts := formulaCellPattern.FindStringSubmatch(ref)
	if parts == nil {
		return "", false
	}
	colAbsolute := parts[1] == "$"
	rowAbsolute := parts[3] == "$"
	col, row, err := excelize.CellNameToCoordinates(parts[2] + parts[4])
	if err != nil {
		return "", false
	}
	if !colAbsolute {
		col += colDelta
	}
	if !rowAbsolute {
		row += rowDelta
	}
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return "", false
	}
	letters := strings.TrimRight(cell, "0123456789")
	rowPart := strings.TrimPrefix(cell, letters)
	if colAbsolute {
		letters = "$" + letters
	}
	if rowAbsolute {
		rowPart = "$" + rowPart
	}
	return letters + rowPart, true
}

func quoteFormulaSheetPrefix(prefix string) string {
	if prefix == "" {
		return ""
	}
	sheetName := strings.TrimSuffix(prefix, "!")
	if strings.HasPrefix(sheetName, "'") {
		return prefix
	}
	return quoteSheetName(sheetName) + "!"
}
