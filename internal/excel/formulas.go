package excel

import (
	"regexp"
	"strings"

	"github.com/xuri/efp"
	"github.com/xuri/excelize/v2"
)

var formulaRangePattern = regexp.MustCompile(`^((?:'[^']*(?:''[^']*)*'|[^!]+)!)?(\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(\$?[A-Za-z]{1,3}\$?[0-9]+))?$`)
var formulaCellPattern = regexp.MustCompile(`^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$`)

func translateFormula(formula string, colDelta, rowDelta int) string {
	if formula == "" || (colDelta == 0 && rowDelta == 0) {
		return formula
	}
	parser := efp.ExcelParser()
	parser.Parse(formula)
	for index := range parser.Tokens.Items {
		token := &parser.Tokens.Items[index]
		if token.TType != efp.TokenTypeOperand || token.TSubType != efp.TokenSubTypeRange {
			continue
		}
		match := formulaRangePattern.FindStringSubmatch(token.TValue)
		if match == nil {
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
