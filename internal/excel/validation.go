package excel

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func AddDataValidation(path, sheetName, rangeRef string, opts DataValidationOptions) (string, error) {
	startCell, endCell, _, _, _, _, err := normalizeRangeOrCellRef(rangeRef)
	if err != nil {
		return "", fmt.Errorf("invalid range: %w", err)
	}
	validationType, err := normalizeValidationType(opts.Type)
	if err != nil {
		return "", err
	}
	operator, err := normalizeValidationOperator(opts.Operator, validationType)
	if err != nil {
		return "", err
	}
	if validationType != "list" && len(opts.Values) > 0 {
		return "", fmt.Errorf("values is only supported for list validation")
	}
	if validationType != "list" && opts.ShowDropDown {
		return "", fmt.Errorf("show_drop_down is only supported for list validation")
	}
	if (validationType == "list" || validationType == "custom") && strings.TrimSpace(opts.Formula2) != "" {
		return "", fmt.Errorf("validation type %q does not accept formula2", validationType)
	}
	if validationType != "list" && validationType != "custom" && operator != "between" && operator != "notBetween" && strings.TrimSpace(opts.Formula2) != "" {
		return "", fmt.Errorf("operator %q does not accept formula2", opts.Operator)
	}
	errorStyle, err := validationErrorStyle(opts.ErrorStyle)
	if err != nil {
		return "", err
	}

	err = withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		dv := excelize.NewDataValidation(opts.AllowBlank)
		dv.Sqref = startCell + ":" + endCell
		dv.ShowDropDown = opts.ShowDropDown
		switch validationType {
		case "list":
			if len(opts.Values) > 0 && strings.TrimSpace(opts.Formula1) != "" {
				return fmt.Errorf("list validation accepts either values or formula1, not both")
			}
			switch {
			case len(opts.Values) > 0:
				if err := dv.SetDropList(opts.Values); err != nil {
					return err
				}
			case strings.TrimSpace(opts.Formula1) != "":
				dv.SetSqrefDropList(opts.Formula1)
			default:
				return fmt.Errorf("list validation requires values or formula1")
			}
		case "custom":
			if strings.TrimSpace(opts.Formula1) == "" {
				return fmt.Errorf("custom validation requires formula1")
			}
			dv.Type = validationType
			dv.Formula1 = opts.Formula1
		default:
			if strings.TrimSpace(opts.Formula1) == "" {
				return fmt.Errorf("%s validation requires formula1", validationType)
			}
			if (operator == "between" || operator == "notBetween") && strings.TrimSpace(opts.Formula2) == "" {
				return fmt.Errorf("operator %q requires formula2", opts.Operator)
			}
			dv.Type = validationType
			dv.Operator = operator
			dv.Formula1 = opts.Formula1
			dv.Formula2 = opts.Formula2
		}
		if opts.ErrorStyle != "" || opts.ErrorTitle != "" || opts.ErrorBody != "" {
			dv.SetError(errorStyle, opts.ErrorTitle, opts.ErrorBody)
		}
		if opts.PromptTitle != "" || opts.PromptBody != "" {
			dv.SetInput(opts.PromptTitle, opts.PromptBody)
		}
		return f.AddDataValidation(sheetName, dv)
	})
	if err != nil {
		return "", err
	}
	return fmt.Sprintf("added data validation to %s!%s:%s", sheetName, startCell, endCell), nil
}

func DeleteDataValidation(path, sheetName, rangeRef string) (string, error) {
	var normalizedRange string
	if strings.TrimSpace(rangeRef) != "" {
		startCell, endCell, _, _, _, _, err := normalizeRangeOrCellRef(rangeRef)
		if err != nil {
			return "", fmt.Errorf("invalid range: %w", err)
		}
		normalizedRange = startCell + ":" + endCell
	}
	err := withWorkbook(path, func(f *excelize.File) error {
		if err := ensureSheetExists(f, sheetName); err != nil {
			return err
		}
		if normalizedRange == "" {
			return f.DeleteDataValidation(sheetName)
		}
		return f.DeleteDataValidation(sheetName, normalizedRange)
	})
	if err != nil {
		return "", err
	}
	if normalizedRange == "" {
		return fmt.Sprintf("deleted all data validations from %s", sheetName), nil
	}
	return fmt.Sprintf("deleted data validation from %s!%s", sheetName, normalizedRange), nil
}

func normalizeValidationType(value string) (string, error) {
	switch strings.ToLower(strings.TrimSpace(value)) {
	case "custom":
		return "custom", nil
	case "date":
		return "date", nil
	case "decimal":
		return "decimal", nil
	case "list":
		return "list", nil
	case "text_length", "textlength":
		return "textLength", nil
	case "time":
		return "time", nil
	case "whole", "whole_number":
		return "whole", nil
	default:
		return "", fmt.Errorf("unsupported validation type %q", value)
	}
}

func normalizeValidationOperator(value, validationType string) (string, error) {
	value = strings.ToLower(strings.TrimSpace(value))
	if validationType == "list" || validationType == "custom" {
		if value != "" {
			return "", fmt.Errorf("validation type %q does not accept an operator", validationType)
		}
		return "", nil
	}
	if value == "" {
		value = "between"
	}
	switch value {
	case "between":
		return "between", nil
	case "equal", "equals":
		return "equal", nil
	case "greater_than", "greaterthan", "gt":
		return "greaterThan", nil
	case "greater_than_or_equal", "greaterthanorequal", "gte":
		return "greaterThanOrEqual", nil
	case "less_than", "lessthan", "lt":
		return "lessThan", nil
	case "less_than_or_equal", "lessthanorequal", "lte":
		return "lessThanOrEqual", nil
	case "not_between", "notbetween":
		return "notBetween", nil
	case "not_equal", "notequal":
		return "notEqual", nil
	default:
		return "", fmt.Errorf("unsupported validation operator %q", value)
	}
}

func validationErrorStyle(value string) (excelize.DataValidationErrorStyle, error) {
	switch strings.ToLower(strings.TrimSpace(value)) {
	case "", "stop":
		return excelize.DataValidationErrorStyleStop, nil
	case "warning":
		return excelize.DataValidationErrorStyleWarning, nil
	case "information", "info":
		return excelize.DataValidationErrorStyleInformation, nil
	default:
		return 0, fmt.Errorf("unsupported error_style %q", value)
	}
}
