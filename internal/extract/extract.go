package extract

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"unicode"

	"github.com/xuri/excelize/v2"
)

const (
	maxExcelColumn = 16384
	maxExcelRow    = 1048576
)

var (
	ErrUnsupportedFile = errors.New("only .xlsx and .xlsm files are supported")
	ErrSheetNotFound   = errors.New("sheet not found")
)

type SheetList struct {
	File   string   `json:"file"`
	Sheets []string `json:"sheets"`
}

type CellResult struct {
	File  string `json:"file"`
	Sheet string `json:"sheet"`
	Row   int    `json:"row"`
	Col   string `json:"col"`
	Value string `json:"value"`
}

type RowCell struct {
	Col   string `json:"col"`
	Value string `json:"value"`
}

type RowResult struct {
	File  string    `json:"file"`
	Sheet string    `json:"sheet"`
	Row   int       `json:"row"`
	Cells []RowCell `json:"cells"`
}

func ListSheets(path string) (SheetList, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return SheetList{}, err
	}
	defer f.Close()

	return SheetList{
		File:   path,
		Sheets: f.GetSheetList(),
	}, nil
}

func ExtractCell(path, sheet string, row int, col string) (CellResult, error) {
	if err := validateRow(row); err != nil {
		return CellResult{}, err
	}

	colNum, colName, err := NormalizeColumn(col)
	if err != nil {
		return CellResult{}, err
	}

	f, err := openWorkbook(path)
	if err != nil {
		return CellResult{}, err
	}
	defer f.Close()

	if !sheetExists(f, sheet) {
		return CellResult{}, fmt.Errorf("%w: %s", ErrSheetNotFound, sheet)
	}

	cellName, err := excelize.CoordinatesToCellName(colNum, row)
	if err != nil {
		return CellResult{}, err
	}

	value, err := cellDisplayValue(f, sheet, cellName)
	if err != nil {
		return CellResult{}, err
	}

	return CellResult{
		File:  path,
		Sheet: sheet,
		Row:   row,
		Col:   colName,
		Value: value,
	}, nil
}

func ExtractRow(path, sheet string, row int, fromCol, toCol string) (RowResult, error) {
	if err := validateRow(row); err != nil {
		return RowResult{}, err
	}

	f, err := openWorkbook(path)
	if err != nil {
		return RowResult{}, err
	}
	defer f.Close()

	if !sheetExists(f, sheet) {
		return RowResult{}, fmt.Errorf("%w: %s", ErrSheetNotFound, sheet)
	}

	rowLen, rowsMaxRow, rowsMaxCol, err := rowStats(f, sheet, row)
	if err != nil {
		return RowResult{}, err
	}
	dimensionMaxRow, dimensionMaxCol, err := sheetDimension(f, sheet)
	if err != nil {
		return RowResult{}, err
	}
	maxRow := max(dimensionMaxRow, rowsMaxRow)
	maxCol := max(dimensionMaxCol, rowsMaxCol)

	fromNum := 1
	if strings.TrimSpace(fromCol) != "" {
		var err error
		fromNum, _, err = NormalizeColumn(fromCol)
		if err != nil {
			return RowResult{}, fmt.Errorf("invalid from column: %w", err)
		}
	}

	var toNum int
	explicitToCol := strings.TrimSpace(toCol) != ""
	switch {
	case explicitToCol:
		var err error
		toNum, _, err = NormalizeColumn(toCol)
		if err != nil {
			return RowResult{}, fmt.Errorf("invalid to column: %w", err)
		}
	case rowLen == 0 && row > maxRow:
		toNum = 0
	case maxCol > 0:
		toNum = maxCol
	case rowLen == 0:
		toNum = 0
	default:
		toNum = rowLen
	}

	if toNum == 0 || (!explicitToCol && fromNum > toNum) {
		return RowResult{File: path, Sheet: sheet, Row: row, Cells: []RowCell{}}, nil
	}
	if fromNum > toNum {
		return RowResult{}, fmt.Errorf("from column must be less than or equal to to column")
	}

	cells := make([]RowCell, 0, toNum-fromNum+1)
	for colNum := fromNum; colNum <= toNum; colNum++ {
		colName, err := excelize.ColumnNumberToName(colNum)
		if err != nil {
			return RowResult{}, err
		}
		cellName, err := excelize.CoordinatesToCellName(colNum, row)
		if err != nil {
			return RowResult{}, err
		}
		value, err := cellDisplayValue(f, sheet, cellName)
		if err != nil {
			return RowResult{}, err
		}
		cells = append(cells, RowCell{Col: colName, Value: value})
	}

	return RowResult{
		File:  path,
		Sheet: sheet,
		Row:   row,
		Cells: cells,
	}, nil
}

func NormalizeColumn(input string) (int, string, error) {
	col := strings.TrimSpace(input)
	if col == "" {
		return 0, "", fmt.Errorf("column is required")
	}

	if allDigits(col) {
		n, err := strconv.Atoi(col)
		if err != nil || n < 1 {
			return 0, "", fmt.Errorf("column number must be greater than 0")
		}
		if n > maxExcelColumn {
			return 0, "", fmt.Errorf("column exceeds Excel limit %d", maxExcelColumn)
		}
		name, err := excelize.ColumnNumberToName(n)
		if err != nil {
			return 0, "", err
		}
		return n, name, nil
	}

	col = strings.ToUpper(col)
	for _, r := range col {
		if r < 'A' || r > 'Z' {
			return 0, "", fmt.Errorf("column must be letters or a positive number")
		}
	}

	n, err := excelize.ColumnNameToNumber(col)
	if err != nil {
		return 0, "", err
	}
	if n < 1 || n > maxExcelColumn {
		return 0, "", fmt.Errorf("column exceeds Excel limit %d", maxExcelColumn)
	}
	return n, col, nil
}

func openWorkbook(path string) (*excelize.File, error) {
	if err := validateFile(path); err != nil {
		return nil, err
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("open workbook: %w", err)
	}
	return f, nil
}

func validateFile(path string) error {
	if strings.TrimSpace(path) == "" {
		return fmt.Errorf("file is required")
	}

	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".xlsx" && ext != ".xlsm" {
		return fmt.Errorf("%w: %s", ErrUnsupportedFile, ext)
	}

	info, err := os.Stat(path)
	if err != nil {
		return fmt.Errorf("stat file: %w", err)
	}
	if info.IsDir() {
		return fmt.Errorf("file is a directory: %s", path)
	}
	return nil
}

func validateRow(row int) error {
	if row < 1 {
		return fmt.Errorf("row must be greater than 0")
	}
	if row > maxExcelRow {
		return fmt.Errorf("row exceeds Excel limit %d", maxExcelRow)
	}
	return nil
}

func sheetExists(f *excelize.File, sheet string) bool {
	for _, name := range f.GetSheetList() {
		if name == sheet {
			return true
		}
	}
	return false
}

func rowStats(f *excelize.File, sheet string, row int) (int, int, int, error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return 0, 0, 0, err
	}

	maxCol := 0
	for _, currentRow := range rows {
		maxCol = max(maxCol, len(currentRow))
	}
	if row > len(rows) {
		return 0, len(rows), maxCol, nil
	}
	return len(rows[row-1]), len(rows), maxCol, nil
}

func sheetDimension(f *excelize.File, sheet string) (int, int, error) {
	dimension, err := f.GetSheetDimension(sheet)
	if err != nil {
		return 0, 0, err
	}
	if dimension == "" {
		return 0, 0, nil
	}

	parts := strings.Split(dimension, ":")
	ref := parts[len(parts)-1]
	col, row, err := excelize.CellNameToCoordinates(ref)
	if err != nil {
		return 0, 0, err
	}
	return row, col, nil
}

func cellDisplayValue(f *excelize.File, sheet, cell string) (string, error) {
	formula, err := f.GetCellFormula(sheet, cell)
	if err != nil {
		return "", fmt.Errorf("read formula %s!%s: %w", sheet, cell, err)
	}
	if formula != "" {
		value, err := f.CalcCellValue(sheet, cell)
		if err != nil {
			return "", fmt.Errorf("calculate formula %s!%s: %w", sheet, cell, err)
		}
		return strings.TrimSpace(value), nil
	}

	value, err := f.GetCellValue(sheet, cell)
	if err != nil {
		return "", fmt.Errorf("read value %s!%s: %w", sheet, cell, err)
	}
	return value, nil
}

func allDigits(s string) bool {
	for _, r := range s {
		if !unicode.IsDigit(r) {
			return false
		}
	}
	return s != ""
}
