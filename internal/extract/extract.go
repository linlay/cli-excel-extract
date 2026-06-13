package extract

import (
	"errors"
	"fmt"
	"math"
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

type WorkbookInfo struct {
	File   string      `json:"file"`
	Sheets []SheetInfo `json:"sheets"`
}

type SheetInfo struct {
	Name   string `json:"name"`
	Range  string `json:"range"`
	MaxRow int    `json:"maxRow"`
	MaxCol int    `json:"maxCol"`
}

type ColCell struct {
	Row   int    `json:"row"`
	Value string `json:"value"`
}

type ColResult struct {
	File  string    `json:"file"`
	Sheet string    `json:"sheet"`
	Col   string    `json:"col"`
	Cells []ColCell `json:"cells"`
}

type RangeCell struct {
	Row   int    `json:"row"`
	Col   string `json:"col"`
	Value string `json:"value"`
}

type RangeResult struct {
	File  string      `json:"file"`
	Sheet string      `json:"sheet"`
	Range string      `json:"range"`
	Cells []RangeCell `json:"cells"`
}

type ReadBatchRequest struct {
	Queries []ReadQuery `json:"queries"`
}

type ReadQuery struct {
	Sheet string `json:"sheet"`
	Row   int    `json:"row"`
	Col   string `json:"col"`
}

type ReadBatchCell struct {
	Sheet string `json:"sheet"`
	Row   int    `json:"row"`
	Col   string `json:"col"`
	Value string `json:"value"`
}

type ReadBatchResult struct {
	File  string          `json:"file"`
	Cells []ReadBatchCell `json:"cells"`
}

type FillRequest struct {
	Updates []FillUpdate `json:"updates"`
}

type FillUpdate struct {
	Sheet string  `json:"sheet"`
	Row   int     `json:"row"`
	Col   string  `json:"col"`
	Type  string  `json:"type"`
	Value *string `json:"value,omitempty"`
}

type FillResult struct {
	File    string `json:"file"`
	Output  string `json:"output"`
	InPlace bool   `json:"inPlace"`
	Updated int    `json:"updated"`
}

type FillValue struct {
	Type  string  `json:"type"`
	Value *string `json:"value,omitempty"`
}

type ValuesRequest struct {
	Values []FillValue `json:"values"`
}

type RangeValuesRequest struct {
	Rows [][]FillValue `json:"rows"`
}

type cellRange struct {
	startCol int
	startRow int
	endCol   int
	endRow   int
	ref      string
}

type preparedFillUpdate struct {
	sheet   string
	cell    string
	typ     string
	text    string
	number  float64
	boolean bool
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

func InspectWorkbook(path string) (WorkbookInfo, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return WorkbookInfo{}, err
	}
	defer f.Close()

	sheets := make([]SheetInfo, 0, len(f.GetSheetList()))
	for _, sheet := range f.GetSheetList() {
		dimensionMaxRow, dimensionMaxCol, err := sheetDimension(f, sheet)
		if err != nil {
			return WorkbookInfo{}, err
		}
		_, rowsMaxRow, rowsMaxCol, err := rowStats(f, sheet, 1)
		if err != nil {
			return WorkbookInfo{}, err
		}
		maxRow := max(dimensionMaxRow, rowsMaxRow)
		maxCol := max(dimensionMaxCol, rowsMaxCol)
		rangeRef := ""
		if maxRow > 0 && maxCol > 0 {
			lastCell, err := excelize.CoordinatesToCellName(maxCol, maxRow)
			if err != nil {
				return WorkbookInfo{}, err
			}
			rangeRef = "A1:" + lastCell
		}
		sheets = append(sheets, SheetInfo{
			Name:   sheet,
			Range:  rangeRef,
			MaxRow: maxRow,
			MaxCol: maxCol,
		})
	}
	return WorkbookInfo{File: path, Sheets: sheets}, nil
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

func ExtractCol(path, sheet string, col string, fromRow, toRow int) (ColResult, error) {
	colNum, colName, err := NormalizeColumn(col)
	if err != nil {
		return ColResult{}, err
	}
	if fromRow < 1 {
		return ColResult{}, fmt.Errorf("from row must be greater than 0")
	}
	if err := validateRow(fromRow); err != nil {
		return ColResult{}, err
	}
	if toRow < 0 {
		return ColResult{}, fmt.Errorf("to row must be greater than 0")
	}
	if toRow > 0 {
		if err := validateRow(toRow); err != nil {
			return ColResult{}, err
		}
		if fromRow > toRow {
			return ColResult{}, fmt.Errorf("from row must be less than or equal to to row")
		}
	}

	f, err := openWorkbook(path)
	if err != nil {
		return ColResult{}, err
	}
	defer f.Close()

	if !sheetExists(f, sheet) {
		return ColResult{}, fmt.Errorf("%w: %s", ErrSheetNotFound, sheet)
	}

	if toRow == 0 {
		dimensionMaxRow, _, err := sheetDimension(f, sheet)
		if err != nil {
			return ColResult{}, err
		}
		if dimensionMaxRow == 0 || fromRow > dimensionMaxRow {
			return ColResult{File: path, Sheet: sheet, Col: colName, Cells: []ColCell{}}, nil
		}
		toRow = dimensionMaxRow
	}

	cells := make([]ColCell, 0, toRow-fromRow+1)
	for row := fromRow; row <= toRow; row++ {
		cellName, err := excelize.CoordinatesToCellName(colNum, row)
		if err != nil {
			return ColResult{}, err
		}
		value, err := cellDisplayValue(f, sheet, cellName)
		if err != nil {
			return ColResult{}, err
		}
		cells = append(cells, ColCell{Row: row, Value: value})
	}
	return ColResult{File: path, Sheet: sheet, Col: colName, Cells: cells}, nil
}

func ExtractRange(path, sheet, rangeRef string) (RangeResult, error) {
	parsed, err := parseA1Range(rangeRef)
	if err != nil {
		return RangeResult{}, err
	}

	f, err := openWorkbook(path)
	if err != nil {
		return RangeResult{}, err
	}
	defer f.Close()

	if !sheetExists(f, sheet) {
		return RangeResult{}, fmt.Errorf("%w: %s", ErrSheetNotFound, sheet)
	}

	cells := make([]RangeCell, 0, parsed.width()*parsed.height())
	for row := parsed.startRow; row <= parsed.endRow; row++ {
		for colNum := parsed.startCol; colNum <= parsed.endCol; colNum++ {
			colName, err := excelize.ColumnNumberToName(colNum)
			if err != nil {
				return RangeResult{}, err
			}
			cellName, err := excelize.CoordinatesToCellName(colNum, row)
			if err != nil {
				return RangeResult{}, err
			}
			value, err := cellDisplayValue(f, sheet, cellName)
			if err != nil {
				return RangeResult{}, err
			}
			cells = append(cells, RangeCell{Row: row, Col: colName, Value: value})
		}
	}
	return RangeResult{File: path, Sheet: sheet, Range: parsed.ref, Cells: cells}, nil
}

func ReadCells(path string, request ReadBatchRequest) (ReadBatchResult, error) {
	if len(request.Queries) == 0 {
		return ReadBatchResult{}, fmt.Errorf("queries must contain at least one cell query")
	}

	f, err := openWorkbook(path)
	if err != nil {
		return ReadBatchResult{}, err
	}
	defer f.Close()

	cells := make([]ReadBatchCell, 0, len(request.Queries))
	for i, query := range request.Queries {
		cell, err := readBatchCell(f, query)
		if err != nil {
			return ReadBatchResult{}, fmt.Errorf("query %d: %w", i+1, err)
		}
		cells = append(cells, cell)
	}
	return ReadBatchResult{File: path, Cells: cells}, nil
}

func FillCells(path string, request FillRequest, output string, overwrite bool) (FillResult, error) {
	return WriteBatch(path, request, output, overwrite)
}

func WriteBatch(path string, request FillRequest, output string, overwrite bool) (FillResult, error) {
	outputPath, inPlace, err := normalizeFillOutput(path, output, overwrite)
	if err != nil {
		return FillResult{}, err
	}

	f, err := openWorkbook(path)
	if err != nil {
		return FillResult{}, err
	}
	defer f.Close()

	updates, err := prepareFillUpdates(f, request)
	if err != nil {
		return FillResult{}, err
	}

	for _, update := range updates {
		if err := applyFillUpdate(f, update); err != nil {
			return FillResult{}, err
		}
	}

	if inPlace {
		if err := f.Save(); err != nil {
			return FillResult{}, fmt.Errorf("save workbook: %w", err)
		}
	} else {
		if err := f.SaveAs(outputPath); err != nil {
			return FillResult{}, fmt.Errorf("save workbook as %s: %w", outputPath, err)
		}
	}

	return FillResult{
		File:    path,
		Output:  outputPath,
		InPlace: inPlace,
		Updated: len(updates),
	}, nil
}

func WriteCell(path, sheet string, row int, col string, value FillValue, output string, overwrite bool) (FillResult, error) {
	return WriteBatch(path, FillRequest{Updates: []FillUpdate{{
		Sheet: sheet,
		Row:   row,
		Col:   col,
		Type:  value.Type,
		Value: value.Value,
	}}}, output, overwrite)
}

func WriteRow(path, sheet string, row int, fromCol string, values ValuesRequest, output string, overwrite bool) (FillResult, error) {
	if err := validateRow(row); err != nil {
		return FillResult{}, err
	}
	fromNum, _, err := NormalizeColumn(fromCol)
	if err != nil {
		return FillResult{}, err
	}
	if len(values.Values) == 0 {
		return FillResult{}, fmt.Errorf("values must contain at least one cell value")
	}

	updates := make([]FillUpdate, 0, len(values.Values))
	for i, value := range values.Values {
		colName, err := excelize.ColumnNumberToName(fromNum + i)
		if err != nil {
			return FillResult{}, err
		}
		updates = append(updates, FillUpdate{Sheet: sheet, Row: row, Col: colName, Type: value.Type, Value: value.Value})
	}
	return WriteBatch(path, FillRequest{Updates: updates}, output, overwrite)
}

func WriteCol(path, sheet string, col string, fromRow int, values ValuesRequest, output string, overwrite bool) (FillResult, error) {
	if err := validateRow(fromRow); err != nil {
		return FillResult{}, err
	}
	_, colName, err := NormalizeColumn(col)
	if err != nil {
		return FillResult{}, err
	}
	if len(values.Values) == 0 {
		return FillResult{}, fmt.Errorf("values must contain at least one cell value")
	}

	updates := make([]FillUpdate, 0, len(values.Values))
	for i, value := range values.Values {
		updates = append(updates, FillUpdate{Sheet: sheet, Row: fromRow + i, Col: colName, Type: value.Type, Value: value.Value})
	}
	return WriteBatch(path, FillRequest{Updates: updates}, output, overwrite)
}

func WriteRange(path, sheet, rangeRef string, values RangeValuesRequest, output string, overwrite bool) (FillResult, error) {
	parsed, err := parseA1Range(rangeRef)
	if err != nil {
		return FillResult{}, err
	}
	if len(values.Rows) != parsed.height() {
		return FillResult{}, fmt.Errorf("range values must contain %d rows", parsed.height())
	}

	updates := make([]FillUpdate, 0, parsed.width()*parsed.height())
	for rowOffset, rowValues := range values.Rows {
		if len(rowValues) != parsed.width() {
			return FillResult{}, fmt.Errorf("range values row %d must contain %d values", rowOffset+1, parsed.width())
		}
		for colOffset, value := range rowValues {
			colName, err := excelize.ColumnNumberToName(parsed.startCol + colOffset)
			if err != nil {
				return FillResult{}, err
			}
			updates = append(updates, FillUpdate{
				Sheet: sheet,
				Row:   parsed.startRow + rowOffset,
				Col:   colName,
				Type:  value.Type,
				Value: value.Value,
			})
		}
	}
	return WriteBatch(path, FillRequest{Updates: updates}, output, overwrite)
}

func ClearRange(path, sheet, rangeRef string, output string, overwrite bool) (FillResult, error) {
	parsed, err := parseA1Range(rangeRef)
	if err != nil {
		return FillResult{}, err
	}
	updates := make([]FillUpdate, 0, parsed.width()*parsed.height())
	for row := parsed.startRow; row <= parsed.endRow; row++ {
		for colNum := parsed.startCol; colNum <= parsed.endCol; colNum++ {
			colName, err := excelize.ColumnNumberToName(colNum)
			if err != nil {
				return FillResult{}, err
			}
			updates = append(updates, FillUpdate{Sheet: sheet, Row: row, Col: colName, Type: "blank"})
		}
	}
	return WriteBatch(path, FillRequest{Updates: updates}, output, overwrite)
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

func normalizeFillOutput(path, output string, overwrite bool) (string, bool, error) {
	output = strings.TrimSpace(output)
	if output == "" {
		return path, true, nil
	}

	if err := validateSupportedExtension(output); err != nil {
		return "", false, err
	}

	same, err := samePath(path, output)
	if err != nil {
		return "", false, err
	}
	if same {
		return "", false, fmt.Errorf("output must be different from input; omit --output to save in place")
	}

	info, err := os.Stat(output)
	if err == nil {
		if info.IsDir() {
			return "", false, fmt.Errorf("output is a directory: %s", output)
		}
		if !overwrite {
			return "", false, fmt.Errorf("output file already exists: %s", output)
		}
		return output, false, nil
	}
	if !errors.Is(err, os.ErrNotExist) {
		return "", false, fmt.Errorf("stat output: %w", err)
	}

	return output, false, nil
}

func validateFile(path string) error {
	if strings.TrimSpace(path) == "" {
		return fmt.Errorf("file is required")
	}

	if err := validateSupportedExtension(path); err != nil {
		return err
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

func validateSupportedExtension(path string) error {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".xlsx" && ext != ".xlsm" {
		return fmt.Errorf("%w: %s", ErrUnsupportedFile, ext)
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

func readBatchCell(f *excelize.File, query ReadQuery) (ReadBatchCell, error) {
	sheet := strings.TrimSpace(query.Sheet)
	if sheet == "" {
		return ReadBatchCell{}, fmt.Errorf("sheet is required")
	}
	if !sheetExists(f, sheet) {
		return ReadBatchCell{}, fmt.Errorf("%w: %s", ErrSheetNotFound, sheet)
	}
	if err := validateRow(query.Row); err != nil {
		return ReadBatchCell{}, err
	}

	colNum, colName, err := NormalizeColumn(query.Col)
	if err != nil {
		return ReadBatchCell{}, err
	}
	cellName, err := excelize.CoordinatesToCellName(colNum, query.Row)
	if err != nil {
		return ReadBatchCell{}, err
	}
	value, err := cellDisplayValue(f, sheet, cellName)
	if err != nil {
		return ReadBatchCell{}, err
	}
	return ReadBatchCell{Sheet: sheet, Row: query.Row, Col: colName, Value: value}, nil
}

func parseA1Range(input string) (cellRange, error) {
	ref := strings.TrimSpace(input)
	if ref == "" {
		return cellRange{}, fmt.Errorf("range is required")
	}
	if strings.Contains(ref, "!") {
		return cellRange{}, fmt.Errorf("range must not include a sheet name")
	}

	parts := strings.Split(ref, ":")
	if len(parts) > 2 {
		return cellRange{}, fmt.Errorf("range must be a cell or rectangular A1 range")
	}

	startCol, startRow, err := parseCellRef(parts[0])
	if err != nil {
		return cellRange{}, err
	}
	endCol, endRow := startCol, startRow
	if len(parts) == 2 {
		endCol, endRow, err = parseCellRef(parts[1])
		if err != nil {
			return cellRange{}, err
		}
	}
	if startCol > endCol || startRow > endRow {
		return cellRange{}, fmt.Errorf("range must run top-left to bottom-right")
	}

	startName, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return cellRange{}, err
	}
	endName, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return cellRange{}, err
	}
	normalized := startName
	if startName != endName {
		normalized = startName + ":" + endName
	}

	return cellRange{startCol: startCol, startRow: startRow, endCol: endCol, endRow: endRow, ref: normalized}, nil
}

func parseCellRef(ref string) (int, int, error) {
	col, row, err := excelize.CellNameToCoordinates(strings.ToUpper(strings.TrimSpace(ref)))
	if err != nil {
		return 0, 0, err
	}
	if col < 1 || col > maxExcelColumn {
		return 0, 0, fmt.Errorf("column exceeds Excel limit %d", maxExcelColumn)
	}
	if err := validateRow(row); err != nil {
		return 0, 0, err
	}
	return col, row, nil
}

func (r cellRange) width() int {
	return r.endCol - r.startCol + 1
}

func (r cellRange) height() int {
	return r.endRow - r.startRow + 1
}

func prepareFillUpdates(f *excelize.File, request FillRequest) ([]preparedFillUpdate, error) {
	if len(request.Updates) == 0 {
		return nil, fmt.Errorf("updates must contain at least one cell update")
	}

	updates := make([]preparedFillUpdate, 0, len(request.Updates))
	for i, update := range request.Updates {
		prepared, err := prepareFillUpdate(f, update)
		if err != nil {
			return nil, fmt.Errorf("update %d: %w", i+1, err)
		}
		updates = append(updates, prepared)
	}
	return updates, nil
}

func prepareFillUpdate(f *excelize.File, update FillUpdate) (preparedFillUpdate, error) {
	sheet := strings.TrimSpace(update.Sheet)
	if sheet == "" {
		return preparedFillUpdate{}, fmt.Errorf("sheet is required")
	}
	if !sheetExists(f, sheet) {
		return preparedFillUpdate{}, fmt.Errorf("%w: %s", ErrSheetNotFound, sheet)
	}
	if err := validateRow(update.Row); err != nil {
		return preparedFillUpdate{}, err
	}

	colNum, _, err := NormalizeColumn(update.Col)
	if err != nil {
		return preparedFillUpdate{}, err
	}
	cell, err := excelize.CoordinatesToCellName(colNum, update.Row)
	if err != nil {
		return preparedFillUpdate{}, err
	}

	typ := strings.ToLower(strings.TrimSpace(update.Type))
	switch typ {
	case "text":
		value, err := requiredFillValue(update.Value)
		if err != nil {
			return preparedFillUpdate{}, err
		}
		return preparedFillUpdate{sheet: sheet, cell: cell, typ: typ, text: value}, nil
	case "number":
		value, err := requiredFillValue(update.Value)
		if err != nil {
			return preparedFillUpdate{}, err
		}
		number, err := strconv.ParseFloat(strings.TrimSpace(value), 64)
		if err != nil || math.IsNaN(number) || math.IsInf(number, 0) {
			return preparedFillUpdate{}, fmt.Errorf("number value must be a finite decimal")
		}
		return preparedFillUpdate{sheet: sheet, cell: cell, typ: typ, number: number}, nil
	case "bool":
		value, err := requiredFillValue(update.Value)
		if err != nil {
			return preparedFillUpdate{}, err
		}
		normalized := strings.ToLower(strings.TrimSpace(value))
		if normalized != "true" && normalized != "false" {
			return preparedFillUpdate{}, fmt.Errorf("bool value must be true or false")
		}
		return preparedFillUpdate{sheet: sheet, cell: cell, typ: typ, boolean: normalized == "true"}, nil
	case "formula":
		value, err := requiredFillValue(update.Value)
		if err != nil {
			return preparedFillUpdate{}, err
		}
		if !strings.HasPrefix(value, "=") {
			return preparedFillUpdate{}, fmt.Errorf("formula value must start with =")
		}
		return preparedFillUpdate{sheet: sheet, cell: cell, typ: typ, text: value}, nil
	case "blank":
		if update.Value != nil {
			return preparedFillUpdate{}, fmt.Errorf("blank update must not include value")
		}
		return preparedFillUpdate{sheet: sheet, cell: cell, typ: typ}, nil
	default:
		return preparedFillUpdate{}, fmt.Errorf("type must be one of text, number, bool, formula, blank")
	}
}

func requiredFillValue(value *string) (string, error) {
	if value == nil {
		return "", fmt.Errorf("value is required")
	}
	return *value, nil
}

func applyFillUpdate(f *excelize.File, update preparedFillUpdate) error {
	clearFormula := func() error {
		if err := f.SetCellFormula(update.sheet, update.cell, ""); err != nil {
			return fmt.Errorf("clear formula %s!%s: %w", update.sheet, update.cell, err)
		}
		return nil
	}

	switch update.typ {
	case "text":
		if err := clearFormula(); err != nil {
			return err
		}
		if err := f.SetCellStr(update.sheet, update.cell, update.text); err != nil {
			return fmt.Errorf("set text %s!%s: %w", update.sheet, update.cell, err)
		}
	case "number":
		if err := clearFormula(); err != nil {
			return err
		}
		if err := f.SetCellValue(update.sheet, update.cell, update.number); err != nil {
			return fmt.Errorf("set number %s!%s: %w", update.sheet, update.cell, err)
		}
	case "bool":
		if err := clearFormula(); err != nil {
			return err
		}
		if err := f.SetCellValue(update.sheet, update.cell, update.boolean); err != nil {
			return fmt.Errorf("set bool %s!%s: %w", update.sheet, update.cell, err)
		}
	case "formula":
		if err := f.SetCellFormula(update.sheet, update.cell, update.text); err != nil {
			return fmt.Errorf("set formula %s!%s: %w", update.sheet, update.cell, err)
		}
	case "blank":
		if err := clearFormula(); err != nil {
			return err
		}
		if err := f.SetCellValue(update.sheet, update.cell, nil); err != nil {
			return fmt.Errorf("clear value %s!%s: %w", update.sheet, update.cell, err)
		}
	}
	return nil
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

func samePath(a, b string) (bool, error) {
	absA, err := filepath.Abs(a)
	if err != nil {
		return false, fmt.Errorf("resolve input path: %w", err)
	}
	absB, err := filepath.Abs(b)
	if err != nil {
		return false, fmt.Errorf("resolve output path: %w", err)
	}
	return filepath.Clean(absA) == filepath.Clean(absB), nil
}
