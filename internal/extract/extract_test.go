package extract

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestNormalizeColumn(t *testing.T) {
	tests := []struct {
		name      string
		input     string
		wantNum   int
		wantName  string
		wantError bool
	}{
		{name: "letters", input: "A", wantNum: 1, wantName: "A"},
		{name: "lower letters", input: "bc", wantNum: 55, wantName: "BC"},
		{name: "number", input: "27", wantNum: 27, wantName: "AA"},
		{name: "spaces", input: " 3 ", wantNum: 3, wantName: "C"},
		{name: "empty", input: "", wantError: true},
		{name: "zero", input: "0", wantError: true},
		{name: "invalid", input: "A1", wantError: true},
		{name: "beyond limit", input: "XFE", wantError: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotNum, gotName, err := NormalizeColumn(tt.input)
			if tt.wantError {
				if err == nil {
					t.Fatalf("expected error")
				}
				return
			}
			if err != nil {
				t.Fatalf("NormalizeColumn returned error: %v", err)
			}
			if gotNum != tt.wantNum || gotName != tt.wantName {
				t.Fatalf("got (%d, %q), want (%d, %q)", gotNum, gotName, tt.wantNum, tt.wantName)
			}
		})
	}
}

func TestWorkbookExtraction(t *testing.T) {
	path := writeTestWorkbook(t)

	sheets, err := ListSheets(path)
	if err != nil {
		t.Fatalf("ListSheets returned error: %v", err)
	}
	if len(sheets.Sheets) != 2 || sheets.Sheets[0] != "Data" || sheets.Sheets[1] != "SR1-1" {
		t.Fatalf("unexpected sheets: %#v", sheets.Sheets)
	}

	cell, err := ExtractCell(path, "Data", 2, "C")
	if err != nil {
		t.Fatalf("ExtractCell returned error: %v", err)
	}
	if cell.Col != "C" || cell.Value != "42.5" {
		t.Fatalf("unexpected cell result: %#v", cell)
	}

	empty, err := ExtractCell(path, "Data", 2, "B")
	if err != nil {
		t.Fatalf("ExtractCell empty returned error: %v", err)
	}
	if empty.Value != "" {
		t.Fatalf("empty cell value = %q, want empty string", empty.Value)
	}

	row, err := ExtractRow(path, "Data", 2, "A", "C")
	if err != nil {
		t.Fatalf("ExtractRow returned error: %v", err)
	}
	if len(row.Cells) != 3 {
		t.Fatalf("row has %d cells, want 3: %#v", len(row.Cells), row.Cells)
	}
	if row.Cells[0].Value != "alpha" || row.Cells[1].Value != "" || row.Cells[2].Value != "42.5" {
		t.Fatalf("unexpected row cells: %#v", row.Cells)
	}

	wideRow, err := ExtractRow(path, "Data", 1, "", "")
	if err != nil {
		t.Fatalf("ExtractRow wide returned error: %v", err)
	}
	if len(wideRow.Cells) != 4 {
		t.Fatalf("wide row has %d cells, want 4: %#v", len(wideRow.Cells), wideRow.Cells)
	}
	if wideRow.Cells[3].Col != "D" || wideRow.Cells[3].Value != "" {
		t.Fatalf("trailing empty column not preserved: %#v", wideRow.Cells)
	}
}

func TestFormulaExtraction(t *testing.T) {
	path := writeFormulaWorkbook(t)

	cell, err := ExtractCell(path, "Data", 2, "C")
	if err != nil {
		t.Fatalf("ExtractCell formula returned error: %v", err)
	}
	if cell.Value != "30" {
		t.Fatalf("formula cell value = %q, want 30", cell.Value)
	}

	formatted, err := ExtractCell(path, "Summary", 2, "B")
	if err != nil {
		t.Fatalf("ExtractCell formatted formula returned error: %v", err)
	}
	if formatted.Value != "30.00" {
		t.Fatalf("formatted formula value = %q, want 30.00", formatted.Value)
	}

	row, err := ExtractRow(path, "Summary", 2, "A", "B")
	if err != nil {
		t.Fatalf("ExtractRow formula returned error: %v", err)
	}
	if len(row.Cells) != 2 || row.Cells[1].Value != "30.00" {
		t.Fatalf("unexpected formula row: %#v", row.Cells)
	}
}

func TestExtractionErrors(t *testing.T) {
	path := writeTestWorkbook(t)

	if _, err := ExtractCell(path, "Missing", 1, "A"); !errors.Is(err, ErrSheetNotFound) {
		t.Fatalf("missing sheet error = %v, want ErrSheetNotFound", err)
	}
	if _, err := ExtractCell(path, "Data", 0, "A"); err == nil {
		t.Fatalf("expected row validation error")
	}
	if _, err := ExtractCell(path, "Data", 1, "0"); err == nil {
		t.Fatalf("expected column validation error")
	}
	if _, err := ListSheets(filepath.Join(t.TempDir(), "book.xls")); !errors.Is(err, ErrUnsupportedFile) {
		t.Fatalf("unsupported file error = %v, want ErrUnsupportedFile", err)
	}
}

func TestWorkbookInfoAndAdditionalReads(t *testing.T) {
	path := writeTestWorkbook(t)

	info, err := InspectWorkbook(path)
	if err != nil {
		t.Fatalf("InspectWorkbook returned error: %v", err)
	}
	if len(info.Sheets) != 2 || info.Sheets[0].Name != "Data" || info.Sheets[0].MaxRow != 2 || info.Sheets[0].MaxCol != 4 {
		t.Fatalf("unexpected workbook info: %#v", info)
	}

	col, err := ExtractCol(path, "Data", "C", 1, 2)
	if err != nil {
		t.Fatalf("ExtractCol returned error: %v", err)
	}
	if col.Col != "C" || len(col.Cells) != 2 || col.Cells[0].Value != "Amount" || col.Cells[1].Value != "42.5" {
		t.Fatalf("unexpected col result: %#v", col)
	}

	rng, err := ExtractRange(path, "Data", "A1:C2")
	if err != nil {
		t.Fatalf("ExtractRange returned error: %v", err)
	}
	if rng.Range != "A1:C2" || len(rng.Cells) != 6 {
		t.Fatalf("unexpected range result: %#v", rng)
	}
	if rng.Cells[0].Value != "Name" || rng.Cells[4].Value != "" || rng.Cells[5].Value != "42.5" {
		t.Fatalf("unexpected range cells: %#v", rng.Cells)
	}

	batch, err := ReadCells(path, ReadBatchRequest{Queries: []ReadQuery{
		{Sheet: "Data", Row: 2, Col: "A"},
		{Sheet: "Data", Row: 2, Col: "C"},
	}})
	if err != nil {
		t.Fatalf("ReadCells returned error: %v", err)
	}
	if len(batch.Cells) != 2 || batch.Cells[0].Col != "A" || batch.Cells[0].Value != "alpha" || batch.Cells[1].Value != "42.5" {
		t.Fatalf("unexpected batch result: %#v", batch)
	}
}

func TestRangeReadValidationErrors(t *testing.T) {
	path := writeTestWorkbook(t)

	tests := []string{"", "Data!A1:B2", "C2:A1", "A0", "XFE1"}
	for _, rangeRef := range tests {
		t.Run(rangeRef, func(t *testing.T) {
			if _, err := ExtractRange(path, "Data", rangeRef); err == nil {
				t.Fatalf("expected range error for %q", rangeRef)
			}
		})
	}
}

func TestWriteHelpers(t *testing.T) {
	path := writeTestWorkbook(t)

	if _, err := WriteCell(path, "Data", 2, "B", FillValue{Type: "text", Value: strPtr("single")}, "", false); err != nil {
		t.Fatalf("WriteCell returned error: %v", err)
	}
	cell, err := ExtractCell(path, "Data", 2, "B")
	if err != nil {
		t.Fatalf("ExtractCell after WriteCell: %v", err)
	}
	if cell.Value != "single" {
		t.Fatalf("WriteCell value = %q, want single", cell.Value)
	}

	if _, err := WriteRow(path, "Data", 3, "A", ValuesRequest{Values: []FillValue{
		{Type: "text", Value: strPtr("row-a")},
		{Type: "number", Value: strPtr("12")},
	}}, "", false); err != nil {
		t.Fatalf("WriteRow returned error: %v", err)
	}
	row, err := ExtractRow(path, "Data", 3, "A", "B")
	if err != nil {
		t.Fatalf("ExtractRow after WriteRow: %v", err)
	}
	if row.Cells[0].Value != "row-a" || row.Cells[1].Value != "12" {
		t.Fatalf("unexpected WriteRow cells: %#v", row.Cells)
	}

	if _, err := WriteCol(path, "Data", "D", 3, ValuesRequest{Values: []FillValue{
		{Type: "bool", Value: strPtr("true")},
		{Type: "text", Value: strPtr("col-d")},
	}}, "", false); err != nil {
		t.Fatalf("WriteCol returned error: %v", err)
	}
	col, err := ExtractCol(path, "Data", "D", 3, 4)
	if err != nil {
		t.Fatalf("ExtractCol after WriteCol: %v", err)
	}
	if !strings.EqualFold(col.Cells[0].Value, "true") || col.Cells[1].Value != "col-d" {
		t.Fatalf("unexpected WriteCol cells: %#v", col.Cells)
	}

	if _, err := WriteRange(path, "Data", "B5:C6", RangeValuesRequest{Rows: [][]FillValue{
		{{Type: "text", Value: strPtr("r1c1")}, {Type: "text", Value: strPtr("r1c2")}},
		{{Type: "number", Value: strPtr("21")}, {Type: "formula", Value: strPtr("=B6+1")}},
	}}, "", false); err != nil {
		t.Fatalf("WriteRange returned error: %v", err)
	}
	rng, err := ExtractRange(path, "Data", "B5:C6")
	if err != nil {
		t.Fatalf("ExtractRange after WriteRange: %v", err)
	}
	if rng.Cells[0].Value != "r1c1" || rng.Cells[1].Value != "r1c2" || rng.Cells[2].Value != "21" || rng.Cells[3].Value != "22" {
		t.Fatalf("unexpected WriteRange cells: %#v", rng.Cells)
	}

	if _, err := ClearRange(path, "Data", "B5:C5", "", false); err != nil {
		t.Fatalf("ClearRange returned error: %v", err)
	}
	cleared, err := ExtractRange(path, "Data", "B5:C5")
	if err != nil {
		t.Fatalf("ExtractRange after ClearRange: %v", err)
	}
	if cleared.Cells[0].Value != "" || cleared.Cells[1].Value != "" {
		t.Fatalf("unexpected cleared cells: %#v", cleared.Cells)
	}
}

func TestWriteRangeValidationDoesNotSave(t *testing.T) {
	path := writeTestWorkbook(t)

	_, err := WriteRange(path, "Data", "A2:B2", RangeValuesRequest{Rows: [][]FillValue{
		{{Type: "text", Value: strPtr("changed")}},
	}}, "", false)
	if err == nil {
		t.Fatalf("expected shape validation error")
	}

	cell, err := ExtractCell(path, "Data", 2, "A")
	if err != nil {
		t.Fatalf("ExtractCell after failed WriteRange: %v", err)
	}
	if cell.Value != "alpha" {
		t.Fatalf("source value after failed WriteRange = %q, want alpha", cell.Value)
	}
}

func TestFillCellsWritesMultipleTypesInPlace(t *testing.T) {
	path := writeFillWorkbook(t)

	result, err := FillCells(path, FillRequest{
		Updates: []FillUpdate{
			{Sheet: "Data", Row: 3, Col: "A", Type: "text", Value: strPtr("已确认")},
			{Sheet: "Data", Row: 3, Col: "B", Type: "number", Value: strPtr("123.45")},
			{Sheet: "Data", Row: 3, Col: "C", Type: "bool", Value: strPtr("true")},
			{Sheet: "Data", Row: 3, Col: "D", Type: "formula", Value: strPtr("=SUM(B3,1)")},
			{Sheet: "Data", Row: 2, Col: "B", Type: "blank"},
		},
	}, "", false)
	if err != nil {
		t.Fatalf("FillCells returned error: %v", err)
	}
	if result.File != path || result.Output != path || !result.InPlace || result.Updated != 5 {
		t.Fatalf("unexpected fill result: %#v", result)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f.Close()

	assertCellValue(t, f, "Data", "A3", "已确认")
	assertCellValue(t, f, "Data", "B3", "123.45")
	boolValue, err := f.GetCellValue("Data", "C3")
	if err != nil {
		t.Fatalf("GetCellValue C3: %v", err)
	}
	if !strings.EqualFold(boolValue, "true") {
		t.Fatalf("bool cell value = %q, want true", boolValue)
	}
	formula, err := f.GetCellFormula("Data", "D3")
	if err != nil {
		t.Fatalf("GetCellFormula D3: %v", err)
	}
	if formula != "=SUM(B3,1)" {
		t.Fatalf("formula = %q, want =SUM(B3,1)", formula)
	}
	assertCellValue(t, f, "Data", "B2", "")
	formula, err = f.GetCellFormula("Data", "B2")
	if err != nil {
		t.Fatalf("GetCellFormula B2: %v", err)
	}
	if formula != "" {
		t.Fatalf("blank cell formula = %q, want empty", formula)
	}
	styleID, err := f.GetCellStyle("Data", "B2")
	if err != nil {
		t.Fatalf("GetCellStyle B2: %v", err)
	}
	if styleID == 0 {
		t.Fatalf("blank should preserve existing style")
	}
}

func TestFillCellsOutputKeepsSourceUnchanged(t *testing.T) {
	path := writeTestWorkbook(t)
	output := filepath.Join(t.TempDir(), "filled.xlsx")

	result, err := FillCells(path, FillRequest{
		Updates: []FillUpdate{
			{Sheet: "Data", Row: 2, Col: "A", Type: "text", Value: strPtr("changed")},
		},
	}, output, false)
	if err != nil {
		t.Fatalf("FillCells returned error: %v", err)
	}
	if result.Output != output || result.InPlace {
		t.Fatalf("unexpected fill result: %#v", result)
	}

	sourceCell, err := ExtractCell(path, "Data", 2, "A")
	if err != nil {
		t.Fatalf("ExtractCell source: %v", err)
	}
	if sourceCell.Value != "alpha" {
		t.Fatalf("source value = %q, want alpha", sourceCell.Value)
	}
	outputCell, err := ExtractCell(output, "Data", 2, "A")
	if err != nil {
		t.Fatalf("ExtractCell output: %v", err)
	}
	if outputCell.Value != "changed" {
		t.Fatalf("output value = %q, want changed", outputCell.Value)
	}
}

func TestFillCellsValidationFailureDoesNotSave(t *testing.T) {
	path := writeTestWorkbook(t)

	_, err := FillCells(path, FillRequest{
		Updates: []FillUpdate{
			{Sheet: "Data", Row: 2, Col: "A", Type: "text", Value: strPtr("changed")},
			{Sheet: "Data", Row: 2, Col: "C", Type: "number", Value: strPtr("not-a-number")},
		},
	}, "", false)
	if err == nil {
		t.Fatalf("expected validation error")
	}

	cell, err := ExtractCell(path, "Data", 2, "A")
	if err != nil {
		t.Fatalf("ExtractCell after failed fill: %v", err)
	}
	if cell.Value != "alpha" {
		t.Fatalf("source value after failed fill = %q, want alpha", cell.Value)
	}
}

func TestFillCellsValidationErrors(t *testing.T) {
	path := writeTestWorkbook(t)
	existingOutput := filepath.Join(t.TempDir(), "existing.xlsx")
	if err := os.WriteFile(existingOutput, []byte("already here"), 0o644); err != nil {
		t.Fatalf("WriteFile existing output: %v", err)
	}

	tests := []struct {
		name    string
		request FillRequest
		output  string
	}{
		{
			name:    "empty updates",
			request: FillRequest{},
		},
		{
			name: "missing sheet",
			request: FillRequest{Updates: []FillUpdate{
				{Row: 1, Col: "A", Type: "text", Value: strPtr("x")},
			}},
		},
		{
			name: "unknown sheet",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Missing", Row: 1, Col: "A", Type: "text", Value: strPtr("x")},
			}},
		},
		{
			name: "invalid row",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 0, Col: "A", Type: "text", Value: strPtr("x")},
			}},
		},
		{
			name: "invalid column",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "0", Type: "text", Value: strPtr("x")},
			}},
		},
		{
			name: "invalid type",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "A", Type: "date", Value: strPtr("x")},
			}},
		},
		{
			name: "missing value",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "A", Type: "text"},
			}},
		},
		{
			name: "invalid number",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "A", Type: "number", Value: strPtr("NaN")},
			}},
		},
		{
			name: "invalid bool",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "A", Type: "bool", Value: strPtr("yes")},
			}},
		},
		{
			name: "formula without equals",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "A", Type: "formula", Value: strPtr("SUM(A1:A2)")},
			}},
		},
		{
			name: "blank with value",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "A", Type: "blank", Value: strPtr("")},
			}},
		},
		{
			name: "existing output without overwrite",
			request: FillRequest{Updates: []FillUpdate{
				{Sheet: "Data", Row: 1, Col: "A", Type: "text", Value: strPtr("x")},
			}},
			output: existingOutput,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if _, err := FillCells(path, tt.request, tt.output, false); err == nil {
				t.Fatalf("expected error")
			}
		})
	}
}

func TestFillCellsOverwriteExistingOutput(t *testing.T) {
	path := writeTestWorkbook(t)
	output := writeTestWorkbook(t)

	_, err := FillCells(path, FillRequest{
		Updates: []FillUpdate{
			{Sheet: "Data", Row: 2, Col: "A", Type: "text", Value: strPtr("overwritten")},
		},
	}, output, true)
	if err != nil {
		t.Fatalf("FillCells overwrite returned error: %v", err)
	}

	cell, err := ExtractCell(output, "Data", 2, "A")
	if err != nil {
		t.Fatalf("ExtractCell output: %v", err)
	}
	if cell.Value != "overwritten" {
		t.Fatalf("output value = %q, want overwritten", cell.Value)
	}
}

func writeTestWorkbook(t *testing.T) string {
	t.Helper()

	f := excelize.NewFile()
	defer f.Close()

	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("SR1-1"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	values := map[string]any{
		"A1": "Name",
		"C1": "Amount",
		"A2": "alpha",
		"C2": 42.5,
		"D2": "extends sheet width",
	}
	for cell, value := range values {
		if err := f.SetCellValue("Data", cell, value); err != nil {
			t.Fatalf("SetCellValue %s: %v", cell, err)
		}
	}

	path := filepath.Join(t.TempDir(), "book.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	return path
}

func writeFillWorkbook(t *testing.T) string {
	t.Helper()

	f := excelize.NewFile()
	defer f.Close()

	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if err := f.SetCellValue("Data", "A2", 10); err != nil {
		t.Fatalf("SetCellValue A2: %v", err)
	}
	if err := f.SetCellFormula("Data", "B2", "=A2+1"); err != nil {
		t.Fatalf("SetCellFormula B2: %v", err)
	}
	style, err := f.NewStyle(&excelize.Style{NumFmt: 2})
	if err != nil {
		t.Fatalf("NewStyle: %v", err)
	}
	if err := f.SetCellStyle("Data", "B2", "B2", style); err != nil {
		t.Fatalf("SetCellStyle B2: %v", err)
	}

	path := filepath.Join(t.TempDir(), "fill.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	return path
}

func writeFormulaWorkbook(t *testing.T) string {
	t.Helper()

	f := excelize.NewFile()
	defer f.Close()

	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Summary"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	values := map[string]any{
		"A2": 10,
		"B2": 20,
	}
	for cell, value := range values {
		if err := f.SetCellValue("Data", cell, value); err != nil {
			t.Fatalf("SetCellValue %s: %v", cell, err)
		}
	}
	if err := f.SetCellFormula("Data", "C2", "A2+B2"); err != nil {
		t.Fatalf("SetCellFormula Data!C2: %v", err)
	}
	if err := f.SetCellValue("Summary", "A2", "total"); err != nil {
		t.Fatalf("SetCellValue Summary!A2: %v", err)
	}
	if err := f.SetCellFormula("Summary", "B2", "'Data'!C2"); err != nil {
		t.Fatalf("SetCellFormula Summary!B2: %v", err)
	}
	style, err := f.NewStyle(&excelize.Style{NumFmt: 2})
	if err != nil {
		t.Fatalf("NewStyle: %v", err)
	}
	if err := f.SetCellStyle("Summary", "B2", "B2", style); err != nil {
		t.Fatalf("SetCellStyle: %v", err)
	}

	path := filepath.Join(t.TempDir(), "formula.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	return path
}

func assertCellValue(t *testing.T, f *excelize.File, sheet, cell, want string) {
	t.Helper()

	value, err := f.GetCellValue(sheet, cell)
	if err != nil {
		t.Fatalf("GetCellValue %s!%s: %v", sheet, cell, err)
	}
	if value != want {
		t.Fatalf("%s!%s = %q, want %q", sheet, cell, value, want)
	}
}

func strPtr(value string) *string {
	return &value
}
