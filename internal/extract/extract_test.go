package extract

import (
	"errors"
	"path/filepath"
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
