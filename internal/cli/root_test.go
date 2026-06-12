package cli

import (
	"bytes"
	"encoding/json"
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/cligrep/cli-excel-extract/internal/extract"
	"github.com/xuri/excelize/v2"
)

func TestCellCommandJSON(t *testing.T) {
	path := writeTestWorkbook(t)

	stdout, _, err := executeCommand("cell", "-f", path, "-s", "Data", "-r", "1", "-c", "B", "--json")
	if err != nil {
		t.Fatalf("command returned error: %v", err)
	}

	var result extract.CellResult
	if err := json.Unmarshal([]byte(stdout), &result); err != nil {
		t.Fatalf("json unmarshal: %v; output=%q", err, stdout)
	}
	if result.File != path || result.Sheet != "Data" || result.Row != 1 || result.Col != "B" || result.Value != "指标名称" {
		t.Fatalf("unexpected result: %#v", result)
	}
}

func TestRowCommandTextAndJSON(t *testing.T) {
	path := writeTestWorkbook(t)

	stdout, _, err := executeCommand("row", "-f", path, "-s", "Data", "-r", "1", "--from-col", "A", "--to-col", "C")
	if err != nil {
		t.Fatalf("text command returned error: %v", err)
	}
	if stdout != "指标序号\t指标名称\t监管标准\n" {
		t.Fatalf("stdout = %q", stdout)
	}

	stdout, _, err = executeCommand("row", "-f", path, "-s", "Data", "-r", "2", "--from-col", "A", "--to-col", "C", "--json")
	if err != nil {
		t.Fatalf("json command returned error: %v", err)
	}

	var result extract.RowResult
	if err := json.Unmarshal([]byte(stdout), &result); err != nil {
		t.Fatalf("json unmarshal: %v; output=%q", err, stdout)
	}
	if len(result.Cells) != 3 || result.Cells[1].Col != "B" || result.Cells[1].Value != "" {
		t.Fatalf("unexpected JSON row result: %#v", result)
	}
}

func TestRowCommandCalculatesFormulaJSON(t *testing.T) {
	path := writeFormulaWorkbook(t)

	stdout, _, err := executeCommand("row", "-f", path, "-s", "Summary", "-r", "2", "--from-col", "A", "--to-col", "B", "--json")
	if err != nil {
		t.Fatalf("command returned error: %v", err)
	}

	var result extract.RowResult
	if err := json.Unmarshal([]byte(stdout), &result); err != nil {
		t.Fatalf("json unmarshal: %v; output=%q", err, stdout)
	}
	if len(result.Cells) != 2 || result.Cells[1].Value != "30.00" {
		t.Fatalf("unexpected formula JSON row result: %#v", result)
	}
}

func TestSheetsCommand(t *testing.T) {
	path := writeTestWorkbook(t)

	stdout, _, err := executeCommand("sheets", "-f", path)
	if err != nil {
		t.Fatalf("command returned error: %v", err)
	}
	if !strings.Contains(stdout, "Data\n") || !strings.Contains(stdout, "校验表\n") {
		t.Fatalf("unexpected sheets output: %q", stdout)
	}
}

func TestVersionCommand(t *testing.T) {
	stdout, _, err := executeCommand("--version")
	if err != nil {
		t.Fatalf("version command returned error: %v", err)
	}
	assertContainsAll(t, stdout, "excel-extract", "version=", "commit=", "date=", "platform=")
}

func TestRootHelpIsSelfContained(t *testing.T) {
	stdout, _, err := executeCommand("--help")
	if err != nil {
		t.Fatalf("help command returned error: %v", err)
	}

	assertContainsAll(t, stdout,
		".xlsx",
		".xlsm",
		"Formula",
		"Examples:",
		"sheets",
		"cell",
		"row",
		"--json",
		"--version",
	)
	if strings.Contains(stdout, "completion") {
		t.Fatalf("root help should hide completion command, got:\n%s", stdout)
	}
}

func TestCellHelpIsSelfContained(t *testing.T) {
	stdout, _, err := executeCommand("cell", "--help")
	if err != nil {
		t.Fatalf("cell help returned error: %v", err)
	}

	assertContainsAll(t, stdout,
		"-f, --file",
		"-s, --sheet",
		"-r, --row",
		"-c, --col",
		"--json",
		"Formula",
		`"value"`,
		"Examples:",
	)
}

func TestRowHelpIsSelfContained(t *testing.T) {
	stdout, _, err := executeCommand("row", "--help")
	if err != nil {
		t.Fatalf("row help returned error: %v", err)
	}

	assertContainsAll(t, stdout,
		"--from-col",
		"--to-col",
		"tab",
		"Empty cells",
		`"cells"`,
		"Examples:",
	)
}

func TestCommandsRejectExtraArgs(t *testing.T) {
	path := writeTestWorkbook(t)

	tests := [][]string{
		{"sheets", "extra", "-f", path},
		{"cell", "extra", "-f", path, "-s", "Data", "-r", "1", "-c", "A"},
		{"row", "extra", "-f", path, "-s", "Data", "-r", "1"},
	}

	for _, args := range tests {
		t.Run(strings.Join(args[:1], " "), func(t *testing.T) {
			if _, _, err := executeCommand(args...); err == nil {
				t.Fatalf("expected extra arg error for args: %v", args)
			}
		})
	}
}

func executeCommand(args ...string) (string, string, error) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer
	cmd := NewRootCommand(&stdout, &stderr)
	cmd.SetArgs(args)
	err := cmd.Execute()
	return stdout.String(), stderr.String(), err
}

func assertContainsAll(t *testing.T, text string, terms ...string) {
	t.Helper()

	for _, term := range terms {
		if !strings.Contains(text, term) {
			t.Fatalf("%s\nmissing %q in:\n%s", fmt.Sprintf("expected text to contain %d terms", len(terms)), term, text)
		}
	}
}

func writeTestWorkbook(t *testing.T) string {
	t.Helper()

	f := excelize.NewFile()
	defer f.Close()

	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("校验表"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	values := map[string]any{
		"A1": "指标序号",
		"B1": "指标名称",
		"C1": "监管标准",
		"A2": 1,
		"C2": "随机备注",
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
	if err := f.SetCellValue("Data", "A2", 10); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellValue("Data", "B2", 20); err != nil {
		t.Fatalf("SetCellValue Data!B2: %v", err)
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
