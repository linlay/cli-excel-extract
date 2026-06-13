package cli

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/cligrep/excelx/internal/extract"
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

func TestFillCommandFromFileJSON(t *testing.T) {
	path := writeTestWorkbook(t)
	output := filepath.Join(t.TempDir(), "filled.xlsx")
	updates := writeUpdatesFile(t, `{"updates":[
		{"sheet":"Data","row":2,"col":"B","type":"text","value":"filled"},
		{"sheet":"Data","row":2,"col":"D","type":"number","value":"99.5"}
	]}`)

	stdout, _, err := executeCommand("fill", "-f", path, "--updates", updates, "--output", output, "--json")
	if err != nil {
		t.Fatalf("fill command returned error: %v", err)
	}

	var result extract.FillResult
	if err := json.Unmarshal([]byte(stdout), &result); err != nil {
		t.Fatalf("json unmarshal: %v; output=%q", err, stdout)
	}
	if result.File != path || result.Output != output || result.InPlace || result.Updated != 2 {
		t.Fatalf("unexpected fill result: %#v", result)
	}

	sourceCell, err := extract.ExtractCell(path, "Data", 2, "B")
	if err != nil {
		t.Fatalf("ExtractCell source: %v", err)
	}
	if sourceCell.Value != "" {
		t.Fatalf("source value = %q, want empty", sourceCell.Value)
	}
	outputCell, err := extract.ExtractCell(output, "Data", 2, "B")
	if err != nil {
		t.Fatalf("ExtractCell output: %v", err)
	}
	if outputCell.Value != "filled" {
		t.Fatalf("output value = %q, want filled", outputCell.Value)
	}
}

func TestFillCommandFromStdinText(t *testing.T) {
	path := writeTestWorkbook(t)
	input := `{"updates":[{"sheet":"Data","row":2,"col":"B","type":"text","value":"stdin"}]}`

	stdout, _, err := executeCommandWithInput(input, "fill", "-f", path, "--updates", "-")
	if err != nil {
		t.Fatalf("fill command returned error: %v", err)
	}
	if stdout != fmt.Sprintf("updated %s: 1 cells\n", path) {
		t.Fatalf("stdout = %q", stdout)
	}

	cell, err := extract.ExtractCell(path, "Data", 2, "B")
	if err != nil {
		t.Fatalf("ExtractCell: %v", err)
	}
	if cell.Value != "stdin" {
		t.Fatalf("cell value = %q, want stdin", cell.Value)
	}
}

func TestReadCommands(t *testing.T) {
	path := writeTestWorkbook(t)

	oldCell, _, err := executeCommand("cell", "-f", path, "-s", "Data", "-r", "1", "-c", "B", "--json")
	if err != nil {
		t.Fatalf("legacy cell returned error: %v", err)
	}
	readCell, _, err := executeCommand("read", "cell", "-f", path, "-s", "Data", "-r", "1", "-c", "B", "--json")
	if err != nil {
		t.Fatalf("read cell returned error: %v", err)
	}
	if readCell != oldCell {
		t.Fatalf("read cell output differs from legacy cell:\nold=%s\nnew=%s", oldCell, readCell)
	}

	stdout, _, err := executeCommand("read", "info", "-f", path)
	if err != nil {
		t.Fatalf("read info returned error: %v", err)
	}
	assertContainsAll(t, stdout, "Data", "校验表")

	stdout, _, err = executeCommand("read", "col", "-f", path, "-s", "Data", "-c", "C", "--from-row", "1", "--to-row", "2")
	if err != nil {
		t.Fatalf("read col returned error: %v", err)
	}
	if stdout != "监管标准\n随机备注\n" {
		t.Fatalf("read col stdout = %q", stdout)
	}

	stdout, _, err = executeCommand("read", "range", "-f", path, "-s", "Data", "--range", "A1:C2")
	if err != nil {
		t.Fatalf("read range returned error: %v", err)
	}
	if stdout != "指标序号\t指标名称\t监管标准\n1\t\t随机备注\n" {
		t.Fatalf("read range stdout = %q", stdout)
	}

	queries := writeJSONFile(t, "queries.json", `{"queries":[
		{"sheet":"Data","row":1,"col":"B"},
		{"sheet":"Data","row":2,"col":"C"}
	]}`)
	stdout, _, err = executeCommand("read", "batch", "-f", path, "--queries", queries, "--json")
	if err != nil {
		t.Fatalf("read batch returned error: %v", err)
	}
	var batch extract.ReadBatchResult
	if err := json.Unmarshal([]byte(stdout), &batch); err != nil {
		t.Fatalf("json unmarshal read batch: %v; output=%q", err, stdout)
	}
	if len(batch.Cells) != 2 || batch.Cells[0].Value != "指标名称" || batch.Cells[1].Value != "随机备注" {
		t.Fatalf("unexpected read batch result: %#v", batch)
	}
}

func TestWriteCommands(t *testing.T) {
	path := writeTestWorkbook(t)

	stdout, _, err := executeCommand("write", "cell", "-f", path, "-s", "Data", "-r", "2", "-c", "B", "--type", "text", "--value", "single")
	if err != nil {
		t.Fatalf("write cell returned error: %v", err)
	}
	if stdout != fmt.Sprintf("updated %s: 1 cells\n", path) {
		t.Fatalf("write cell stdout = %q", stdout)
	}
	cell, err := extract.ExtractCell(path, "Data", 2, "B")
	if err != nil {
		t.Fatalf("ExtractCell after write cell: %v", err)
	}
	if cell.Value != "single" {
		t.Fatalf("write cell value = %q, want single", cell.Value)
	}

	output := filepath.Join(t.TempDir(), "row-output.xlsx")
	values := writeJSONFile(t, "values.json", `{"values":[
		{"type":"text","value":"row-a"},
		{"type":"number","value":"12"}
	]}`)
	stdout, _, err = executeCommand("write", "row", "-f", path, "-s", "Data", "-r", "3", "--from-col", "A", "--values", values, "--output", output, "--json")
	if err != nil {
		t.Fatalf("write row returned error: %v", err)
	}
	var rowResult extract.FillResult
	if err := json.Unmarshal([]byte(stdout), &rowResult); err != nil {
		t.Fatalf("json unmarshal write row: %v; output=%q", err, stdout)
	}
	if rowResult.Output != output || rowResult.InPlace || rowResult.Updated != 2 {
		t.Fatalf("unexpected write row result: %#v", rowResult)
	}
	row, err := extract.ExtractRow(output, "Data", 3, "A", "B")
	if err != nil {
		t.Fatalf("ExtractRow after write row: %v", err)
	}
	if row.Cells[0].Value != "row-a" || row.Cells[1].Value != "12" {
		t.Fatalf("unexpected write row cells: %#v", row.Cells)
	}

	colValues := writeJSONFile(t, "col-values.json", `{"values":[
		{"type":"bool","value":"true"},
		{"type":"text","value":"col-b"}
	]}`)
	if _, _, err := executeCommand("write", "col", "-f", path, "-s", "Data", "-c", "D", "--from-row", "3", "--values", colValues); err != nil {
		t.Fatalf("write col returned error: %v", err)
	}
	col, err := extract.ExtractCol(path, "Data", "D", 3, 4)
	if err != nil {
		t.Fatalf("ExtractCol after write col: %v", err)
	}
	if !strings.EqualFold(col.Cells[0].Value, "true") || col.Cells[1].Value != "col-b" {
		t.Fatalf("unexpected write col cells: %#v", col.Cells)
	}

	rangeValues := writeJSONFile(t, "range-values.json", `{"rows":[
		[{"type":"text","value":"r1c1"},{"type":"text","value":"r1c2"}],
		[{"type":"number","value":"21"},{"type":"formula","value":"=B6+1"}]
	]}`)
	if _, _, err := executeCommand("write", "range", "-f", path, "-s", "Data", "--range", "B5:C6", "--values", rangeValues); err != nil {
		t.Fatalf("write range returned error: %v", err)
	}
	rng, err := extract.ExtractRange(path, "Data", "B5:C6")
	if err != nil {
		t.Fatalf("ExtractRange after write range: %v", err)
	}
	if rng.Cells[0].Value != "r1c1" || rng.Cells[3].Value != "22" {
		t.Fatalf("unexpected write range cells: %#v", rng.Cells)
	}

	if _, _, err := executeCommand("write", "clear", "-f", path, "-s", "Data", "--range", "B5:C5"); err != nil {
		t.Fatalf("write clear returned error: %v", err)
	}
	cleared, err := extract.ExtractRange(path, "Data", "B5:C5")
	if err != nil {
		t.Fatalf("ExtractRange after write clear: %v", err)
	}
	if cleared.Cells[0].Value != "" || cleared.Cells[1].Value != "" {
		t.Fatalf("unexpected cleared cells: %#v", cleared.Cells)
	}

	updates := writeUpdatesFile(t, `{"updates":[{"sheet":"Data","row":8,"col":"A","type":"text","value":"batch"}]}`)
	if _, _, err := executeCommand("write", "batch", "-f", path, "--updates", updates); err != nil {
		t.Fatalf("write batch returned error: %v", err)
	}
	batchCell, err := extract.ExtractCell(path, "Data", 8, "A")
	if err != nil {
		t.Fatalf("ExtractCell after write batch: %v", err)
	}
	if batchCell.Value != "batch" {
		t.Fatalf("write batch value = %q, want batch", batchCell.Value)
	}
}

func TestVersionCommand(t *testing.T) {
	stdout, _, err := executeCommand("--version")
	if err != nil {
		t.Fatalf("version command returned error: %v", err)
	}
	assertContainsAll(t, stdout, "excelx", "version=", "commit=", "date=", "platform=")
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
		"fill",
		"read",
		"write",
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

func TestFillHelpIsSelfContained(t *testing.T) {
	stdout, _, err := executeCommand("fill", "--help")
	if err != nil {
		t.Fatalf("fill help returned error: %v", err)
	}

	assertContainsAll(t, stdout,
		"-f, --file",
		"--updates",
		"--output",
		"--overwrite",
		"--json",
		"in place",
		"text",
		"number",
		"bool",
		"formula",
		"blank",
		"Examples:",
	)
}

func TestReadWriteHelpIsSelfContained(t *testing.T) {
	tests := [][]string{
		{"read", "--help"},
		{"read", "batch", "--help"},
		{"write", "--help"},
		{"write", "cell", "--help"},
		{"write", "range", "--help"},
		{"write", "batch", "--help"},
	}

	for _, args := range tests {
		t.Run(strings.Join(args, " "), func(t *testing.T) {
			stdout, _, err := executeCommand(args...)
			if err != nil {
				t.Fatalf("%v returned error: %v", args, err)
			}
			assertContainsAll(t, stdout, "Examples:")
		})
	}
}

func TestCommandsRejectExtraArgs(t *testing.T) {
	path := writeTestWorkbook(t)
	updates := writeUpdatesFile(t, `{"updates":[{"sheet":"Data","row":1,"col":"A","type":"text","value":"x"}]}`)

	tests := [][]string{
		{"sheets", "extra", "-f", path},
		{"cell", "extra", "-f", path, "-s", "Data", "-r", "1", "-c", "A"},
		{"row", "extra", "-f", path, "-s", "Data", "-r", "1"},
		{"fill", "extra", "-f", path, "--updates", updates},
		{"read", "cell", "extra", "-f", path, "-s", "Data", "-r", "1", "-c", "A"},
		{"write", "batch", "extra", "-f", path, "--updates", updates},
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
	return executeCommandWithInput("", args...)
}

func executeCommandWithInput(input string, args ...string) (string, string, error) {
	var stdout bytes.Buffer
	var stderr bytes.Buffer
	cmd := NewRootCommand(&stdout, &stderr)
	cmd.SetIn(strings.NewReader(input))
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

func writeUpdatesFile(t *testing.T, contents string) string {
	t.Helper()

	return writeJSONFile(t, "updates.json", contents)
}

func writeJSONFile(t *testing.T, name, contents string) string {
	t.Helper()

	path := filepath.Join(t.TempDir(), name)
	if err := os.WriteFile(path, []byte(contents), 0o644); err != nil {
		t.Fatalf("WriteFile %s: %v", name, err)
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
