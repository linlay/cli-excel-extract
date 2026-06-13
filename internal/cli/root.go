package cli

import (
	"encoding/json"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/cligrep/excelx/internal/buildinfo"
	"github.com/cligrep/excelx/internal/extract"
	"github.com/spf13/cobra"
)

func NewRootCommand(out, errOut io.Writer) *cobra.Command {
	root := &cobra.Command{
		Use:     "excelx",
		Short:   "Read and write .xlsx and .xlsm workbooks",
		Version: buildinfo.String(),
		Long: `excelx reads and writes .xlsx and .xlsm workbooks.

Supported files:
  - .xlsx and .xlsm only; legacy .xls is not supported.
  - .xlsm macros are not executed. Write operations modify workbook content only.

Command model:
  - read commands never modify files.
  - write commands save in place by default; use --output to write a copy.
  - Legacy sheets, cell, row, and fill commands remain supported.

Coordinates:
  - Rows use Excel-style 1-based numbers: 1, 2, 3.
  - Columns accept either Excel letters (A, B, AA) or 1-based numbers (1, 2, 27).

Values and output:
  - Formula cells return calculated display values when Excelize can evaluate the formula.
  - Empty cells are returned as empty strings.
  - Batch read and write commands accept JSON files, or - for stdin.
  - Text output is suitable for terminal use; row text output is tab-separated.
  - Add --json for machine-readable output.
  - Errors are printed to stderr and return a non-zero exit code.`,
		Example: `  excelx read sheets -f report.xlsm
  excelx read cell -f report.xlsm -s SR1 -r 6 -c C
  excelx read range -f report.xlsm -s SR1 --range A1:H20 --json
  excelx write cell -f report.xlsx -s SR1 -r 6 -c C --type text --value done
  excelx write batch -f report.xlsx --updates updates.json --output filled.xlsx
  excelx --version`,
		Args:          cobra.NoArgs,
		SilenceUsage:  true,
		SilenceErrors: true,
		CompletionOptions: cobra.CompletionOptions{
			DisableDefaultCmd: true,
		},
	}
	root.SetOut(out)
	root.SetErr(errOut)
	root.SetVersionTemplate("{{.Name}} {{.Version}}\n")

	root.AddCommand(newSheetsCommand(out))
	root.AddCommand(newCellCommand(out))
	root.AddCommand(newRowCommand(out))
	root.AddCommand(newFillCommand(out))
	root.AddCommand(newReadCommand(out))
	root.AddCommand(newWriteCommand(out))

	return root
}

func newSheetsCommand(out io.Writer) *cobra.Command {
	var file string
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "sheets -f <file>",
		Short: "List workbook sheets",
		Long: `List all sheet names in a .xlsx or .xlsm workbook.

Required:
  -f, --file    Workbook path.

Output:
  - Text mode prints one sheet name per line.
  - JSON mode prints: {"file":"report.xlsm","sheets":["封面","SR1"]}`,
		Example: `  excelx sheets -f report.xlsm
  excelx sheets -f report.xlsm --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ListSheets(file)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			for _, sheet := range result.Sheets {
				fmt.Fprintln(out, sheet)
			}
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of text")
	_ = cmd.MarkFlagRequired("file")
	return cmd
}

func newCellCommand(out io.Writer) *cobra.Command {
	var file, sheet, col string
	var row int
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "cell -f <file> -s <sheet> -r <row> -c <col>",
		Short: "Extract one cell",
		Long: `Extract one cell from a .xlsx or .xlsm workbook.

Required:
  -f, --file     Workbook path.
  -s, --sheet    Sheet name, exactly as it appears in the workbook.
  -r, --row      1-based row number.
  -c, --col      Column letter or 1-based column number, e.g. B or 2.

Values:
  - Formula cells return calculated display values when supported.
  - Empty cells return an empty string.

Output:
  - Text mode prints only the cell value.
  - JSON mode prints: {"file":"report.xlsm","sheet":"SR1","row":6,"col":"C","value":"43,077,363.00"}`,
		Example: `  excelx cell -f report.xlsm -s SR1 -r 6 -c C
  excelx cell -f report.xlsm -s SR1 -r 6 -c 3
  excelx cell -f report.xlsm -s SR1 -r 6 -c C --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ExtractCell(file, sheet, row, col)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			fmt.Fprintln(out, result.Value)
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().IntVarP(&row, "row", "r", 0, "1-based row number")
	cmd.Flags().StringVarP(&col, "col", "c", "", "column letter or 1-based number")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("row")
	_ = cmd.MarkFlagRequired("col")
	return cmd
}

func newRowCommand(out io.Writer) *cobra.Command {
	var file, sheet, fromCol, toCol string
	var row int
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "row -f <file> -s <sheet> -r <row>",
		Short: "Extract one row",
		Long: `Extract one row from a .xlsx or .xlsm workbook.

Required:
  -f, --file     Workbook path.
  -s, --sheet    Sheet name, exactly as it appears in the workbook.
  -r, --row      1-based row number.

Column range:
  - Without --from-col and --to-col, output the workbook's used row width.
  - Use --from-col and --to-col to limit the inclusive column range.
  - Columns accept letters (A, H, AA) or 1-based numbers (1, 8, 27).

Values and output:
  - Formula cells return calculated display values when supported.
  - Empty cells are preserved as empty strings.
  - Text mode joins values with tab characters.
  - JSON mode prints: {"file":"report.xlsm","sheet":"SR1","row":6,"cells":[{"col":"C","value":"43,077,363.00"}]}`,
		Example: `  excelx row -f report.xlsm -s SR1 -r 6
  excelx row -f report.xlsm -s SR1 -r 6 --from-col A --to-col H
  excelx row -f report.xlsm -s SR1 -r 6 --from-col 1 --to-col 8 --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ExtractRow(file, sheet, row, fromCol, toCol)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			values := make([]string, 0, len(result.Cells))
			for _, cell := range result.Cells {
				values = append(values, cell.Value)
			}
			fmt.Fprintln(out, strings.Join(values, "\t"))
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().IntVarP(&row, "row", "r", 0, "1-based row number")
	cmd.Flags().StringVar(&fromCol, "from-col", "", "first column letter or 1-based number")
	cmd.Flags().StringVar(&toCol, "to-col", "", "last column letter or 1-based number")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of tab-separated text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("row")
	return cmd
}

func newFillCommand(out io.Writer) *cobra.Command {
	var file, updatesPath, output string
	var overwrite, jsonOutput bool

	cmd := &cobra.Command{
		Use:   "fill -f <file> --updates <updates.json|->",
		Short: "Fill workbook cells from JSON updates",
		Long: `Fill one or more cells in a .xlsx or .xlsm workbook from JSON updates.

Required:
  -f, --file       Workbook path.
      --updates    JSON update file path, or - to read JSON from stdin.

Save behavior:
  - By default, fill saves changes in place to --file.
  - Use --output to save a modified copy and leave --file unchanged.
  - Existing --output files are rejected unless --overwrite is set.
  - All updates are validated before any cell is changed or saved.

Update JSON:
  {"updates":[
    {"sheet":"SR1","row":4,"col":"B","type":"text","value":"已确认"},
    {"sheet":"SR1","row":4,"col":"C","type":"number","value":"123.45"},
    {"sheet":"SR1","row":4,"col":"D","type":"bool","value":"true"},
    {"sheet":"SR1","row":4,"col":"E","type":"formula","value":"=SUM(A4:C4)"},
    {"sheet":"SR1","row":4,"col":"F","type":"blank"}
  ]}

Types:
  - text writes value as text.
  - number writes a finite decimal number.
  - bool writes true or false.
  - formula writes a formula and requires value to start with =.
  - blank clears cell value and formula while preserving style.`,
		Example: `  excelx fill -f report.xlsx --updates updates.json
  excelx fill -f report.xlsx --updates updates.json --output filled.xlsx
  cat updates.json | excelx fill -f report.xlsx --updates - --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			request, err := readFillRequest(updatesPath, cmd.InOrStdin())
			if err != nil {
				return err
			}
			result, err := extract.FillCells(file, request, output, overwrite)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			fmt.Fprintf(out, "updated %s: %d cells\n", result.Output, result.Updated)
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVar(&updatesPath, "updates", "", "JSON update file path, or - for stdin")
	cmd.Flags().StringVar(&output, "output", "", "write modified workbook to this path instead of saving in place")
	cmd.Flags().BoolVar(&overwrite, "overwrite", false, "allow --output to replace an existing file")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("updates")
	return cmd
}

func newReadCommand(out io.Writer) *cobra.Command {
	cmd := &cobra.Command{
		Use:   "read",
		Short: "Read workbook content without modifying files",
		Long: `Read .xlsx or .xlsm workbook content without modifying source files.

Use read subcommands for sheets, workbook info, one cell, one row, one column,
one rectangular range, or a JSON batch of arbitrary cell queries.`,
		Example: `  excelx read sheets -f report.xlsx
  excelx read info -f report.xlsx --json
  excelx read cell -f report.xlsx -s SR1 -r 4 -c B
  excelx read row -f report.xlsx -s SR1 -r 4 --from-col A --to-col H
  excelx read col -f report.xlsx -s SR1 -c B --from-row 1 --to-row 20
  excelx read range -f report.xlsx -s SR1 --range A1:H20
  excelx read batch -f report.xlsx --queries queries.json --json`,
		Args: cobra.NoArgs,
	}
	cmd.AddCommand(newReadSheetsCommand(out))
	cmd.AddCommand(newReadInfoCommand(out))
	cmd.AddCommand(newReadCellCommand(out))
	cmd.AddCommand(newReadRowCommand(out))
	cmd.AddCommand(newReadColCommand(out))
	cmd.AddCommand(newReadRangeCommand(out))
	cmd.AddCommand(newReadBatchCommand(out))
	return cmd
}

func newReadSheetsCommand(out io.Writer) *cobra.Command {
	var file string
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "sheets -f <file>",
		Short: "List workbook sheets",
		Long: `List all sheet names in a .xlsx or .xlsm workbook.

This read command never modifies the source file.`,
		Example: `  excelx read sheets -f report.xlsx
  excelx read sheets -f report.xlsx --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ListSheets(file)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			for _, sheet := range result.Sheets {
				fmt.Fprintln(out, sheet)
			}
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of text")
	_ = cmd.MarkFlagRequired("file")
	return cmd
}

func newReadInfoCommand(out io.Writer) *cobra.Command {
	var file string
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "info -f <file>",
		Short: "Show workbook structure",
		Long: `Show workbook structure for a .xlsx or .xlsm workbook.

Text output is tab-separated: sheet name, used range, max row, max column.`,
		Example: `  excelx read info -f report.xlsx
  excelx read info -f report.xlsx --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.InspectWorkbook(file)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			for _, sheet := range result.Sheets {
				rangeRef := sheet.Range
				if rangeRef == "" {
					rangeRef = "-"
				}
				fmt.Fprintf(out, "%s\t%s\t%d\t%d\n", sheet.Name, rangeRef, sheet.MaxRow, sheet.MaxCol)
			}
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of text")
	_ = cmd.MarkFlagRequired("file")
	return cmd
}

func newReadCellCommand(out io.Writer) *cobra.Command {
	var file, sheet, col string
	var row int
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "cell -f <file> -s <sheet> -r <row> -c <col>",
		Short: "Read one cell",
		Long: `Read one cell from a .xlsx or .xlsm workbook.

Text output prints only the cell value. JSON output reuses the stable cell fields:
{"file":"report.xlsx","sheet":"SR1","row":4,"col":"B","value":"text"}.`,
		Example: `  excelx read cell -f report.xlsx -s SR1 -r 4 -c B
  excelx read cell -f report.xlsx -s SR1 -r 4 -c 2 --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ExtractCell(file, sheet, row, col)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			fmt.Fprintln(out, result.Value)
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().IntVarP(&row, "row", "r", 0, "1-based row number")
	cmd.Flags().StringVarP(&col, "col", "c", "", "column letter or 1-based number")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("row")
	_ = cmd.MarkFlagRequired("col")
	return cmd
}

func newReadRowCommand(out io.Writer) *cobra.Command {
	var file, sheet, fromCol, toCol string
	var row int
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "row -f <file> -s <sheet> -r <row>",
		Short: "Read one row",
		Long: `Read one row from a .xlsx or .xlsm workbook.

Use --from-col and --to-col to limit the inclusive column range. Text output is tab-separated.`,
		Example: `  excelx read row -f report.xlsx -s SR1 -r 4
  excelx read row -f report.xlsx -s SR1 -r 4 --from-col A --to-col H --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ExtractRow(file, sheet, row, fromCol, toCol)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			values := make([]string, 0, len(result.Cells))
			for _, cell := range result.Cells {
				values = append(values, cell.Value)
			}
			fmt.Fprintln(out, strings.Join(values, "\t"))
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().IntVarP(&row, "row", "r", 0, "1-based row number")
	cmd.Flags().StringVar(&fromCol, "from-col", "", "first column letter or 1-based number")
	cmd.Flags().StringVar(&toCol, "to-col", "", "last column letter or 1-based number")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of tab-separated text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("row")
	return cmd
}

func newReadColCommand(out io.Writer) *cobra.Command {
	var file, sheet, col string
	var fromRow, toRow int
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "col -f <file> -s <sheet> -c <col>",
		Short: "Read one column",
		Long: `Read one column from a .xlsx or .xlsm workbook.

Use --from-row and --to-row to limit the inclusive row range. If --to-row is omitted,
the command reads through the worksheet used range.`,
		Example: `  excelx read col -f report.xlsx -s SR1 -c B
  excelx read col -f report.xlsx -s SR1 -c B --from-row 1 --to-row 20 --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ExtractCol(file, sheet, col, fromRow, toRow)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			for _, cell := range result.Cells {
				fmt.Fprintln(out, cell.Value)
			}
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().StringVarP(&col, "col", "c", "", "column letter or 1-based number")
	cmd.Flags().IntVar(&fromRow, "from-row", 1, "first 1-based row number")
	cmd.Flags().IntVar(&toRow, "to-row", 0, "last 1-based row number; default is worksheet used range")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("col")
	return cmd
}

func newReadRangeCommand(out io.Writer) *cobra.Command {
	var file, sheet, rangeRef string
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "range -f <file> -s <sheet> --range <A1:B2>",
		Short: "Read a rectangular range",
		Long: `Read a rectangular A1 range from a .xlsx or .xlsm workbook.

The range must not include a sheet name. Text output is a TSV matrix.`,
		Example: `  excelx read range -f report.xlsx -s SR1 --range A1:H20
  excelx read range -f report.xlsx -s SR1 --range B4 --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ExtractRange(file, sheet, rangeRef)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			writeRangeText(out, result)
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().StringVar(&rangeRef, "range", "", "A1 cell or rectangular range, e.g. A1:H20")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of TSV text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("range")
	return cmd
}

func newReadBatchCommand(out io.Writer) *cobra.Command {
	var file, queriesPath string
	var jsonOutput bool

	cmd := &cobra.Command{
		Use:   "batch -f <file> --queries <queries.json|->",
		Short: "Read arbitrary cells from JSON queries",
		Long: `Read arbitrary cells from a .xlsx or .xlsm workbook using JSON queries.

The query JSON shape is: {"queries":[{"sheet":"SR1","row":4,"col":"B"}]}.`,
		Example: `  excelx read batch -f report.xlsx --queries queries.json --json
  cat queries.json | excelx read batch -f report.xlsx --queries -`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			request, err := readBatchRequest(queriesPath, cmd.InOrStdin())
			if err != nil {
				return err
			}
			result, err := extract.ReadCells(file, request)
			if err != nil {
				return err
			}
			if jsonOutput {
				return writeJSON(out, result)
			}
			for _, cell := range result.Cells {
				fmt.Fprintf(out, "%s\t%d\t%s\t%s\n", cell.Sheet, cell.Row, cell.Col, cell.Value)
			}
			return nil
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVar(&queriesPath, "queries", "", "JSON query file path, or - for stdin")
	cmd.Flags().BoolVar(&jsonOutput, "json", false, "write JSON output instead of tab-separated text")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("queries")
	return cmd
}

func newWriteCommand(out io.Writer) *cobra.Command {
	cmd := &cobra.Command{
		Use:   "write",
		Short: "Write workbook content",
		Long: `Write .xlsx or .xlsm workbook content.

Write commands save changes in place by default. Use --output to write a modified copy
and leave --file unchanged. Existing --output files are rejected unless --overwrite is set.`,
		Example: `  excelx write cell -f report.xlsx -s SR1 -r 4 -c B --type text --value done
  excelx write row -f report.xlsx -s SR1 -r 4 --from-col B --values values.json
  excelx write range -f report.xlsx -s SR1 --range B4:D6 --values range-values.json
  excelx write clear -f report.xlsx -s SR1 --range B4:D6
  excelx write batch -f report.xlsx --updates updates.json --output filled.xlsx`,
		Args: cobra.NoArgs,
	}
	cmd.AddCommand(newWriteCellCommand(out))
	cmd.AddCommand(newWriteRowCommand(out))
	cmd.AddCommand(newWriteColCommand(out))
	cmd.AddCommand(newWriteRangeCommand(out))
	cmd.AddCommand(newWriteClearCommand(out))
	cmd.AddCommand(newWriteBatchCommand(out))
	return cmd
}

type saveFlags struct {
	output    string
	overwrite bool
	json      bool
}

func newWriteCellCommand(out io.Writer) *cobra.Command {
	var file, sheet, col, typ, value string
	var row int
	var save saveFlags

	cmd := &cobra.Command{
		Use:   "cell -f <file> -s <sheet> -r <row> -c <col> --type <type>",
		Short: "Write one cell",
		Long: `Write one cell in a .xlsx or .xlsm workbook.

Supported types are text, number, bool, formula, and blank. Non-blank types require --value.`,
		Example: `  excelx write cell -f report.xlsx -s SR1 -r 4 -c B --type text --value "已确认"
  excelx write cell -f report.xlsx -s SR1 -r 4 -c C --type number --value 123.45
  excelx write cell -f report.xlsx -s SR1 -r 4 -c D --type blank --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.WriteCell(file, sheet, row, col, fillValueFromFlags(cmd, typ, value), save.output, save.overwrite)
			if err != nil {
				return err
			}
			return writeFillResult(out, result, save.json)
		},
	}
	bindWriteCellFlags(cmd, &file, &sheet, &row, &col, &typ, &value)
	bindSaveFlags(cmd, &save)
	return cmd
}

func newWriteRowCommand(out io.Writer) *cobra.Command {
	var file, sheet, fromCol, valuesPath string
	var row int
	var save saveFlags

	cmd := &cobra.Command{
		Use:   "row -f <file> -s <sheet> -r <row> --from-col <col> --values <values.json|->",
		Short: "Write consecutive cells in one row",
		Long: `Write consecutive cells in one row using typed JSON values.

The values JSON shape is: {"values":[{"type":"text","value":"A"},{"type":"number","value":"123"}]}.`,
		Example: `  excelx write row -f report.xlsx -s SR1 -r 4 --from-col B --values values.json
  cat values.json | excelx write row -f report.xlsx -s SR1 -r 4 --from-col B --values - --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			values, err := readValuesRequest(valuesPath, cmd.InOrStdin())
			if err != nil {
				return err
			}
			result, err := extract.WriteRow(file, sheet, row, fromCol, values, save.output, save.overwrite)
			if err != nil {
				return err
			}
			return writeFillResult(out, result, save.json)
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().IntVarP(&row, "row", "r", 0, "1-based row number")
	cmd.Flags().StringVar(&fromCol, "from-col", "", "first column letter or 1-based number")
	cmd.Flags().StringVar(&valuesPath, "values", "", "JSON values file path, or - for stdin")
	bindSaveFlags(cmd, &save)
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("row")
	_ = cmd.MarkFlagRequired("from-col")
	_ = cmd.MarkFlagRequired("values")
	return cmd
}

func newWriteColCommand(out io.Writer) *cobra.Command {
	var file, sheet, col, valuesPath string
	var fromRow int
	var save saveFlags

	cmd := &cobra.Command{
		Use:   "col -f <file> -s <sheet> -c <col> --from-row <row> --values <values.json|->",
		Short: "Write consecutive cells in one column",
		Long: `Write consecutive cells in one column using typed JSON values.

The values JSON shape is: {"values":[{"type":"text","value":"A"},{"type":"blank"}]}.`,
		Example: `  excelx write col -f report.xlsx -s SR1 -c B --from-row 4 --values values.json`,
		Args:    cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			values, err := readValuesRequest(valuesPath, cmd.InOrStdin())
			if err != nil {
				return err
			}
			result, err := extract.WriteCol(file, sheet, col, fromRow, values, save.output, save.overwrite)
			if err != nil {
				return err
			}
			return writeFillResult(out, result, save.json)
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().StringVarP(&col, "col", "c", "", "column letter or 1-based number")
	cmd.Flags().IntVar(&fromRow, "from-row", 0, "first 1-based row number")
	cmd.Flags().StringVar(&valuesPath, "values", "", "JSON values file path, or - for stdin")
	bindSaveFlags(cmd, &save)
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("col")
	_ = cmd.MarkFlagRequired("from-row")
	_ = cmd.MarkFlagRequired("values")
	return cmd
}

func newWriteRangeCommand(out io.Writer) *cobra.Command {
	var file, sheet, rangeRef, valuesPath string
	var save saveFlags

	cmd := &cobra.Command{
		Use:   "range -f <file> -s <sheet> --range <A1:B2> --values <values.json|->",
		Short: "Write a rectangular range",
		Long: `Write a rectangular A1 range using a typed JSON value matrix.

The values JSON shape is: {"rows":[[{"type":"text","value":"A"}],[{"type":"bool","value":"true"}]]}.`,
		Example: `  excelx write range -f report.xlsx -s SR1 --range B4:D6 --values range-values.json`,
		Args:    cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			values, err := readRangeValuesRequest(valuesPath, cmd.InOrStdin())
			if err != nil {
				return err
			}
			result, err := extract.WriteRange(file, sheet, rangeRef, values, save.output, save.overwrite)
			if err != nil {
				return err
			}
			return writeFillResult(out, result, save.json)
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().StringVar(&rangeRef, "range", "", "A1 cell or rectangular range, e.g. B4:D6")
	cmd.Flags().StringVar(&valuesPath, "values", "", "JSON range values file path, or - for stdin")
	bindSaveFlags(cmd, &save)
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("range")
	_ = cmd.MarkFlagRequired("values")
	return cmd
}

func newWriteClearCommand(out io.Writer) *cobra.Command {
	var file, sheet, rangeRef string
	var save saveFlags

	cmd := &cobra.Command{
		Use:   "clear -f <file> -s <sheet> --range <A1:B2>",
		Short: "Clear a cell or rectangular range",
		Long:  `Clear cell values and formulas in an A1 cell or rectangular range while preserving styles.`,
		Example: `  excelx write clear -f report.xlsx -s SR1 --range B4:D6
  excelx write clear -f report.xlsx -s SR1 --range B4 --output cleared.xlsx`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			result, err := extract.ClearRange(file, sheet, rangeRef, save.output, save.overwrite)
			if err != nil {
				return err
			}
			return writeFillResult(out, result, save.json)
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(&sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().StringVar(&rangeRef, "range", "", "A1 cell or rectangular range, e.g. B4:D6")
	bindSaveFlags(cmd, &save)
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("range")
	return cmd
}

func newWriteBatchCommand(out io.Writer) *cobra.Command {
	var file, updatesPath string
	var save saveFlags

	cmd := &cobra.Command{
		Use:   "batch -f <file> --updates <updates.json|->",
		Short: "Write arbitrary cells from JSON updates",
		Long: `Write arbitrary cells in a .xlsx or .xlsm workbook using JSON updates.

This command uses the same JSON shape as the legacy fill command:
{"updates":[{"sheet":"SR1","row":4,"col":"B","type":"text","value":"done"}]}.`,
		Example: `  excelx write batch -f report.xlsx --updates updates.json
  excelx write batch -f report.xlsx --updates updates.json --output filled.xlsx --json`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			request, err := readFillRequest(updatesPath, cmd.InOrStdin())
			if err != nil {
				return err
			}
			result, err := extract.WriteBatch(file, request, save.output, save.overwrite)
			if err != nil {
				return err
			}
			return writeFillResult(out, result, save.json)
		},
	}
	cmd.Flags().StringVarP(&file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVar(&updatesPath, "updates", "", "JSON update file path, or - for stdin")
	bindSaveFlags(cmd, &save)
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("updates")
	return cmd
}

func writeJSON(out io.Writer, value any) error {
	encoder := json.NewEncoder(out)
	encoder.SetEscapeHTML(false)
	return encoder.Encode(value)
}

func bindSaveFlags(cmd *cobra.Command, save *saveFlags) {
	cmd.Flags().StringVar(&save.output, "output", "", "write modified workbook to this path instead of saving in place")
	cmd.Flags().BoolVar(&save.overwrite, "overwrite", false, "allow --output to replace an existing file")
	cmd.Flags().BoolVar(&save.json, "json", false, "write JSON output instead of text")
}

func bindWriteCellFlags(cmd *cobra.Command, file, sheet *string, row *int, col, typ, value *string) {
	cmd.Flags().StringVarP(file, "file", "f", "", "workbook path (.xlsx or .xlsm)")
	cmd.Flags().StringVarP(sheet, "sheet", "s", "", "sheet name")
	cmd.Flags().IntVarP(row, "row", "r", 0, "1-based row number")
	cmd.Flags().StringVarP(col, "col", "c", "", "column letter or 1-based number")
	cmd.Flags().StringVar(typ, "type", "", "value type: text, number, bool, formula, or blank")
	cmd.Flags().StringVar(value, "value", "", "value to write; omit for --type blank")
	_ = cmd.MarkFlagRequired("file")
	_ = cmd.MarkFlagRequired("sheet")
	_ = cmd.MarkFlagRequired("row")
	_ = cmd.MarkFlagRequired("col")
	_ = cmd.MarkFlagRequired("type")
}

func fillValueFromFlags(cmd *cobra.Command, typ, value string) extract.FillValue {
	var valuePtr *string
	if cmd.Flags().Changed("value") {
		valuePtr = &value
	}
	return extract.FillValue{Type: typ, Value: valuePtr}
}

func writeFillResult(out io.Writer, result extract.FillResult, jsonOutput bool) error {
	if jsonOutput {
		return writeJSON(out, result)
	}
	fmt.Fprintf(out, "updated %s: %d cells\n", result.Output, result.Updated)
	return nil
}

func writeRangeText(out io.Writer, result extract.RangeResult) {
	currentRow := 0
	values := make([]string, 0)
	flush := func() {
		if currentRow != 0 {
			fmt.Fprintln(out, strings.Join(values, "\t"))
		}
	}
	for _, cell := range result.Cells {
		if currentRow == 0 {
			currentRow = cell.Row
		}
		if cell.Row != currentRow {
			flush()
			currentRow = cell.Row
			values = values[:0]
		}
		values = append(values, cell.Value)
	}
	flush()
}

func readFillRequest(path string, stdin io.Reader) (extract.FillRequest, error) {
	var request extract.FillRequest
	if err := decodeJSONInput(path, stdin, "updates", &request); err != nil {
		return extract.FillRequest{}, err
	}
	return request, nil
}

func readBatchRequest(path string, stdin io.Reader) (extract.ReadBatchRequest, error) {
	var request extract.ReadBatchRequest
	if err := decodeJSONInput(path, stdin, "queries", &request); err != nil {
		return extract.ReadBatchRequest{}, err
	}
	return request, nil
}

func readValuesRequest(path string, stdin io.Reader) (extract.ValuesRequest, error) {
	var request extract.ValuesRequest
	if err := decodeJSONInput(path, stdin, "values", &request); err != nil {
		return extract.ValuesRequest{}, err
	}
	return request, nil
}

func readRangeValuesRequest(path string, stdin io.Reader) (extract.RangeValuesRequest, error) {
	var request extract.RangeValuesRequest
	if err := decodeJSONInput(path, stdin, "range values", &request); err != nil {
		return extract.RangeValuesRequest{}, err
	}
	return request, nil
}

func decodeJSONInput(path string, stdin io.Reader, name string, value any) error {
	var input io.Reader
	if strings.TrimSpace(path) == "-" {
		input = stdin
	} else {
		file, err := os.Open(path)
		if err != nil {
			return fmt.Errorf("open %s: %w", name, err)
		}
		defer file.Close()
		input = file
	}

	decoder := json.NewDecoder(input)
	decoder.DisallowUnknownFields()

	if err := decoder.Decode(value); err != nil {
		return fmt.Errorf("decode %s JSON: %w", name, err)
	}
	var extra any
	if err := decoder.Decode(&extra); err != io.EOF {
		if err == nil {
			return fmt.Errorf("decode %s JSON: multiple JSON values are not supported", name)
		}
		return fmt.Errorf("decode %s JSON: %w", name, err)
	}
	return nil
}
