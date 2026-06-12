package cli

import (
	"encoding/json"
	"fmt"
	"io"
	"strings"

	"github.com/cligrep/cli-excel-extract/internal/buildinfo"
	"github.com/cligrep/cli-excel-extract/internal/extract"
	"github.com/spf13/cobra"
)

func NewRootCommand(out, errOut io.Writer) *cobra.Command {
	root := &cobra.Command{
		Use:     "excel-extract",
		Short:   "Extract values from .xlsx and .xlsm workbooks",
		Version: buildinfo.String(),
		Long: `excel-extract reads .xlsx and .xlsm workbooks and prints selected content.

Supported files:
  - .xlsx and .xlsm only; legacy .xls is not supported.
  - .xlsm files are read-only: macros are not executed and source files are not modified.

Coordinates:
  - Rows use Excel-style 1-based numbers: 1, 2, 3.
  - Columns accept either Excel letters (A, B, AA) or 1-based numbers (1, 2, 27).

Values and output:
  - Formula cells return calculated display values when Excelize can evaluate the formula.
  - Empty cells are returned as empty strings.
  - Text output is suitable for terminal use; row text output is tab-separated.
  - Add --json for machine-readable output.
  - Errors are printed to stderr and return a non-zero exit code.`,
		Example: `  excel-extract sheets -f report.xlsm
  excel-extract sheets -f report.xlsm --json
  excel-extract cell -f report.xlsm -s SR1 -r 6 -c C
  excel-extract cell -f report.xlsm -s SR1 -r 6 -c 3 --json
  excel-extract row -f report.xlsm -s SR1 -r 6
  excel-extract row -f report.xlsm -s SR1 -r 6 --from-col A --to-col H --json
  excel-extract --version`,
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
		Example: `  excel-extract sheets -f report.xlsm
  excel-extract sheets -f report.xlsm --json`,
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
		Example: `  excel-extract cell -f report.xlsm -s SR1 -r 6 -c C
  excel-extract cell -f report.xlsm -s SR1 -r 6 -c 3
  excel-extract cell -f report.xlsm -s SR1 -r 6 -c C --json`,
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
		Example: `  excel-extract row -f report.xlsm -s SR1 -r 6
  excel-extract row -f report.xlsm -s SR1 -r 6 --from-col A --to-col H
  excel-extract row -f report.xlsm -s SR1 -r 6 --from-col 1 --to-col 8 --json`,
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

func writeJSON(out io.Writer, value any) error {
	encoder := json.NewEncoder(out)
	encoder.SetEscapeHTML(false)
	return encoder.Encode(value)
}
