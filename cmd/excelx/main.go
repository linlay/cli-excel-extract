package main

import (
	"fmt"
	"os"

	"github.com/cligrep/excelx/internal/cli"
)

func main() {
	cmd := cli.NewRootCommand(os.Stdout, os.Stderr)
	if err := cmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
