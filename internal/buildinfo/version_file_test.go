package buildinfo

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestVersionFile(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("..", "..", "VERSION"))
	if err != nil {
		t.Fatalf("read VERSION: %v", err)
	}

	version := strings.TrimSpace(string(data))
	if version != "v0.1.0" {
		t.Fatalf("VERSION = %q, want v0.1.0", version)
	}
}
