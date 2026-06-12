package buildinfo

import (
	"fmt"
	"runtime"
)

var (
	Version = "dev"
	Commit  = "unknown"
	Date    = "unknown"
)

func Platform() string {
	return runtime.GOOS + "/" + runtime.GOARCH
}

func String() string {
	return fmt.Sprintf("version=%s commit=%s date=%s platform=%s", Version, Commit, Date, Platform())
}
