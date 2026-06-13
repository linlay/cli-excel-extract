#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIST_DIR="${ROOT_DIR}/dist"
VERSION_FILE="${ROOT_DIR}/VERSION"

if [[ ! -f "${VERSION_FILE}" ]]; then
  echo "missing VERSION file: ${VERSION_FILE}" >&2
  exit 1
fi

VERSION="$(tr -d '[:space:]' < "${VERSION_FILE}")"
if [[ -z "${VERSION}" ]]; then
  echo "VERSION file is empty: ${VERSION_FILE}" >&2
  exit 1
fi

if git -C "${ROOT_DIR}" rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  COMMIT="$(git -C "${ROOT_DIR}" rev-parse --short HEAD 2>/dev/null || printf 'unknown')"
else
  COMMIT="unknown"
fi
DATE="$(date -u '+%Y-%m-%dT%H:%M:%SZ')"
LDFLAGS="-s -w -X github.com/cligrep/excelx/internal/buildinfo.Version=${VERSION} -X github.com/cligrep/excelx/internal/buildinfo.Commit=${COMMIT} -X github.com/cligrep/excelx/internal/buildinfo.Date=${DATE}"

build_target() {
  local goos="$1"
  local goarch="$2"
  local output="$3"

  echo "Building ${goos}/${goarch}: ${output}"
  (
    cd "${ROOT_DIR}"
    CGO_ENABLED=0 GOOS="${goos}" GOARCH="${goarch}" \
      go build -trimpath -ldflags="${LDFLAGS}" -o "${output}" ./cmd/excelx
  )
}

rm -rf "${DIST_DIR}"
mkdir -p "${DIST_DIR}/darwin-arm64" "${DIST_DIR}/windows-amd64"

build_target "darwin" "arm64" "${DIST_DIR}/darwin-arm64/excelx"
build_target "windows" "amd64" "${DIST_DIR}/windows-amd64/excelx.exe"

echo
echo "Package outputs:"
echo "  ${DIST_DIR}/darwin-arm64/excelx"
echo "  ${DIST_DIR}/windows-amd64/excelx.exe"
echo
echo "Version:"
echo "  version=${VERSION} commit=${COMMIT} date=${DATE}"
