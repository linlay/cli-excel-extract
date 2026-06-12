# excel-extract 操作手册

`excel-extract` 是一个使用 Go 1.26、Cobra 和 Excelize 构建的命令行工具，用于从 `.xlsx` 和 `.xlsm` 文件中提取内容。它支持列出工作表、读取单个单元格、读取整行或指定列范围，并可输出纯文本或 JSON。

## 支持范围

- 支持文件：`.xlsx`、`.xlsm`
- 不支持文件：旧版二进制 `.xls`
- `.xlsm` 只读取已有单元格内容，不执行宏，不修改原文件
- 公式单元格默认返回计算后的显示值；计算由内置 Excelize 完成，不依赖 Excel 或 LibreOffice
- 行号使用 Excel 习惯的 1-based 编号
- 列号支持两种写法：字母列名如 `A`、`BC`，或数字列号如 `1`、`55`

## 构建

构建当前平台可执行程序：

```bash
CGO_ENABLED=0 go build -trimpath -ldflags="-s -w" -o dist/excel-extract ./cmd/excel-extract
```

构建完成后，可执行程序路径为：

```bash
dist/excel-extract
```

最终运行只需要这个可执行程序，不需要额外脚本、配置文件或样例 Excel 文件。

查看版本：

```bash
./dist/excel-extract --version
```

输出示例：

```text
excel-extract version=v0.1.0 commit=unknown date=unknown platform=darwin/arm64
```

## 跨平台打包

项目内置打包脚本会分别构建 macOS Darwin arm64 和 Windows amd64：

```bash
./scripts/package.sh
```

输出文件：

```text
dist/darwin-arm64/excel-extract
dist/windows-amd64/excel-extract.exe
```

每个目标目录里都是单一 CLI 可执行程序，不需要额外运行时文件。

打包脚本会把版本信息写进可执行程序：

- `version`：读取根目录 `VERSION` 文件，当前为 `v0.1.0`
- `commit`：使用短 commit hash，没有 git 信息时为 `unknown`
- `date`：UTC 构建时间
- `platform`：运行时平台，如 `darwin/arm64` 或 `windows/amd64`

发布新版本时，修改 `VERSION` 文件后重新执行：

```bash
./scripts/package.sh
```

## 查看帮助

```bash
./dist/excel-extract --help
./dist/excel-extract --version
./dist/excel-extract sheets --help
./dist/excel-extract cell --help
./dist/excel-extract row --help
```

帮助内容包含支持文件类型、行列规则、公式计算行为、文本/JSON 输出格式和可复制示例。通常只看 `--help` 就能完成操作。

## 列出工作表

纯文本输出：

```bash
./dist/excel-extract sheets -f "/path/to/report.xlsm"
```

JSON 输出：

```bash
./dist/excel-extract sheets -f "/path/to/report.xlsm" --json
```

JSON 格式：

```json
{"file":"/path/to/report.xlsm","sheets":["封面","校验表","SR1"]}
```

## 提取单个单元格

按字母列号读取：

```bash
./dist/excel-extract cell -f "/path/to/report.xlsm" -s SR1 -r 4 -c B
```

按数字列号读取：

```bash
./dist/excel-extract cell -f "/path/to/report.xlsm" -s SR1 -r 4 -c 2
```

JSON 输出：

```bash
./dist/excel-extract cell -f "/path/to/report.xlsm" -s SR1 -r 4 -c B --json
```

JSON 格式：

```json
{"file":"/path/to/report.xlsm","sheet":"SR1","row":4,"col":"B","value":"指标名称"}
```

## 提取一行

读取整行：

```bash
./dist/excel-extract row -f "/path/to/report.xlsm" -s SR1 -r 4
```

纯文本输出使用制表符分隔，便于复制到表格或管道处理。

读取指定列范围：

```bash
./dist/excel-extract row -f "/path/to/report.xlsm" -s SR1 -r 4 --from-col A --to-col H
```

列范围也可以使用数字：

```bash
./dist/excel-extract row -f "/path/to/report.xlsm" -s SR1 -r 4 --from-col 1 --to-col 8
```

JSON 输出：

```bash
./dist/excel-extract row -f "/path/to/report.xlsm" -s SR1 -r 4 --json
```

JSON 格式：

```json
{
  "file": "/path/to/report.xlsm",
  "sheet": "SR1",
  "row": 4,
  "cells": [
    {"col": "A", "value": "指标序号"},
    {"col": "B", "value": "指标名称"}
  ]
}
```

空单元格会以空字符串 `""` 返回。默认读取整行时，工具会尽量保留工作表已用范围内的空列；指定 `--from-col` 和 `--to-col` 时，以用户指定范围为准。

如果行内包含公式单元格，输出会使用公式计算后的显示值。例如单元格公式引用其他 sheet 时，工具会沿公式链计算并返回最终结果。

## 样例验证

使用本机样例文件：

```bash
SAMPLE="/Users/linlay/Downloads/期货公司监管报表（月报）1.3.2版 (2)_全表随机数据.xlsm"
```

列出工作表：

```bash
./dist/excel-extract sheets -f "$SAMPLE" --json
```

提取 `SR1!B4`：

```bash
./dist/excel-extract cell -f "$SAMPLE" -s SR1 -r 4 -c B
```

预期输出：

```text
指标名称
```

提取 `SR1` 第 4 行 JSON：

```bash
./dist/excel-extract row -f "$SAMPLE" -s SR1 -r 4 --json
```

## 错误处理

以下情况会返回非零退出码，并在标准错误输出中打印原因：

- 文件路径为空或文件不存在
- 文件扩展名不是 `.xlsx` 或 `.xlsm`
- 指定的 sheet 不存在
- 行号小于 1 或超过 Excel 行数限制
- 列号为空、格式非法或超过 Excel 列数限制
- `--from-col` 大于 `--to-col`

## 测试

运行全部测试：

```bash
go test ./...
```

测试覆盖列号解析、参数校验、JSON 输出、空单元格处理，以及临时 `.xlsx` 的 `sheets`、`cell`、`row` 命令行为。
# cli-excel-extract
