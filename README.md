# excel-extract 操作手册

`excel-extract` 是一个使用 Go 1.26、Cobra 和 Excelize 构建的命令行工具，用于从 `.xlsx` 和 `.xlsm` 文件中读取和写入内容。新命令模型使用 `read` 表示不修改文件，使用 `write` 表示会修改原文件或生成新文件；旧版 `sheets`、`cell`、`row`、`fill` 命令继续兼容。

## 支持范围

- 支持文件：`.xlsx`、`.xlsm`
- 不支持文件：旧版二进制 `.xls`
- `.xlsm` 不执行宏；`read` 命令不修改原文件，`write`/`fill` 命令只修改 workbook 内容
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
./dist/excel-extract read --help
./dist/excel-extract read range --help
./dist/excel-extract write --help
./dist/excel-extract write batch --help
```

帮助内容包含支持文件类型、行列规则、公式计算行为、文本/JSON 输出格式和可复制示例。通常只看 `--help` 就能完成操作。

## 读取内容

`read` 命令只读取，不修改源文件。

列出工作表：

```bash
./dist/excel-extract read sheets -f "/path/to/report.xlsm"
```

查看 workbook 结构。文本输出依次是 sheet 名、已用范围、最大行、最大列：

```bash
./dist/excel-extract read info -f "/path/to/report.xlsm"
```

读取单元格：

```bash
./dist/excel-extract read cell -f "/path/to/report.xlsm" -s SR1 -r 4 -c B
./dist/excel-extract read cell -f "/path/to/report.xlsm" -s SR1 -r 4 -c 2 --json
```

读取一行或指定列范围：

```bash
./dist/excel-extract read row -f "/path/to/report.xlsm" -s SR1 -r 4
./dist/excel-extract read row -f "/path/to/report.xlsm" -s SR1 -r 4 --from-col A --to-col H --json
```

读取一列或指定行范围：

```bash
./dist/excel-extract read col -f "/path/to/report.xlsm" -s SR1 -c B
./dist/excel-extract read col -f "/path/to/report.xlsm" -s SR1 -c B --from-row 1 --to-row 20 --json
```

读取矩形区域。文本输出是 TSV 矩阵：

```bash
./dist/excel-extract read range -f "/path/to/report.xlsm" -s SR1 --range A1:H20
./dist/excel-extract read range -f "/path/to/report.xlsm" -s SR1 --range B4 --json
```

任意单元格批量读取：

```bash
./dist/excel-extract read batch -f "/path/to/report.xlsm" --queries queries.json --json
```

`queries.json`：

```json
{
  "queries": [
    {"sheet": "SR1", "row": 4, "col": "B"},
    {"sheet": "SR1", "row": 5, "col": "C"}
  ]
}
```

读取结果 JSON 示例：

```json
{"file":"/path/to/report.xlsm","sheet":"SR1","row":4,"col":"B","value":"指标名称"}
```

```json
{
  "file": "/path/to/report.xlsm",
  "sheet": "SR1",
  "range": "A1:B2",
  "cells": [
    {"row": 1, "col": "A", "value": "指标序号"},
    {"row": 1, "col": "B", "value": "指标名称"}
  ]
}
```

空单元格会以空字符串 `""` 返回。公式单元格输出 Excelize 能计算出的显示值。

## 写入内容

`write` 命令默认原地修改 `--file`。如果要保留源文件，使用 `--output` 写入新文件；`--output` 指向已存在文件时默认报错，确认覆盖时添加 `--overwrite`。

写一个单元格：

```bash
./dist/excel-extract write cell -f "/path/to/report.xlsx" -s SR1 -r 4 -c B --type text --value "已确认"
./dist/excel-extract write cell -f "/path/to/report.xlsx" -s SR1 -r 4 -c C --type number --value 123.45
./dist/excel-extract write cell -f "/path/to/report.xlsx" -s SR1 -r 4 -c D --type blank --json
```

写一行连续区域：

```bash
./dist/excel-extract write row -f "/path/to/report.xlsx" -s SR1 -r 4 --from-col B --values values.json
```

写一列连续区域：

```bash
./dist/excel-extract write col -f "/path/to/report.xlsx" -s SR1 -c B --from-row 4 --values values.json
```

`values.json`：

```json
{
  "values": [
    {"type": "text", "value": "A"},
    {"type": "number", "value": "123"},
    {"type": "blank"}
  ]
}
```

写矩形区域：

```bash
./dist/excel-extract write range -f "/path/to/report.xlsx" -s SR1 --range B4:D6 --values range-values.json
```

`range-values.json`：

```json
{
  "rows": [
    [
      {"type": "text", "value": "A"},
      {"type": "number", "value": "123"}
    ],
    [
      {"type": "bool", "value": "true"},
      {"type": "formula", "value": "=SUM(B4:C4)"}
    ]
  ]
}
```

清空单元格或矩形区域，保留样式：

```bash
./dist/excel-extract write clear -f "/path/to/report.xlsx" -s SR1 --range B4:D6
```

任意单元格批量写入：

```bash
./dist/excel-extract write batch -f "/path/to/report.xlsx" --updates updates.json
./dist/excel-extract write batch -f "/path/to/report.xlsx" --updates updates.json --output "/path/to/filled.xlsx"
```

`updates.json`：

```json
{
  "updates": [
    {"sheet": "SR1", "row": 4, "col": "B", "type": "text", "value": "已确认"},
    {"sheet": "SR1", "row": 4, "col": "C", "type": "number", "value": "123.45"},
    {"sheet": "SR1", "row": 4, "col": "D", "type": "bool", "value": "true"},
    {"sheet": "SR1", "row": 4, "col": "E", "type": "formula", "value": "=SUM(A4:C4)"},
    {"sheet": "SR1", "row": 4, "col": "F", "type": "blank"}
  ]
}
```

从标准输入读取 JSON：

```bash
cat updates.json | ./dist/excel-extract write batch -f "/path/to/report.xlsx" --updates - --json
```

支持的写入类型：

- `text`：按文本写入，公式样式文本也不会执行。
- `number`：写入有限十进制数字。
- `bool`：写入 `true` 或 `false`。
- `formula`：写入公式，`value` 必须以 `=` 开头。
- `blank`：清空单元格值和公式，保留样式；不能带 `value`。

写入文本输出示例：

```text
updated /path/to/report.xlsx: 5 cells
```

写入 JSON 输出示例：

```json
{"file":"/path/to/report.xlsx","output":"/path/to/report.xlsx","inPlace":true,"updated":5}
```

## 兼容旧命令

以下旧命令继续可用，行为与新命令对应：

```bash
./dist/excel-extract sheets -f "/path/to/report.xlsm"
./dist/excel-extract cell -f "/path/to/report.xlsm" -s SR1 -r 4 -c B
./dist/excel-extract row -f "/path/to/report.xlsm" -s SR1 -r 4 --from-col A --to-col H
./dist/excel-extract fill -f "/path/to/report.xlsx" --updates updates.json
```

对应的新命令：

```bash
./dist/excel-extract read sheets -f "/path/to/report.xlsm"
./dist/excel-extract read cell -f "/path/to/report.xlsm" -s SR1 -r 4 -c B
./dist/excel-extract read row -f "/path/to/report.xlsm" -s SR1 -r 4 --from-col A --to-col H
./dist/excel-extract write batch -f "/path/to/report.xlsx" --updates updates.json
```

## 样例验证

使用本机样例文件：

```bash
SAMPLE="/Users/linlay/Downloads/期货公司监管报表（月报）1.3.2版 (2)_全表随机数据.xlsm"
```

列出工作表：

```bash
./dist/excel-extract read sheets -f "$SAMPLE" --json
```

提取 `SR1!B4`：

```bash
./dist/excel-extract read cell -f "$SAMPLE" -s SR1 -r 4 -c B
```

预期输出：

```text
指标名称
```

提取 `SR1` 第 4 行 JSON：

```bash
./dist/excel-extract read row -f "$SAMPLE" -s SR1 -r 4 --json
```

## 错误处理

以下情况会返回非零退出码，并在标准错误输出中打印原因：

- 文件路径为空或文件不存在
- 文件扩展名不是 `.xlsx` 或 `.xlsm`
- 指定的 sheet 不存在
- 行号小于 1 或超过 Excel 行数限制
- 列号为空、格式非法或超过 Excel 列数限制
- `--from-col` 大于 `--to-col`
- `read range` / `write range` 的 A1 范围非法、反向或包含 sheet 名
- JSON 为空、字段缺失、类型非法或包含未知字段
- `number`、`bool`、`formula`、`blank` 值不符合类型要求
- `write --output` 文件已存在且未指定 `--overwrite`

## 测试

运行全部测试：

```bash
go test ./...
```

测试覆盖列号解析、A1 范围解析、参数校验、JSON 输出、空单元格处理、批量读写校验和保存策略，以及临时 `.xlsx` 的新旧命令行为。
# cli-excel-extract
