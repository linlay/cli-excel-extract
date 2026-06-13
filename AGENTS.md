# AGENTS.md

本项目是一个 Go 1.26 CLI，用于从 `.xlsx` 和 `.xlsm` 文件中读取和写入单元格、行、列、区域和工作表信息。后续维护请遵守以下约定。

## 项目约束

- 最终交付形态必须是单一 CLI 可执行程序；跨平台打包时每个目标目录只放一个可执行程序。
- 运行时不得依赖脚本、外部配置文件、样例 Excel 或本地服务。
- 只支持 `.xlsx` 和 `.xlsm`；不要在未明确要求时加入旧版 `.xls` 支持。
- `.xlsm` 不执行宏；`read` 命令不修改源文件，`write`/兼容写入命令只修改 workbook 内容。
- CLI 使用 Cobra，Excel 读取使用 Excelize。
- `--help` 必须足够自解释；用户或 AI 不看 README 也应该能完成常见操作。
- `--version` 必须输出 version、commit、date 和 platform。
- 版本号唯一来源是根目录 `VERSION` 文件，格式必须带 `v` 前缀，例如 `v0.1.0`。

## 代码结构

- `cmd/excelx/main.go`：程序入口，只负责创建并执行 root command。
- `internal/cli`：Cobra 命令、参数绑定、文本和 JSON 输出。
- `internal/extract`：文件校验、列号/区域解析、sheet/cell/row/col/range 读写逻辑。

新增能力时优先把业务逻辑放在 `internal/extract`，让 `internal/cli` 保持薄封装，方便测试。

## 命令接口兼容性

保持这些命令稳定：

```bash
excelx sheets -f <file> [--json]
excelx cell -f <file> -s <sheet> -r <row> -c <col> [--json]
excelx row -f <file> -s <sheet> -r <row> [--from-col <col>] [--to-col <col>] [--json]
excelx fill -f <file> --updates <updates.json|-> [--output <file>] [--overwrite] [--json]
```

JSON 字段名属于对外接口，修改前需要同步更新 README 和测试。

新能力优先挂在 `excelx read ...` 和 `excelx write ...` 命令树下；旧顶层命令用于兼容，不要移除。

Help 文案和 `--version` 输出也属于对外接口。新增或修改命令参数时，同步更新子命令 `Long`、`Example`、测试和 README。

## 开发检查

修改 Go 代码后执行：

```bash
gofmt -w cmd/excelx/main.go internal/cli/*.go internal/extract/*.go internal/buildinfo/*.go
go test ./...
./scripts/package.sh
```

`scripts/package.sh` 需要从根目录 `VERSION` 读取 version，并通过 `-ldflags` 注入 `internal/buildinfo` 的 version、commit 和 date。非 git 目录下构建时 commit 允许使用 `unknown`，但 version 不允许退回 `dev`。

如果修改了 CLI 行为，还需要用样例 `.xlsm` 手工验收 `read sheets`、`read cell`、`read row --json`，以及 `write ... --output`。

## 文档要求

- `README.md` 是用户操作手册，应优先描述如何构建、运行、查看输出和处理错误。
- 变更命令参数、输出格式或支持文件类型时，必须更新 README。
- 不要把构建产物、临时文件、`.DS_Store` 或样例数据提交到仓库。
