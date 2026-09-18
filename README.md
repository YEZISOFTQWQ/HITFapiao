# HITFapiao

Windows 发票批量处理程序：读取 PDF 发票、确认材料名称、填写材料验收单，并生成 A4 双联发票打印汇总。

## 功能

- 解析电子发票中的发票号码、日期、购买方、供应商和商品明细。
- 从文件名和发票明细提取材料名称，并通过 GUI 逐条快捷确认。
- 按含税金额填写材料名称、单位、数量、含税单价、含税金额和供应商。
- 验收日期读取生成时的系统日期。
- 非标准发票格式、购买方名称错误或纳税人识别号错误时，将验收单明细标红并在界面提示。
- 同一供应商累计金额超过 1000 元时，将供应商单元格标黄并在界面提示。
- 超过 10 条材料时，在同一个工作表内增加行；打印时每 10 条分页并重复表头。
- 经办人、验收人、负责人和凭证号保留空白，供打印后手写。
- 将原发票归档为 `序号.材料名称[金额].pdf`。
- 将归档发票按数字序号排序，生成 A4 竖版、每页上下两张的打印汇总 PDF。
- 已有归档允许缺号和同号；非标准命名 PDF 会被排除并在 GUI 中提示。

## 仓库与隐私

仓库只保存运行源码、构建脚本和说明文件，不保存私有测试、真实发票、Excel 模板、生成结果、EXE 或 ZIP。

本地运行前：

1. 准备自己的发票 PDF 文件夹。
2. 把材料验收单模板放在项目根目录并命名为 `模板.xls`，或在 GUI 中选择其他 `.xls` 模板。
3. `tests/`、`demo/`、所有 PDF/XLS/XLSX、`输出/` 和 `release/` 均已加入 `.gitignore`。

## 源码运行

电脑需安装 Microsoft Excel。首次运行：

```powershell
.\setup.ps1
.\run.bat
```

`setup.ps1` 会在项目目录创建 `.venv` 并检查依赖。GUI 默认从程序目录寻找 `模板.xls`；发票目录需要用户选择。

## 命令行

```powershell
.\.venv\Scripts\python.exe -m invoice_acceptance.cli `
  --invoice-dir D:\待处理发票 `
  --template D:\材料验收单模板.xls `
  --output D:\输出\材料验收单.xls
```

可选参数包括 `--unit-name` 和 `--expense-card`。使用 `--parse-only` 可只查看 JSON 解析结果。需要逐条人工确认材料名称时请使用 GUI。

只从已经人工整理过的发票归档重新生成打印汇总：

```powershell
.\.venv\Scripts\python.exe -m invoice_acceptance.cli `
  --combine-existing D:\输出\材料验收单_发票
```

## 填写口径

- 金额采用发票含税金额，验收单合计应与全部发票的价税合计一致。
- 含税单价按“含税金额 ÷ 数量”计算。
- 解析失败时默认停止生成，防止遗漏发票。

## 构建 Windows Release

先在项目根目录放置本地 `模板.xls`，然后运行：

```powershell
.\build_release.ps1
```

构建结果位于本地 `release/`：

- `HITFapiao.exe`：Windows 64 位单文件 GUI。
- `模板.xls`：运行所需的本地模板副本。
- `HITFapiao-Windows-x64.zip`：可上传到 GitHub Release 的压缩包。

Release ZIP 不包含任何发票或 `demo` 数据。解压后应让 `HITFapiao.exe` 和 `模板.xls` 保持在同一目录。
