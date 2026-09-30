# Fintech CSV → Excel

## 第三步：input3 地址 ID 替换（0929）

将以下两个真实文件放入项目的 `input3/` 文件夹（文件需由你提供）：

- `Transaction Columns for Fintech RFI_0929.xlsx`
- `address_mapping_to_Alert.csv`

关闭该 Excel 后，在项目目录的 Windows PowerShell 中执行：

```powershell
py -3 -m pip install -r requirements.txt
python map_address_details.py
```

也可以双击 `run_address_mapping_windows.bat`（需要 `python` 在 PATH 中）。若 `py -3` 提示未安装 Python，但 `python --version` 正常，则安装依赖时使用 `python -m pip install -r requirements.txt`。这一步直接处理 0929 文件，无需重新运行前两步。

程序用 CSV 的 `raw_addr_id` 分别匹配 Excel 的 `Debtor Address Line 1`、`Creditor Address Line 1`、`Initiating Party Address Line 1`，将匹配单元格替换为同一记录的 `clean_addr_with_nmbrs` 完整文本，不拆分到 Line 2/3。其余 12 个地址列、其他列、行顺序、工作表和单元格格式保留。

表头位于第一行；表头和匹配 ID 忽略首尾空格，ID 区分大小写并保留前导零。Excel ID 必须为文本，数字或公式会报错。未匹配 ID 保留原值，空单元格保留；CSV 空 ID 跳过，匹配到空 `clean_addr_with_nmbrs` 时清空目标单元格。重复 ID 对应完全相同的清洗地址可复用，对应不同清洗地址则报错，原 Excel 不变。

成功时生成 `output/Transaction Columns for Fintech RFI_0929.xlsx`，保留 `input3/` 中的原始 Excel。终端分别显示三列的匹配、未匹配及空值数量。重复运行会重新读取原始 ID 并替换输出文件。输入验证失败时保留已有输出。真实文件被 Git 忽略。

默认 CSV 编码为 UTF-8（兼容 BOM）。其他编码或多个符合条件的工作表可显式指定：

```powershell
python map_address_details.py --encoding gb18030 --sheet "Transactions"
```

可使用 `--input3 "D:\your-folder\input3"` 指定存放这两个文件的其他目录，使用 `--output "D:\your-folder\output\result.xlsx"` 指定输出文件路径；输出路径不能与输入 Excel 相同。

将 `input/` 中所有 CSV 的数据行合并到 `output/Transaction Columns for Fintech RFI_0925_v1.xlsx`，固定输出 45 列。当前交付的 Excel 只有表头，没有示例或业务数据。

## Synthetic dataset notebook

`generate_synthetic_transactions.ipynb` is an English-only notebook that generates 1,000,000 fictional fintech AFC test transactions in `output/synthetic_wire_transactions_1m.csv`. It uses exactly the 12 screenshot columns, in the supplied order. It does not read any real input data. Generated CSVs and the quality report are excluded from Git.

Open the notebook in VS Code with the Jupyter extension or in Jupyter, select a Python 3.11 kernel, and run all cells. Install `ipykernel` into your selected environment if needed (`python -m pip install ipykernel`). The generator itself uses only the standard library. Set `PROJECT_DIR` if the notebook cannot locate the project automatically. The revised notebook sets `OVERWRITE = True` to replace the previous synthetic file; change it to `False` to protect existing output.

Valid dates range from 2025-01-01 through 2026-09-30. Product values are limited to Wires, Domestic ACH, and International ACH. Every `trxn_id` is nonblank and unique. Every amount is generated as a positive Python float and remains nonblank; CSV readers must parse the numeric token as a float because CSV has no native type metadata. Exactly 25% of transaction rows have one simulated PayPal, Payoneer, TTT MoneyCorp, Wise, or Stripe leg and one ordinary individual or business leg. The other 75% have no fintech leg. All names are synthetic and marked accordingly. The default `NOISE_RATE = 0.08` injects date, account, and country variations while protecting IDs, amounts, products, and core party names. Country noise includes `97`, `HH`, `ZZ`, and blanks in either country column.

Independently, `MISSING_ROW_RATE = 0.03` selects rows and blanks one to `MAX_MISSING_FIELDS = 3` eligible date, account, or country fields. Empty fields are written as actual blanks. The original 12-column schema is unchanged: no city columns or other fields are added. The English notebook explains generation, injection order, normalization, and missing-value handling; the report includes observed blank counts per column and observed country-noise counts.

The notebook streams generation and a full quality pass. `output/synthetic_afc_quality_report.json` reports injected noise, normalization counts, and rejection reasons. By default, no cleaned dataset is exported. Set `WRITE_CLEAN_OUTPUTS = True` to also export `synthetic_afc_clean.csv` and `synthetic_afc_quarantine.csv`; both preserve the original 12-column schema. Quarantined records retain original values; aggregate rejection reasons are in the separate JSON report. Invalid dates and ambiguous values are quarantined instead of guessed. The raw file always retains its noise. These are data-quality test rules, not financial-crime determinations. The 12-column synthetic file does not include all fields required by `convert.py` and should not be passed directly to that converter.

## Extended synthetic dataset: 291 screenshot fields

The final section of `generate_synthetic_transactions.ipynb`, titled **Extended AFC Dataset: 291 Screenshot Fields**, is independent of the earlier 12-column generator. Run its cells from the configuration cell onward to create `output/synthetic_afc_extended_1m.csv` and `output/synthetic_afc_extended_quality_report.json`. No real inputs are needed. The notebook is written in English and supports Python 3.11.

The new CSV transposes the six reference screenshots: the 291 field names from column A become the single CSV header row, and the next 1,000,000 rows contain fictional AFC fintech transactions. The exact ordering comprises 19 initial fields, 19 groups of 14 party fields, and 6 trailing fields. The earlier 12-column CSV remains separate.

`WIDE_NOISE_RATE = 0.15` selects exactly 15% of **transaction rows**, configurable between 10% and 20%, for noise. Scenarios include city `NOT FOUND`/`STREET`, country `97`/`HH`/`ZZ`, eligible blanks, invalid dates, account and entity errors, and invalid flags/codes. IDs, both amount fields, all product fields, and orig/bene raw/clean/master names are protected. Every ID is nonblank and unique, both amounts are positive floats, every product text field is limited to Wires, Domestic ACH, or International ACH, and exactly 25% of rows have one fintech leg paired with an ordinary individual or business. Naturally empty optional/intermediary groups are modeled separately and do not count as injected noise. All valid dates remain within 2025-01-01 to 2026-09-30.

Generation and the complete read-back audit stream records, with a temporary SQLite index for duplicate detection. Allow several GB of disk space and several minutes to run. The CSV retains noise and its exact schema; field lists, blank counts, normalization counts, and quality issues are recorded only in the separate JSON report. `WIDE_OVERWRITE = True` replaces an existing extended output. Codes and fixed FX factors are synthetic test assumptions, not definitions from the source system or live rates. Generated data and reports remain excluded from Git.

## Distinct values for both synthetic datasets

The final notebook section, **Distinct Values for Both Generated Datasets**, scans every column of both the 12-column and 291-column CSVs. Run its configuration, tests, and execution cells after generating the datasets. It does not change either dataset.

For each dataset it writes `<input_stem>_distinct_summary.csv` (one row per column, counts including/excluding blank, blank/nonblank rows, total rows, and a small preview) and `<input_stem>_distinct_values.csv` (every distinct value and its occurrence count, without truncating high-cardinality columns) into `output/`. Empty fields are explicitly distinguished from literal strings such as `NULL`. Case, whitespace, and leading zeros remain significant. Exact occurrence totals are checked against the row count for every column. The notebook prints counts and previews for all columns.

The profiler streams rows and spills high-cardinality counters into a disposable SQLite database. Complete value reports can be large, so allow additional disk space and time. `DISTINCT_OVERWRITE = True` replaces reports only. Source CSVs are never rewritten; all report files remain excluded from Git.

## 使用

本项目支持 Python 3.11 和 openpyxl。不需要 Node.js、Codex 或安装 Microsoft Excel。

### Windows

1. 从 https://www.python.org/downloads/windows/ 安装 Python，启用 Python launcher，并勾选 Add Python to PATH。
2. 在 GitHub 仓库点击 Code → Download ZIP，解压到本地文件夹。不要直接在 ZIP 内运行。
3. 把真实 CSV 放进解压后项目的 `input` 文件夹。
4. 关闭已打开的同名输出 Excel，双击 `run_windows.bat`。首次运行需要联网安装 openpyxl。
5. 在 `output` 文件夹查看输出 Excel。窗口会显示成功提示或错误信息，按任意键关闭。

也可以打开项目目录的 PowerShell，执行：

```powershell
py -3.11 -m pip install -r requirements.txt
py -3.11 convert.py
```

在 Conda 环境中，先运行 `python --version` 确认是 3.11，再运行 `python -m pip install -r requirements.txt` 和 `python convert.py`。无需 `py` launcher。如果提示 Permission denied，请先关闭输出 Excel。公司电脑无法安装依赖时，请联系 IT 配置 Python 和 openpyxl。

### macOS / Linux

首次运行安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

1. 把真实 CSV 文件放进 `input/`（只读取该目录，不递归子目录）。
2. 在项目目录运行：

```bash
python3 convert.py
```

也可在任意工作目录运行脚本的绝对路径。默认 input/output 始终位于项目目录中。

其他选项：

```bash
python3 convert.py --encoding gb18030
python3 convert.py --input "/path/to/input" --output "/path/to/result.xlsx"
python3 convert.py --empty-template
```

`--empty-template` 忽略 input，只写表头。每次成功运行会替换指定输出文件；如需保留历史版本，请使用不同的 `--output` 路径。输入验证失败时不会替换原输出。

## 映射和数据规则

- `mapping.py` 是唯一的列顺序和映射定义。输出使用最终清单中的 `Creditor Account ID`、`Ultimate Creditor Account ID` 和带右括号的交易 ID 列名。
- Ultimate Debtor 使用 `ult_orig_*` 五个来源字段；Ultimate Creditor 使用 `ult_bene_*` 五个来源字段。Debtor 和 Creditor 分别使用 `orig_*` 和 `bene_*`。这些来源列必须存在，单元格值可以为空。
- 按 CSV 文件名排序合并，保留文件内行顺序。每条数据行的客户名为所属文件名去掉最后的 `.csv` 扩展名。
- 所有输入字段按文本读取，账号和交易 ID 的前导零不会丢失。日期直接保留 `value_date` 原文，避免猜测不同日期格式。
- `amount` 按十进制数除以 100，Excel 中为数值；空金额保留空白，零金额保留零。接受小数、负数及科学计数法，不接受千分位逗号或货币符号；不进行汇率转换或自行四舍五入。超出 Excel 15 位有效数字的金额会报错。
- 映射中的 Address Line 2/3 一律为空。空字符串和仅含空白的单元格输出为空；`NULL`、`NA` 等非空文本保持原样。
- 标准逗号分隔 CSV，支持带引号的逗号、换行和 UTF-8 BOM，默认编码为 UTF-8。列名必须准确匹配，允许额外列。缺失必需列、重复表头、字段数量不一致或非法金额均报错。物理空行由 CSV 解析器跳过；有完整字段但值均为空的记录仍作为一条数据行处理。
- 输入文本以 `=` 开头时作为文字写入，不作为 Excel 公式执行。
- 名称检查：将客户名、Debtor Name 和 Creditor Name 忽略大小写、去除所有空白字符后比较。任一名称包含完整客户名即匹配；两者均不包含时，输出整行 A:AS（包括空白单元格）填充黄色。两者均为空也标黄。比如 `Billy Kim` 可以匹配 `BILLYKIM` 或 `Payment for BILLY KIM Ltd`。标点不会被删除，不做拼写近似匹配；原始名称保持不变。
- 标黄在每次导出时计算。手动修改输出 Excel 后颜色不会自动更新，需要重新运行转换。终端显示标黄行数；空白模板只有表头，无黄色数据行。
- 单表最多 1,048,575 条数据行。当前实现会在内存中合并数据，超大文件的可处理规模取决于本机内存。

## 文件

- `convert.py`：CSV 读取、验证、金额计算和命令行入口。
- `mapping.py`：45 列顺序及映射。
- `requirements.txt`：Python 依赖。
- `run_windows.bat`：Windows 双击运行入口。
- `input/`：放入真实 CSV；交付时为空。
- `output/`：Excel 输出。

Excel 导出已集成到 `convert.py`，使用 openpyxl；项目不再依赖本机 Codex 运行时或 `node_modules`。

## 验证范围

已检查空白模板的 45 列名称、顺序和无数据行状态，并预览表头。未创建 synthetic input dataset。尚未用真实 CSV 验证多文件转换，也未在 Windows 实机运行；提供真实数据后可运行转换并核对输入/输出行数。

## 第二步：从 input2 补充地址

已有 `Transaction Columns for Fintech RFI_0928.xlsx` 时，直接运行这一步，不需要重新运行 `convert.py`。

1. 将 `Batch1.csv` 至 `Batch8.csv` 全部放入 `input2/`。
2. 将现有 Excel 放入 `output/Transaction Columns for Fintech RFI_0928.xlsx`，运行前关闭 Excel。
3. 在 Windows 的 Python 3.11 / Conda 终端执行：

```powershell
python -m pip install -r requirements.txt
python fill_addresses.py
```

Excel 在其他位置时：

```powershell
python fill_addresses.py --workbook "C:\Users\yourname\Documents\Transaction Columns for Fintech RFI_0928.xlsx"
```

默认 CSV 为 UTF-8（兼容 BOM）；需要时加 `--encoding gb18030`。有多个交易工作表时加 `--sheet "Transactions"`。

程序先按 Batch1 至 Batch8 的顺序生成 `input2/batch_combined.csv`，只保留一份表头，保留各批次全部记录及额外列。不把旧的 batch_combined.csv 再次合并。八个文件必须全部存在，每个文件都必须包含 `cinq_trxn_id` 以及下面 15 个来源列。不同批次的额外列取并集，缺少的额外列留空。

| Excel 输出列 | Batch 来源列 |
|---|---|
| Ultimate Debtor Address Line 1 | ult_deb_address_line1 |
| Ultimate Debtor Address Line 2 | ult_deb_address_line2 |
| Ultimate Debtor Address Line 3 | ult_deb_address_line3 |
| Debtor Address Line 1 | orig_addr_line1 |
| Debtor Address Line 2 | orig_addr_line2 |
| Debtor Address Line 3 | orig_addr_line3 |
| Creditor Address Line 1 | secondary_bene_addr_line1 |
| Creditor Address Line 2 | secondary_bene_addr_line2 |
| Creditor Address Line 3 | secondary_bene_addr_line3 |
| Ultimate Creditor Address Line 1 | ult_cred_address_line1 |
| Ultimate Creditor Address Line 2 | ult_cred_address_line2 |
| Ultimate Creditor Address Line 3 | ult_cred_address_line3 |
| Initiating Party Address Line 1 | ip_address_line1 |
| Initiating Party Address Line 2 | ip_address_line2 |
| Initiating Party Address Line 3 | ip_address_line3 |

地址取自生成的 batch_combined.csv。Excel 的 `Transaction ID (UETR or Other Unique ID)` 与 `cinq_trxn_id` 按文本精确匹配，去掉首尾空白，保留大小写、内部字符和前导零。数值类型或公式类型的 Excel 交易 ID 会报错，避免精度损失造成错误匹配。

只更新原工作表的 15 个地址列，覆盖旧地址；其他列、行顺序、重复交易记录和原来的黄色高亮保留。匹配不到、空交易 ID、来源空地址均留空，不回退到旧 input 的地址。空 CSV 交易 ID 不参与匹配。重复 CSV ID 的 15 个地址完全相同时可复用；地址不同则报错并保留原 Excel，合并 CSV 仍可用于排查。不会因为重复 ID 增加 Excel 的行数。

保存前自动创建带时间戳的原文件备份，再更新指定 Excel。终端显示各批次行数、匹配/未匹配行数。batch_combined.csv、真实批次、Excel 及备份不会作为新文件上传 GitHub。

当前工作区未提供 8 个真实 Batch 或 0928 Excel，因此地址填充尚未实际执行，未生成模拟数据。
