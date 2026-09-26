# Fintech CSV → Excel

将 `input/` 中所有 CSV 的数据行合并到 `output/Transaction Columns for Fintech RFI_0925_v1.xlsx`，固定输出 45 列。当前交付的 Excel 只有表头，没有示例或业务数据。

## 使用

需要 Python 3.9+（推荐 Python 3.12 或更新版本）和 openpyxl。不需要 Node.js、Codex 或安装 Microsoft Excel。

### Windows

1. 从 https://www.python.org/downloads/windows/ 安装 Python，启用 Python launcher，并勾选 Add Python to PATH。
2. 在 GitHub 仓库点击 Code → Download ZIP，解压到本地文件夹。不要直接在 ZIP 内运行。
3. 把真实 CSV 放进解压后项目的 `input` 文件夹。
4. 关闭已打开的同名输出 Excel，双击 `run_windows.bat`。首次运行需要联网安装 openpyxl。
5. 在 `output` 文件夹查看输出 Excel。窗口会显示成功提示或错误信息，按任意键关闭。

也可以打开项目目录的 PowerShell，执行：

```powershell
py -3 -m pip install -r requirements.txt
py -3 convert.py
```

如果提示找不到 `py`，尝试用 `python` 替代，或重新安装 Python 并启用 launcher。如果提示 Permission denied，请先关闭输出 Excel。公司电脑无法安装依赖时，请联系 IT 配置 Python 和 openpyxl。

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
- 按 CSV 文件名排序合并，保留文件内行顺序。每条数据行的客户名为所属文件名去掉最后的 `.csv` 扩展名。
- 所有输入字段按文本读取，账号和交易 ID 的前导零不会丢失。日期直接保留 `value_date` 原文，避免猜测不同日期格式。
- `amount` 按十进制数除以 100，Excel 中为数值；空金额保留空白，零金额保留零。接受小数、负数及科学计数法，不接受千分位逗号或货币符号；不进行汇率转换或自行四舍五入。超出 Excel 15 位有效数字的金额会报错。
- 映射中的 Address Line 2/3 一律为空。空字符串和仅含空白的单元格输出为空；`NULL`、`NA` 等非空文本保持原样。
- 标准逗号分隔 CSV，支持带引号的逗号、换行和 UTF-8 BOM，默认编码为 UTF-8。列名必须准确匹配，允许额外列。缺失必需列、重复表头、字段数量不一致或非法金额均报错。物理空行由 CSV 解析器跳过；有完整字段但值均为空的记录仍作为一条数据行处理。
- 输入文本以 `=` 开头时作为文字写入，不作为 Excel 公式执行。
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
