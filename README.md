# Fintech CSV → Excel

将 `input/` 中所有 CSV 的数据行合并到 `output/Transaction Columns for Fintech RFI_0925_v1.xlsx`，固定输出 45 列。当前交付的 Excel 只有表头，没有示例或业务数据。

## 使用

本机已连接 Codex 的 `@oai/artifact-tool` 依赖。需要 Python 3.9+ 和 Node.js 22+，Python 不需要安装第三方包。

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
- `export.mjs`：Excel 导出、表头样式及冻结首行/首列。
- `input/`：放入真实 CSV；交付时为空。
- `output/`：Excel 输出。

`node_modules` 是指向本机 Codex 运行时依赖的链接。迁移到另一台电脑时，需要让 `@oai/artifact-tool` 可被该项目的 Node.js 加载；不要直接复制本机依赖链接。

## 验证范围

已生成并检查空白模板的 45 列名称、顺序和无数据行状态，并预览表头。未创建 synthetic input dataset。尚未用真实 CSV 验证多文件转换；提供真实数据后可运行转换并核对输入/输出行数。
