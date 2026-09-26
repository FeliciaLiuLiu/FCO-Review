#!/usr/bin/env python3
"""Combine customer CSVs into the requested Excel workbook. Python 3.9+."""
import argparse
import csv
import json
import math
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
import re
import subprocess
import sys

from mapping import MAPPING, HEADERS

ROOT = Path(__file__).resolve().parent
DEFAULT_OUTPUT = ROOT / 'output' / 'Transaction Columns for Fintech RFI_0925_v1.xlsx'

def amount_value(raw, location):
    if not raw.strip():
        return None
    if not re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?', raw.strip()):
        raise ValueError(f'{location}: amount 不是有效数字')
    try:
        with localcontext() as context:
            context.prec = max(28, len(raw) + 4)
            value = Decimal(raw.strip()) / Decimal(100)
        number = float(value)
        if not math.isfinite(number) or Decimal(str(number)) != value:
            raise ValueError(f'{location}: amount 超出 Excel 可安全表示的数值范围')
        if len(value.normalize().as_tuple().digits) > 15:
            raise ValueError(f'{location}: amount 超过 Excel 的 15 位有效数字精度')
        return number
    except (InvalidOperation, OverflowError):
        raise ValueError(f'{location}: amount 不是有效数字') from None

def read_rows(folder, encoding):
    files = sorted((p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == '.csv'),
                   key=lambda p: p.name)
    if not files:
        raise ValueError('input 文件夹没有 CSV；如需空白模板，请使用 --empty-template')
    required = {source for _, source in MAPPING if source and source != '$filename'}
    rows = []
    counts = []
    for path in files:
        count = 0
        with path.open('r', encoding=encoding, newline='') as handle:
            reader = csv.DictReader(handle, strict=True)
            fields = reader.fieldnames
            if not fields or len(fields) != len(set(fields)):
                raise ValueError(f'{path.name}: CSV 表头为空或存在重复列名')
            missing = sorted(required - set(fields))
            if missing:
                raise ValueError(f'{path.name}: 缺少列: {", ".join(missing)}')
            for record in reader:
                location = f'{path.name}, CSV 行 {reader.line_num}'
                if None in record or any(v is None for v in record.values()):
                    raise ValueError(f'{location}: 字段数量与表头不一致')
                output = []
                for _, source in MAPPING:
                    if source == '$filename':
                        value = path.stem
                    elif source is None:
                        value = None
                    elif source == 'amount':
                        value = amount_value(record[source], location)
                    else:
                        value = record[source] if record[source].strip() else None
                    output.append(value)
                rows.append(output)
                count += 1
                if len(rows) > 1048575:
                    raise ValueError('数据超过单个 Excel 工作表允许的 1,048,575 条数据行')
        counts.append((path.name, count))
    return rows, counts

def main():
    parser = argparse.ArgumentParser(description='将 input 中客户 CSV 合并为 45 列 Excel')
    parser.add_argument('--input', type=Path, default=ROOT / 'input')
    parser.add_argument('--output', type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument('--encoding', default='utf-8-sig', help='默认 UTF-8，兼容 BOM；可指定 gb18030')
    parser.add_argument('--empty-template', action='store_true', help='忽略 CSV，仅生成表头')
    parser.add_argument('--preview', action='store_true', help='生成表头预览供检查')
    args = parser.parse_args()
    try:
        rows, counts = ([], []) if args.empty_template else read_rows(args.input, args.encoding)
        payload = {'headers': HEADERS, 'rows': rows, 'output': str(args.output.resolve()),
                   'preview': args.preview}
        subprocess.run(['node', str(ROOT / 'export.mjs')], input=json.dumps(payload, ensure_ascii=False),
                       text=True, check=True, cwd=ROOT)
        for name, count in counts:
            print(f'{name}: {count} 行')
        print(f'完成：{len(rows)} 条数据，{len(HEADERS)} 列。输出：{args.output.resolve()}')
    except (ValueError, OSError, UnicodeError, csv.Error, subprocess.CalledProcessError) as error:
        print(f'转换失败：{error}', file=sys.stderr)
        return 1
    return 0

if __name__ == '__main__':
    sys.exit(main())
