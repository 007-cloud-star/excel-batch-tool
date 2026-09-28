#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 批量处理工具（作品样品）
================================
一个开箱即用的 Excel 批量处理小工具，适合电商、行政、财务等需要
反复整理表格的人群。零学习成本，一条命令完成常见脏活累活。

功能
----
merge    合并：把一个文件夹下所有 .xlsx 的第一个工作表纵向合并
clean    清洗：删除空行、去除文本首尾空格、统一日期列格式
dedup    去重：按一列或多列组合去重（保留第一条）
summary  汇总：按分组列对数值列求和，并统计订单数

安装
----
pip install openpyxl

使用示例
--------
python excel_batch_tool.py merge ./data -o 合并结果.xlsx
python excel_batch_tool.py clean 原始数据.xlsx -o 清洗后.xlsx --date-cols 下单日期,付款日期
python excel_batch_tool.py dedup 清洗后.xlsx -o 去重后.xlsx --keys 手机号
python excel_batch_tool.py summary 去重后.xlsx -o 销售汇总.xlsx --group 商品名称 --sum 销售额
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook, Workbook


def read_rows(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def write_rows(path, rows):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def cmd_merge(args):
    folder = Path(args.input)
    files = sorted(folder.glob("*.xlsx"))
    if not files:
        print(f"文件夹 {folder} 下没有找到 xlsx 文件")
        sys.exit(1)
    header, out = None, []
    for f in files:
        rows = [r for r in read_rows(f) if any(c not in (None, "") for c in r)]
        if not rows:
            continue
        if header is None:
            header = rows[0]
            out.append(header)
        out.extend(rows[1:])
    write_rows(args.output, out)
    print(f"已合并 {len(files)} 个文件，共 {len(out) - 1} 行数据 -> {args.output}")


def _clean_value(v, date_idx, col_idx):
    if isinstance(v, str):
        v = v.strip()
        if col_idx in date_idx:
            for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%d/%m/%Y"):
                try:
                    return datetime.strptime(v, fmt).date().isoformat()
                except ValueError:
                    pass
    if isinstance(v, datetime):
        return v.date().isoformat()
    return v


def cmd_clean(args):
    rows = [r for r in read_rows(args.input) if any(c not in (None, "") for c in r)]
    if not rows:
        print("输入文件没有数据")
        sys.exit(1)
    header = [str(c).strip() if isinstance(c, str) else c for c in rows[0]]
    date_cols = [c.strip() for c in (args.date_cols or "").split(",") if c.strip()]
    date_idx = {i for i, h in enumerate(header) if h in date_cols}
    cleaned = [header]
    for r in rows[1:]:
        cleaned.append(tuple(_clean_value(v, date_idx, i) for i, v in enumerate(r)))
    write_rows(args.output, cleaned)
    print(f"清洗完成：{len(rows) - 1} -> {len(cleaned) - 1} 行 -> {args.output}")


def cmd_dedup(args):
    rows = read_rows(args.input)
    header = list(rows[0])
    keys = [k.strip() for k in args.keys.split(",") if k.strip()]
    idx = [header.index(k) for k in keys]
    seen, out = set(), [header]
    for r in rows[1:]:
        sig = tuple(r[i] for i in idx)
        if sig not in seen:
            seen.add(sig)
            out.append(r)
    write_rows(args.output, out)
    print(f"去重完成：{len(rows) - 1} -> {len(out) - 1} 行（按 {keys} 去重）-> {args.output}")


def cmd_summary(args):
    rows = read_rows(args.input)
    header = list(rows[0])
    gi, si = header.index(args.group), header.index(args.sum)
    totals, counts = {}, {}
    for r in rows[1:]:
        try:
            val = float(r[si] or 0)
        except (TypeError, ValueError):
            val = 0
        totals[r[gi]] = totals.get(r[gi], 0) + val
        counts[r[gi]] = counts.get(r[gi], 0) + 1
    out = [[args.group, f"{args.sum}汇总", "订单数"]]
    for k in sorted(totals, key=str):
        out.append([k, round(totals[k], 2), counts[k]])
    out.append(["合计", round(sum(totals.values()), 2), sum(counts.values())])
    write_rows(args.output, out)
    print(f"汇总完成：{len(totals)} 个分组 -> {args.output}")


def main():
    ap = argparse.ArgumentParser(description="Excel 批量处理工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("merge", help="合并文件夹下所有 xlsx")
    p.add_argument("input", help="存放 xlsx 的文件夹")
    p.add_argument("-o", "--output", required=True, help="输出文件")

    p = sub.add_parser("clean", help="清洗数据：去空行/去空格/统一日期")
    p.add_argument("input", help="输入 xlsx")
    p.add_argument("-o", "--output", required=True, help="输出文件")
    p.add_argument("--date-cols", default="", help="日期列名，逗号分隔")

    p = sub.add_parser("dedup", help="按指定列去重")
    p.add_argument("input", help="输入 xlsx")
    p.add_argument("-o", "--output", required=True, help="输出文件")
    p.add_argument("--keys", required=True, help="去重依据列名，逗号分隔")

    p = sub.add_parser("summary", help="按分组列汇总数值列")
    p.add_argument("input", help="输入 xlsx")
    p.add_argument("-o", "--output", required=True, help="输出文件")
    p.add_argument("--group", required=True, help="分组列名")
    p.add_argument("--sum", required=True, help="求和数值列名")

    args = ap.parse_args()
    {"merge": cmd_merge, "clean": cmd_clean,
     "dedup": cmd_dedup, "summary": cmd_summary}[args.cmd](args)


if __name__ == "__main__":
    main()
