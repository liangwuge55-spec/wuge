#!/usr/bin/env python3
"""读取 Excel 文件，去重后输出为新文件。"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="读取 Excel → 去重 → 输出新文件")
    parser.add_argument("input", type=Path, help="输入 Excel 文件路径，例如: data.xlsx")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="输出 Excel 文件路径（默认: 原文件名 + _dedup.xlsx）",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        default=0,
        help="要处理的工作表名称或索引，默认第一个工作表",
    )
    parser.add_argument(
        "-k",
        "--keep",
        choices=("first", "last", "none"),
        default="first",
        help="重复值保留策略：first(保留第一条) / last(保留最后一条) / none(全部删掉)",
    )
    parser.add_argument(
        "--subset",
        nargs="+",
        default=None,
        help="指定用于去重的列名；不填则按整行去重",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if not args.input.exists():
        raise FileNotFoundError(f"输入文件不存在: {args.input}")

    output = args.output or args.input.with_name(f"{args.input.stem}_dedup.xlsx")

    keep_value: str | bool
    keep_value = False if args.keep == "none" else args.keep

    df = pd.read_excel(args.input, sheet_name=args.sheet)
    before = len(df)

    deduped = df.drop_duplicates(subset=args.subset, keep=keep_value)
    after = len(deduped)

    deduped.to_excel(output, index=False)

    print(f"处理完成：{args.input} -> {output}")
    print(f"原始行数: {before}")
    print(f"去重后行数: {after}")
    print(f"删除重复行数: {before - after}")


if __name__ == "__main__":
    main()
