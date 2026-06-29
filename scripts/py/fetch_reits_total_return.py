"""Fetch CSI REITs total return data and export it for the VBA workflow."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


SYMBOL = "932047"
INDEX_NAME = "中证REITs全收益"
BASE_DATE = datetime(2025, 7, 17)
START_DATE = "20250717"


def prompt_end_date() -> str:
    prompt = "请输入结束日期，格式为 yyyymmdd，例如 20260625："

    try:
        import tkinter as tk
        from tkinter import simpledialog

        root = tk.Tk()
        root.withdraw()
        value = simpledialog.askstring("REITs全收益数据", prompt)
        root.destroy()
    except Exception:
        value = input(prompt)

    if value is None:
        raise ValueError("用户取消输入。")
    return value.strip()


def parse_end_date(value: str) -> str:
    try:
        parsed = datetime.strptime(value, "%Y%m%d")
    except ValueError as exc:
        raise ValueError("结束日期格式错误，请输入 yyyymmdd，例如 20260625。") from exc

    if parsed < BASE_DATE:
        raise ValueError("结束日期不能早于基准日期 20250717。")
    return parsed.strftime("%Y%m%d")


def calc_annualized_return(row: pd.Series, base_actual_date: pd.Timestamp, base_close: float) -> float:
    row_date = row["日期_dt"]
    close = row["收盘"]
    days = (row_date - base_actual_date).days
    if days == 0:
        return 0.0
    if days > 0:
        return (close - base_close) / base_close * (365.0 / days)
    return (base_close - close) / close * (365.0 / abs(days))


def fetch_data(end_date: str) -> pd.DataFrame:
    try:
        import akshare as ak
    except ImportError as exc:
        raise RuntimeError("缺少 akshare，请先运行：uv add akshare 或 pip install akshare") from exc

    df = ak.stock_zh_index_hist_csindex(
        symbol=SYMBOL,
        start_date=START_DATE,
        end_date=end_date,
    )
    if df.empty:
        raise RuntimeError("akshare 未返回数据，请检查网络、指数代码或日期范围。")

    df = df.copy()
    df["日期_dt"] = pd.to_datetime(df["日期"])
    df_sorted = df.sort_values("日期_dt")

    base_rows = df_sorted[df_sorted["日期_dt"] <= BASE_DATE]
    if base_rows.empty:
        base_rows = df_sorted.iloc[[(df_sorted["日期_dt"] - BASE_DATE).abs().idxmin()]]

    base_close = float(base_rows.iloc[-1]["收盘"])
    base_actual_date = base_rows.iloc[-1]["日期_dt"]

    df["REITs期间年化收益率"] = df.apply(
        calc_annualized_return,
        axis=1,
        base_actual_date=base_actual_date,
        base_close=base_close,
    )
    df["REITs收盘"] = df["收盘"]

    selected_columns = ["日期", "指数代码", "指数中文全称", "REITs收盘", "REITs期间年化收益率"]
    return df[selected_columns].sort_values("日期")


def export_data(df: pd.DataFrame, end_date: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{end_date}-{INDEX_NAME}.xlsx"
    df.to_excel(output_path, index=False)
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="导出中证REITs全收益数据。")
    parser.add_argument("--end-date", help="结束日期，格式为 yyyymmdd，例如 20260625。")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).resolve().parents[2],
        help="输出目录，默认项目根目录。",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    try:
        end_date = parse_end_date(args.end_date or prompt_end_date())
        df = fetch_data(end_date)
        output_path = export_data(df, end_date, args.output_dir)
    except Exception as exc:
        print(f"导出失败：{exc}", file=sys.stderr)
        return 1

    print(f"数据已导出至：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
