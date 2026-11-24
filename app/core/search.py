# -*- coding: utf-8 -*-
"""
search.py — 关键词检索 (AND/OR) + 时间筛选 + 导出
"""
import pandas as pd


def _period_mask(df, start, end):
    start = pd.Timestamp(start)
    end = pd.Timestamp(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return (df["时间"] >= start) & (df["时间"] <= end)


def keyword_search(df, keywords, mode="or", start=None, end=None):
    """
    keywords: ["食品", "药品"]
    mode: "and"=同时包含 / "or"=包含任一
    返回: 命中的 DataFrame (按时间排序)
    """
    keywords = [k.strip() for k in keywords if k and k.strip()]
    if not keywords:
        return df.iloc[0:0]

    mask = None
    for k in keywords:
        m = df["内容"].str.contains(k, na=False, regex=False)
        mask = m if mask is None else (mask & m if mode == "and" else mask | m)

    if start and end:
        mask &= _period_mask(df, start, end)

    return df[mask].sort_values("时间").reset_index(drop=True)


def export_df(df, path, extra_sheets=None):
    """导出 xlsx; extra_sheets: {sheet名: DataFrame}"""
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="明细")
        if extra_sheets:
            for name, sheet_df in extra_sheets.items():
                sheet_df.to_excel(w, index=False, sheet_name=name[:31])
