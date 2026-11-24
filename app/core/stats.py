# -*- coding: utf-8 -*-
"""
stats.py — 统计: 总量 / 同比 / 分分局 / 月度趋势
"""
import pandas as pd


def _period_mask(df, start, end):
    """闭区间 [start, end]，end 含当天 23:59:59"""
    start = pd.Timestamp(start)
    end = pd.Timestamp(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return (df["时间"] >= start) & (df["时间"] <= end)


def _last_year_range(start, end):
    """去年同期区间"""
    s = pd.Timestamp(start) - pd.DateOffset(years=1)
    e = pd.Timestamp(end) - pd.DateOffset(years=1)
    return s, e


def yoy(cur, prev):
    """同比增减%: prev=0 时返回 None (界面显示'—')"""
    if prev == 0:
        return None
    return round((cur - prev) / prev * 100, 1)


def overview(df, start, end):
    """总量 + 分平台 + 同比"""
    cur = df[_period_mask(df, start, end)]
    ls, le = _last_year_range(start, end)
    prev = df[_period_mask(df, ls, le)]

    platforms = {}
    for name in ["12315", "市12345", "省12345"]:
        c = int((cur["平台"] == name).sum())
        p = int((prev["平台"] == name).sum())
        platforms[name] = {"cur": c, "prev": p, "yoy": yoy(c, p)}

    return {
        "range": {"start": str(start), "end": str(end),
                  "prev_start": str(ls.date()), "prev_end": str(le.date())},
        "total": {"cur": len(cur), "prev": len(prev), "yoy": yoy(len(cur), len(prev))},
        "platforms": platforms,
    }


def by_fenju(df, start, end):
    """分分局: 本期数 + 去年同期 + 同比, 按本期数降序"""
    cur = df[_period_mask(df, start, end)]
    ls, le = _last_year_range(start, end)
    prev = df[_period_mask(df, ls, le)]

    c_cnt = cur["分局"].value_counts()
    p_cnt = prev["分局"].value_counts()
    names = list(dict.fromkeys(list(c_cnt.index) + list(p_cnt.index)))
    rows = []
    for n in names:
        c, p = int(c_cnt.get(n, 0)), int(p_cnt.get(n, 0))
        rows.append({"分局": n, "cur": c, "prev": p, "yoy": yoy(c, p)})
    rows.sort(key=lambda r: r["cur"], reverse=True)
    return rows


def monthly_trend(df, start, end):
    """按月统计: 本期 vs 去年同期 (画柱状/折线图用)"""
    cur = df[_period_mask(df, start, end)].copy()
    ls, le = _last_year_range(start, end)
    prev = df[_period_mask(df, ls, le)].copy()

    cur["月份"] = cur["时间"].dt.strftime("%Y-%m")
    prev["月份"] = prev["时间"].dt.strftime("%Y-%m")
    c = cur["月份"].value_counts().sort_index()
    p = prev["月份"].value_counts().sort_index()
    return {
        "cur_months": list(c.index), "cur_counts": [int(x) for x in c.values],
        "prev_months": list(p.index), "prev_counts": [int(x) for x in p.values],
    }
