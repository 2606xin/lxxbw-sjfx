# -*- coding: utf-8 -*-
"""api.py — 界面与核心逻辑之间的桥接层（不依赖 webview，可独立测试）"""
import os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
from paths import resource_path
from core import loader, stats, search


def export_dir():
    """导出目录：固定在程序旁的 导出结果/ 文件夹（无 GUI 对话框依赖）
    - exe 版: exe 所在目录/导出结果
    - 开发版: 项目根目录/导出结果
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(resource_path(".")))  # app/ 的上级
    d = os.path.join(base, "导出结果")
    os.makedirs(d, exist_ok=True)
    return d


def _stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


class Api:
    def __init__(self):
        self._files = []      # [{name, platform, raw, meta}]
        self._df = None       # 合并 meta（含 _fid/_rid 回溯原始行）
        self._window = None

    # ================= 文件导入 =================
    def _load(self, paths):
        """加载路径列表 → 重建 self._files / self._df"""
        files, merged, reports = loader.load_files(list(paths))
        self._files = files
        # 给 merged 加 _fid/_rid 以便回溯到原始表行
        frames = []
        for fid, f in enumerate(files):
            m = f["meta"].copy()
            m["_fid"] = fid
            m["_rid"] = range(len(m))
            frames.append(m)
        self._df = pd.concat(frames, ignore_index=True) if frames else \
            pd.DataFrame(columns=["平台", "单号", "时间", "分局", "内容", "_fid", "_rid"])
        self._df = self._df.sort_values("时间").reset_index(drop=True)
        return reports

    def pick_files(self):
        """多选 Excel 文件导入"""
        import webview
        paths = self._window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=True,
            file_types=("Excel 文件 (*.xls;*.xlsx)", "所有文件 (*.*)"))
        if not paths:
            return {"cancelled": True}
        reports = self._load(list(paths))
        return {"cancelled": False, "reports": reports, "summary": self._summary()}

    def pick_folder(self):
        """选择文件夹，导入其中全部 Excel"""
        import webview
        folder = self._window.create_file_dialog(webview.FOLDER_DIALOG)
        if not folder:
            return {"cancelled": True}
        folder = folder[0] if isinstance(folder, (tuple, list)) else folder
        paths = loader.list_excel_files(folder)
        if not paths:
            return {"cancelled": False, "reports": [], "summary": self._summary(),
                    "msg": f"文件夹里没有 Excel 文件: {folder}"}
        reports = self._load(paths)
        return {"cancelled": False, "reports": reports, "summary": self._summary()}

    def _summary(self):
        if self._df is None or not len(self._df):
            return {"total": 0}
        return {
            "total": len(self._df),
            "min_time": self._df["时间"].min().strftime("%Y-%m-%d"),
            "max_time": self._df["时间"].max().strftime("%Y-%m-%d"),
            "platforms": {k: int(v) for k, v in self._df["平台"].value_counts().items()},
            "files": len(self._files),
        }

    def get_state(self):
        return {"loaded": self._df is not None and len(self._df) > 0,
                "summary": self._summary()}

    def get_filter_options(self):
        """返回筛选面板可选项：平台、分局"""
        if self._df is None or not len(self._df):
            return {"platforms": [], "fenju": []}
        return {
            "platforms": [str(x) for x in self._df["平台"].unique()],
            "fenju": [str(x) for x in self._df["分局"].unique()],
        }

    # ================= 统计 =================
    def _require_df(self):
        if self._df is None or not len(self._df):
            raise RuntimeError("请先选择并加载 Excel 文件")
        return self._df

    def get_dashboard(self, start, end):
        df = self._require_df()
        return {
            "overview": stats.overview(df, start, end),
            "fenju": stats.by_fenju(df, start, end),
            "trend": stats.monthly_trend(df, start, end),
        }

    # ================= 多条件筛选 =================
    def _filter_mask(self, platforms=None, fenju=None, start=None, end=None,
                     keywords=None, mode="or"):
        """构造筛选布尔掩码（全条件 AND 组合）"""
        df = self._df
        mask = pd.Series(True, index=df.index)
        if platforms:
            mask &= df["平台"].isin(platforms)
        if fenju:
            mask &= df["分局"].isin(fenju)
        if start and end:
            s = pd.Timestamp(start)
            e = pd.Timestamp(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            mask &= (df["时间"] >= s) & (df["时间"] <= e)
        if keywords:
            keys = [k.strip() for k in keywords if k and k.strip()]
            if keys:
                for i, k in enumerate(keys):
                    m = df["内容"].str.contains(k, na=False, regex=False)
                    if i == 0:
                        kw_mask = m
                    else:
                        kw_mask = kw_mask & m if mode == "and" else kw_mask | m
                mask &= kw_mask
        return mask

    def filter_rows(self, platforms=None, fenju=None, start=None, end=None,
                    keywords=None, mode="or"):
        """多条件筛选 → {count, loss_sum, loss_rows, rows(预览200), truncated}
        loss_sum: 筛选结果中 12315「挽回经济损失」合计（元）
        """
        df = self._require_df()
        if not self._files:
            return {"count": 0, "loss_sum": 0.0, "loss_rows": 0,
                    "rows": [], "truncated": False}
        hit = df[self._filter_mask(platforms, fenju, start, end, keywords, mode)]
        loss_series = pd.to_numeric(hit.get("挽回损失", 0), errors="coerce").fillna(0.0)
        loss_sum = float(loss_series.sum())
        loss_rows = int((loss_series > 0).sum())
        preview = hit.head(200).copy()
        preview["时间"] = preview["时间"].dt.strftime("%Y-%m-%d %H:%M")
        preview["内容"] = preview["内容"].str.slice(0, 120)
        if "挽回损失" in preview.columns:
            preview["挽回损失"] = preview["挽回损失"].fillna(0).astype(float).round(2)
        cols = [c for c in ["平台", "单号", "时间", "分局", "挽回损失", "内容"] if c in preview.columns]
        return {"count": len(hit), "loss_sum": round(loss_sum, 2), "loss_rows": loss_rows,
                "truncated": len(hit) > 200, "rows": preview[cols].to_dict("records")}

    # ================= 原格式导出 =================
    def export_filtered(self, platforms=None, fenju=None, start=None, end=None,
                        keywords=None, mode="or"):
        """按筛选条件导出，保持原始 Excel 格式：
        每个平台一个 sheet，列=该平台原表头，行=原始数据行"""
        self._require_df()
        if not self._files:
            return {"ok": False, "msg": "没有已加载的数据"}
        hit = self._df[self._filter_mask(platforms, fenju, start, end, keywords, mode)]
        if not len(hit):
            return {"ok": False, "msg": "筛选结果为空，没有可导出的数据"}

        tag = "-".join(keywords) if keywords else "全部"
        path = os.path.join(export_dir(), f"筛选导出_{tag}_{_stamp()}.xlsx")

        # 按平台分组；每个平台在其原始表上取命中行（保持原表头/原内容）
        by_platform = {}
        for fid, f in enumerate(self._files):
            hit_f = hit[hit["_fid"] == fid]
            if len(hit_f):
                ridx = hit_f["_rid"].values
                raw_sub = f["raw"].iloc[ridx].copy()
                by_platform.setdefault(f["platform"], []).append(raw_sub)

        with pd.ExcelWriter(path, engine="openpyxl") as w:
            for pname, parts in by_platform.items():
                sheet = pd.concat(parts, ignore_index=True)
                sheet_name = pname[:31]
                sheet.to_excel(w, index=False, sheet_name=sheet_name)
        total = sum(len(sheet) for parts_list in by_platform.values()
                    for sheet in parts_list)
        return {"ok": True, "path": path, "count": int(total)}

    # ================= 导出（既有功能） =================
    def open_export_dir(self):
        """在系统文件管理器中打开导出目录"""
        try:
            if sys.platform.startswith("win"):
                os.startfile(export_dir())
            elif sys.platform == "darwin":
                os.system(f'open "{export_dir()}"')
            else:
                os.system(f'xdg-open "{export_dir()}"')
            return True
        except Exception:
            return False

    def export_search(self, keywords, mode, start, end):
        df = self._require_df()
        hit = search.keyword_search(df, keywords, mode, start, end)
        if not len(hit):
            return {"ok": False, "msg": "没有可导出的数据"}
        tag = "-".join(keywords)
        path = os.path.join(export_dir(), f"检索_{tag}_{start}_{end}_{_stamp()}.xlsx")
        out = hit.copy()
        out["时间"] = out["时间"].dt.strftime("%Y-%m-%d %H:%M:%S")
        search.export_df(out, path)
        return {"ok": True, "path": path, "count": len(out)}

    def export_stats(self, start, end):
        df = self._require_df()
        path = os.path.join(export_dir(), f"工单统计_{start}_{end}_{_stamp()}.xlsx")
        ov = stats.overview(df, start, end)
        fj = stats.by_fenju(df, start, end)
        tr = stats.monthly_trend(df, start, end)

        ov_rows = [{"平台": "合计", "本期": ov["total"]["cur"],
                    "去年同期": ov["total"]["prev"], "同比%": ov["total"]["yoy"]}]
        for k, v in ov["platforms"].items():
            ov_rows.append({"平台": k, "本期": v["cur"],
                            "去年同期": v["prev"], "同比%": v["yoy"]})
        ov_df = pd.DataFrame(ov_rows)
        fj_df = pd.DataFrame(
            [{"分局": r["分局"], "本期": r["cur"], "去年同期": r["prev"], "同比%": r["yoy"]}
             for r in fj])
        months = sorted(set(tr["cur_months"]) | set(tr["prev_months"]))
        cm = dict(zip(tr["cur_months"], tr["cur_counts"]))
        pm = dict(zip(tr["prev_months"], tr["prev_counts"]))
        tr_df = pd.DataFrame(
            [{"月份": m, "本期": cm.get(m, 0), "去年同期": pm.get(m, 0)} for m in months])

        search.export_df(ov_df, path, extra_sheets={"分分局": fj_df, "分月": tr_df})
        return {"ok": True, "path": path}