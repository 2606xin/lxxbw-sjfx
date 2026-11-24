# -*- coding: utf-8 -*-
"""
loader.py — Excel 加载 + 平台识别 + 清洗 + 统一字段映射
统一输出列: 平台 / 单号 / 时间 / 分局 / 内容
同时保留每个文件的原始表 (raw)，用于"原格式导出"
"""
import os
import pandas as pd

# ---------- 平台定义 ----------
# signature: 表头中"同时出现"即判定为该平台（用区分度高的列名）
PLATFORMS = {
    "12315": {
        "signature": {"登记编号", "处理机构", "登记日期"},
        "time": "登记日期",
        "fenju": "处理机构",
        "content": ["诉求内容", "具体问题"],   # 两列都搜（已确认）
        "order": "登记编号",
        "loss": "挽回经济损失(单位:元)",       # 仅 12315 有此列
    },
    "市12345": {
        "signature": {"工单编号", "主办单位", "下单时间"},
        "time": "下单时间",
        "fenju": "主办单位",
        "content": ["工单内容"],
        "order": "工单编号",
        "loss": None,
    },
    "省12345": {
        "signature": {"工单编号", "最新交办下级部门", "来电时间"},
        "time": "来电时间",
        "fenju": "最新交办下级部门",
        "content": ["诉求内容"],
        "order": "工单编号",
        "loss": None,
    },
}

# 分局全名 → 短名：去掉这些前缀
FJ_PREFIXES = [
    "江苏省市场监督管理局->无锡市市场监督管理局->梁溪区市场监督管理局->",
    "无锡市梁溪区市场监督管理局",
    "梁溪区市场监督管理局",
    "无锡市梁溪区市场监管局",
    "梁溪区市场监管局",
]
DEFAULT_FENJU = "消保科"   # 分局空白归属

EXCEL_EPOCH = pd.Timestamp("1899-12-30")  # Excel 日期序列原点

EXCEL_EXTS = (".xls", ".xlsx")


def detect_platform(header):
    cols = set(str(c).strip() for c in header)
    for name, cfg in PLATFORMS.items():
        if cfg["signature"].issubset(cols):
            return name
    return None


def parse_time(series):
    """兼容三种时间形态: 字符串 / 日期对象 / Excel 序列数"""
    s = series
    # 已是 datetime
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    # 数值(Excel 序列数): >20000 视为序列数
    num = pd.to_numeric(s, errors="coerce")
    is_serial = num.notna() & (num > 20000) & (num < 80000)
    out = pd.Series(pd.NaT, index=s.index)
    if is_serial.any():
        out[is_serial] = EXCEL_EPOCH + pd.to_timedelta(num[is_serial], unit="D")
    if (~is_serial).any():
        out[~is_serial] = pd.to_datetime(s[~is_serial], format="mixed", errors="coerce")
    return out


def clean_fenju(value):
    """分局名归一: 去前缀、取路径末段、去空白 → 短名; 空白 → 消保科"""
    if value is None:
        return DEFAULT_FENJU
    v = str(value).strip()
    if not v or v.lower() in ("nan", "none"):
        return DEFAULT_FENJU
    if "->" in v:                       # 12315 全称路径取末段
        v = v.split("->")[-1].strip()
    for p in FJ_PREFIXES:               # 12345 全名去前缀
        if v.startswith(p):
            v = v[len(p):]
            break
    v = v.strip()
    return v if v else DEFAULT_FENJU


def read_excel_any(path):
    """按扩展名读取 (支持真 .xls 与 .xlsx), 全部按字符串读"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        return pd.read_excel(path, dtype=str, engine="xlrd")
    return pd.read_excel(path, dtype=str, engine="openpyxl")


def load_file(path):
    """加载单个文件 → (文件对象 dict, 报告 dict)
    文件对象: {name, platform, raw(原始表, 全列), meta(统一5列)}
    """
    report = {"file": os.path.basename(path), "platform": None, "rows": 0,
              "ok": False, "msg": ""}
    try:
        raw = read_excel_any(path)
    except Exception as e:
        report["msg"] = f"读取失败: {e}"
        return None, report

    platform = detect_platform(raw.columns)
    if not platform:
        report["msg"] = "表头无法识别（不是 12315/市12345/省12345 的导出表）"
        return None, report
    cfg = PLATFORMS[platform]

    time = parse_time(raw[cfg["time"]])
    fenju = raw[cfg["fenju"]].map(clean_fenju)
    content = raw[cfg["content"][0]].fillna("")
    for extra in cfg["content"][1:]:
        content = content + "\n" + raw[extra].fillna("")
    content = content.str.strip()

    # 挽回经济损失：仅 12315 有该列；其他平台按 0
    if cfg.get("loss") and cfg["loss"] in raw.columns:
        loss = pd.to_numeric(raw[cfg["loss"]], errors="coerce").fillna(0.0)
    else:
        loss = pd.Series(0.0, index=raw.index)

    meta = pd.DataFrame({
        "平台": platform,
        "单号": raw[cfg["order"]].fillna("").astype(str).str.strip(),
        "时间": time,
        "分局": fenju,
        "内容": content,
        "挽回损失": loss,
    })
    bad_time = int(meta["时间"].isna().sum())
    keep = meta["时间"].notna()
    meta = meta[keep].reset_index(drop=True)
    raw = raw[keep].reset_index(drop=True)   # 原始表同步过滤

    fobj = {"name": os.path.basename(path), "platform": platform,
            "raw": raw, "meta": meta}
    report.update(platform=platform, rows=len(meta), ok=True)
    if bad_time:
        report["msg"] = f"有 {bad_time} 行时间无法解析，已跳过"
    return fobj, report


def load_files(paths):
    """加载多个文件 → (文件对象列表, 合并meta DataFrame, [报告...])"""
    files, reports = [], []
    seen = set()   # (平台, 单号) 去重
    for p in paths:
        fobj, rep = load_file(p)
        reports.append(rep)
        if fobj is None:
            continue
        # 单号去重（同一工单重复导出时只保留首个文件中的）
        mask = ~fobj["meta"].apply(
            lambda r: (r["平台"], r["单号"]) in seen, axis=1)
        fobj["meta"] = fobj["meta"][mask].reset_index(drop=True)
        fobj["raw"] = fobj["raw"][mask].reset_index(drop=True)
        seen.update(zip(fobj["meta"]["平台"], fobj["meta"]["单号"]))
        if len(fobj["meta"]):
            files.append(fobj)

    if files:
        merged = pd.concat([f["meta"] for f in files], ignore_index=True)
        merged = merged.sort_values("时间").reset_index(drop=True)
    else:
        merged = pd.DataFrame(columns=["平台", "单号", "时间", "分局", "内容", "挽回损失"])
    return files, merged, reports


def list_excel_files(folder):
    """列出文件夹下所有 Excel 文件（按文件名排序）"""
    out = []
    for name in sorted(os.listdir(folder)):
        if os.path.splitext(name)[1].lower() in EXCEL_EXTS:
            out.append(os.path.join(folder, name))
    return out
