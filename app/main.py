# -*- coding: utf-8 -*-
"""
梁溪区工单分析工具 — 主程序入口 (pywebview)
运行: python main.py    打包: 见 打包说明.md
"""
import os, sys, traceback, datetime


def _base_dir():
    """程序所在基准目录（exe 旁 / 绿色包根目录）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _log(file_name):
    p = os.path.join(_base_dir(), file_name)
    try:
        with open(p, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] ")
    except Exception:
        pass
    return p


def slog(msg):
    """启动阶段日志：run.log 记录每一步，出问题看它就知道死在哪"""
    try:
        with open(os.path.join(_base_dir(), "run.log"), "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass


def _install_crash_logger():
    """任何未捕获异常 → 写 error.log（闪退也能定位原因）"""
    def _hook(t, v, tb):
        try:
            with open(os.path.join(_base_dir(), "error.log"), "a", encoding="utf-8") as f:
                f.write(f"\n[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] 未捕获异常\n")
                traceback.print_exception(t, v, tb, file=f)
        except Exception:
            pass
        traceback.print_exception(t, v, tb)
    sys.excepthook = _hook
    try:
        import threading
        threading.excepthook = lambda a: _hook(a.exc_type, a.exc_value, a.exc_traceback)
    except Exception:
        pass


_install_crash_logger()

# 嵌入式/绿色版：_pth 模式下脚本目录不自动入 path，显式加入，
# 让同目录的 api.py / core/ / paths.py 可被导入（对 exe 无害）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

slog(f"=== 启动: python {sys.version}, frozen={getattr(sys, 'frozen', False)}, "
     f"argv={sys.argv[:2]}, cwd={os.getcwd()}")

try:
    import webview
    slog("导入 webview 成功")
except Exception:
    slog(f"导入 webview 失败:\n{traceback.format_exc()}")
    raise

try:
    from api import Api
    from paths import resource_path
    slog("导入 api/paths 成功")
except Exception:
    slog(f"导入 api/paths 失败:\n{traceback.format_exc()}")
    raise


def pick_gui():
    """: Windows 上按可用性选择后端: edgechromium(WebView2), 否则 mshtml
    优先用环境变量 PYWEBVIEW_GUI 强制指定（便于调试）。"""
    if os.environ.get("PYWEBVIEW_GUI"):
        return os.environ["PYWEBVIEW_GUI"]
    if not sys.platform.startswith("win"):
        return None                    # macOS/Linux 用默认
    # 只有装了 pythonnet 才用 edgechromium，否则用 mshtml（免额外依赖）
    try:
        import pythonnet  # noqa: F401
        return "edgechromium"
    except Exception:
        return "mshtml"


def main():
    api = Api()
    html = resource_path("ui/index.html")
    if not os.path.exists(html):
        slog(f"界面文件缺失: {html}")
        raise RuntimeError(f"界面文件缺失: {html}")
    slog(f"界面文件就绪: {html}")
    gui = pick_gui()
    slog(f"GUI 后端: {gui}")
    win = webview.create_window(
        "梁溪区工单分析工具", url=html, js_api=api,
        width=1280, height=860, min_size=(1024, 700))
    api._window = win
    slog("窗口创建成功，进入消息循环")
    webview.start(gui=gui)
    slog("=== 正常退出")


if __name__ == "__main__":
    main()
