# 临时冒烟测试：预载模拟数据并弹出窗口（开发用）
import sys, os, glob, webview
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from api import Api
from core import loader
from paths import resource_path

api = Api()
dfs = sorted(glob.glob(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '模拟数据', '*')))
api._load(dfs)
print('预载', len(api._df), '条 /', len(api._files), '个文件')
html = resource_path('ui/index.html')
win = webview.create_window('梁溪区工单分析工具(冒烟测试)', url=html, js_api=api, width=1280, height=860)
api._window = win
webview.start()
