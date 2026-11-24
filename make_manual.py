# -*- coding: utf-8 -*-
"""生成《梁溪区工单分析工具 使用说明》Word 文档"""
import docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

PRIMARY = RGBColor(0x24, 0x57, 0xF5)
DARK = RGBColor(0x0F, 0x17, 0x2A)
MUTED = RGBColor(0x55, 0x65, 0x7D)

doc = Document()

# 全局字体
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

def set_cn(run, font='微软雅黑'):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font)

def h1(text):
    p = doc.add_heading(level=1)
    r = p.add_run(text); set_cn(r)
    r.font.size = Pt(16); r.font.color.rgb = DARK; r.font.bold = True
    return p

def h2(text):
    p = doc.add_heading(level=2)
    r = p.add_run(text); set_cn(r)
    r.font.size = Pt(13); r.font.color.rgb = PRIMARY; r.font.bold = True
    return p

def para(text, bold=False, size=10.5, color=None, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text); set_cn(r)
    r.font.size = Pt(size); r.font.bold = bold
    if color: r.font.color.rgb = color
    return p

def bullet(text, bold_head=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    if bold_head:
        r = p.add_run(bold_head); set_cn(r); r.font.bold = True; r.font.size = Pt(10.5)
    r = p.add_run(text); set_cn(r); r.font.size = Pt(10.5)
    return p

def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(text); set_cn(r)
    r.font.size = Pt(9.5); r.font.color.rgb = MUTED
    return p

# ============ 封面 ============
for _ in range(6): doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('梁溪区工单分析工具'); set_cn(r)
r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = DARK
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('使 用 说 明'); set_cn(r)
r.font.size = Pt(20); r.font.color.rgb = PRIMARY
doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('内网单机版 v2.1　·　数据不出内网　·　完全离线运行'); set_cn(r)
r.font.size = Pt(11); r.font.color.rgb = MUTED
doc.add_page_break()

# ============ 一、工具简介 ============
h1('一、工具简介')
para('本工具用于梁溪区市场监督管理局三个平台（12315、市12345、省12345）工单数据的统计分析，'
     '在本地电脑上运行，不联网、不调用任何外部服务，所有分析均在导入的 Excel 数据上完成，'
     '满足内网数据安全要求。')
para('主要功能：', bold=True)
bullet('导入三个平台导出的 Excel 工单表（支持一次多个文件或整个文件夹），自动识别平台、合并、按单号去重；')
bullet('同期对比统计：任意时间段（年 / 半年 / 季度 / 月 / 自定义区间）的工单总量、分平台量、同比增减；')
bullet('分分局统计：各分局（含消保科）工单量与同比；')
bullet('综合筛选：平台、分局、时间范围、关键词（支持「或 / 且」组合）任意条件组合筛选；')
bullet('挽回经济损失统计：筛选结果自动汇总 12315 平台的挽回损失金额；')
bullet('导出：统计报表、筛选明细（保持平台原始 Excel 格式）。')

# ============ 二、安装与启动 ============
h1('二、安装与启动')
h2('2.1 运行环境')
bullet('操作系统：Windows 10 / 11（64 位）；')
bullet('无需安装 Python 或其他任何软件，解压即用；')
bullet('需要系统自带组件：Microsoft Edge WebView2 与 .NET Framework 4.7.2+（Win10/11 一般已自带）。')
h2('2.2 启动步骤')
bullet('第 1 步：', '解压「梁溪区工单分析工具」文件夹到任意位置（如桌面）；')
bullet('第 2 步：', '双击文件夹内的「启动.bat」；')
bullet('第 3 步：', '等待 3~5 秒，程序窗口自动打开。')
note('提示：首次启动略慢属正常现象。若杀毒软件弹出拦截提示，请选择「信任 / 允许运行」'
     '（PyInstaller / 绿色 Python 程序的常见误报，非病毒）。')

# ============ 三、界面总览 ============
h1('三、界面总览')
try:
    doc.add_picture('docs/shots/overview.png', width=Cm(16))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    note('▲ 主界面（示意数据）')
except Exception as e:
    print('截图嵌入失败:', e)
para('界面分为三个区域：', bold=True)
bullet('左侧导航栏：', '「数据总览 / 综合筛选 / 数据导入 / 导出结果」四个入口，点击平滑跳转到对应区域；')
bullet('主操作区：', '页面右上角是数据导入按钮；下方依次为统计卡片、综合筛选、图表、分分局统计表、筛选结果；')
bullet('时间范围选择器：', '位于「综合筛选」卡片内，统计与筛选共用同一时间范围。')

# ============ 四、导入数据 ============
h1('四、导入数据')
h2('4.1 两种导入方式')
bullet('「导入整个文件夹」（推荐）：', '选择存放 Excel 的文件夹，自动导入其中全部 Excel 文件；')
bullet('「选择文件（可多选）」：', '在对话框中按住 Ctrl 逐个点选，一次选多个文件。')
h2('4.2 支持的数据')
bullet('平台：12315、市12345、省12345（程序按表头自动识别，无需手工指定）；')
bullet('格式：.xlsx 与老版 .xls 均支持；')
bullet('数量：一次可导入任意多个文件——系统单次导出上限 1 万条时，可分多次导出后一并导入，'
       '程序自动合并并按「平台 + 单号」去重，重复导出不会重复计数。')
h2('4.3 导入后')
para('导入完成后，顶部显示已加载的总条数、文件数与时间跨度；「数据导入」卡片列出每个文件的识别结果'
     '（平台、条数）；无法识别的文件以红色标注原因。')

# ============ 五、同期对比统计 ============
h1('五、同期对比统计')
para('在「综合筛选」卡片中选择时间范围（快捷按钮：全部 / 最近一年 / 前一年 / 上半年 / 下半年，'
     '或自定义起止日期），统计看板自动刷新：')
bullet('KPI 卡片：', '全部平台合计与三个平台各自的「本期工单量、去年同期、同比增减%」（红色▲上涨、绿色▼下降）；')
bullet('月度工单量图：', '本期逐月柱状图与去年同期折线对比；')
bullet('分局图表与表格：', '各分局（含消保科）本期 / 去年同期工单量对比与同比。')
note('说明：分局空白或无法归口的工单统一计入「消保科」。')
h2('导出统计报表')
para('点击「分分局统计」卡片右侧的「导出统计报表」按钮，生成 Excel 文件'
     '（含三个工作表：明细 / 分分局 / 分月），自动保存到「导出结果」文件夹。')

# ============ 六、综合筛选 ============
h1('六、综合筛选')
para('四个条件可任意组合（同时满足），点「筛选」按钮执行：')
bullet('平台：', '勾选要统计的平台（不勾 = 全部）；')
bullet('分局：', '勾选分局（不勾 = 全部，已选数量实时显示）；')
bullet('时间范围：', '与统计共用；')
bullet('关键词：', '在工单内容中搜索，多个词用空格分隔，可选「任一（或）」或「全部（且）」，'
       '例如输入「食品 药品」选「且」表示内容同时含两词。')
h2('筛选结果')
para('筛选结果卡片显示：命中工单条数、挽回经济损失合计（大字红色）、有金额工单条数'
     '（仅 12315 平台导出表含此字段，市 / 省12345 无此列不计入），以及前 200 条明细预览'
     '（含每条工单的挽回损失金额）。')

# ============ 七、导出 ============
h1('七、导出')
h2('7.1 导出位置')
para('所有导出文件自动保存到程序文件夹旁的「导出结果」文件夹中，文件名自动带上时间戳，'
     '不会互相覆盖。点击「打开导出文件夹」按钮可直接打开该文件夹。')
note('程序不弹出「另存为」对话框，导出后界面会提示完整保存路径。')
h2('7.2 筛选结果导出（保持原格式）')
para('点击「导出筛选结果（原格式）」按钮，导出的 Excel 按平台分工作表'
     '（12315 / 市12345 / 省12345 各一个），每个工作表保留该平台原始表的全部列与原始内容'
     '（与业务系统导出格式完全一致），只截取符合当前筛选条件的行，可直接用于二次加工或上报。')

# ============ 八、常见问题 ============
h1('八、常见问题（FAQ）')
faq = [
    ('双击「启动.bat」没反应？',
     '等 3~5 秒（首次加载较慢）；仍无反应请检查杀毒软件是否拦截，添加信任后重试；'
     '确认解压后文件夹完整（包含 python、site-packages、app 三个文件夹）。'),
    ('导入后提示「表头无法识别」？',
     '该文件不是三个平台的标准导出表。请从业务系统重新导出，不要改动表头行。'),
    ('统计里没有某分局？',
     '分局列表来自导入的数据；确认所选时间范围内该分局有工单。空白分局计入「消保科」。'),
    ('挽回损失金额为什么是 0？',
     '只有 12315 平台导出表含「挽回经济损失(单位:元)」列；市 / 省12345 的导出表没有该字段，'
     '因此只统计 12315。'),
    ('导出的文件在哪里？',
     '程序文件夹旁的「导出结果」文件夹；也可点界面上的「打开导出文件夹」。'),
    ('数据会传出内网吗？',
     '不会。程序完全离线运行，不联网、不调用任何云端服务，所有计算在本机完成。'),
    ('想升级为单文件 exe？',
     '在任意一台装了 Python 的 Windows 电脑上，进入 app 文件夹双击 build_exe.bat，'
     '生成 dist\\梁溪区工单分析工具.exe 单文件程序。'),
]
for q, a in faq:
    para('Q：' + q, bold=True, space_after=2)
    para('A：' + a, space_after=8)

# ============ 附 ============
h1('附：版本信息')
bullet('版本：v2.1（2025-11）')
bullet('适配平台：12315（48 列）、市12345（39 列）、省12345（54 列）标准导出表')
bullet('运行形态：Windows 绿色免安装 / 可选 PyInstaller 单文件 exe')

doc.save('docs/梁溪区工单分析工具-使用说明.docx')
print('已生成 docs/梁溪区工单分析工具-使用说明.docx')
