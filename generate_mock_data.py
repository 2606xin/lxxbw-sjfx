# -*- coding: utf-8 -*-
"""
模拟数据生成器：按 3 个样表结构，批量生成测试用 Excel
- 12315.xlsx      (48列, 字符串日期 2026-07-03 09:34:45)
- 市12345.xls     (39列, 真xls二进制, 日期 2023/10/24 12:45:44)
- 省12345.xlsx    (54列, 来电时间为字符串, 其余时间列为 Excel 序列数)
年份: 2024 全年 / 2025 全年(部分拆多文件) / 2026 1-7月
"""
import os, random, datetime as dt

random.seed(42)
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "模拟数据")
os.makedirs(OUT, exist_ok=True)

# ---------------- 基础素材 ----------------
FENJU = ["崇安寺分局", "通江分局", "广益分局", "扬名一分局", "扬名二分局",
         "山北分局", "惠山一分局", "惠山二分局", "清名桥分局", "金匮分局"]
FENJU_W = [12, 10, 11, 9, 9, 10, 8, 8, 12, 7]   # 权重
BLANK_FJ = 0.05                                   # 空白分局比例→消保科

MERCHANTS = ["鑫和", "瑞丰", "康之源", "悦美", "星动力", "博学", "蜀香", "宜佳",
             "驰腾", "美居家", "光速", "衣品", "天天鲜", "好邻居", "百味", "惠民"]
SHOPS = ["超市", "餐饮店", "药房", "摄影工作室", "健身中心", "培训机构",
         "美容院", "汽车4S店", "装修公司", "数码店", "服装店", "便利店"]
ROADS = ["中山路", "清扬路", "人民中路", "解放东路", "县前街", "广瑞路",
         "芦庄路", "运河东路", "凤宾路", "学前街", "永乐路", "金城路"]
FOODS = ["牛奶", "火腿肠", "面包", "食用油", "速冻水饺", "饮料", "糕点", "熟食"]
SURNAMES = list("张王李赵刘陈杨黄周吴徐孙马朱胡郭何")

TOPICS = [
    # (话题, 适用店铺类型, 模板列表, 权重)  —— 店铺类型与话题相关联，更真实
    ("食品过期",   ["超市", "便利店"],
        ["在{m}购买的{f}已超过保质期，要求退款赔偿", "在{m}买到过期{f}，商家拒绝处理",
         "在{m}购买的食品已超过保质期，生产日期疑似篡改"], 12),
    ("食品异物",   ["超市", "便利店", "餐饮店"],
        ["在{m}购买的{f}中吃出异物，要求赔偿", "在{m}购买的食品中吃出异物，要求赔偿"], 8),
    ("食品变质",   ["超市", "便利店"],
        ["在{m}购买的{f}已变质发霉，食用后身体不适", "在{m}购买的食品已变质发霉"], 6),
    ("药品问题",   ["药房"],
        ["在{m}购买的药品价格虚高，要求查处", "在{m}买到无批准文号的药品，怀疑是假药",
         "{m}违规销售处方药，要求处理"], 8),
    ("保健食品",   ["药房", "超市"],
        ["在{m}购买保健食品，商家夸大宣传功效，且搭售药品，要求退款"], 4),  # 同时含"食品""药品"
    ("摄影写真",   ["摄影工作室"],
        ["在{m}拍摄写真，商家额外收取底片费用，要求退还", "{m}摄影套餐虚假宣传，底片不全部交付"], 5),
    ("化妆品",     ["美容院", "超市", "便利店"],
        ["在{m}购买的化妆品使用后过敏，要求赔偿", "{m}销售无中文标识的进口化妆品"], 5),
    ("电梯安全",   ["超市"],
        ["{a}小区电梯频繁故障，存在安全隐患，要求督促维保"], 5),
    ("培训退费",   ["培训机构"],
        ["{m}培训机构关门，剩余课时费未退，要求协调退费"], 6),
    ("健身卡",     ["健身中心"],
        ["在{m}办理的健身卡，商家停业，要求退款"], 5),
    ("餐饮卫生",   ["餐饮店"],
        ["{m}餐饮店卫生条件差，后厨发现蟑螂", "在{m}就餐后腹泻，怀疑食物不新鲜"], 6),
    ("价格问题",   ["超市", "餐饮店", "便利店"],
        ["{m}不明码标价，结账时加价，要求查处"], 6),
    ("虚假宣传",   ["超市", "美容院", "培训机构", "健身中心"],
        ["{m}虚假宣传，实际服务与宣传不符，要求处理"], 6),
    ("汽车消费",   ["汽车4S店"],
        ["在{m}购买汽车，4S店收取金融服务费不合理"], 3),
    ("家装建材",   ["装修公司"],
        ["{m}装修公司拖延工期，材料以次充好"], 4),
    ("数码产品",   ["数码店"],
        ["在{m}购买的手机为翻新机，要求退一赔三"], 4),
    ("服装鞋帽",   ["服装店"],
        ["在{m}购买的衣服存在质量问题，商家拒绝退换"], 3),
]
TOPIC_W = [t[3] for t in TOPICS]

def merchant(shops=None):
    return random.choice(MERCHANTS) + random.choice(shops if shops else SHOPS)

def address():
    return "梁溪区" + random.choice(ROADS) + str(random.randint(1, 500)) + "号"

def content():
    name, shops, tpls, _ = random.choices(TOPICS, weights=TOPIC_W)[0]
    tpl = random.choice(tpls)
    m = merchant(shops)
    return name, m, tpl.format(m=m, f=random.choice(FOODS), a=address())

def person():
    return random.choice(SURNAMES) + random.choice(["某", "先生", "女士"])

def phone():
    return random.choice(["13", "15", "18"]) + "".join(random.choices("0123456789", k=9))

MONTH_W = [0.9, 0.7, 1.5, 1.0, 1.0, 0.95, 1.0, 1.0, 0.95, 1.0, 1.3, 1.1]  # 3月315 / 11月双11

def gen_times(year, n, months):
    ws = [MONTH_W[m - 1] for m in months]
    picks = random.choices(months, weights=ws, k=n)
    out = []
    for m in picks:
        if m == 12:
            dmax = 31
        else:
            dmax = (dt.date(year, m + 1, 1) - dt.timedelta(days=1)).day
        out.append(dt.datetime(year, m, random.randint(1, dmax),
                               random.randint(8, 20), random.randint(0, 59), random.randint(0, 59)))
    out.sort()
    return out

def pick_fenju():
    if random.random() < BLANK_FJ:
        return ""
    return random.choices(FENJU, weights=FENJU_W)[0]

def full_fj(short):
    return "无锡市梁溪区市场监督管理局" + short if short else ""

def excel_serial(t):
    return (t - dt.datetime(1899, 12, 30)).total_seconds() / 86400.0

def done_time(t):
    """办结时间 = 登记后 3~25 天（可能超出当月，无妨）"""
    return t + dt.timedelta(days=random.randint(3, 25), hours=random.randint(0, 8))

# ---------------- 12315 (48列) ----------------
H_12315 = ["登记编号","类型","登记日期","接收方式","提供方姓名","提供方联系方式","提供方联系地址","性别",
"企业名称","统一社会信用代码","注册号","企业电话","企业地址","销售方式","客体类别","客体类别全称",
"商品/服务名称","订单号","争议发生时间","消费金额(单位:元)","诉求内容","问题类别","问题类别全称","咨询范围",
"事发地","具体问题","登记人","登记机构","登记机构全称","登记部门","处理人","处理机构","处理机构全称",
"处理部门","初查反馈期限","办结期限","初查反馈时间","办结时间","初查受理状态","办理情况状态","反馈内容",
"信息来源","挽回经济损失(单位:元)","案值(单位:元)","罚款金额(单位:元)","没收金额(单位:元)","调查处理结果","电商平台"]

def row_12315(t):
    typ = "投诉" if random.random() < 0.55 else "举报"
    tname, m, c = content()
    fj = pick_fenju()
    sq, wt = "", ""
    if typ == "投诉":
        sq = c + "。诉求目的：要求相关单位协调处理。"
        if random.random() < 0.2:
            wt = c
    else:
        wt = c
        if random.random() < 0.1:
            sq = c
    dt_done = done_time(t)
    row = [""] * 48
    row[0] = "2132021300" + "".join(random.choices("0123456789", k=15))
    row[1] = typ
    row[2] = t.strftime("%Y-%m-%d %H:%M:%S")
    row[3] = random.choice(["来电", "来函", "网络", "微信小程序", "APP"])
    row[4] = person(); row[5] = phone(); row[7] = random.choice(["男", "女"])
    row[8] = "梁溪区" + m; row[12] = address()
    row[13] = random.choice(["现场", "网络"])
    row[15] = tname; row[16] = "商品（产品）->" + tname
    row[19] = str(random.choice([0, 50, 200, 800, 3000]))
    row[20] = sq; row[25] = wt
    row[26] = person()[0] + "某"
    row[27] = "梁溪区市场监督管理局"
    row[28] = "江苏省市场监督管理局->无锡市市场监督管理局->梁溪区市场监督管理局"
    row[30] = person()[0] + "某"
    row[31] = fj
    row[32] = ("江苏省市场监督管理局->无锡市市场监督管理局->梁溪区市场监督管理局->" + fj) if fj else ""
    row[34] = (t + dt.timedelta(days=7)).strftime("%Y-%m-%d")
    row[36] = dt_done.strftime("%Y-%m-%d %H:%M:%S"); row[37] = row[36]
    row[38] = random.choice(["不立案", "立案"]) if typ == "举报" else ""
    row[39] = "已办结"
    row[40] = "经查，已短信告知投诉人处理结果。" if typ == "投诉" else "经查，举报事项不予立案，理由：已短信告知"
    row[41] = "全国12315互联网平台"
    # 挽回经济损失：约 30% 工单有金额（50~5000元）
    row[42] = str(random.choice([0, 0, 0, 0, 0, 0, 0,
               100, 200, 350, 500, 800, 1200, 2000, 3500, 5000]))
    row[47] = random.choice(["", "拼多多", "淘宝", "京东", "抖音"]) if row[13] == "网络" else ""
    return row

# ---------------- 市12345 (39列) ----------------
H_SHI = ["序号","工单编号","工单来源","一级分类","二级分类","主题工单","工单标题","工单内容","下单时间",
"主办单位","二级单位","诉求目的","诉求地址","业务类型","诉求类型","市民姓名","联系号码","工单类型",
"网单类型","过程满意","最终满意","三级分类","处理状态","截止时间","话务员","预审岗下派时间",
"二级单位申请办结时间","最终处理意见","最终处理时间","最终处理方式","成员单位意见","省平台编号",
"标签名称","三级单位","来电号码","预审岗","区域","四级分类","五级分类"]

def row_shi(t, idx):
    tname, m, c = content()
    fj = pick_fenju()
    dt_done = done_time(t)
    opinion = ("梁溪区市场监督管理局回复：无锡市梁溪区市场监督管理局" + fj +
               "回复：经核实，已现场检查并责令整改，已与诉求人沟通反馈。") if fj else \
              "梁溪区市场监督管理局回复：经核实，已现场检查并责令整改，已与诉求人沟通反馈。"
    row = [""] * 39
    row[0] = str(idx)
    row[1] = "WX" + str(t.year) + "".join(random.choices("0123456789", k=9))
    row[2] = random.choice(["来电", "微信", "网站", "APP"])
    row[3] = "经济综合"; row[4] = random.choice(["消费维权", "市场监管", "物价"])
    row[6] = "关于" + tname + "的相关问题"
    row[7] = "服务对象来电反映" + c + "。（服务对象要求联系，请做好沟通。）"
    row[8] = t.strftime("%Y/%m/%d %H:%M:%S")
    row[9] = full_fj(fj)
    row[10] = "梁溪区"
    row[11] = "服务对象希望相关单位核实处理。（服务对象要求联系，请做好沟通。）"
    row[12] = address()
    row[13] = random.choice(["投诉类", "举报类", "咨询类"])
    row[15] = person(); row[16] = phone()
    row[17] = "新问题"
    row[19] = random.choice(["满意", "基本满意", "不满意"])
    row[20] = random.choice(["满意", "基本满意", "不满意"])
    row[21] = "其他消费维权"
    row[22] = "已办结"
    row[23] = dt_done.strftime("%Y/%m/%d %H:%M:%S")
    row[27] = opinion
    row[28] = dt_done.strftime("%Y/%m/%d %H:%M:%S")
    row[29] = "办结"
    row[30] = opinion
    row[34] = phone()
    row[36] = "梁溪区"
    return row

# ---------------- 省12345 (54列) ----------------
H_SHENG = ["工单编号","来电号码","联系号码","备用联系电话","服务对象类型","来电时间","工单状态","诉求地址",
"一级归口","二级归口","三级归口","一级归口（新）","二级归口（新）","三级归口（新）","四级归口（新）","五级归口（新）",
"诉求内容","诉求目的","紧急类型","信息保护","是否回访","信息来源","诉求类型","是否热点","诉求标签","签收状态",
"是否超期","是否引用信息","创建工单时间","是否督办","是否满意","上级交办时间","平台退回时间","主协办标签",
"本级计划签收时间","本级签收时间","本级计划完成时间","本级答复时间","本级退回时间","处理意见","最新处理方式",
"交办下级主办部门","交办下级协办部门","最新交办下级部门","最新交办下级时间","下级签收时间","下级计划签收时间",
"下级退回时间","下级退回部门","下级答复时间","下级答复部门","下级计划完成时间","下级处理意见","下级部门处理是否包含附件"]

def row_sheng(t):
    tname, m, c = content()
    fj = pick_fenju()
    dt_done = done_time(t)
    opinion = "您好，已于" + dt_done.strftime("%Y-%m-%d %H:%M:%S") + "通过电话方式联系您；经与商家核实，已协调处理，其表示认可，故申请办结。"
    row = [""] * 54
    row[0] = "DH0200" + "".join(random.choices("0123456789", k=13))
    row[1] = phone(); row[2] = phone()
    row[4] = random.choice(["个人", "个人", "企业"])
    row[5] = t.strftime("%Y-%m-%d %H:%M:%S")          # 字符串
    row[6] = random.choice(["工单办结-回访办结", "工单办结", "办理中"])
    row[7] = "无锡市梁溪区"
    row[8] = "经济综合"; row[9] = "消费维权"; row[10] = tname
    row[16] = "服务对象来电反映" + c + "。"
    row[17] = "服务对象来电希望相关单位核实处理。（服务对象要求联系，请做好沟通。）"
    row[18] = random.choice(["一般", "一般", "紧急"])
    row[20] = "是"; row[21] = random.choice(["电话", "微信", "网站"])
    row[22] = "投诉举报"; row[23] = "否"; row[25] = "已签收"
    row[26] = random.choice(["未超期", "未超期", "已超期"])
    row[28] = round(excel_serial(t), 10)               # Excel 序列数
    row[29] = "未督办"
    row[30] = random.choice(["满意", "基本满意", "不满意", ""])
    row[31] = round(excel_serial(t + dt.timedelta(hours=2)), 10)
    row[34] = round(excel_serial(t + dt.timedelta(days=1)), 10)
    row[35] = round(excel_serial(t + dt.timedelta(hours=5)), 10)
    row[36] = round(excel_serial(dt_done), 10)
    row[37] = round(excel_serial(dt_done - dt.timedelta(hours=3)), 10)
    row[39] = opinion
    row[40] = "已办结"
    if fj:
        fj2 = full_fj(fj)
        other = full_fj(random.choice([x for x in FENJU if x != fj]))
        row[41] = fj2 + ("," + other if random.random() < 0.15 else "")   # 主办可能多个
        row[43] = fj2                                                       # 最新交办下级部门=唯一权威
        row[49] = fj2
    row[44] = round(excel_serial(t + dt.timedelta(hours=6)), 10)
    row[45] = round(excel_serial(t + dt.timedelta(hours=10)), 10)
    row[49] = row[49] or ""
    row[50] = round(excel_serial(dt_done), 10)
    row[52] = opinion
    row[53] = "否"
    return row

# ---------------- 写出 ----------------
def write_xlsx(path, header, rows):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)

def write_xls(path, header, rows):
    import xlwt
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("sheet")
    for j, h in enumerate(header):
        ws.write(0, j, h)
    for i, r in enumerate(rows, 1):
        for j, v in enumerate(r):
            ws.write(i, j, v)
    wb.save(path)

def split_and_write(platform, year_tag, header, rows, writer, ext, chunk=4000):
    parts = [rows[i:i + chunk] for i in range(0, len(rows), chunk)]
    files = []
    for k, part in enumerate(parts, 1):
        suffix = f"_{k}" if len(parts) > 1 else ""
        fn = f"{platform}_{year_tag}{suffix}.{ext}"
        p = os.path.join(OUT, fn)
        writer(p, header, part)
        files.append((fn, len(part)))
    return files

PLAN = [
    ("12315",  H_12315, row_12315, write_xlsx, "xlsx", {2024: 3500, 2025: 3900, 2026: 2400}),
    ("市12345", H_SHI,   row_shi,   write_xls,  "xls",  {2024: 4200, 2025: 4700, 2026: 2700}),
    ("省12345", H_SHENG, row_sheng, write_xlsx, "xlsx", {2024: 2600, 2025: 2900, 2026: 1700}),
]

def main():
    manifest = []
    for platform, header, rowfn, writer, ext, years in PLAN:
        for year, n in years.items():
            months = list(range(1, 13)) if year < 2026 else list(range(1, 8))
            times = gen_times(year, n, months)
            rows = []
            for i, t in enumerate(times, 1):
                if platform == "市12345":
                    rows.append(rowfn(t, i))
                else:
                    rows.append(rowfn(t))
            tag = f"{year}年" if year < 2026 else "2026年1-7月"
            for fn, cnt in split_and_write(platform, tag, header, rows, writer, ext):
                manifest.append((fn, cnt))
                print(f"  ✓ {fn}  ({cnt} 行)")
    total = sum(c for _, c in manifest)
    print(f"\n共生成 {len(manifest)} 个文件 / {total} 条数据 → {OUT}")

if __name__ == "__main__":
    main()
