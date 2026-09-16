# -*- coding: utf-8 -*-
"""图片 → Visio 矢量重绘：布局数据模板

用法：
  1. 复制本文件，重命名为 gen_<图名>_data.py
  2. 按『看图拆解方法论』改造 PAGES 部分（改文字、数量、配色、尺寸）
  3. 运行：python gen_<图名>_data.py      → 输出 _visio_data.json
  4. 绘制：powershell -ExecutionPolicy Bypass -File build_visio.ps1 \
               -JsonPath _visio_data.json -OutPath 输出图.vsdx
  5. 自检：powershell -ExecutionPolicy Bypass -File export_png.ps1 \
               -VsdxPath 输出图.vsdx -OutDir 预览      （然后读 PNG 与原图对比）

坐标系：页面左上角为原点，x 向右、top 向下，单位英寸。
（Visio 内部的 Y 轴翻转由 build_visio.ps1 自动处理，本文件不用管。）
"""
import json, os, sys
sys.stdout.reconfigure(encoding='utf-8')

# ============================ 页面与配色 ============================
PW, PH = 23.4, 16.5          # A2 横向（英寸）。内容多可改 33.1×23.4（A1）

NAVY   = (31, 78, 121)       # 层名 / 主标题 / 页眉
BLUE   = (46, 117, 182)      # 卡片标题条
MBLUE  = (91, 155, 213)      # 次级条
LBLUE  = (155, 195, 230)     # 边框
XLBLUE = (222, 235, 247)     # 容器背景
XXL    = (240, 247, 253)     # 大容器背景
WHITE  = (255, 255, 255)
DARK   = (26, 26, 26)
GRAY   = (89, 89, 89)
RED    = (192, 80, 77)       # 安全 / 警示
GREEN  = (84, 130, 53)       # 运营 / 监管
ORANGE = (237, 125, 49)      # 交通管理
PURPLE = (112, 48, 160)      # 监管方
YELLOW = (237, 194, 64)      # 制造方
PALEGR = (226, 239, 218)     # 浅绿底
PALEOR = (252, 228, 214)     # 浅橙底

PAGES = []


# ============================ 通用 helper ============================
def newpage(name):
    """新建一页"""
    PAGES.append({"name": name, "shapes": [], "conns": []})
    return PAGES[-1]


def add(pg, x, top, w, h, text="", fill=WHITE, line=LBLUE, fs=7.5, bold=False,
        fc=DARK, align=0, valign=1, lw=0.75, rnd=0.05, dash=1, oval=False):
    """基础矩形/椭圆。align: 0左 1中 2右；valign: 0上 1中 2下；
    dash: 1实线 0无边框 ≥2透传Visio线型(2=虚线)；oval=True 画椭圆。"""
    pg["shapes"].append(dict(
        x=round(x, 3), y=round(top, 3), w=round(w, 3), h=round(h, 3),
        text=text, fill=list(fill), line=list(line), fs=fs, bold=bold,
        fc=list(fc), align=align, valign=valign, lw=lw, rnd=rnd, dash=dash,
        oval=bool(oval)))


def card(pg, x, top, w, title, items, head=BLUE, body=XXL, edge=LBLUE,
         fs=7.5, head_fs=9.0, line_h=0.25, pad=0.12, hh=0.42):
    """标题条 + 内容框（内容按行换行拼接）。返回总高度，供 Y 累加。"""
    body_h = line_h * len(items) + pad * 2
    h = hh + body_h
    add(pg, x, top, w, hh, title, fill=head, line=head, fs=head_fs, bold=True,
        fc=WHITE, align=1, valign=1, lw=0.75, rnd=0.03)
    add(pg, x, top + hh, w, body_h, "\n".join(items), fill=body, line=edge,
        fs=fs, bold=False, fc=DARK, align=0, valign=0, lw=0.75, rnd=0.03)
    return h


def section(pg, x, top, w, h, title, body_text, head_color=BLUE, body_color=XXL):
    """带标题条的分区块（用于"端"层那种多段堆叠）。返回高度。"""
    add(pg, x, top, w, h, "", fill=body_color, line=head_color, lw=0.8, rnd=0.04)
    add(pg, x + 0.08, top + 0.05, w - 0.16, 0.32, title,
        fill=head_color, line=head_color, fs=8.6, bold=True, fc=WHITE,
        align=1, valign=1, lw=0.4, rnd=0.02)
    add(pg, x + 0.08, top + 0.40, w - 0.16, h - 0.45, body_text,
        fill=WHITE, line=head_color, fs=7.5, bold=False, fc=DARK,
        align=1, valign=1, lw=0.5, rnd=0.02)
    return h


def sidebar(pg, x, top, w, h, title, items, head=NAVY, edge=NAVY,
            fs=9.0, head_fs=11.0, pad=0.1):
    """左右竖栏：标题 + 纵向堆叠的若干项。返回高度。"""
    add(pg, x, top, w, h, "", fill=XXL, line=edge, lw=1.0, rnd=0.05)
    add(pg, x + 0.1, top + 0.1, w - 0.2, 0.95, title, fill=head, line=head,
        fs=head_fs, bold=True, fc=WHITE, align=1, valign=1, lw=0.5, rnd=0.04)
    item_h = (h - 1.30) / max(len(items), 1)
    y = top + 1.15
    for t in items:
        add(pg, x + 0.1, y, w - 0.2, item_h - 0.15, t, fill=WHITE, line=edge,
            fs=fs, bold=True, fc=head, align=1, valign=1, lw=0.8, rnd=0.04)
        y += item_h
    return h


def cols(n, gap=0.2, x0=0.4, total=None):
    """返回 n 列等宽的 x 坐标列表和列宽"""
    if total is None:
        total = PW - 0.8
    w = (total - gap * (n - 1)) / n
    return [x0 + i * (w + gap) for i in range(n)], w


def footer(pg, text, fs=6.5):
    """页脚说明（注明来源、对原图的补全/修正）"""
    add(pg, 0.4, PH - 0.4, PW - 0.8, 0.3, text,
        fill=WHITE, line=WHITE, fs=fs, fc=GRAY, align=1, valign=1, lw=0, dash=0)


def conn(pg, x1, y1, x2, y2, color=GRAY, lw=1.0, dash=1, arrow="end", label="", fs=7.0):
    """连线（可选）。arrow: end / start / both / none。
    注意：dash 是 Visio LinePattern 编号，0=不显示线（箭头也会消失），普通实线用 1。"""
    pg["conns"].append(dict(x1=round(x1, 3), y1=round(y1, 3), x2=round(x2, 3),
                            y2=round(y2, 3), color=list(color), lw=lw,
                            dash=dash, arrow=arrow, label=label, fs=fs,
                            fc=list(GRAY)))


# ===================================================================
#  示例：云 / 网 / 端 三层架构（把这里替换成你拆解出来的内容）
# ===================================================================
p = newpage("P1  示例：云网端三层架构")

# ---------- 顶部：场景（等分横排） ----------
add(p, 0.4, 0.45, 2.4, 0.6, "场景", fill=NAVY, line=NAVY, fs=11, bold=True,
    fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)
sx0 = 2.95
sw = (PW - 0.8 - 2.55 - 0.2 * 3) / 4
for i, s in enumerate(["载人出行", "低空物流", "公共服务", "文化旅游"]):
    add(p, sx0 + i * (sw + 0.2), 0.45, sw, 0.6, s, fill=MBLUE, line=MBLUE,
        fs=12, bold=True, fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)

# ---------- 主体分区 ----------
body_top, body_bot = 1.20, 14.30
body_h = body_bot - body_top

LX, LW = 0.4, 2.5                 # 左竖栏
RX, RW = 20.5, 2.5                # 右竖栏
CX = LX + LW + 0.2                # 主体起点
CW = RX - CX - 0.2                # 主体宽度

ROW_H = [3.0, 2.4, body_h - 3.0 - 2.4 - 0.1 * 2]   # 云 / 网 / 端
LBLX, LBLW = CX, 0.7              # 层名标签
CONX = CX + LBLW + 0.1            # 内容区起点
CONW = CW - LBLW - 0.1

# 层名标签（云 / 网 / 端）
yy = body_top
for (nm, sub), hh in zip([("云", "应用系统"),
                          ("网", "数据与使能\n支撑"),
                          ("端", "机载终端与\n基础设施")], ROW_H):
    add(p, LBLX, yy, LBLW, hh, "", fill=XLBLUE, line=BLUE, lw=1.0, rnd=0.05)
    add(p, LBLX, yy, LBLW, 0.5, nm, fill=NAVY, line=NAVY, fs=22, bold=True,
        fc=WHITE, align=1, valign=1, lw=0.5, rnd=0.04)
    add(p, LBLX + 0.06, yy + 0.55, LBLW - 0.12, hh - 0.6, sub, fill=XLBLUE,
        line=XLBLUE, fs=10, bold=True, fc=NAVY, align=1, valign=1, lw=0, dash=0)
    yy += hh + 0.1

# ---------- 行1：云（3 张大卡） ----------
yy = body_top
cards = [
    ("运营管理系统", GREEN, PALEGR, ["飞行计划", "飞控跟踪", "语控控制", "信息发布"], 2),
    ("低空交通管理和服务系统", ORANGE, PALEOR,
     ["空域管理", "流量管理", "情报服务", "气象服务", "低空交通管制", "低空数据服务"], 3),
    ("低空监管系统", GREEN, PALEGR,
     ["身份认证", "信息处置", "机场认证", "违法处置", "事故调查", "设备处置"], 3),
]
cw_each = (CONW - 0.2 * 2) / 3
for i, (title, color, bcol, items, nrows) in enumerate(cards):
    x = CONX + i * (cw_each + 0.2)
    add(p, x, yy, cw_each, ROW_H[0], "", fill=bcol, line=color, lw=1.2, rnd=0.05)
    add(p, x + 0.12, yy + 0.10, cw_each - 0.24, 0.4, title, fill=color, line=color,
        fs=9.5, bold=True, fc=WHITE, align=1, valign=1, lw=0.5, rnd=0.03)
    # 内部小格：按 nrows 行 × 2 列排列
    item_h = (ROW_H[0] - 0.6) / nrows
    half = (cw_each - 0.30) / 2
    for k, t in enumerate(items):
        r, c = divmod(k, 2)
        add(p, x + 0.12 + c * (half + 0.06), yy + 0.6 + r * item_h + 0.05,
            half, item_h - 0.10, t, fill=WHITE, line=color, fs=8.5, bold=True,
            fc=DARK, align=1, valign=1, lw=0.5, rnd=0.03)

# ---------- 行2：网（满宽横条 + 接入格） ----------
yy = body_top + ROW_H[0] + 0.1
add(p, CONX, yy, CONW, 0.7, "通信、导航、监视及信息服务", fill=MBLUE, line=MBLUE,
    fs=10, bold=True, fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)
add(p, CONX, yy + 0.8, CONW, 0.7, "数据交换网", fill=MBLUE, line=MBLUE,
    fs=10, bold=True, fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)
aw = (CONW - 0.15 * 3) / 4
for i, s in enumerate(["通信接入", "导航接入", "监视接入", "其他接入"]):
    add(p, CONX + i * (aw + 0.15), yy + 1.6, aw, 0.7, s, fill=XLBLUE, line=BLUE,
        fs=8.5, bold=True, fc=NAVY, align=1, valign=1, lw=0.6, rnd=0.04)

# ---------- 行3：端（左右两半，各若干段） ----------
yy = body_top + ROW_H[0] + 0.1 + ROW_H[1] + 0.1
end_h = ROW_H[2]
ew = (CONW - 0.2) / 2
left_secs = [("载具平台", "无人机 / eVTOL / 通航飞行器"),
             ("机载通信", "公网 / 卫星 / 自组网 / 数据链"),
             ("多源导航", "惯导 / 视觉 / 卫星 / 高度计")]
right_secs = [("通信基础设施", "移动网 / 卫星网 / 地空专网"),
              ("导航基础设施", "RTK / 地基增强 / 抗干扰监视"),
              ("监视基础设施", "RID 接收 / 低空雷达 / 光电红外")]
sec_h = (end_h - 0.55) / 3
add(p, CONX, yy, ew, 0.45, "机载终端与航电系统", fill=NAVY, line=NAVY,
    fs=10, bold=True, fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)
add(p, CONX + ew + 0.2, yy, ew, 0.45, "信息物理基础设施", fill=NAVY, line=NAVY,
    fs=10, bold=True, fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)
sy = yy + 0.5
for (lt, lb), (rt, rb) in zip(left_secs, right_secs):
    section(p, CONX, sy, ew, sec_h, lt, lb, BLUE, XXL)
    section(p, CONX + ew + 0.2, sy, ew, sec_h, rt, rb, BLUE, XXL)
    sy += sec_h

# ---------- 左竖栏：标准体系 ----------
sidebar(p, LX, body_top, LW, body_h, "标准体系",
        ["数据与服务支撑类标准", "智能网联类标准", "基础设施类标准"],
        head=NAVY, edge=NAVY)

# ---------- 右竖栏：安全体系 ----------
sidebar(p, RX, body_top, RW, body_h, "安全体系",
        ["装备体系安全", "网络数据安全", "电磁频谱安全"],
        head=RED, edge=RED)

# ---------- 底部：参与方 ----------
ptop = 14.50
add(p, 0.4, ptop - 0.4, 1.5, 0.35, "参与方", fill=GRAY, line=GRAY, fs=8.5,
    bold=True, fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)
pw = (PW - 0.8 - 0.2 * 4) / 5
for i, (t, c) in enumerate([("制造方", YELLOW), ("交管服务方", ORANGE),
                            ("监管方", PURPLE), ("运营方", GREEN),
                            ("基础设施服务方", BLUE)]):
    add(p, 0.4 + i * (pw + 0.2), ptop, pw, 0.6, t, fill=c, line=c, fs=9,
        bold=True, fc=WHITE, align=1, valign=1, lw=0.6, rnd=0.04)

# ---------- 页脚 ----------
footer(p, "示例架构图（模板自带）｜ 改造时请替换为拆解出的真实内容，"
          "并在页脚注明图片来源与对原图的补全/修正说明")

# ============================ 输出 ============================
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_visio_data.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump({"pageW": PW, "pageH": PH, "pages": PAGES}, f, ensure_ascii=False)
print("pages:", len(PAGES), "| shapes:", sum(len(x["shapes"]) for x in PAGES),
      "| conns:", sum(len(x["conns"]) for x in PAGES))
print("->", out)
