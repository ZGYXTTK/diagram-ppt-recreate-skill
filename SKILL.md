---
name: "diagram-ppt-recreate"
description: "Recreates a diagram image as an editable PPT (native shapes/charts) with a drawing tutorial and comparison images. Invoke when a user gives a diagram image to recreate/redraw in PPT or wants a PPT drawing tutorial."
---

# 示意图 PPT 复现（Diagram PPT Recreate）

把用户提供的示意图原图复现为一页**原生可编辑**的 PowerPoint 文件，同时产出一份**手把手绘制教程**和**前后对比图**。工作流已在真实研报图表上完整验证（框图类：几何对齐达像素级，平均灰度差 < 8/255）。

## 触发场景

用户提供示意图原图（截图、扫描图、PNG/JPG），并提出以下任一诉求：
- "把这张图复现成 PPT / 照着这张图画 / 重绘这张示意图"
- "我想自己学怎么在 PPT 里画这张图 / 给我一份画图教程"
- "这张图能不能做成可编辑的"

## 适用范围

| 类型 | 判断特征 | 复现方式 | 支持 |
|---|---|---|---|
| 框图类 | 矩形 + 箭头 + 文字（产业链图、流程图、架构图、时间轴） | python-pptx 原生形状 | 完整流程 |
| 数据图表类 | 折线、柱状、饼图、组合图等带坐标轴的统计图 | python-pptx 原生图表对象 | 完整流程 |
| 图片型 | 照片、复杂插画、三维渲染 | — | 明确告知不支持，说明原因 |

类型判断在阶段 1 完成；一张原图兼有两种成分时（如柱状图 + 框图注释），分别走两条支线再合成一页。

## 运行流程总览

```
阶段0 轻交互确认 → 阶段1 原图分析 → 阶段2 教程制作 → 阶段3 PPT复现 → 阶段4 验证与对比 → 交付
```

一次性交付全套产物，除阶段 0 外不设人工确认点。

## 阶段 0：轻交互确认（1–2 个问题）

只问影响产出的关键项，其余决策全部自动：

1. 配色偏好：还原原图（默认）/ 用户提供品牌色或指定色
2. 仅数据图表类追加：数据来源——从原图像素测算（默认，需提示存在读数误差）/ 用户提供权威数据表

## 阶段 1：原图分析

临时脚本放系统临时目录，只读原图，不污染输出目录。依赖：`pip install pillow numpy scipy`。

### 框图类

1. **对象盘点**：读图列出栏头、流程框、块箭头、细箭头、文本框数量与连接走线（分叉汇合 / 纯串联 / 一分二等）。
2. **坐标实测**：用 `scripts/measure_diagram.py`——色系掩码 + 腐蚀断桥 + 连通域，输出每个矩形的像素 bbox，换算画布厘米。换算规则：宽度铺满画布、垂直居中留白，`X(cm) = x_px × (画布宽cm ÷ 图宽px)`，`Y(cm) = y_px × 同比例 + 留白`。画布默认跟随原图比例（16:9 时为 33.867 × 19.05 cm）。
3. **色值实测**：对每个矩形取**众数色**（抗文字污染，比中位数可靠）；边缘 1px 与中心色差 >10 判定有描边；底边/右缘外侧暗带判定阴影方向与延伸。
4. **字号测算**：文字行墨高换算 `pt = 墨高px × 72 ÷ dpi ÷ 0.92`，取 0.5 磅粒度；同时记录每个框的文字行数。
5. **文字誊录**：逐框抄录全部文字（含全半角符号、换行位置），教程与复现共用这份文字表。
6. **阴影参数映射**：暗带延伸约 7px（150dpi）对应 PowerPoint 预设"外部 → 右下斜偏移"（角度 45°、距离约 3.5 磅、模糊 4 磅、透明度 60%）。

**降级规则**：分辨率低或模糊导致掩码/连通域失败时，改用目测估计，且教程与交付说明中所有此类数值必须标注"估算值"。

### 数据图表类

1. 类型判断与构成（系列数、坐标轴、网格、图例、数据标签）。
2. **数据测算**：定位轴刻度像素 → 建立像素-数值映射 → 逐数据点取值；柱状图量柱高，折线图量点位，饼图量角度。结果标注"由原图测算，存在读数误差"。
3. 色值、字号方法同框图类。

## 阶段 2：教程制作（HTML）

按 doc-writing-guide 与 html-report 规范产出自包含 HTML。**每个数值必须让读者能直接照输**。

- **框图类章节骨架**（按图复杂度增删）：开始之前（定心丸：全程只用哪几类元素）→ 读懂原图（走线与对齐规律）→ 画布与准备（幻灯片大小、设置形状格式窗格、配色卡）→ 分阶段绘制（栏头 → 各列/各区块 → 箭头，每章末尾设"此时应看到"核对点）→ 收尾（选择窗格清点、叠加核对法、组合、导出）→ 常见问题 → 附录（色值表 / 坐标尺寸表 / 对象清单 / 全部文字表）。
- **图表类章节骨架**：准备数据表 → 插入图表并选类型 → 逐项设置（系列颜色、坐标轴、数据标签、图例）→ 对照检查 → 常见问题。
- 菜单路径按 **Microsoft 365 中文版**书写，WPS 与旧版 Office 差异用括注；同一路径全书只完整写一次。
- 文风：每章首句为事实或决定；规避"进行/开展 + 名词"；标题不带标点；无 emoji。
- 保存到用户输出目录：`示意图复现教程_<主题>.html` + `教程素材_<主题>/`（原图整图与局部特写裁切图）。

## 阶段 3：PPT 复现（python-pptx）

按教程附录数据逐对象生成，`pip install python-pptx`。关键实现模板：

```python
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

prs = Presentation()
prs.slide_width, prs.slide_height = Cm(33.867), Cm(19.05)
slide = prs.slides.add_slide(prs.slide_layouts[6])

def add_shadow(shape):
    """外部阴影·右下斜偏移：角度45° 距离3.5磅 模糊4磅 透明度60%"""
    xml = ('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<a:outerShdw blurRad="50800" dist="44450" dir="2700000" algn="tl" rotWithShape="0">'
           '<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>')
    shape._element.spPr.append(etree.fromstring(xml))

def add_box(x, y, w, h, rgb, lines, size, bold, tcolor):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor(*rgb)
    sp.line.fill.background()                     # 无轮廓
    add_shadow(sp)
    tf = sp.text_frame
    tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(lines):              # 多行文字 = 多段落
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = line
        f = r.font; f.size, f.bold, f.name = Pt(size), bold, "微软雅黑"
        f.color.rgb = RGBColor(*tcolor)
        rPr = r._r.get_or_add_rPr()
        ea = rPr.makeelement(qn('a:ea'), {'typeface': '微软雅黑'}); rPr.append(ea)

def add_arrow(x, y0, y1):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x), Cm(y0), Cm(x), Cm(y1))
    conn.line.color.rgb = RGBColor(0, 0, 0); conn.line.width = Pt(1)
    ln = conn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
```

- 数据图表类用 `slide.shapes.add_chart(...)` 生成原生图表对象，类别与系列数据来自阶段 1 测算表，系列颜色按实测色值设置。
- 生成后**必须读回校验**：对象总数、抽查坐标尺寸（容差 ±0.02cm）、颜色、文字全量、阴影标签数、箭头端点、连接线方向。

## 阶段 4：验证与对比

1. **渲染导出**：`scripts/export_and_compare.py` 用本机 PowerPoint COM（只读、无窗口）把幻灯片导出为 PNG（宽 2470）。依赖 `pip install pywin32`；本机需装有 PowerPoint，不可用时改用 LibreOffice 无头导出并在交付说明中注明。
2. **对比图两张**：并排图（原图在上、渲染在下，带标签）；50% 叠加图（渲染图裁至内容区后缩放到原图尺寸混合）。
3. **量化差异**：平均绝对差（/255）与差异 >30 的像素占比，打印进交付说明。
4. **修正循环**：几何错位（边框重影、箭头偏移）→ 修正坐标重新生成，最多 3 轮；文字级双影属字体光栅化差异，可接受并在交付说明注明。

## 固定产物清单

1. `示意图复现教程_<主题>.html` + `教程素材_<主题>/`
2. `示意图复现_<主题>.pptx`（原生可编辑）
3. `对比图_并排_<宽>x<高>.png`、`对比图_叠加_<宽>x<高>.png`（文件名含实际尺寸）
4. 交付说明：验证结论、量化差异、已知细微差异（文字光栅化等）

## 脚本工具（scripts/）

| 脚本 | 用途 |
|---|---|
| `scripts/measure_diagram.py` | 框图实测：主要色系矩形坐标、众数色值、字号、阴影延伸，输出 JSON + 控制台报告 |
| `scripts/export_and_compare.py` | PowerPoint COM 导出幻灯片 + 合成并排/叠加对比图 + 量化差异 |
