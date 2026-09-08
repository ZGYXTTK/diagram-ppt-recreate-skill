# 修改示例（EXAMPLES）

面向想**定制本 skill** 的开发者。每个示例给出：改哪个文件、改动位置、改前/改后代码。所有行号以当前代码为准，改代码后如涉及对外行为，请同步更新 `SKILL.md` 对应描述（`SKILL.md` 是 Agent 侧的工作流契约）。

## 示例 1：画布从 16:9 改为 4:3

**场景**：目标幻灯片是 4:3（25.4 × 19.05 cm）。涉及两个脚本的默认值，以及导出高度的宽高比。

**`scripts/measure_diagram.py`**（参数默认值）：

```python
# 改前
ap.add_argument("--slide-w", type=float, default=33.867, help="画布宽 cm，默认 16:9 宽屏")
# 改后
ap.add_argument("--slide-w", type=float, default=25.4, help="画布宽 cm，默认 4:3")
```

**`scripts/export_and_compare.py`**（两处）：

```python
# 改前：参数默认值与导出高度按 9/16 计算
ap.add_argument("--slide-w-cm", type=float, default=33.867)
ap.add_argument("--slide-h-cm", type=float, default=19.05)
...
pres.Slides(1).Export(os.path.abspath(png_path), "PNG", width, int(width * 9 / 16))

# 改后
ap.add_argument("--slide-w-cm", type=float, default=25.4)
ap.add_argument("--slide-h-cm", type=float, default=19.05)
...
pres.Slides(1).Export(os.path.abspath(png_path), "PNG", width, int(width * 3 / 4))
```

**同步文档**：`SKILL.md` 阶段 1 中"16:9 时为 33.867 × 19.05 cm"一句改为"4:3 时为 25.4 × 19.05 cm"。

## 示例 2：更换默认字体

**场景**：模板用微软雅黑，想换成思源黑体或公司规范字体。

**`SKILL.md`** 阶段 3 实现模板（生成 PPT 时代码按此书写，字体出现两处）：

```python
# 改前
f = r.font; f.size, f.bold, f.name = Pt(size), bold, "微软雅黑"
ea = rPr.makeelement(qn('a:ea'), {'typeface': '微软雅黑'}); rPr.append(ea)

# 改后（以思源黑体为例）
f = r.font; f.size, f.bold, f.name = Pt(size), bold, "思源黑体 CN"
ea = rPr.makeelement(qn('a:ea'), {'typeface': '思源黑体 CN'}); rPr.append(ea)
```

注意：西文字体（`latin`）与东亚字体（`a:ea`）要一起改，否则中文用新字体、数字英文仍是旧字体。

## 示例 3：调整矩形识别灵敏度

**场景**：某类图里小色块总是漏识别。不必改代码——运行时传参即可：

```bash
python scripts/measure_diagram.py 原图.png --min-share 0.01 --min-box-px 30 --out boxes.json
```

若想让某类原图**默认**就更灵敏，改 `scripts/measure_diagram.py` 的默认值：

```python
# 改前
ap.add_argument("--min-share", type=float, default=0.02, help="色系最少占比（合并后）")
ap.add_argument("--min-box-px", type=int, default=50, help="矩形最小宽 px")
# 改后
ap.add_argument("--min-share", type=float, default=0.01, help="色系最少占比（合并后）")
ap.add_argument("--min-box-px", type=int, default=30, help="矩形最小宽 px")
```

副作用提醒：`--min-share` 过低会把抗锯齿边缘、浅色装饰条误判为色系；建议一次只动一个参数并对比 `boxes.json` 的框数量与纯度（`purity`）。

## 示例 4：修改阴影参数映射

**场景**：原图阴影更长更淡，想换一种 PowerPoint 阴影预设。涉及 `SKILL.md` 的实现模板 XML 与 `measure_diagram.py` 的提示文案，两处需一致。

**`SKILL.md`** 阶段 3 的 `add_shadow`（单位：EMU，1 磅 = 12700；`dir` 单位为 1/60000 度）：

```python
# 改前：外部阴影·右下斜偏移 45°，距离 3.5 磅，模糊 4 磅，不透明度 40%（=透明度 60%）
xml = ('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
       '<a:outerShdw blurRad="50800" dist="44450" dir="2700000" algn="tl" rotWithShape="0">'
       '<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>')

# 改后：距离 5 磅（5×12700=63500 EMU）、模糊 6 磅（6×12700=76200 EMU）
xml = ('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
       '<a:outerShdw blurRad="76200" dist="63500" dir="2700000" algn="tl" rotWithShape="0">'
       '<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>')
```

**`scripts/measure_diagram.py`** 末尾的提示文案同步改写，保证实测像素能对上新预设：

```python
# 改前
report["shadow_hint"] = ("阴影向下/向右各延伸约7px(150dpi)对应PowerPoint预设"
                         "外部-右下斜偏移(45°/3.5磅/模糊4磅/透明度60%)；实测值见各框 shadow_down_px")
# 改后
report["shadow_hint"] = ("阴影向下/向右各延伸约7px(150dpi)对应PowerPoint预设"
                         "外部-右下斜偏移(45°/5磅/模糊6磅/透明度60%)；实测值见各框 shadow_down_px")
```

## 示例 5：调整导出宽度与差异判定阈值

**场景**：想要更高清的对比图，或收紧"差异像素"的判定。

**`scripts/export_and_compare.py`**：

```python
# 改前：导出宽度 2470；并排图渲染缩放上限 1235；判定阈值 30
ap.add_argument("--width", type=int, default=2470)
...
W = min(orig.width, 1235)
...
print(f"差异>30 的像素占比: {(diff.max(axis=2) > 30).mean() * 100:.2f}%")

# 改后：导出宽度 3200（也可运行时 --width 3200）；阈值收紧到 20
ap.add_argument("--width", type=int, default=3200)
...
print(f"差异>20 的像素占比: {(diff.max(axis=2) > 20).mean() * 100:.2f}%")
```

注意：`--width` 过大时 COM 导出变慢且受幻灯片原始分辨率限制；阈值收紧后文字抗锯齿像素会更多计入"差异"，判读时以"几何边框是否重影"为准而非绝对占比。

## 示例 6：叠加图内容区的两种处理方式

**场景**：复现 PPT 时没有采用"垂直居中留白"画法（内容顶格）。

```bash
# 默认：内容区 = 整个渲染页（顶格画法直接用默认即可）
python scripts/export_and_compare.py 复现.pptx 原图.png 输出目录

# 居中留白画法：自动计算内容区比例
python scripts/export_and_compare.py 复现.pptx 原图.png 输出目录 --auto-content

# 自定义：内容区从渲染页纵向 12% 处开始、占 76% 高
python scripts/export_and_compare.py 复现.pptx 原图.png 输出目录 --top-frac 0.12 --height-frac 0.76
```

三者互斥：`--auto-content` 优先于 `--top-frac`/`--height-frac`（见 `main()` 中的分支顺序）。

## 示例 7：新增一类图的识别（改 SKILL.md 而非脚本）

**场景**：想支持"表格型"原图。表格没有纯色矩形可测，脚本层无需改动，扩展点在 `SKILL.md`：

1. "适用范围"表新增一行（类型、判断特征、复现方式=python-pptx 表格对象、支持=完整流程）。
2. "阶段 1"新增小节：定位表格线像素 → 推算行列边界 → 逐单元格誊录文字与对齐方式。
3. "阶段 2/3"各补对应章节骨架与实现模板（`slide.shapes.add_table(...)`）。
4. 重新在真实图表上走一遍阶段 0–4 验证，把验证结论写回文档。

这是本 skill 的推荐扩展方式：`SKILL.md` 承载方法论，脚本只承担可程序化的实测/渲染环节。
