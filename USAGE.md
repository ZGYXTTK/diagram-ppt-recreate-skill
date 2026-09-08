# 使用说明（USAGE）

本文面向两类使用者：**直接调用本 skill 的 Agent**（读 `SKILL.md` 即可，本文是其补充速查）与**手动运行脚本的用户**。两个脚本的参数、输出与判定标准均以当前代码为准。

---

## 1. 环境准备

| 依赖 | 安装命令 | 用途 |
|---|---|---|
| pillow / numpy / scipy | `pip install pillow numpy scipy` | 框图实测（measure_diagram.py） |
| pywin32 / pillow / numpy | `pip install pywin32 pillow numpy` | 渲染导出与对比（export_and_compare.py，仅 Windows） |
| python-pptx | `pip install python-pptx` | 阶段 3 生成 PPT |

- `export_and_compare.py` 依赖本机安装的 **Microsoft PowerPoint**（COM 组件）；不可用时按 `SKILL.md` 阶段 4 的降级规则改用 LibreOffice 无头导出，并在交付说明中注明。
- 临时分析脚本放系统临时目录，只读原图，不污染输出目录。

## 2. measure_diagram.py — 框图实测

### 用法

```bash
python scripts/measure_diagram.py <原图.png> [--slide-w 33.867] [--slide-h 19.05] [--out boxes.json] [--q 4] [--min-share 0.02] [--min-box-px 50]
```

### 参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `image`（位置参数） | — | 原图路径，PNG/JPG |
| `--slide-w` | 33.867 | PPT 画布宽（cm），默认 16:9 宽屏 |
| `--slide-h` | 19.05 | PPT 画布高（cm） |
| `--out` | boxes.json | 输出 JSON 路径 |
| `--q` | 4 | 色值量化粒度；渐变/纹理填充时可调大 |
| `--min-share` | 0.02 | 色系最少占比（合并后）；漏识别小色块时调低 |
| `--min-box-px` | 50 | 矩形最小宽（px）；漏识别窄条时可调低 |

### 工作原理

1. **识别主要纯色填充**：量化统计图中占比最高的若干纯色（贪心合并近邻色桶），排除近白背景（均值 >235）与近黑文字（均值 <30）。
2. **逐色系找矩形**：色系掩码 → 3×3 腐蚀 3 轮断桥（断开文字造成的粘连）→ 连通域 → 同行等高碎片回并（缝隙内仍有本色填充说明是被文字挤断的同一框；缝隙为纯背景则是真并列）。
3. **逐框属性**：众数色值（比中位数抗文字污染）＋纯度、文字行数与字号（暗色/亮色文字自适应，`pt = 墨高px × 72 ÷ dpi ÷ 0.92`，0.5 磅粒度）、面积最大 3 框的向下阴影延伸与深度。
4. **画布换算**：宽度铺满画布、垂直居中留白，`X(cm) = x_px × (画布宽cm ÷ 图宽px)`，`Y(cm) = y_px × 同比例 + 留白`。原图宽高比与画布差异 >15% 时提示传入匹配的 `--slide-h`。
5. **越界断言**：输出前校验每个矩形都在画布内（容差 0.05cm），越界即报错。

### 输出

- `boxes.json`：`image`（宽高、dpi、比例、留白）、`slide`（画布）、`color_groups`（按占比排序；每组含 `seed_rgb`、`share`、`boxes` 列表——每框含 `px` 像素 bbox、`x_cm/y_cm/w_cm/h_cm`、`rgb`/`hex`、`purity`、`text_lines`、`font_pt`、大框另有 `shadow_down_px`/`shadow_depth`）。
- 控制台：图片信息、识别的色系数、每框汇总。
- `shadow_hint`：阴影映射提示——延伸约 7px（150dpi）对应 PowerPoint 预设"外部 → 右下斜偏移"（45°/3.5磅/模糊4磅/透明度60%）。

## 3. export_and_compare.py — 渲染导出与对比

### 用法

```bash
python scripts/export_and_compare.py <复现.pptx> <原图.png> <输出目录> \
    [--top-frac 0.0] [--height-frac 1.0] [--auto-content] \
    [--slide-w-cm 33.867] [--slide-h-cm 19.05] [--width 2470]
```

### 参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `pptx` / `original` / `outdir`（位置参数） | — | 复现文件、原图、输出目录 |
| `--top-frac` / `--height-frac` | 0.0 / 1.0 | 渲染页中"内容区"纵向起止比例；复现时画布内留白居中就用 `--auto-content` |
| `--auto-content` | 关 | 按画布与原图宽高比自动算内容区比例（需 `--slide-w-cm`/`--slide-h-cm`） |
| `--slide-w-cm` / `--slide-h-cm` | 33.867 / 19.05 | 画布尺寸，与 `--auto-content` 配合 |
| `--width` | 2470 | 导出 PNG 宽度（px） |

### 输出

- `render_slide.png`：PowerPoint COM 只读、无窗口导出的幻灯片渲染图。
- `对比图_并排_<宽>x<高>.png`：原图在上、渲染在下，带标签；渲染图缩到原图宽度（上限 1235px，LANCZOS）。
- `对比图_叠加_<宽>x<高>.png`：渲染图裁至内容区、缩放到原图尺寸后与原图 50% 混合。
- 控制台量化差异：**平均绝对差（/255）** 与 **差异>30 灰阶的像素占比**。

### 判读标准

| 现象 | 结论 | 处理 |
|---|---|---|
| 几何边框出现重影、箭头偏移 | 坐标错误 | 修正坐标重新生成，最多 3 轮 |
| 仅文字处轻微双影 | 字体光栅化差异 | 可接受，交付说明中注明 |

## 4. 在 Agent 中触发的完整流程

1. 用户提供示意图 + 表达复现/学习诉求（触发场景见 `SKILL.md`）。
2. 阶段 0：轻交互确认（配色偏好；图表类追加数据来源），1–2 个问题。
3. 阶段 1：框图类跑 `measure_diagram.py` 实测；图表类定位轴刻度建立像素-数值映射测算数据。
4. 阶段 2–3：生成教程 HTML 与 python-pptx 复现（实现模板见 `SKILL.md`）。
5. 阶段 4：跑 `export_and_compare.py` 出对比图与量化差异，必要时修正循环。
6. 交付固定产物清单（教程 + pptx + 两张对比图 + 交付说明）。

## 5. 常见问题

- **识别出的矩形比肉眼看到的少**：先调低 `--min-share`（如 0.01）与 `--min-box-px`（如 30）；渐变填充调大 `--q`。
- **一个框被识别成两个**：正常情况下同行回并逻辑已处理；仍发生时检查 `--q` 是否过大导致色系分裂。
- **叠加图整体错位**：确认复现时是否垂直居中留白——是则加 `--auto-content`，否则手动传 `--top-frac`/`--height-frac`。
- **COM 报错无法导出**：确认 PowerPoint 已安装且未被策略禁用 COM 自动化；或改用 LibreOffice 无头导出并在交付说明注明。
- **dpi 读取失败**：脚本回退 150dpi；对字号测算有影响时，可用图片属性中的真实 dpi 重新截图。
