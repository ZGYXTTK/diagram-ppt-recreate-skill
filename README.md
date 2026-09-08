# diagram-ppt-recreate — 示意图 PPT 复现 Agent Skill

把一张示意图原图（截图 / 扫描图 / PNG / JPG）复现为一页**原生可编辑**的 PowerPoint 文件，同时产出**手把手绘制教程（HTML）**与**前后对比图**。已在真实研报图表上完整验证：框图类几何对齐达像素级，平均灰度差 < 8/255。

## 它能做什么

| 图类型 | 判断特征 | 复现方式 | 支持 |
|---|---|---|---|
| 框图类 | 矩形 + 箭头 + 文字（产业链图、流程图、架构图、时间轴） | python-pptx 原生形状 | 完整流程 |
| 数据图表类 | 折线、柱状、饼图等带坐标轴的统计图 | python-pptx 原生图表对象 | 完整流程 |
| 图片型 | 照片、复杂插画、三维渲染 | — | 明确告知不支持 |

**不是**把图片贴进 PPT，而是逐对象重建：每个矩形、每条箭头、每行文字都是原生形状，可继续编辑。

## 快速开始

### 在 Agent（TRAE 等）中使用

把本仓库克隆到技能目录后，Agent 读取 `SKILL.md` 即可获得完整工作流。触发话术示例：

- "把这张图复现成 PPT"
- "我想自己学怎么在 PPT 里画这张图"
- "这张图能不能做成可编辑的"

### 单独使用两个脚本

```bash
# 1. 实测框图：输出每个矩形的像素/厘米坐标、众数色值、字号建议
python scripts/measure_diagram.py 原图.png --out boxes.json

# 2. 渲染导出 + 对比（需本机装有 Microsoft PowerPoint）
python scripts/export_and_compare.py 复现.pptx 原图.png 输出目录 --auto-content
```

详细参数与产物说明见 [USAGE.md](USAGE.md)，二次开发/定制示例见 [EXAMPLES.md](EXAMPLES.md)。

## 工作流总览

```
阶段0 轻交互确认 → 阶段1 原图分析 → 阶段2 教程制作 → 阶段3 PPT复现 → 阶段4 验证与对比 → 交付
```

- **阶段 1（原图分析）**：`scripts/measure_diagram.py` 用色系掩码 + 腐蚀断桥 + 连通域实测每个矩形的像素 bbox，换算为 PPT 画布厘米（默认 16:9，33.867 × 19.05 cm）；同时实测众数色值（抗文字污染）、字号（墨高换算 `pt = 墨高px × 72 ÷ dpi ÷ 0.92`）与阴影延伸。
- **阶段 2（教程制作）**：产出自包含 HTML 教程，每个数值可直接照输，含"此时应看到"核对点与附录（色值表 / 坐标尺寸表 / 对象清单 / 全部文字表）。
- **阶段 3（PPT 复现）**：按实测数据用 python-pptx 逐对象生成，生成后必须读回校验（坐标容差 ±0.02cm）。
- **阶段 4（验证对比）**：PowerPoint COM 渲染导出 PNG，合成并排图与 50% 叠加图，打印量化差异；几何错位最多修正 3 轮。

## 安装依赖

```bash
pip install pillow numpy scipy        # measure_diagram.py
pip install pywin32 pillow numpy      # export_and_compare.py（仅 Windows + PowerPoint）
pip install python-pptx               # 阶段 3 生成 PPT
```

## 仓库结构

```
├── SKILL.md                    # Agent 侧工作流契约（阶段 0–4 全流程规范）
├── README.md                   # 本文件
├── USAGE.md                    # 使用说明：脚本参数、产物、常见问题
├── EXAMPLES.md                 # 修改示例：画布、字体、灵敏度、阴影等定制
└── scripts/
    ├── measure_diagram.py      # 框图实测：坐标/色值/字号/阴影 → JSON
    └── export_and_compare.py   # PowerPoint 渲染导出 + 并排/叠加对比图 + 量化差异
```

## 固定产物清单

1. `示意图复现教程_<主题>.html` + `教程素材_<主题>/`（原图整图与局部特写）
2. `示意图复现_<主题>.pptx`（原生可编辑）
3. `对比图_并排_<宽>x<高>.png`、`对比图_叠加_<宽>x<高>.png`
4. 交付说明：验证结论、量化差异（平均绝对差 /255、差异>30 像素占比）、已知细微差异

## 已知边界

- 原图分辨率过低或填充为渐变/纹理时，实测降级为目测估计，且此类数值全部标注"估算值"。
- 渲染对比中文字处轻微双影属字体光栅化差异，可接受；几何边框重影则需要修正坐标。
- 画布默认跟随原图比例；原图宽高比与画布差异大时脚本会提示传入匹配的 `--slide-h`。
