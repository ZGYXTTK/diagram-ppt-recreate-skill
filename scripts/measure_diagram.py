# -*- coding: utf-8 -*-
"""
measure_diagram.py — 框图类示意图实测工具
实测主要色系矩形/块箭头的坐标（换算为 PPT 画布厘米）、众数色值、字号、文字行数、阴影延伸。

用法:
  python measure_diagram.py <原图.png> [--slide-w 33.867] [--slide-h 19.05] [--out boxes.json]

依赖: pip install pillow numpy scipy

说明:
  - 自动识别图中占比最高的若干纯色填充（量化统计），逐色做掩码 + 腐蚀断桥 + 连通域找矩形
  - 若原图为渐变/纹理填充，请按控制台提示调大 --q 或降低 --min-share 后重试
  - 输出 JSON 含: 每个色系的矩形列表(px 与 cm)、众数 RGB、字号建议、文字行数、阴影参数
"""
import argparse
import json
import os
import numpy as np
from PIL import Image
from scipy import ndimage

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("--slide-w", type=float, default=33.867, help="画布宽 cm，默认 16:9 宽屏")
    ap.add_argument("--slide-h", type=float, default=19.05, help="画布高 cm")
    ap.add_argument("--out", default="boxes.json")
    ap.add_argument("--q", type=int, default=4, help="色值量化粒度")
    ap.add_argument("--min-share", type=float, default=0.02, help="色系最少占比（合并后）")
    ap.add_argument("--min-box-px", type=int, default=50, help="矩形最小宽 px")
    args = ap.parse_args()

    img = Image.open(args.image).convert("RGB")
    W, H = img.size
    arr = np.asarray(img).astype(int)
    R, G, B = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
    lum = arr.mean(axis=2)
    try:
        dpi = img.info.get("dpi", (150, 150))[0] or 150
    except Exception:
        dpi = 150

    scale = args.slide_w / W
    draw_h = H * scale
    off_y = (args.slide_h - draw_h) / 2
    print(f"图片 {W}x{H}px, dpi≈{dpi:.0f}, 比例 {scale:.6f} cm/px, "
          f"绘制高 {draw_h:.2f}cm, 顶部留白 {off_y:.3f}cm")
    if abs(draw_h - args.slide_h) / args.slide_h > 0.15:
        print("[提示] 原图宽高比与画布差异大，建议用 --slide-h 传入匹配的画布高度")

    # ---- 1. 识别主要纯色填充（量化桶贪心合并，排除近白背景与近黑文字）----
    flat = arr.reshape(-1, 3)
    quant = (flat // args.q) * args.q
    view = np.ascontiguousarray(quant).view([("r", int), ("g", int), ("b", int)]).ravel()
    vals, counts = np.unique(view, return_counts=True)
    order = np.argsort(-counts)
    total = len(flat)
    dominant = []
    used = np.zeros(len(vals), bool)
    for i in order[:60]:
        if used[i]:
            continue
        c = np.array(vals[i].tolist())
        if c.mean() > 235 or c.mean() < 30:      # 近白背景 / 近黑文字
            continue
        near = np.all(np.abs(np.stack([np.array(v.tolist()) for v in vals]) - c) <= args.q * 2, axis=1)
        merged_share = float(counts[near].sum() / total)
        used |= near
        if merged_share < args.min_share:
            continue
        dominant.append((c, merged_share))
    dominant.sort(key=lambda t: -t[1])
    print(f"识别到 {len(dominant)} 个主要填充色系")

    def to_cm(x, y, w, h):
        return dict(x_cm=round(x * scale, 2), y_cm=round(y * scale + off_y, 2),
                    w_cm=round(w * scale, 2), h_cm=round(h * scale, 2))

    def mode_color(x0, y0, x1, y1):
        reg = arr[y0:y1, x0:x1].reshape(-1, 3)
        qq = (reg // args.q) * args.q
        v2 = np.ascontiguousarray(qq).view([("r", int), ("g", int), ("b", int)]).ravel()
        vals2, counts2 = np.unique(v2, return_counts=True)
        top = np.array(vals2[np.argmax(counts2)].tolist())
        m = np.all(np.abs(qq - top) <= args.q, axis=1)
        med = np.median(reg[m], axis=0).astype(int)
        return [int(med[0]), int(med[1]), int(med[2])], float(counts2.max() / len(reg))

    def text_stats(x0, y0, x1, y1, bg_lum):
        """文字行数与字号建议（暗色文字或亮色文字自适应）"""
        reg = lum[y0 + 3:y1 - 3, x0 + 8:x1 - 8]
        dark_text = (np.median(reg) > 128)
        m = (reg < bg_lum - 60) if dark_text else (reg > bg_lum + 60)
        rows = m.sum(axis=1)
        bands, s = [], None
        for i, v in enumerate(rows):
            if v >= 2 and s is None:
                s = i
            elif v < 2 and s is not None:
                bands.append((s, i)); s = None
        if s is not None:
            bands.append((s, len(rows)))
        bands = [(a, b) for a, b in bands if b - a >= 6]
        sizes = [round((b - a) * 72 / dpi / 0.92 * 2) / 2 for a, b in bands]
        return len(bands), sizes

    def merge_split_fragments(frags, cmask, max_gap=45):
        """同行等高碎片若缝隙内仍有本色填充（被文字挤断）则并回一个框；缝隙为纯背景则是真并列"""
        rows = {}
        for b in frags:
            key = (round(b["y0"] / 4), round(b["h"] / 4))
            rows.setdefault(key, []).append(b)
        out = []
        for group in rows.values():
            if len(group) < 2:
                out.extend(group)
                continue
            group.sort(key=lambda b: b["x0"])
            merged = [dict(group[0])]
            for b in group[1:]:
                prev = merged[-1]
                g0 = prev["x0"] + prev["w"]
                g1 = b["x0"]
                gap = g1 - g0
                if 0 < gap <= max_gap:
                    strip = cmask[prev["y0"]:prev["y0"] + prev["h"], g0:g1]
                    if strip.mean() >= 0.25:      # 缝隙含本色填充 -> 被文字挤断，并回
                        prev["w"] = b["x0"] + b["w"] - prev["x0"]
                        continue
                merged.append(dict(b))
            out.extend(merged)
        return out

    # ---- 2. 逐色系找矩形（先收碎片 -> 同行伪切分合并 -> 再统一算属性）----
    report = {"image": dict(w=W, h=H, dpi=float(dpi), scale=scale, off_y=off_y),
              "slide": dict(w=args.slide_w, h=args.slide_h), "color_groups": []}
    all_boxes = []
    for c, share in dominant:
        mask = (np.abs(R - c[0]) <= args.q * 2) & (np.abs(G - c[1]) <= args.q * 2) & (np.abs(B - c[2]) <= args.q * 2)
        er = ndimage.binary_erosion(mask, structure=np.ones((3, 3)), iterations=3)
        lab, n = ndimage.label(er)
        frags = []
        for sl in ndimage.find_objects(lab):
            y0 = max(sl[0].start - 3, 0); y1 = min(sl[0].stop + 3, H)
            x0 = max(sl[1].start - 3, 0); x1 = min(sl[1].stop + 3, W)
            w_, h_ = x1 - x0, y1 - y0
            if w_ >= args.min_box_px and h_ >= 16 and mask[y0:y1, x0:x1].mean() > 0.55:
                frags.append(dict(x0=int(x0), y0=int(y0), w=int(w_), h=int(h_)))
        boxes = []
        for f in merge_split_fragments(frags, mask):
            x0, y0, w_, h_ = f["x0"], f["y0"], f["w"], f["h"]
            rgb, purity = mode_color(x0, y0, x0 + w_, y0 + h_)
            n_lines, sizes = text_stats(x0, y0, x0 + w_, y0 + h_, bg_lum=float(c.mean()))
            boxes.append(dict(px=dict(x0=x0, y0=y0, w=w_, h=h_),
                              **to_cm(x0, y0, w_, h_),
                              rgb=rgb, hex="#%02X%02X%02X" % tuple(rgb),
                              purity=round(purity, 2), text_lines=n_lines, font_pt=sizes))
            all_boxes.append(boxes[-1])
        if not boxes:
            continue
        boxes.sort(key=lambda b: (b["px"]["y0"], b["px"]["x0"]))
        report["color_groups"].append(dict(seed_rgb=[int(c[0]), int(c[1]), int(c[2])],
                                           share=round(share, 3), boxes=boxes))

    # ---- 3. 阴影探测（对面积最大的 3 个框）----
    big = sorted(all_boxes, key=lambda b: b["px"]["w"] * b["px"]["h"], reverse=True)[:3]
    bg = float(np.median(lum[0:max(H // 20, 10), :]))
    for b in big:
        px = b["px"]
        x0, y0, w_, h_ = px["x0"], px["y0"], px["w"], px["h"]
        prof = [float(lum[y0 + h_ + d, x0 + 10:x0 + w_ - 10].mean()) for d in range(0, 20)
                if y0 + h_ + d < H]
        reach = 0
        for i, v in enumerate(prof):
            if bg - v > 8:
                reach = i + 1
        b["shadow_down_px"] = reach
        b["shadow_depth"] = round(bg - min(prof[:reach]), 1) if reach else 0
    report["shadow_hint"] = ("阴影向下/向右各延伸约7px(150dpi)对应PowerPoint预设"
                             "外部-右下斜偏移(45°/3.5磅/模糊4磅/透明度60%)；实测值见各框 shadow_down_px")

    # ---- 4. 断言与输出 ----
    for b in all_boxes:
        c = b
        assert 0 <= c["x_cm"] and c["x_cm"] + c["w_cm"] <= args.slide_w + 0.05, f"越界 {c}"
        assert 0 <= c["y_cm"] and c["y_cm"] + c["h_cm"] <= args.slide_h + 0.05, f"越界 {c}"
    out = args.out if os.path.isabs(args.out) else os.path.join(os.getcwd(), args.out)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    n = sum(len(g["boxes"]) for g in report["color_groups"])
    print(f"共实测 {n} 个矩形，全部未越出画布；已输出 {out}")
    for g in report["color_groups"]:
        print(f"  色系 {g['seed_rgb']} (占比{g['share']:.0%}): {len(g['boxes'])} 框")

if __name__ == "__main__":
    main()
