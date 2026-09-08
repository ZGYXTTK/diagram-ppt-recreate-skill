# -*- coding: utf-8 -*-
"""
export_and_compare.py — 用本机 PowerPoint 渲染导出幻灯片，并合成前后对比图

用法:
  python export_and_compare.py <复现.pptx> <原图.png> <输出目录> \
      [--top-frac 0.0] [--height-frac 1.0] [--width 2470] [--slide-w-cm 33.867] [--slide-h-cm 19.05]

参数:
  --top-frac / --height-frac  渲染页中"内容区"的纵向起止比例。若复现时内容区在画布内
                              垂直留白居中（教程默认画法），传 --auto-content 即可自动计算。
  --auto-content              依据画布与原图宽高比自动计算内容区比例（需 --slide-w-cm/--slide-h-cm）

依赖: pip install pywin32 pillow numpy；本机需安装 Microsoft PowerPoint
输出:
  <输出目录>/render_slide.png            幻灯片渲染图
  <输出目录>/对比图_并排_<宽>x<高>.png
  <输出目录>/对比图_叠加_<宽>x<高>.png
  控制台打印量化差异（平均绝对差、差异>30灰阶的像素占比）
"""
import argparse
import os
import tempfile

import numpy as np
from PIL import Image, ImageDraw, ImageFont


def export_slide(pptx_path, png_path, width):
    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(os.path.abspath(pptx_path), ReadOnly=True, WithWindow=False)
    try:
        pres.Slides(1).Export(os.path.abspath(png_path), "PNG", width, int(width * 9 / 16))
    finally:
        pres.Close()
        app.Quit()


def find_font(size):
    for p in [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc",
              r"C:\Windows\Fonts\simhei.ttf"]:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("original")
    ap.add_argument("outdir")
    ap.add_argument("--top-frac", type=float, default=None)
    ap.add_argument("--height-frac", type=float, default=None)
    ap.add_argument("--auto-content", action="store_true")
    ap.add_argument("--slide-w-cm", type=float, default=33.867)
    ap.add_argument("--slide-h-cm", type=float, default=19.05)
    ap.add_argument("--width", type=int, default=2470)
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    render_path = os.path.join(args.outdir, "render_slide.png")
    export_slide(args.pptx, render_path, args.width)

    orig = Image.open(args.original).convert("RGB")
    render = Image.open(render_path).convert("RGB")
    W = min(orig.width, 1235)
    r_small = render.resize((W, round(render.height * W / render.width)), Image.LANCZOS)

    # ---- 内容区比例 ----
    if args.auto_content:
        img_draw_h = args.slide_w_cm * orig.height / orig.width
        off_frac = max((args.slide_h_cm - img_draw_h) / 2 / args.slide_h_cm, 0)
        top_frac, height_frac = off_frac, img_draw_h / args.slide_h_cm
        print(f"[auto-content] top_frac={top_frac:.4f}, height_frac={height_frac:.4f}")
    else:
        top_frac = args.top_frac if args.top_frac is not None else 0.0
        height_frac = args.height_frac if args.height_frac is not None else 1.0

    # ---- 并排对比图 ----
    pad, gap, label_h = 20, 14, 44
    canvas = Image.new("RGB", (W + pad * 2, pad * 2 + label_h * 2 + orig.height + gap + r_small.height),
                       (246, 244, 240))
    d = ImageDraw.Draw(canvas)
    f_label = find_font(26)
    accent, ink = (116, 30, 0), (41, 36, 32)
    y = pad
    d.text((pad, y), "原图（用户提供的示意图）", font=f_label, fill=accent); y += label_h
    canvas.paste(orig, (pad, y)); y += orig.height + gap
    d.text((pad, y), "复现版（PPT 渲染导出）", font=f_label, fill=accent); y += label_h
    canvas.paste(r_small, (pad, y))
    d.rectangle([pad - 1, y - 1, pad + W, y + r_small.height], outline=ink, width=1)
    side_path = os.path.join(args.outdir, f"对比图_并排_{canvas.width}x{canvas.height}.png")
    canvas.save(side_path)
    print(f"并排对比图 {canvas.size} -> {side_path}")

    # ---- 50% 叠加对比图 ----
    top = round(top_frac * render.height)
    h_c = round(height_frac * render.height)
    content = render.crop((0, top, render.width, top + h_c)).resize(orig.size, Image.LANCZOS)
    overlay = Image.blend(orig, content, 0.5)
    overlay_path = os.path.join(args.outdir, f"对比图_叠加_{orig.width}x{orig.height}.png")
    overlay.save(overlay_path)
    print(f"叠加对比图 {overlay.size} -> {overlay_path}")

    # ---- 量化差异 ----
    a = np.asarray(orig).astype(int)
    b = np.asarray(content).astype(int)
    diff = np.abs(a - b)
    print(f"平均绝对差: {diff.mean():.2f} / 255")
    print(f"差异>30 的像素占比: {(diff.max(axis=2) > 30).mean() * 100:.2f}%")
    print("判定参考: 几何边框出现重影 -> 坐标需修正；仅文字处双影 -> 字体光栅化差异，可接受")


if __name__ == "__main__":
    main()
