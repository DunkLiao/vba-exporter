# -*- coding: utf-8 -*-
"""產生 VBAExporter 的應用程式圖示（icon.ico + 視窗內嵌用 PNG 的 base64）。

執行後輸出：
- icon.ico   ：多尺寸圖示，給 PyInstaller 打包用
- icon_48.png：48x48 PNG
- 標準輸出最後一行 ICON_B64:... ：貼進 export_vba.py 當視窗圖示用
"""
import base64
import os

from PIL import Image, ImageDraw, ImageFont

GREEN = (33, 115, 70, 255)   # Excel 綠
WHITE = (255, 255, 255, 255)
MASTER = 1024                # 先畫大圖再縮小，邊緣才會平滑
OUT_DIR = os.path.dirname(os.path.abspath(__file__))

try:
    LANCZOS = Image.Resampling.LANCZOS
except AttributeError:  # 舊版 Pillow
    LANCZOS = Image.LANCZOS


def load_font(size):
    for name in ("arialbd.ttf", "segoeuib.ttf", "calibrib.ttf"):
        path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", name)
        if os.path.isfile(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def build_master():
    img = Image.new("RGBA", (MASTER, MASTER), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # 綠色圓角方形底
    d.rounded_rectangle([8, 8, MASTER - 8, MASTER - 8], radius=210, fill=GREEN)
    # "VBA" 文字
    f = load_font(int(MASTER * 0.34))
    d.text((MASTER / 2, MASTER * 0.38), "VBA", font=f, fill=WHITE, anchor="mm")
    # 下方白色匯出箭頭（下載意象）
    ax = MASTER / 2
    shaft_w = MASTER * 0.095
    d.rectangle([ax - shaft_w / 2, MASTER * 0.56,
                 ax + shaft_w / 2, MASTER * 0.72], fill=WHITE)
    d.polygon([
        (ax - MASTER * 0.145, MASTER * 0.71),
        (ax + MASTER * 0.145, MASTER * 0.71),
        (ax, MASTER * 0.855),
    ], fill=WHITE)
    return img


def main():
    master = build_master()

    ico_path = os.path.join(OUT_DIR, "icon.ico")
    master.resize((256, 256), LANCZOS).save(
        ico_path,
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48),
               (64, 64), (128, 128), (256, 256)])

    png_path = os.path.join(OUT_DIR, "icon_48.png")
    master.resize((48, 48), LANCZOS).save(png_path)
    with open(png_path, "rb") as fp:
        b64 = base64.b64encode(fp.read()).decode("ascii")

    print(f"已產生 {ico_path}")
    print(f"已產生 {png_path}")
    print("ICON_B64:" + b64)


if __name__ == "__main__":
    main()
