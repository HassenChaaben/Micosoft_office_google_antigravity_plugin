"""
Generate high-resolution branding assets for Gemini Antigravity Office Plugin.
Produces sleek, minimalist black Google Gemini 4-pointed sparkle icons
with 100% transparent backgrounds and razor-sharp subpixel antialiasing.
Outputs SVG, multi-resolution PNGs (16, 32, 48, 64, 80, 128, 256), ICO, and installer BMP.
"""

import os
import math
import numpy as np
from PIL import Image, ImageDraw, ImageFont

def generate_svg(output_path):
    svg_content = """<?xml version="1.0" encoding="UTF-8"?>
<svg width="256" height="256" viewBox="0 0 256 256" fill="none" xmlns="http://www.w3.org/2000/svg">
  <!-- Minimalist Sleek Black Gemini 4-Pointed Sparkle Star on Transparent Background -->
  <path d="M128 16 C128 84 172 128 240 128 C172 128 128 172 128 240 C128 170 84 128 16 128 C84 128 128 84 128 16 Z"
        fill="#000000"/>
</svg>
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(svg_content)
    print(f"Generated SVG at {output_path}")

def render_icon(size, color=(0, 0, 0)):
    """
    Renders the sleek black Google Gemini 4-pointed sparkle on a 100% transparent background.
    Uses high-resolution supersampling with signed distance curvature and Lanczos downsampling.
    """
    is_small = (size <= 32)
    scale = 8 if is_small else 4
    W = size * scale
    H = size * scale
    cx = W / 2.0
    cy = H / 2.0

    # At tiny sizes (16x16 / 32x32), slightly fuller waist for razor-sharp ribbon visibility
    p = 0.62 if is_small else 0.57
    padding = 0.05 if is_small else 0.07
    R = (min(W, H) / 2.0) * (1.0 - padding)

    y, x = np.ogrid[:H, :W]
    dx = (x - cx) / R
    dy = (y - cy) / R

    ax = np.abs(dx)
    ay = np.abs(dy)
    val = (ax ** p) + (ay ** p)

    # Subpixel antialiasing edge transition
    edge_width = (1.5 * scale) / (R * p)
    alpha = np.clip(1.0 - (val - 1.0) / edge_width, 0.0, 1.0)

    rgba = np.zeros((H, W, 4), dtype=np.uint8)
    rgba[..., 0] = color[0]
    rgba[..., 1] = color[1]
    rgba[..., 2] = color[2]
    rgba[..., 3] = (alpha * 255).astype(np.uint8)

    hi_img = Image.fromarray(rgba, mode="RGBA")
    final_img = hi_img.resize((size, size), Image.Resampling.LANCZOS)
    return final_img

def generate_banner(output_path):
    width, height = 497, 314  # Standard Inno Setup wizard banner
    img = Image.new("RGB", (width, height), (15, 23, 42))
    draw = ImageDraw.Draw(img)

    # Sleek dark radial-horizontal background gradient
    for y in range(height):
        ratio = y / height
        r = int(15 * (1 - ratio) + 24 * ratio)
        g = int(23 * (1 - ratio) + 32 * ratio)
        b = int(42 * (1 - ratio) + 56 * ratio)
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    # White/silver emblem on banner for high contrast against dark background
    icon_128 = render_icon(128, color=(255, 255, 255))
    img.paste(icon_128, (30, (height - 128) // 2), icon_128)

    # Typography
    try:
        font_title = ImageFont.truetype("arialbd.ttf", 26)
        font_sub = ImageFont.truetype("arial.ttf", 15)
        font_detail = ImageFont.truetype("arial.ttf", 13)
        font_small = ImageFont.truetype("arial.ttf", 11)
    except:
        font_title = ImageFont.load_default()
        font_sub = ImageFont.load_default()
        font_detail = ImageFont.load_default()
        font_small = ImageFont.load_default()

    draw.text((180, 95), "Gemini Antigravity", fill=(255, 255, 255), font=font_title)
    draw.text((180, 132), "Native AI Engine for Microsoft Office", fill=(56, 189, 248), font=font_sub)
    draw.text((180, 158), "Word, PowerPoint & Excel Autonomous Copilot", fill=(203, 213, 225), font=font_detail)
    draw.text((180, 198), "Google DeepMind Antigravity Edition • v1.0.0", fill=(148, 163, 184), font=font_small)

    img.save(output_path, "BMP")
    print(f"Generated Installer Banner at {output_path}")

def main():
    assets_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "assets"))
    os.makedirs(assets_dir, exist_ok=True)

    # 1. Vector SVG (Sleek black)
    svg_path = os.path.join(assets_dir, "logo.svg")
    generate_svg(svg_path)

    # 2. Multi-size transparent PNGs (Sleek black)
    sizes = [16, 32, 48, 64, 80, 128, 256]
    images = {}
    for s in sizes:
        png_path = os.path.join(assets_dir, f"icon-{s}.png")
        img = render_icon(s, color=(0, 0, 0))
        img.save(png_path, "PNG")
        images[s] = img
        print(f"Generated {png_path} ({s}x{s})")

    # 3. Multi-resolution Favicon / ICO
    ico_path = os.path.join(assets_dir, "favicon.ico")
    images[256].save(
        ico_path,
        format="ICO",
        sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    )
    print(f"Generated ICO at {ico_path}")

    # 4. Inno Setup Installer Banner BMP
    banner_path = os.path.join(assets_dir, "installer_banner.bmp")
    generate_banner(banner_path)

if __name__ == "__main__":
    main()
