"""
Generate high-resolution branding assets for Gemini Antigravity Office Plugin.
Produces authentic, publication-grade Google Gemini 4-pointed sparkle icons
with transparent backgrounds, vibrant multi-stop gradients, and crisp scaling.
Outputs SVG, multi-resolution PNGs (16, 32, 64, 80, 128, 256), ICO, and installer BMP.
"""

import os
import math
import numpy as np
from PIL import Image, ImageDraw, ImageFont

def generate_svg(output_path):
    svg_content = """<?xml version="1.0" encoding="UTF-8"?>
<svg width="256" height="256" viewBox="0 0 256 256" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <!-- Google Gemini Official Radiant Gradient -->
    <linearGradient id="geminiGrad" x1="10%" y1="0%" x2="90%" y2="100%">
      <stop offset="0%" stop-color="#00E5FF"/>
      <stop offset="25%" stop-color="#1A73E8"/>
      <stop offset="55%" stop-color="#7C3AED"/>
      <stop offset="82%" stop-color="#D946EF"/>
      <stop offset="100%" stop-color="#FF4B72"/>
    </linearGradient>

    <!-- Inner Luminous Highlight -->
    <radialGradient id="coreGlow" cx="50%" cy="50%" r="42%">
      <stop offset="0%" stop-color="#FFFFFF" stop-opacity="0.90"/>
      <stop offset="40%" stop-color="#E9D5FF" stop-opacity="0.45"/>
      <stop offset="80%" stop-color="#7C3AED" stop-opacity="0.08"/>
      <stop offset="100%" stop-color="#7C3AED" stop-opacity="0.0"/>
    </radialGradient>

    <!-- Subtle Ambient Bloom -->
    <filter id="bloom" x="-25%" y="-25%" width="150%" height="150%">
      <feGaussianBlur stdDeviation="10" result="blur"/>
      <feComposite in="SourceGraphic" in2="blur" operator="over"/>
    </filter>
  </defs>

  <!-- Ambient Glow Behind Sparkle -->
  <path d="M128 14 C128 86 170 128 242 128 C170 128 128 170 128 242 C128 170 86 128 14 128 C86 128 128 86 128 14 Z"
        fill="url(#geminiGrad)" opacity="0.30" filter="url(#bloom)"/>

  <!-- Master Gemini 4-Pointed Sparkle Star -->
  <path d="M128 16 C128 84 172 128 240 128 C172 128 128 172 128 240 C128 170 84 128 16 128 C84 128 128 84 128 16 Z"
        fill="url(#geminiGrad)"/>

  <!-- Core Brilliant Luminescence -->
  <circle cx="128" cy="128" r="54" fill="url(#coreGlow)"/>
</svg>
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(svg_content)
    print(f"Generated SVG at {output_path}")

def render_icon(size):
    """
    Renders the Google Gemini 4-pointed sparkle on a 100% transparent background.
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
    edge_width = (1.6 * scale) / (R * p)
    alpha = np.clip(1.0 - (val - 1.0) / edge_width, 0.0, 1.0)

    # Gradient interpolation along diagonal from top-left/cyan to bottom-right/coral
    t = np.clip((dx * 0.65 - dy * 0.75 + 1.2) / 2.4, 0.0, 1.0)

    stops = [
        (0.00, np.array([0, 229, 255])),   # Electric Cyan (#00E5FF)
        (0.28, np.array([26, 115, 232])),  # Google Royal Blue (#1A73E8)
        (0.58, np.array([124, 58, 237])),  # Deep Indigo-Purple (#7C3AED)
        (0.82, np.array([217, 70, 239])),  # Vivid Fuchsia-Magenta (#D946EF)
        (1.00, np.array([255, 75, 114])),  # Radiant Coral-Rose (#FF4B72)
    ]

    r_channel = np.zeros_like(t)
    g_channel = np.zeros_like(t)
    b_channel = np.zeros_like(t)

    for i in range(len(stops) - 1):
        t0, c0 = stops[i]
        t1, c1 = stops[i + 1]
        m = (t >= t0) & (t <= t1)
        frac = (t[m] - t0) / (t1 - t0)
        # Cosine smoothstep for rich color blending
        smooth_frac = 0.5 * (1.0 - np.cos(math.pi * frac))

        r_channel[m] = c0[0] + (c1[0] - c0[0]) * smooth_frac
        g_channel[m] = c0[1] + (c1[1] - c0[1]) * smooth_frac
        b_channel[m] = c0[2] + (c1[2] - c0[2]) * smooth_frac

    # Center luminous bloom
    dist_center = np.sqrt(dx ** 2 + dy ** 2)
    core_glow = np.clip(1.0 - (dist_center / 0.35), 0.0, 1.0) ** 1.9
    core_boost = core_glow * (0.50 if is_small else 0.42)

    r_channel = np.clip(r_channel + (255 - r_channel) * core_boost, 0, 255)
    g_channel = np.clip(g_channel + (255 - g_channel) * core_boost, 0, 255)
    b_channel = np.clip(b_channel + (255 - b_channel) * core_boost, 0, 255)

    rgba = np.zeros((H, W, 4), dtype=np.uint8)
    rgba[..., 0] = r_channel.astype(np.uint8)
    rgba[..., 1] = g_channel.astype(np.uint8)
    rgba[..., 2] = b_channel.astype(np.uint8)
    rgba[..., 3] = (alpha * 255).astype(np.uint8)

    hi_img = Image.fromarray(rgba, mode="RGBA")
    final_img = hi_img.resize((size, size), Image.Resampling.LANCZOS)
    return final_img

def generate_banner(output_path):
    width, height = 497, 314  # Standard Inno Setup wizard banner
    img = Image.new("RGB", (width, height), (11, 15, 25))
    draw = ImageDraw.Draw(img)

    # Sleek dark radial-horizontal background gradient
    for y in range(height):
        ratio = y / height
        r = int(11 * (1 - ratio) + 19 * ratio)
        g = int(15 * (1 - ratio) + 24 * ratio)
        b = int(25 * (1 - ratio) + 45 * ratio)
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    # Ambient soft radial glow behind the icon
    glow_size = 200
    glow_img = Image.new("RGBA", (glow_size, glow_size), (0, 0, 0, 0))
    glow_draw = ImageDraw.Draw(glow_img)
    for rad in range(glow_size // 2, 0, -2):
        alpha = int(45 * (1.0 - (rad / (glow_size / 2.0))) ** 2.0)
        glow_draw.ellipse(
            [glow_size // 2 - rad, glow_size // 2 - rad, glow_size // 2 + rad, glow_size // 2 + rad],
            fill=(124, 58, 237, alpha)
        )
    img.paste(glow_img, (30 - 36, (height - 128) // 2 - 36), glow_img)

    # Render high-res 128px Gemini star on the banner
    icon_128 = render_icon(128)
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

    # 1. Vector SVG
    svg_path = os.path.join(assets_dir, "logo.svg")
    generate_svg(svg_path)

    # 2. Multi-size transparent PNGs
    sizes = [16, 32, 48, 64, 80, 128, 256]
    images = {}
    for s in sizes:
        png_path = os.path.join(assets_dir, f"icon-{s}.png")
        img = render_icon(s)
        img.save(png_path, "PNG")
        images[s] = img
        print(f"Generated {png_path} ({s}x{s})")

    # 3. Multi-resolution Favicon / ICO
    ico_path = os.path.join(assets_dir, "favicon.ico")
    # Save ICO containing all standard Windows sizes
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
