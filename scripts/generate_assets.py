"""
Generate high-resolution branding assets for Gemini Antigravity Office Plugin.
Combines Google Gemini multi-point sparkle and Antigravity styling.
Outputs SVG, multi-resolution PNGs (16, 32, 64, 80, 128), ICO, and installer BMP.
"""

import os
import math
from PIL import Image, ImageDraw, ImageFont

def generate_svg(output_path):
    svg_content = """<?xml version="1.0" encoding="UTF-8"?>
<svg width="256" height="256" viewBox="0 0 256 256" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="geminiGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#1A73E8"/>
      <stop offset="50%" stop-color="#8E24AA"/>
      <stop offset="100%" stop-color="#00BCD4"/>
    </linearGradient>
    <linearGradient id="sparkleGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FFFFFF"/>
      <stop offset="50%" stop-color="#E1BEE7"/>
      <stop offset="100%" stop-color="#80D8FF"/>
    </linearGradient>
    <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
      <feGaussianBlur stdDeviation="6" result="blur" />
      <feComposite in="SourceGraphic" in2="blur" operator="over" />
    </filter>
  </defs>

  <!-- Background Rounded Shield -->
  <rect x="16" y="16" width="224" height="224" rx="56" fill="#0B0F19"/>
  <rect x="18" y="18" width="220" height="220" rx="54" stroke="url(#geminiGrad)" stroke-width="4" fill="none" opacity="0.8"/>

  <!-- Antigravity Ring / Orbit -->
  <ellipse cx="128" cy="128" rx="84" ry="34" stroke="url(#geminiGrad)" stroke-width="6" transform="rotate(-25 128 128)" stroke-dasharray="14 8" opacity="0.85"/>

  <!-- Center Gemini 4-Point Star Sparkle -->
  <!-- Outer glow -->
  <path d="M128 44 C128 92 164 128 212 128 C164 128 128 164 128 212 C128 164 92 128 44 128 C92 128 128 92 128 44 Z"
        fill="url(#geminiGrad)" filter="url(#glow)" opacity="0.9"/>
  
  <!-- Inner brilliant sparkle -->
  <path d="M128 58 C128 96 158 128 198 128 C158 128 128 160 128 198 C128 160 98 128 58 128 C98 128 128 96 128 58 Z"
        fill="url(#sparkleGrad)"/>

  <!-- Antigravity Core Glyph (A-triangle inverted levitation) -->
  <circle cx="128" cy="128" r="14" fill="#0B0F19"/>
  <polygon points="128,120 134,132 122,132" fill="#00E5FF"/>
</svg>
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(svg_content)
    print(f"Generated SVG at {output_path}")

def render_icon(size):
    # High resolution supersampling
    scale = 4
    canvas_size = size * scale
    img = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    margin = 8 * scale
    rect_box = [margin, margin, canvas_size - margin, canvas_size - margin]
    corner_radius = int(canvas_size * 0.22)

    # Background rounded rect (Dark slate)
    draw.rounded_rectangle(rect_box, radius=corner_radius, fill=(15, 23, 42, 255), outline=(56, 189, 248, 220), width=int(2.5 * scale))

    # Center coordinates
    cx = canvas_size // 2
    cy = canvas_size // 2

    # Draw Antigravity Orbit Ring (ellipse)
    orbit_rx = int(canvas_size * 0.36)
    orbit_ry = int(canvas_size * 0.16)
    orbit_box = [cx - orbit_rx, cy - orbit_ry, cx + orbit_rx, cy + orbit_ry]
    # Draw arc
    draw.arc(orbit_box, start=30, end=330, fill=(168, 85, 247, 200), width=int(2.5 * scale))

    # Draw Gemini Star: 4 points curve
    star_radius = int(canvas_size * 0.38)
    inner_pinch = int(canvas_size * 0.08)

    points = []
    # 4-pointed star curve approximation
    steps = 72
    for i in range(steps):
        angle = 2 * math.pi * i / steps
        # Distance formula for 4-point astroid star
        # r(theta) = R / ( (|cos|^p + |sin|^p)^(1/p) )
        c = abs(math.cos(angle * 2))
        r = inner_pinch + (star_radius - inner_pinch) * (c ** 1.8)
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        points.append((x, y))

    draw.polygon(points, fill=(59, 130, 246, 255))

    # Inner bright star
    inner_points = []
    inner_radius = int(star_radius * 0.7)
    for i in range(steps):
        angle = 2 * math.pi * i / steps
        c = abs(math.cos(angle * 2))
        r = (inner_pinch * 0.6) + (inner_radius - inner_pinch * 0.6) * (c ** 2.0)
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        inner_points.append((x, y))

    draw.polygon(inner_points, fill=(236, 72, 153, 255))

    # Center Core Glow
    draw.ellipse([cx - int(4 * scale), cy - int(4 * scale), cx + int(4 * scale), cy + int(4 * scale)], fill=(255, 255, 255, 255))

    # Resample down to requested size with high quality Lanczos filter
    final_img = img.resize((size, size), Image.Resampling.LANCZOS)
    return final_img

def generate_banner(output_path):
    width, height = 497, 314  # Standard Inno Setup wizard banner
    img = Image.new("RGB", (width, height), (15, 23, 42))
    draw = ImageDraw.Draw(img)

    # Draw sleek gradient background
    for y in range(height):
        ratio = y / height
        r = int(15 * (1 - ratio) + 30 * ratio)
        g = int(23 * (1 - ratio) + 27 * ratio)
        b = int(42 * (1 - ratio) + 75 * ratio)
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    # Render a 128px icon on the banner
    icon_128 = render_icon(128)
    img.paste(icon_128, (30, (height - 128) // 2), icon_128)

    # Draw Title & Tagline text
    try:
        font_title = ImageFont.truetype("arialbd.ttf", 26)
        font_sub = ImageFont.truetype("arial.ttf", 15)
        font_small = ImageFont.truetype("arial.ttf", 12)
    except:
        font_title = ImageFont.load_default()
        font_sub = ImageFont.load_default()
        font_small = ImageFont.load_default()

    draw.text((180, 100), "Gemini Antigravity", fill=(255, 255, 255), font=font_title)
    draw.text((180, 135), "Native AI Engine for Microsoft Office", fill=(56, 189, 248), font=font_sub)
    draw.text((180, 160), "Word & PowerPoint Autonomous Copilot", fill=(203, 213, 225), font=font_sub)
    draw.text((180, 200), "v1.0.0 Production Edition", fill=(148, 163, 184), font=font_small)

    img.save(output_path, "BMP")
    print(f"Generated Installer Banner at {output_path}")

def main():
    assets_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "assets"))
    os.makedirs(assets_dir, exist_ok=True)

    # 1. SVG
    svg_path = os.path.join(assets_dir, "logo.svg")
    generate_svg(svg_path)

    # 2. PNG sizes: 16, 32, 64, 80, 128
    sizes = [16, 32, 64, 80, 128]
    images = {}
    for s in sizes:
        png_path = os.path.join(assets_dir, f"icon-{s}.png")
        img = render_icon(s)
        img.save(png_path, "PNG")
        images[s] = img
        print(f"Generated {png_path} ({s}x{s})")

    # 3. Favicon / ICO
    ico_path = os.path.join(assets_dir, "favicon.ico")
    images[64].save(ico_path, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
    print(f"Generated ICO at {ico_path}")

    # 4. Installer Banner BMP
    banner_path = os.path.join(assets_dir, "installer_banner.bmp")
    generate_banner(banner_path)

if __name__ == "__main__":
    main()
