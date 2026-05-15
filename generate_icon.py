#!/usr/bin/env python3
"""Generate a Liquid Glass style app icon for MIoT Platform Tool - Green, final version."""
from PIL import Image, ImageDraw, ImageFilter
import os
import subprocess
import math

SIZE = 1024
RADIUS = 228
cx, cy = 512, 512

def lerp(a, b, t):
    return int(a + (b - a) * t)

# === LAYER 1: Green gradient base ===
base = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(base)

# Apple-style fresh green gradient (top-left bright -> bottom-right deep)
GREEN_TOP = (100, 235, 130)      # Bright lime-green at top
GREEN_MID = (52, 199, 89)        # Standard apple green
GREEN_BOT = (25, 115, 55)        # Deep forest green at bottom

for y in range(SIZE):
    for x in range(SIZE):
        ty = y / SIZE
        tx = x / SIZE
        # Diagonal gradient: top-left brighter
        t = (ty * 0.7 + tx * 0.3)
        
        # Slight radial brightening toward center
        dist = math.sqrt((tx - 0.5) ** 2 + (ty - 0.5) ** 2)
        center_boost = max(0, 1 - dist * 1.5) ** 2
        
        r = lerp(GREEN_TOP[0], GREEN_BOT[0], t) + int(center_boost * 15)
        g = lerp(GREEN_TOP[1], GREEN_BOT[1], t)
        b = lerp(GREEN_TOP[2], GREEN_BOT[2], t) + int(center_boost * 10)
        draw.point((x, y), fill=(r, g, b, 255))

# Mask to rounded rect
mask = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
mask_draw = ImageDraw.Draw(mask)
mask_draw.rounded_rectangle([(0, 0), (SIZE - 1, SIZE - 1)], radius=RADIUS, fill=(255, 255, 255, 255))
base = Image.composite(base, Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0)), mask)

# === LAYER 2: Glass top highlight (main specular) ===
glass = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
glass_draw = ImageDraw.Draw(glass)

# Wide soft highlight covering upper portion
for i in range(400):
    p = i / 400
    alpha = int(65 * (1 - p) ** 2.3)
    y = 35 + i
    # Elliptical highlight shape
    x_spread = int(RADIUS * 1.4 * math.sin(p * math.pi / 2) * (1.02 - p * 0.15))
    glass_draw.line([(cx - x_spread, y), (cx + x_spread, y)], fill=(255, 255, 255, alpha))

# Sharp bright specular near top-center
for i in range(90):
    p = i / 90
    alpha = int(150 * (1 - p) ** 1.6)
    y = 55 + i * 0.3
    x_spread = int(220 * math.sqrt(p))
    glass_draw.line([(cx - x_spread, y), (cx + x_spread, y)], fill=(255, 255, 255, alpha))

glass = glass.filter(ImageFilter.GaussianBlur(radius=14))
glass = Image.composite(glass, Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0)), mask)

# === LAYER 3: Inner border ===
border = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
bdr = ImageDraw.Draw(border)

for offset in range(5):
    a_top = int(150 - offset * 30)
    ir = RADIUS - 10 - offset * 3
    if ir > 10:
        bdr.rounded_rectangle(
            [(8 + offset * 3, 8 + offset * 3), (SIZE - 9 - offset * 3, SIZE - 9 - offset * 3)],
            radius=ir, outline=(255, 255, 255, a_top), width=2
        )

for offset in range(5):
    a_bot = int(50 - offset * 11)
    ir = RADIUS - 10 - offset * 3
    if ir > 10:
        bdr.rounded_rectangle(
            [(10 + offset * 3, 10 + offset * 3), (SIZE - 7 - offset * 3, SIZE - 7 - offset * 3)],
            radius=ir, outline=(15, 70, 38, a_bot), width=1
        )

border = Image.composite(border, Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0)), mask)

# === LAYER 4: M Logo ===
m_layer = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
m_d = ImageDraw.Draw(m_layer)

# Mijia-like proportions: tall vertical bars, clear V roof, smile bottom
# The M fills most of the icon area
bw = 85          # bar width (thick)
bl_x = cx - 270  # left bar outer x
br_x = cx + 270  # right bar outer x
bt_y = cy - 60   # bar top y
bb_y = cy + 280  # bar bottom y  
pk_y = cy - 340  # peak apex y (V-shape top)
sm_y = cy + 240  # smile lowest point

m_pts = [
    # Left bar top-outer
    (bl_x, bt_y),
    # Left bar up to V start
    (bl_x, bt_y - 25),
    # Left V-slope outer
    (cx - 60, pk_y),
    # Peak flat top
    (cx, pk_y - 14),
    (cx + 60, pk_y),
    # Right V-slope down
    (br_x, bt_y - 25),
    # Right bar top-outer
    (br_x, bt_y),
    # Right bar down outer
    (br_x, bb_y),
    # Right bar inner-bottom (smile starts)
    (br_x - bw, bb_y),
    # Smile curve up-left
    (br_x - bw - 70, bb_y - 95),
    (cx + 105, bb_y - 180),
    # Smile center (lowest point)
    (cx, sm_y),
    # Smile curve up-right
    (cx - 105, bb_y - 180),
    (bl_x + bw + 70, bb_y - 95),
    # Left bar inner-bottom
    (bl_x + bw, bb_y),
    # Left bar inner up
    (bl_x + bw, bt_y),
]

m_d.polygon(m_pts, fill=(255, 255, 255, 238))

# Glass sheen on upper part of M (subtle brightness on the letter's top half)
for i in range(320):
    p = i / 320
    alpha = int(30 * (1 - p) ** 2.8)
    y = pk_y - 14 + i
    if y < cy:
        m_d.line(
            [(bl_x + bw // 2, y), (br_x - bw // 2, y)],
            fill=(255, 255, 255, alpha)
        )

# Subtle shadow on lower inside of M bars
for i in range(200):
    p = i / 200
    alpha = int(15 * p ** 1.8)
    y = cy + 40 + i
    m_d.line(
        [(bl_x + bw, y), (br_x - bw, y)],
        fill=(180, 225, 195, alpha)
    )

# === LAYER 5: Bottom refraction glow ===
refract = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
rf_d = ImageDraw.Draw(refract)

for i in range(120):
    t = i / 120
    y = SIZE - 120 + int(i * 0.5)
    alpha = int(22 * (1 - abs(t - 0.5) * 2) ** 1.6)
    rc = int(170 + 60 * math.sin(t * math.pi * 3.5))
    gc = int(245 + 10 * math.cos(t * math.pi * 2))
    bc = int(190 + 65 * math.sin(t * math.pi * 3.5 + 1.5))
    sp = int(450 * (t if t < 0.5 else 1 - t))
    rf_d.line([(cx - sp, y), (cx + sp, y)], fill=(rc, gc, bc, alpha))

refract = refract.filter(ImageFilter.GaussianBlur(radius=20))
refract = Image.composite(refract, Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0)), mask)

# === COMPOSITE ALL ===
result = base.copy()
result = Image.alpha_composite(result, glass)
result = Image.alpha_composite(result, refract)
result = Image.alpha_composite(result, m_layer)
result = Image.alpha_composite(result, border)

# Save PNG
out_dir = os.path.dirname(os.path.abspath(__file__))
png_path = os.path.join(out_dir, "icon_1024.png")
result.save(png_path, "PNG")
print(f"Saved: {png_path}")

# Generate .icns
iconset = os.path.join(out_dir, "icon.iconset")
os.makedirs(iconset, exist_ok=True)

sizes = [16, 32, 64, 128, 256, 512]
for s in sizes:
    resized = result.resize((s, s), Image.LANCZOS)
    resized.save(os.path.join(iconset, f"icon_{s}x{s}.png"))
    s2 = min(s * 2, 1024)
    resized2 = result.resize((s2, s2), Image.LANCZOS)
    resized2.save(os.path.join(iconset, f"icon_{s}x{s}@2x.png"))

icns_path = os.path.join(out_dir, "icon.icns")
res = subprocess.run(["iconutil", "-c", "icns", iconset, "-o", icns_path],
                     capture_output=True, text=True)
print(f"Saved: {icns_path}" if res.returncode == 0 else f"Error: {res.stderr}")

import shutil
shutil.rmtree(iconset)
print("Done!")
