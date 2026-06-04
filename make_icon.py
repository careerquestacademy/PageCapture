"""Generate assets/icon.ico for the build."""
from PIL import Image, ImageDraw
import os

os.makedirs('assets', exist_ok=True)

sizes  = [16, 32, 48, 64, 128, 256]
frames = []

for size in sizes:
    img  = Image.new('RGBA', (size, size), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    m    = max(2, size // 8)
    draw.rounded_rectangle([m, m, size-m, size-m],
                            radius=size//5,
                            fill=(203, 166, 247, 255))
    lw = max(1, size // 10)
    fs = int(size * 0.55)
    cx, cy = size//2, size//2
    x0 = cx - fs//4
    x1 = cx + fs//4
    y0 = cy - fs//2 + size//10
    y1 = cy + fs//2 - size//10
    ym = cy - size//16
    draw.rectangle([x0, y0, x0+lw, y1],    fill='white')
    draw.rectangle([x0, y0, x1,    y0+lw], fill='white')
    draw.rectangle([x0, ym, x1,    ym+lw], fill='white')
    draw.rectangle([x1-lw, y0, x1, ym+lw], fill='white')
    frames.append(img)

frames[0].save('assets/icon.ico', format='ICO',
               sizes=[(s,s) for s in sizes],
               append_images=frames[1:])
print("Created assets/icon.ico")
