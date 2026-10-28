"""Renders the legacy (pre-Android 8) launcher icons + 512px store icon for FB IDs Vault."""
from PIL import Image, ImageDraw
import os, math
S = 4
def render(size, round_=False, scale=1.0):
    N = size * S
    k = N / 108.0
    img = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    # gradient background
    g = Image.new("RGBA", (N, N))
    px = g.load()
    c1, c2 = (0x3F, 0x68, 0xFF), (0x1A, 0x2F, 0xB8)
    for y in range(N):
        for x in range(N):
            t = min(1, max(0, ((x / N) * 0.35 + (y / N) * 0.65)))
            px[x, y] = tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3)) + (255,)
    mask = Image.new("L", (N, N), 0)
    md = ImageDraw.Draw(mask)
    if round_: md.ellipse([0, 0, N - 1, N - 1], fill=255)
    else: md.rounded_rectangle([0, 0, N - 1, N - 1], radius=int(N * 0.22), fill=255)
    img.paste(g, (0, 0), mask)
    d = ImageDraw.Draw(img)
    def P(x, y):
        return ((x - 54) * scale + 54) * k, ((y - 54) * scale + 54) * k
    def bez(p0, p1, p2, p3, n=24):
        return [tuple((1-t)**3*p0[i] + 3*(1-t)**2*t*p1[i] + 3*(1-t)*t**2*p2[i] + t**3*p3[i] for i in range(2)) for t in [j/n for j in range(n+1)]]
    pts = [(54, 27), (77, 35.5), (77, 54)]
    pts += bez((77, 54), (77, 67.5), (67, 76.5), (54, 83))[1:]
    pts += bez((54, 83), (41, 76.5), (31, 67.5), (31, 54))[1:]
    pts += [(31, 35.5)]
    d.polygon([P(*p) for p in pts], fill=(255, 255, 255, 255))
    blue = (0x24, 0x42, 0xD9, 255)
    x0, y0 = P(41, 43); x1, y1 = P(67, 65)
    d.rounded_rectangle([x0, y0, x1, y1], radius=4 * scale * k, fill=blue)
    def circ(cx, cy, r, fill):
        a, b = P(cx - r, cy - r); c, e = P(cx + r, cy + r); d.ellipse([a, b, c, e], fill=fill)
    circ(48.5, 51.1, 2.9, (255, 255, 255, 255))
    d.pieslice([*P(43.6, 55.6), *P(53.4, 65.2)], 180, 360, fill=(255, 255, 255, 255))
    for (ya, w) in ((48.4, 7), (53.2, 7), (58, 4)):
        a, b = P(56.5, ya); c, e = P(56.5 + w, ya + 2)
        d.rounded_rectangle([a, b, c, e], radius=k * scale, fill=(255, 255, 255, 255))
    circ(54, 71.6, 2.4, blue)
    return img.resize((size, size), Image.LANCZOS)

dens = {"mdpi": 48, "hdpi": 72, "xhdpi": 96, "xxhdpi": 144, "xxxhdpi": 192}
res = "../app/src/main/res"
for n, s in dens.items():
    render(s, False, 1.32).save(f"{res}/mipmap-{n}/ic_launcher.webp", "WEBP", quality=95)
    render(s, True, 1.2).save(f"{res}/mipmap-{n}/ic_launcher_round.webp", "WEBP", quality=95)
render(512, False, 1.32).save("ic_launcher_512.png")
