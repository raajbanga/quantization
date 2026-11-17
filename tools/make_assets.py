"""Image assets for the deck (photos/pictures only; charts are native PowerPoint)."""
import sys
from pathlib import Path
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)

# Same synthetic landscape as the notebook (Part 2)
H, W = 480, 720
yy, xx = np.mgrid[0:H, 0:W] / np.array([H, W])[:, None, None]
img = np.zeros((H, W, 3))
img[..., 0] = 0.95 - 0.35 * yy
img[..., 1] = 0.55 + 0.25 * yy
img[..., 2] = 0.85 - 0.10 * yy
sun = ((xx - 0.7) ** 2 + ((yy - 0.35) * 1.5) ** 2) < 0.012
img[sun] = [1.0, 0.85, 0.35]
hills = yy > 0.62 + 0.08 * np.sin(xx * 9)
img[hills] = np.stack([0.15 + 0.2 * yy, 0.35 + 0.25 * (1 - yy), 0.25 + 0 * yy], -1)[hills]

def posterize(im, bits):
    n = 2 ** bits - 1
    return np.round(im * n) / n

for bits in (8, 4, 2, 1):
    plt.imsave(OUT / f"pic_{bits}bit.png", np.clip(posterize(img, bits), 0, 1))
print("assets ->", OUT)
