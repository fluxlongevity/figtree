import os
import numpy as np
from PIL import Image

assets_dir = r"c:\Users\andre\Documents\Codex\2026-06-16\files-mentioned-by-the-user-rosto\outputs\pagina-pedido-cadernos\assets"
files = [
    "miolo-pautado-com-data-dias-01.webp",
    "miolo-liso-com-data-dias.webp",
    "miolo-pontilhado-com-data.webp",
    "miolo-quadriculado-com-data.webp"
]

for name in files:
    path = os.path.join(assets_dir, name)
    if not os.path.exists(path):
        print(f"{name} missing")
        continue
    
    img = Image.open(path).convert("L")  # Convert to grayscale
    arr = np.array(img)
    
    # Let's count pixels that belong to the white notebook paper.
    # The concrete background has brightness values around 180.
    # The white paper is very bright, above 225.
    threshold = 225
    bright_mask = arr > threshold
    
    rows = np.any(bright_mask, axis=1)
    cols = np.any(bright_mask, axis=0)
    
    if not np.any(rows) or not np.any(cols):
        print(f"{name}: No notebook detected with threshold {threshold}")
        continue
        
    ymin, ymax = np.where(rows)[0][0], np.where(rows)[0][-1]
    xmin, xmax = np.where(cols)[0][0], np.where(cols)[0][-1]
    
    width = xmax - xmin
    height = ymax - ymin
    print(f"{name}: bbox x=({xmin}, {xmax}) y=({ymin}, {ymax}), notebook pages size = {width}x{height} pixels")
