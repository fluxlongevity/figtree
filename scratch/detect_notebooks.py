import os
import numpy as np
from PIL import Image

brain_dir = r"C:\Users\andre\.gemini\antigravity\brain\b9450792-f598-4337-8276-e3f1c829029c"
images = [
    "miolo_pautado_mockup_1782477042333.png",
    "miolo_liso_mockup_1782477095022.png",
    "miolo_pontilhado_mockup_1782477149903.png",
    "miolo_quadriculado_mockup_1782477204567.png",
    "miolo_pautado_data_mockup_1782477584867.png"
]

for img_name in images:
    path = os.path.join(brain_dir, img_name)
    if not os.path.exists(path):
        print(f"{img_name} missing")
        continue
    
    img = Image.open(path).convert("L")  # Convert to grayscale
    arr = np.array(img)
    
    # The notebook pages are bright (high pixel values, e.g. > 220)
    # The concrete background is darker (medium pixel values, e.g. 150-190)
    # Let's find rows and columns that contain the bright pages
    threshold = 210
    bright_mask = arr > threshold
    
    # Find bounding box
    rows = np.any(bright_mask, axis=1)
    cols = np.any(bright_mask, axis=0)
    
    if not np.any(rows) or not np.any(cols):
        print(f"{img_name}: No notebook detected with threshold {threshold}")
        continue
        
    ymin, ymax = np.where(rows)[0][0], np.where(rows)[0][-1]
    xmin, xmax = np.where(cols)[0][0], np.where(cols)[0][-1]
    
    width = xmax - xmin
    height = ymax - ymin
    print(f"{img_name}: bbox x=({xmin}, {xmax}) y=({ymin}, {ymax}), size={width}x{height}")
