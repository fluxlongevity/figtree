import os
from PIL import Image

assets_dir = r"c:\Users\andre\Documents\Codex\2026-06-16\files-mentioned-by-the-user-rosto\outputs\pagina-pedido-cadernos\assets"
files = [
    "miolo-pautado-sem-data-dias.webp",
    "miolo-pautado-com-data-dias-01.webp",
    "miolo-liso-com-data-dias.webp",
    "miolo-liso-sem-data-dias.webp",
    "miolo-pontilhado-sem-data.webp",
    "miolo-pontilhado-com-data.webp",
    "miolo-quadriculado-sem-data.webp",
    "miolo-quadriculado-com-data.webp"
]

for name in files:
    path = os.path.join(assets_dir, name)
    if os.path.exists(path):
        with Image.open(path) as img:
            print(f"{name}: size={img.size}")
    else:
        print(f"{name} missing")
