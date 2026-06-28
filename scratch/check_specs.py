file_path = r"c:\Users\andre\Documents\Codex\2026-06-16\files-mentioned-by-the-user-rosto\outputs\pagina-pedido-cadernos\index.html"

with open(file_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

for idx in range(1348, 1420):
    if idx < len(lines):
        print(f"{idx+1}: {lines[idx].strip()}")
