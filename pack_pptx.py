"""
Standalone script to pack unpacked PPTX directory into a .pptx file.
Bypasses type hint issues in the skills pack.py.
"""
import zipfile
from pathlib import Path

input_dir = Path(r"C:\Users\20561\Desktop\不常用\tset_playwright\unpacked_pptx")
output_path = Path(r"C:\Users\20561\Desktop\不常用\tset_playwright\AI大矩阵分享_modified.pptx")

with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
    for f in sorted(input_dir.rglob("*")):
        if f.is_file():
            arcname = f.relative_to(input_dir)
            zf.write(f, arcname)

print(f"Packed to {output_path}")
print(f"File size: {output_path.stat().st_size:,} bytes")
