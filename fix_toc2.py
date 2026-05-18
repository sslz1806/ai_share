"""
Rebuild slide2.xml TOC from scratch - proper XML approach.
"""
import re
from pathlib import Path
import xml.etree.ElementTree as ET

SLIDES = Path(r"C:\Users\20561\Desktop\不常用\tset_playwright\unpacked_pptx\ppt\slides")
path = SLIDES / "slide2.xml"
xml = path.read_text(encoding="utf-8")

# Register namespaces
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')

ns = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

tree = ET.parse(str(path))
root = tree.getroot()
spTree = root.find('.//p:cSld/p:spTree', ns)

# Get all shapes
shapes = list(spTree)

print(f"Total shapes: {len(shapes)}")

# Analyze shapes
for i, shape in enumerate(shapes):
    # Get text content
    texts = []
    for t in shape.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    text = ''.join(texts)[:40]

    # Get type and Y
    off = shape.find('.//a:off', ns)
    y = off.get('y') if off is not None else 'N/A'

    prst = shape.find('.//a:prstGeom', ns)
    stype = prst.get('prst') if prst is not None else ''

    cNvPr = shape.find('.//p:cNvPr', ns)
    id_val = cNvPr.get('id') if cNvPr is not None else 'N/A'
    name_val = cNvPr.get('name') if cNvPr is not None else 'N/A'

    print(f"  {i}: id={id_val}, name=\"{name_val}\", y={y}, type={stype}, \"{text}\"")
