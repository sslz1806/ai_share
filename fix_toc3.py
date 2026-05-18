"""
Properly rebuild TOC from original XML. Uses ElementTree for precise control.
"""
from pathlib import Path
import re
import xml.etree.ElementTree as ET

SLIDES = Path(r"C:\Users\20561\Desktop\不常用\tset_playwright\unpacked_pptx\ppt\slides")
ns = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# Register ALL namespaces
ET.register_namespace('a', ns['a'])
ET.register_namespace('p', ns['p'])
ET.register_namespace('r', ns['r'])

# Read the ORIGINAL (freshly extracted) slide2
tree = ET.parse(str(SLIDES / "slide2_original.xml"))
root = tree.getroot()
spTree = root.find('.//p:cSld/p:spTree', ns)
shapes = list(spTree)

# New entries data: (number, title, subtitle, circle_y)
entries = [
    ("01", "AI 进化史速览", "大模型发展历程 &amp; 对工作的影响", "1280160"),
    ("02", "实用干货分享", "模型对比 · 工具推荐 · VibeCoding", "1965960"),
    ("03", "实用命令 &amp; 内置 Skills", "Claude Code 实用命令 · 内置 Skills 速查", "2651760"),
    ("04", "Web 自动化 &amp; Agent", "BrowserUse · Playwright · Agent Teams", "3337560"),
    ("05", "Agent 约束(harness)", "记忆体系 · Superpowers · 强制工作流", "4023360"),
]

def find_text_in_shape(shape, text):
    """Check if a shape contains a specific text string."""
    for t in shape.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text and text in t.text:
            return True
    return False

def is_ellipse(shape):
    """Check if shape is an ellipse."""
    prst = shape.find('.//a:prstGeom', ns)
    return prst is not None and prst.get('prst') == 'ellipse'

def get_y(shape):
    """Get Y position of shape."""
    off = shape.find('.//a:off', ns)
    return off.get('y') if off is not None else None

def set_y(shape, new_y):
    """Set Y position of all off elements in shape."""
    for off in shape.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}off'):
        off.set('y', new_y)

def set_text(shape, new_text):
    """Set the text content of a shape (replacing existing text)."""
    for t in shape.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        t.text = new_text
        break  # Only set first text element

# First: identify and modify entries 1-4 (existing shapes)
# We know the original Y values for each entry's shapes
old_entry_ys = [
    ("1463040", "1417320", "1737360"),  # entry 1 circle_y, title_y, subtitle_y
    ("2377440", "2331720", "2651760"),  # entry 2
    ("3291840", "3246120", "3566160"),  # entry 3
    ("4206240", "4160520", "4480560"),  # entry 4
]

for i, (num, title, subtitle, new_circle_y) in enumerate(entries[:4]):
    old_circle_y, old_title_y, old_sub_y = old_entry_ys[i]
    new_title_y = str(int(new_circle_y) - 45720)
    new_sub_y = str(int(new_circle_y) + 274320)

    # Find and modify the 4 shapes for this entry
    found_circle = False
    found_number = False
    found_title = False
    found_subtitle = False

    for shape in shapes:
        y = get_y(shape)
        if y is None:
            continue

        if y == old_circle_y and is_ellipse(shape):
            # Circle shape - change Y
            set_y(shape, new_circle_y)
            found_circle = True

        elif y == old_circle_y and find_text_in_shape(shape, num):
            # Number text shape - change Y, keep text
            set_y(shape, new_circle_y)
            found_number = True

        elif y == old_title_y:
            # Title text shape - change Y and text
            set_y(shape, new_title_y)
            set_text(shape, title)
            found_title = True

        elif y == old_sub_y:
            # Subtitle text shape - change Y and text
            set_y(shape, new_sub_y)
            set_text(shape, subtitle)
            found_subtitle = True

    if not all([found_circle, found_number, found_title, found_subtitle]):
        print(f"  Entry {num}: MISSING SHAPES! c={found_circle} n={found_number} t={found_title} s={found_subtitle}")
    else:
        print(f"  Entry {num}: ✓ all 4 shapes updated")

# Print all shapes after modification for verification
print("\nShapes after modification:")
for i, shape in enumerate(shapes):
    texts = []
    for t in shape.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    text = ''.join(texts)[:30]
    y = get_y(shape)
    e = "ellipse" if is_ellipse(shape) else ""
    print(f"  {i}: y={y} {e} \"{text}\"")

# ===== Add entry 5 (05 Agent约束) =====
# Entry 5 shapes need to be inserted between last subtitle (shape 19) and page number (shape 20)
c5 = "4023360"
t5 = "3977640"  # 4023360 - 45720
s5 = "4297680"  # 4023360 + 274320

# Build 4 new shape elements
from xml.etree.ElementTree import SubElement
import copy

# Helper: clone a shape from an existing entry and modify it
# Use shape 16 (circle) as template
template_circle = shapes[16]  # Entry 4 circle (ellipse)
template_number = shapes[17]  # Entry 4 number text
template_title = shapes[18]   # Entry 4 title
template_subtitle = shapes[19]  # Entry 4 subtitle

# Clone each template and modify
new_circle = copy.deepcopy(template_circle)
new_number = copy.deepcopy(template_number)
new_title = copy.deepcopy(template_title)
new_subtitle = copy.deepcopy(template_subtitle)

# Update Y positions
set_y(new_circle, c5)
set_y(new_number, c5)
set_y(new_title, t5)
set_y(new_subtitle, s5)

# Update IDs and names
new_circle.find('.//p:cNvPr', ns).set('id', '21')
new_circle.find('.//p:cNvPr', ns).set('name', 'Shape 16b')
new_number.find('.//p:cNvPr', ns).set('id', '22')
new_number.find('.//p:cNvPr', ns).set('name', 'Text 16b')
new_title.find('.//p:cNvPr', ns).set('id', '23')
new_title.find('.//p:cNvPr', ns).set('name', 'Text 16c')
new_subtitle.find('.//p:cNvPr', ns).set('id', '24')
new_subtitle.find('.//p:cNvPr', ns).set('name', 'Text 16d')

# Update text content
set_text(new_number, "05")
set_text(new_title, "Agent 约束(harness)")
set_text(new_subtitle, "记忆体系 · Superpowers · 强制工作流")

# Insert before page number (last shape in spTree)
page_num_shape = shapes[-1]
spTree.remove(page_num_shape)  # Remove from end
# Add entry 5 shapes, then page number back
spTree.append(new_circle)
spTree.append(new_number)
spTree.append(new_title)
spTree.append(new_subtitle)
spTree.append(page_num_shape)

# Update page number text from "2 / 26" to "2 / 29"
for t in page_num_shape.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
    if t.text and '/ 26' in t.text:
        t.text = t.text.replace('/ 26', '/ 29')
        print(f"  Page number updated: \"{t.text}\"")

print("  Entry 5 (05 Agent约束) added successfully!")

# Save final
tree.write(str(SLIDES / "slide2.xml"), encoding="utf-8", xml_declaration=True)
print("  Final slide2.xml saved!")

# Verify the complete TOC
print("\n=== FINAL TOC VERIFICATION ===")
tree2 = ET.parse(str(SLIDES / "slide2.xml"))
for shape in tree2.findall('.//p:sp', ns):
    texts = []
    for t in shape.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text:
            texts.append(t.text.strip())
    text = ' '.join(texts)[:40]
    y = get_y(shape)
    e = "●" if is_ellipse(shape) else ""
    if text or e:
        print(f"  y={y} {e} \"{text}\"")
