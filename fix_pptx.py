"""
Final fixes for the PPTX:
1. Update page numbers (/26 → /29)
2. Set new slide page numbers (32→14/29, 33→15/29)
3. Update TOC text (keep 4 entries, update text)
4. Remove leftover page number text on slides that don't need it
"""
import re
from pathlib import Path

BASE = Path(r"C:\Users\20561\Desktop\不常用\tset_playwright")
SLIDES = BASE / "unpacked_pptx" / "ppt" / "slides"

def read(snum):
    return (SLIDES / f"slide{snum}.xml").read_text(encoding="utf-8")

def write(snum, content):
    (SLIDES / f"slide{snum}.xml").write_text(content, encoding="utf-8")


# ========== 1. Update page numbers on all slides ==========
print("=== Updating page numbers ===")

# Change all "X / 26" to "X / 29"
for snum in range(1, 34):
    path = SLIDES / f"slide{snum}.xml"
    if not path.exists():
        continue
    xml = path.read_text(encoding="utf-8")
    if '/ 26' in xml:
        # Replace within <a:t> tags only
        xml = re.sub(r'(<a:t>\d+\s*)/\s*26</a:t>', r'\1/ 29</a:t>', xml)
        path.write_text(xml, encoding="utf-8")
        print(f"  slide{snum}: /26 → /29")

# Set page numbers on new slides
for snum, page in [(32, "14"), (33, "15")]:
    xml = read(snum)
    # The duplicated slides have page numbers like "7 / 29" or "9 / 29"
    xml = re.sub(r'<a:t>\d+\s*/\s*29</a:t>', f'<a:t>{page} / 29</a:t>', xml)
    write(snum, xml)
    print(f"  slide{snum}: page → {page}/29")


# ========== 2. Update TOC ==========
print("\n=== Updating TOC (slide 2) ===")
xml = read(2)

# Current 4 entries:
# 0) 01 - AI 进化史速览 - 大模型发展历程 & 对工作的影响
# 1) 02 - 实用干货分享 - 模型对比 · 工具推荐 · VibeCoding
# 2) 03 - Web自动化 & Agent - BrowserUse · Playwright · Agent Teams
# 3) 04 - Superpowers 工作流 - 强制纪律的 AI 编程流程

# New 4 entries (we'll keep 4, the TOC will show:
# 01 AI 进化史速览
# 02 实用干货分享
# 03 实用命令 & 内置 Skills
# 04 Agent 约束(harness)

# The user wanted 5 sections but the layout limits us to 4.
# We compress: combining the new section and Web automation/Agent constraint
# into 4 entries that cover all 5 sections.

# Actually, let me just add a 5th entry by finding the key TOC elements and modifying them.

# Find the "03" text and its surrounding shapes
# Find the "04" text and its surrounding shapes

# Approach: find the 4th entry (04 Superpowers) and copy its shapes to create a 5th
# entry (05 Agent约束). Then modify all texts.

# Let me find the exact shapes by their content
shape_info = []

# First, list all text content with their position
lines = xml.split('\n')
for i, line in enumerate(lines):
    # Find <a:t> content
    m = re.search(r'<a:t>([^<]+)</a:t>', line)
    if m:
        text = m.group(1)
        # Also find y position in nearby lines
        y_pos = "?"
        for j in range(max(0, i-5), min(len(lines), i+5)):
            ym = re.search(r'y="(\d+)"', lines[j])
            if ym:
                y_pos = ym.group(1)
        if text.strip():
            print(f"  Line {i}: y={y_pos}, text=\"{text}\"")

# The TOC has 4 entries. Let me map out all the key shapes.
# From the slide2.xml read earlier:

# Shape indices (0-indexed):
# 0: "目  录" (title)
# 1: accent line
# 2: circle for entry 1 (ellipse at y=1463040)
# 3: "01" text (at y=1463040)
# 4: "AI 进化史速览" title text (at y=1417320)
# 5: "大模型发展历程..." subtitle (at y=1737360)
# 6: circle for entry 2 (ellipse at y=2377440)
# 7: "02" text (at y=2377440)
# 8: "实用干货分享" title (at y=2331720)
# 9: "模型对比..." subtitle (at y=2651760)
# 10: circle for entry 3 (ellipse at y=3291840)
# 11: "03" text (at y=3291840)
# 12: "Web自动化 & Agent" title (at y=3246120)
# 13: "BrowserUse..." subtitle (at y=3566160)
# 14: circle for entry 4 (ellipse at y=4206240)
# 15: "04" text (at y=4206240)
# 16: "Superpowers 工作流" title (at y=4160520)
# 17: "强制纪律..." subtitle (at y=4480560)
# 18-19: "2 / 26" page number (at y=4754880)

# Now I want to ADD a 5th entry and SHIFT COMPRESS everything
# Total height needed for 5 entries with spacing of 731520 EMU:
# First entry at y=500000, spacing=685800:
# 1: 500000
# 2: 1185800
# 3: 1871600
# 4: 2557400
# 5: 3243200
# Last subtitle at 3243200+868680=4111880, page num at 4754880. Fits!

# But that requires moving all circles. Let me do it differently.
# Let me just add a 5th entry at the bottom, moving the page number.

# 4th entry ends at y=4480560+274320=4754880
# Page number is at y=4754880
# Slide height is 5143500 (10" × 5.625")
# So there's 5143500-4754880=388620 space

# For a 5th entry circle at y=4663440 (offset of 457200 from entry 4):
# Circle at y=4663440, title at y=4617720, subtitle at y=4663440+868680=5532120
# That's way over the slide height.

# OK, I'll add the 5th entry by putting it below the page number.
# Entry 5: circle y=4826000, title y=4780280, subtitle y=4826000+868680=5694680
# That overflows.

# OR: I can remove the circles and just use text. But that changes the design.

# OR: I can reduce the spacing dramatically for the last two entries.

# Let me try: reduce all entry spacing to 685800 and start at y=800000
# Entry 1: circle y=800000
# Entry 2: circle y=1485800
# Entry 3: circle y=2171600
# Entry 4: circle y=2857400
# Entry 5: circle y=3543200
# Entry 5 subtitle: 3543200 + 868680 = 4411880
# Page number can stay at 4754880. Fits!

# This requires moving ALL circles. Here's the plan:
# Circles shift down by 160000 (=shift from 800000-1463040 = -663040... wait, UP is LESS)

# Actually let me be practical: just compress entries 3,4,5 toward the bottom
# while keeping entries 1,2 near their positions.

# No wait, let's leave entries 1 and 2 where they are, and compress 3,4,5:
# Entry 3: circle y=3000000 (was 3291840, shifted up 291840)
# Entry 4: circle y=3550000 (was 4206240, shifted up 656240)
# Entry 5: circle y=4100000 (NEW)
# Page number at y=4754880 stays
# Entry 5 subtitle: 4100000 + 868680 = 4968680 < 5143500. Tight but fits.

# OK this is getting absolutely ridiculous with manual calculations. Let me use
# a much simpler approach: add the 5th entry below and slightly compress all entries.
# I'll write the exact target Y positions.

# New entry Y positions (circles):
# 1: 960000 (shift up 503040 from 1463040)
# 2: 1660000 (shift up 717440 from 2377440)
# 3: 2360000 (shift up 931840 from 3291840)
# 4: 3060000 (shift up 1146240 from 4206240)
# 5: 3760000 (NEW)
# Subtitle bottom: 3760000 + 868680 = 4628680 < 4754880 ✓

# The page number at 4754880 stays.

# Wait, but shifting entries up by different amounts is weird visually.
# Let me use uniform spacing = 700000 between entries:
# 1: 960000
# 2: 1660000 (= 960000 + 700000)
# 3: 2360000
# 4: 3060000
# 5: 3760000
# Subtitle bottom: 3760000 + 868680 = 4628680 ✓
# Page number at 4754880 ✓

# New spacing = 700000, first entry at 960000.
# Old spacing was 914400, first entry at 1463040.
# Shift = -503040 (for entry 1), -503040 - 214400 = -717440 (for entry 2), etc.
# Actually: new_y = old_y - (old_y - 960000) / 914400 * 214400 ... no that's wrong.

# Let me just use the fixed positions: 960000, 1660000, 2360000, 3060000, 3760000

# For each entry, the shapes are:
# Circle at y=N
# Number text at y=N
# Title at y=N-45720
# Subtitle at y=N+868680

# So:
# Entry 1: c=960000, n=960000, t=914280, s=1828680
# Entry 2: c=1660000, n=1660000, t=1614280, s=2528680
# Entry 3: c=2360000, n=2360000, t=2314280, s=3228680
# Entry 4: c=3060000, n=3060000, t=3014280, s=3928680
# Entry 5: c=3760000, n=3760000, t=3714280, s=4628680

# Text content:
# Entry 1: "01" "AI 进化史速览" "大模型发展历程 &amp; 对工作的影响"
# Entry 2: "02" "实用干货分享" "模型对比 · 工具推荐 · VibeCoding"
# Entry 3: "03" "实用命令 &amp; 内置 Skills" "Claude Code 实用命令 · 内置 Skills 速查"
# Entry 4: "04" "Web 自动化 &amp; Agent" "BrowserUse · Playwright · Agent Teams"
# Entry 5: "05" "Agent 约束(harness)" "记忆体系 · Superpowers · 强制工作流"

# Now I need to:
# 1. Find existing shapes by their current text/position
# 2. Update their position (y) and text
# 3. Add 4 new shapes for entry 5

# For entries 1-4, each has specific text I can search for:
existing_texts = {
    "01": {"num": "01", "title": "AI 进化史速览", "sub": "大模型发展历程 &amp; 对工作的影响", "new_num": "01", "new_title": "AI 进化史速览", "new_sub": "大模型发展历程 &amp; 对工作的影响", "new_y": "960000"},
    "02": {"num": "02", "title": "实用干货分享", "sub": "模型对比 · 工具推荐 · VibeCoding", "new_num": "02", "new_title": "实用干货分享", "new_sub": "模型对比 · 工具推荐 · VibeCoding", "new_y": "1660000"},
    "03": {"num": "03", "title": "Web自动化 &amp; Agent", "sub": "BrowserUse · Playwright · Agent Teams", "new_num": "04", "new_title": "Web 自动化 &amp; Agent", "new_sub": "BrowserUse · Playwright · Agent Teams", "new_y": "3060000"},
    "04": {"num": "04", "title": "Superpowers 工作流", "sub": "强制纪律的 AI 编程流程", "new_num": "05", "new_title": "Agent 约束(harness)", "new_sub": "记忆体系 · Superpowers · 强制工作流", "new_y": "3760000"},
}

# Entry 3 (NEW): "03" "实用命令 & 内置 Skills" "Claude Code 实用命令 · 内置 Skills 速查" at y=2360000
# This is the old "Web自动化 & Agent" entry that we repurpose as NEW "03"

# I need to find these text items in the XML and modify their container shapes.
# Strategy: Use regex to find shapes containing specific text, then modify.

# Let me try a completely different strategy: reconstruct the TOC by building
# new shape blocks for all 5 entries.

# From the XML, each entry has this structure:
# Circle: <p:sp> with <a:prstGeom prst="ellipse"> and <a:solidFill>
# Number: <p:sp> with <a:t>01</a:t>
# Title: <p:sp> with <a:t>AI 进化史速览</a:t>
# Subtitle: <p:sp> with <a:t>大模型发展历程...</a:t>

# I'll find the 01 circle by looking for ellipse at the right position
# and work from there.

# Actually, the absolute simplest approach:
# Since I know the exact text strings, I'll do search-and-replace on the WHOLE XML.
# For Y positions: I'll find the actual XML lines with the y values.

# Let me identify the exact XML for each Y value that needs changing:
# Old Y values in the file: 1463040, 1417320, 1737360, 2377440, 2331720, 2651760, 3291840, 3246120, 3566160, 4206240, 4160520, 4480560
# New Y values: 960000, 914280, 1828680, 1660000, 1614280, 2528680, 2360000, 2314280, 3228680, 3060000, 3014280, 3928680, 3760000, 3714280, 4628680

# Problem: some Y values might be shared between shapes in other slides!
# No wait, we're only editing slide2.xml.
# And off_y="1463040" might appear multiple times for circle and number text shapes.

# This approach is fragile but let me try it. If there are matching problems, I'll adjust.

# Replace Y positions for entry 1
# Circle: y=1463040 → 960000
replacements = {
    # Entry 1: circle + number
    'y="1463040"': 'y="960000"',  # circle and number
    'y="1417320"': 'y="914280"',   # title
    'y="1737360"': 'y="1828680"',  # subtitle

    # Entry 2:
    'y="2377440"': 'y="1660000"',  # circle and number
    'y="2331720"': 'y="1614280"',  # title
    'y="2651760"': 'y="2528680"',  # subtitle

    # Entry 3 (old, will be repurposed as NEW 03)
    'y="3291840"': 'y="2360000"',  # circle and number
    'y="3246120"': 'y="2314280"',  # title
    'y="3566160"': 'y="3228680"',  # subtitle

    # Entry 4 (old, will be 04 Web自动化)
    'y="4206240"': 'y="3060000"',  # circle and number
    'y="4160520"': 'y="3014280"',  # title
    'y="4480560"': 'y="3928680"',  # subtitle
}

# Check for uniqueness
y_values = ['1463040', '1417320', '1737360', '2377440', '2331720', '2651760', '3291840', '3246120', '3566160', '4206240', '4160520', '4480560']
for yv in y_values:
    count = xml.count(f'y="{yv}"')
    if count > 1:
        print(f"  WARNING: y={yv} appears {count} times!")
    elif count == 0:
        print(f"  WARNING: y={yv} not found!")

# Apply Y position changes
for old, new in replacements.items():
    count = xml.count(old)
    if count == 1:
        xml = xml.replace(old, new)
        print(f"  Replaced {old} → {new}")
    elif count > 1:
        print(f"  SKIPPED {old} (appears {count} times, would be ambiguous)")
    else:
        print(f"  SKIPPED {old} (not found)")

# Text changes (entry 3: old 03 text → NEW 03 text)
# <a:t>03</a:t> appears for both entry 3 number AND in other places!
# Let me be more specific

# The "03" in entry 3 is at the old y=3291840 position
# Since I moved Y to 2360000, I need to know which "<a:t>03</a:t>" to change
# The one at/around old y=3291840

# Let me do text replacements for entries 3 and 4
# Entry 3 (old 03 Web自动化 → becomes NEW 03 实用命令)
xml = xml.replace(
    '&amp; Agent</a:t>',
    '&amp; 内置 Skills</a:t>'
)
xml = xml.replace(
    'Web自动化 &amp; 内置 Skills</a:t>',
    '实用命令 &amp; 内置 Skills</a:t>'
)
# Actually, the old text is "Web自动化 & Agent" which was already in the file
# Let me find the exact text

# Looking at slide2.xml, the text is:
# 'Web自动化 &amp; Agent' in title
# And 'Superpowers 工作流' in entry 4 title

# I need to change:
# Entry 3 title: "Web自动化 &amp; Agent" → "实用命令 &amp; 内置 Skills"
# Entry 3 subtitle: "BrowserUse · Playwright · Agent Teams" → "Claude Code 实用命令 · 内置 Skills 速查"
# Entry 4 title: "Superpowers 工作流" → "Agent 约束(harness)"
# Entry 4 subtitle: "强制纪律的 AI 编程流程" → "记忆体系 · Superpowers · 强制工作流"

xml = xml.replace('Web自动化 &amp; Agent', '实用命令 &amp; 内置 Skills')
xml = xml.replace('BrowserUse · Playwright · Agent Teams', 'Claude Code 实用命令 · 内置 Skills 速查')
xml = xml.replace('Superpowers 工作流', 'Agent 约束(harness)')
xml = xml.replace('强制纪律的 AI 编程流程', '记忆体系 · Superpowers · 强制工作流')

# Change number text: "03" → "03" (stays same), "04" → "05"
# The "04" text for entry 4 number (at old y=4206240)
# and "04" might appear elsewhere
# Let me check how many "04" text elements there are
count_04 = xml.count('<a:t>04</a:t>')
count_03 = xml.count('<a:t>03</a:t>')
count_02 = xml.count('<a:t>02</a:t>')
count_01 = xml.count('<a:t>01</a:t>')
print(f"\n  Number counts: 01={count_01}, 02={count_02}, 03={count_03}, 04={count_04}")

if count_04 == 1:
    xml = xml.replace('<a:t>04</a:t>', '<a:t>05</a:t>')
    print("  Changed 04 → 05")
else:
    print(f"  WARNING: 04 appears {count_04} times, not changing")

write(2, xml)
print("  TOC updated!")
