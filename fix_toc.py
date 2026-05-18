"""
Fix TOC slide2:
1. Update all Y positions (circles, numbers, titles, subtitles) for 5 entries
2. Add 5th entry (04 Web 自动化 & Agent)
3. Fix text for all entries
4. Update page number on TOC
"""
import re
from pathlib import Path

SLIDES = Path(r"C:\Users\20561\Desktop\不常用\tset_playwright\unpacked_pptx\ppt\slides")
xml = (SLIDES / "slide2.xml").read_text(encoding="utf-8")

# 5 entries with their Y positions (spacing=685800, first at 1280160)
# Entry: num, title, subtitle, circle_y
entries = [
    ("01", "AI 进化史速览", "大模型发展历程 &amp; 对工作的影响", "1280160"),
    ("02", "实用干货分享", "模型对比 · 工具推荐 · VibeCoding", "1965960"),
    ("03", "实用命令 &amp; 内置 Skills", "Claude Code 实用命令 · 内置 Skills 速查", "2651760"),
    ("04", "Web 自动化 &amp; Agent", "BrowserUse · Playwright · Agent Teams", "3337560"),
    ("05", "Agent 约束(harness)", "记忆体系 · Superpowers · 强制工作流", "4023360"),
]

# Helper: compute title_y and subtitle_y from circle_y
def get_ys(circle_y):
    cy = int(circle_y)
    return (str(cy - 45720), str(cy + 274320))

# ===== STEP 1: Fix entries 1-4 (existing shapes) =====
# The current state: circles at old Y values, titles/subtitles at wrong Y values
# I need to overwrite all Y positions

# Old Y values that need changing (for circles = both circle and number shapes):
old_circle_ys = ["1463040", "2377440", "3291840", "4206240"]
new_circle_ys = ["1280160", "1965960", "2651760", "3337560"]

# For each entry, replace old Y with new Y (appears 2x: circle + number)
for old_y, new_y in zip(old_circle_ys, new_circle_ys):
    # Replace all occurrences of y="old_y" with y="new_y"
    count = xml.count(f'y="{old_y}"')
    if count == 2:
        xml = xml.replace(f'y="{old_y}"', f'y="{new_y}"')
        print(f"  Circle/number y={old_y} → {new_y} (×2)")
    else:
        print(f"  WARNING: y={old_y} appears {count} times (expected 2)")

# For titles and subtitles, the old values that were already changed
# are now at wrong positions. Let me recalculate.
# Old title Y values that were changed by fix_pptx.py:
old_title_ys = {
    "914280": "1271160",      # was 1417320→914280, should be 1280160-45720=1234440
    "1614280": "1918440",     # was 2331720→1614280, should be 1965960-45720=1920240
    "2314280": "2606040",     # was 3246120→2314280, should be 2651760-45720=2606040
    "3014280": "3291840",     # was 4160520→3014280, should be 3337560-45720=3291840
}
old_subtitle_ys = {
    "1828680": "1554480",     # was 1737360→1828680, should be 1280160+274320=1554480
    "2528680": "2240280",     # was 2651760→2528680, should be 1965960+274320=2240280
    "3228680": "2926080",     # was 3566160→3228680, should be 2651760+274320=2926080
    "3928680": "3611880",     # was 4480560→3928680, should be 3337560+274320=3611880
}

# Fix titles: replace wrong values with correct ones
for wrong_y, correct_y in old_title_ys.items():
    xml = xml.replace(f'y="{wrong_y}"', f'y="{correct_y}"')
    print(f"  Title y={wrong_y} → {correct_y}")

# Fix subtitles: replace wrong values with correct ones
for wrong_y, correct_y in old_subtitle_ys.items():
    xml = xml.replace(f'y="{wrong_y}"', f'y="{correct_y}"')
    print(f"  Subtitle y={wrong_y} → {correct_y}")


# ===== STEP 2: Add 5th entry (04 Web自动化) =====
# The current TOC has entries at indices: 01, 02, 03, 05
# Missing: 04 (Web自动化)
# I need to add 4 shapes between old-03(03) and old-04(05)

# The 5th entry (old-04, now 05) shapes are at the right positions.
# The 4th entry (04 Web自动化) needs to be inserted.

# Looking at the XML, entry 3 (03) ends around line 495 (subtitle close tag)
# and entry 4 (05) starts around line 496 (Shape 14 circle)

# I'll insert new shapes before the existing "05" shapes.
# Entry 4 needs:
# 1. Circle shape (ellipse)
# 2. Number text shape ("04")
# 3. Title text shape ("Web 自动化 & Agent")
# 4. Subtitle text shape ("BrowserUse · Playwright · Agent Teams")

# Y positions for entry 4:
c4 = "3337560"
t4 = "3291840"
s4 = "3611880"

# Build the 4 shapes as XML strings
circle4_xml = f'''      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="100" name="Shape 14b"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="640080" y="{c4}"/>
            <a:ext cx="548640" cy="548640"/>
          </a:xfrm>
          <a:prstGeom prst="ellipse">
            <a:avLst/>
          </a:prstGeom>
          <a:solidFill>
            <a:srgbClr val="2B5A8C"/>
          </a:solidFill>
          <a:ln/>
        </p:spPr>
      </p:sp>'''

number4_xml = f'''      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="101" name="Text 14b"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="640080" y="{c4}"/>
            <a:ext cx="548640" cy="548640"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
          <a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="ctr">
              <a:buNone/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1400" b="1" dirty="0">
                <a:solidFill>
                  <a:srgbClr val="FFFFFF"/>
                </a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
                <a:cs typeface="Arial" pitchFamily="34" charset="-120"/>
              </a:rPr>
              <a:t>04</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" sz="1400" dirty="0"/>
          </a:p>
        </p:txBody>
      </p:sp>'''

title4_xml = f'''      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="102" name="Text 14c"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="1463040" y="{t4}"/>
            <a:ext cx="5486400" cy="320040"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
          <a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0">
              <a:buNone/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1800" b="1" dirty="0">
                <a:solidFill>
                  <a:srgbClr val="1A202C"/>
                </a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
                <a:cs typeface="Arial" pitchFamily="34" charset="-120"/>
              </a:rPr>
              <a:t>Web 自动化 &amp; Agent</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" sz="1800" dirty="0"/>
          </a:p>
        </p:txBody>
      </p:sp>'''

subtitle4_xml = f'''      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="103" name="Text 14d"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="1463040" y="{s4}"/>
            <a:ext cx="5486400" cy="274320"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
          <a:noFill/>
          <a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0">
              <a:buNone/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1200" dirty="0">
                <a:solidFill>
                  <a:srgbClr val="4A5568"/>
                </a:solidFill>
                <a:latin typeface="Calibri" pitchFamily="34" charset="0"/>
                <a:ea typeface="Calibri" pitchFamily="34" charset="-122"/>
                <a:cs typeface="Calibri" pitchFamily="34" charset="-120"/>
              </a:rPr>
              <a:t>BrowserUse · Playwright · Agent Teams</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" sz="1200" dirty="0"/>
          </a:p>
        </p:txBody>
      </p:sp>'''

entry4_xml = "\n" + circle4_xml + "\n" + number4_xml + "\n" + title4_xml + "\n" + subtitle4_xml + "\n"

# Insert entry 4 before the "05" shapes
# Find the first "05" shape. The "05" number text is '<a:t>05</a:t>'
# within a shape that starts right before it.
# Actually, I'll look for the closing subtitle of entry 3, then insert after it.
# Entry 3 subtitle ends its shape with </p:sp>, right before entry 4 (05) circle starts.

# Let me find the exact insertion point
# The subtitle of entry 3 has text "Claude Code 实用命令 · 内置 Skills 速查"
# After its closing </p:sp>, the next shapes are entry 5 (05)
# I'll insert entry 4 shapes after the subtitle of entry 3

# Actually, let me insert right before the "05" circle shape
# The 05 circle starts with: <p:sp>\n        <p:nvSpPr>\n          <p:cNvPr id="16" name="Shape 14"/>
# Let me find this unique pattern

insertion_marker = '<p:cNvPr id="16" name="Shape 14"/>'
if insertion_marker in xml:
    # Find the <p:sp> that starts this shape
    # Go backwards from the marker to find the preceding <p:sp>
    idx = xml.index(insertion_marker)
    # Find the <p:sp> that occurs before this marker
    # The <p:sp> is on a line before the marker
    line_start = xml.rfind('\n', 0, idx)
    line_before = xml.rfind('\n', 0, line_start)
    # Check if between line_before and idx there's a <p:sp>
    before_marker = xml[line_before:idx]
    # Find the opening <p:sp> of the 05 circle
    psp_pos = before_marker.rfind('<p:sp>')
    if psp_pos >= 0:
        abs_pos = line_before + psp_pos
        # Insert entry 4 shapes before this position
        xml = xml[:abs_pos] + entry4_xml + xml[abs_pos:]
        print(f"\n  Inserted entry 4 (04 Web自动化) shapes before 05 entry")
    else:
        print("  WARNING: Could not find <p:sp> before 05 circle")
else:
    print("  WARNING: Could not find 05 circle marker")


# ===== STEP 3: Fix entry 5 (old 04, now 05) Y positions =====
# Entry 5 circle should be at 4023360, but it's at 4206240 (unchanged)
# The number is also at 4206240
c5 = "4023360"
count = xml.count(f'y="4206240"')
if count > 0:
    xml = xml.replace(f'y="4206240"', f'y="{c5}"')
    print(f"  Entry 5 circle/number: y=4206240 → {c5} (×{count})")

# Entry 5 title was moved to 3014280 by fix_pptx.py, should be 3977640
# Entry 5 subtitle was moved to 3928680, should be 4297680
xml = xml.replace('y="3014280"', 'y="3977640"')
xml = xml.replace('y="3928680"', 'y="4297680"')
print("  Entry 5 title: 3014280 → 3977640")
print("  Entry 5 subtitle: 3928680 → 4297680")


# ===== STEP 4: Verify the result =====
# Check that each entry's shapes have correct Y relationships
print("\n=== Verification ===")
entries_check = [
    ("01", "1280160", "1234440", "1554480"),
    ("02", "1965960", "1920240", "2240280"),
    ("03", "2651760", "2606040", "2926080"),
    ("04", "3337560", "3291840", "3611880"),
    ("05", "4023360", "3977640", "4297680"),
]

for num, exp_c, exp_t, exp_s in entries_check:
    c_count = xml.count(f'y="{exp_c}"')
    t_count = xml.count(f'y="{exp_t}"')
    s_count = xml.count(f'y="{exp_s}"')
    status = "✓" if (c_count >= 2 and t_count >= 1 and s_count >= 1) else "✗"
    print(f"  Entry {num}: circle={exp_c}(×{c_count}) title={exp_t}(×{t_count}) subtitle={exp_s}(×{s_count}) {status}")


# ===== STEP 5: Write the file =====
(SLIDES / "slide2.xml").write_text(xml, encoding="utf-8")
print("\n  slide2.xml written successfully!")
