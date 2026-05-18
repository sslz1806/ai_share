"""
修改 AI大矩阵分享.pptx：
1. 添加3个新幻灯片作为"03 实用命令 & 内置 Skills"章节
2. 将原03→04，04→05并改名
3. 更新TOC和所有页面编号
"""
import re
import shutil
from pathlib import Path

BASE = Path(r"C:\Users\20561\Desktop\不常用\tset_playwright")
UNPACKED = BASE / "unpacked_pptx"
SLIDES = UNPACKED / "ppt" / "slides"
SLIDES_RELS = SLIDES / "_rels"
PPT_RELS = UNPACKED / "ppt" / "_rels" / "presentation.xml.rels"
CONTENT_TYPES = UNPACKED / "[Content_Types].xml"
PRESENTATION = UNPACKED / "ppt" / "presentation.xml"


def read(path):
    return path.read_text(encoding="utf-8")


def write(path, content):
    path.write_text(content, encoding="utf-8")


def next_slide_number():
    """Get next available slide number."""
    existing = [int(m.group(1)) for f in SLIDES.glob("slide*.xml")
                if (m := re.match(r"slide(\d+)\.xml", f.name))]
    return max(existing) + 1


def add_to_content_types(slide_num):
    """Register new slide in [Content_Types].xml."""
    ct = read(CONTENT_TYPES)
    entry = f'<Override PartName="/ppt/slides/slide{slide_num}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    if f"/ppt/slides/slide{slide_num}.xml" not in ct:
        ct = ct.replace("</Types>", f"  {entry}\n</Types>")
        write(CONTENT_TYPES, ct)
        print(f"  [Content_Types] added slide{slide_num}.xml")


def add_to_presentation_rels(slide_num):
    """Register new slide in ppt/_rels/presentation.xml.rels. Returns the rId."""
    rels = read(PPT_RELS)
    rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    next_rid = max(rids) + 1
    rid = f"rId{next_rid}"
    entry = f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{slide_num}.xml"/>'
    if f"slides/slide{slide_num}.xml" not in rels:
        rels = rels.replace("</Relationships>", f"  {entry}\n</Relationships>")
        write(PPT_RELS, rels)
        print(f"  presentation.xml.rels added {rid} → slide{slide_num}.xml")
    return rid


def duplicate_slide(src_num, dst_num):
    """Duplicate a slide file and its rels."""
    shutil.copy2(SLIDES / f"slide{src_num}.xml", SLIDES / f"slide{dst_num}.xml")
    src_rels = SLIDES_RELS / f"slide{src_num}.xml.rels"
    dst_rels = SLIDES_RELS / f"slide{dst_num}.xml.rels"
    if src_rels.exists():
        shutil.copy2(src_rels, dst_rels)
        rels_content = read(dst_rels)
        rels_content = re.sub(
            r'\s*<Relationship[^>]*Type="[^"]*notesSlide"[^>]*/>\s*',
            "\n",
            rels_content,
        )
        write(dst_rels, rels_content)
    add_to_content_types(dst_num)
    rid = add_to_presentation_rels(dst_num)

    # Also add to slide layout rels
    slide_xml = read(SLIDES / f"slide{dst_num}.xml")
    layout_id = None
    m = re.search(r'<p:sldLayoutId[^>]*id="(\d+)"', slide_xml)
    if m:
        layout_id = m.group(1)

    return rid


def update_slide_order(new_order):
    """
    Update the <p:sldIdLst> in presentation.xml with new slide order.
    new_order: list of (slide_id_num, rid_string) tuples
    """
    pres = read(PRESENTATION)

    entries = []
    for slide_id, rid in new_order:
        entries.append(f'    <p:sldId id="{slide_id}" r:id="{rid}"/>')

    new_sldIdLst = '  <p:sldIdLst>\n' + '\n'.join(entries) + '\n  </p:sldIdLst>'
    pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', new_sldIdLst, pres, flags=re.DOTALL)
    write(PRESENTATION, pres)
    print("  presentation.xml <p:sldIdLst> updated")


def get_current_order():
    """Get current slide order from presentation.xml."""
    pres = read(PRESENTATION)
    matches = re.findall(r'<p:sldId[^>]*id="(\d+)"[^>]*r:id="(rId\d+)"', pres)
    return [(int(id_), rid) for id_, rid in matches]


def get_rid_for_slide(slide_num):
    """Find the rId for a given slide file."""
    rels = read(PPT_RELS)
    m = re.search(r'Id="(rId\d+)"[^>]*Target="slides/slide' + str(slide_num) + r'\.xml"', rels)
    if m:
        return m.group(1)
    return None


def update_text_in_xml(xml, replacements):
    """
    Update text content in slide XML.
    replacements: dict of {old_text: new_text}
    """
    for old, new in replacements.items():
        # Handle both regular text and XML-escaped text
        xml = xml.replace(old, new)
    return xml


def update_page_numbers(xml, total_old, total_new, offset=0):
    """
    Update page numbers like "X / 26" to "X+offset / total_new".
    Also handles case where the X part might be different.
    """
    def replace_page_num(m):
        num = int(m.group(1))
        new_num = num + offset
        return f'{new_num} / {total_new}'

    pattern = rf'(\d+) \/ {total_old}'
    return re.sub(pattern, replace_page_num, xml)


# =============================================
# MAIN
# =============================================

# Step 1: Get current slide order mapping
current_order = get_current_order()
print("Current slide order (id → file mapping via rels):")
for sid, rid in current_order:
    rels = read(PPT_RELS)
    m = re.search(f'{rid}"[^>]*Target="slides/(slide\\d+\\.xml)"', rels)
    file = m.group(1) if m else "?"
    print(f"  id={sid}, {rid} → {file}")

# Map the rels to slide files
def get_slide_file_for_rid(rid):
    rels = read(PPT_RELS)
    m = re.search(f'{rid}"[^>]*Target="slides/(slide\\d+\\.xml)"', rels)
    return m.group(1) if m else None

def get_slide_num_for_rid(rid):
    f = get_slide_file_for_rid(rid)
    if f:
        m = re.search(r'slide(\d+)', f)
        return int(m.group(1)) if m else None
    return None

print("\nMapping rId to slide numbers:")
mapping = {}
for sid, rid in current_order:
    snum = get_slide_num_for_rid(rid)
    mapping[rid] = snum
    print(f"  {rid} → slide{snum}.xml (TOC id={sid})")

# Step 2: Create 3 new slides (31, 32, 33)
print("\n=== Creating new slides ===")

# New slide 31: Section divider (duplicate slide6 - 02 divider)
print("Creating slide31 (section divider)...")
duplicate_slide(6, 31)

# New slide 32: Content - 实用命令速查 (duplicate slide7 - model comparison layout)
print("Creating slide32 (commands)...")
duplicate_slide(7, 32)

# New slide 33: Content - 内置 Skills (duplicate slide9 - VibeCoding layout)
print("Creating slide33 (skills)...")
duplicate_slide(9, 33)

# Step 3: Get new rIds for the new slides
rid31 = get_rid_for_slide(31)
rid32 = get_rid_for_slide(32)
rid33 = get_rid_for_slide(33)
print(f"  slide31 → {rid31}")
print(f"  slide32 → {rid32}")
print(f"  slide33 → {rid33}")

# Step 4: Reorder slides
# Current order (simplified):
# 1-13 (unchanged), then slide14, 15-20, 21-22, 23, 24-28, 29-30
#
# New order:
# 1-13 (unchanged), SLIDE31, SLIDE32, SLIDE33, slide14, 15-20, 21-22, 23, 24-28, 29-30
print("\n=== Reordering slides ===")

# Build new order list
new_order = []
# Keep first 13 slides (rId2-rId14 which map to slide1-slide13)
# We need to figure out which rIds map to slides 1-13
for sid, rid in current_order:
    snum = get_slide_num_for_rid(rid)
    if snum and snum <= 13:
        new_order.append((sid, rid))

# Insert 3 new slides with new sequential IDs
# Use the highest existing sid + 1, +2, +3
max_sid = max(sid for sid, _ in current_order)
new_sid1 = max_sid + 1
new_sid2 = max_sid + 2
new_sid3 = max_sid + 3
new_order.append((new_sid1, rid31))
new_order.append((new_sid2, rid32))
new_order.append((new_sid3, rid33))

# Add remaining slides (14-30)
for sid, rid in current_order:
    snum = get_slide_num_for_rid(rid)
    if snum and snum >= 14:
        new_order.append((sid, rid))

print(f"New order ({len(new_order)} slides):")
for i, (sid, rid) in enumerate(new_order):
    snum = get_slide_num_for_rid(rid)
    print(f"  Position {i+1}: id={sid}, {rid} → slide{snum}.xml")

# Update presentation.xml with new order
update_slide_order(new_order)

# Step 5: Edit slide XMLs
print("\n=== Editing slide XMLs ===")

# 5a: New slide31 - Section divider "03 实用命令 & 内置 Skills"
slide31_xml = read(SLIDES / "slide31.xml")
# Change "02" → "03"
# Change "实用干货分享" → "实用命令 & 内置 Skills"
# Change subtitle
slide31_xml = update_text_in_xml(slide31_xml, {
    "02": "03",
    "实用干货分享": "实用命令 & 内置 Skills",
    "模型选择 · 工具对比 · 实用技巧": "Claude Code 实用命令 · 内置 Skills 速查"
})
write(SLIDES / "slide31.xml", slide31_xml)
print("  slide31: section divider text updated")

# 5b: New slide32 - 实用命令速查
slide32_xml = read(SLIDES / "slide32.xml")

# First make sure we see what's in the model title
# The title shape should be at the top
# Let's identify and modify specific text

# Replace the title
slide32_xml = slide32_xml.replace(
    '模型选择与对比',
    'Claude Code 实用命令'
)

# Replace the section title and content
# Left column: 实用命令
slide32_xml = slide32_xml.replace(
    '🇨🇳 国内模型',
    '🖥️ 常用命令'
)

# Replace the model card content with command descriptions
slide32_xml = slide32_xml.replace(
    'DeepSeek V4 Pro / Flash',
    '/debug · /compact · /clear'
)
slide32_xml = slide32_xml.replace(
    '数学 & 推理能力很强，Agent 能力大幅提升。V4 Pro + Superpower ≈ 0.8× Opus 4.7，价格仅为 1/7',
    '/debug — 系统调试模式，查看运行状态'
)

slide32_xml = slide32_xml.replace(
    'MiniMax 2.7',
    '/model · /config'
)
slide32_xml = slide32_xml.replace(
    '最便宜的模型，适合处理自动化任务等固定流程，省钱首选',
    '/model — 切换 AI 模型  ·  /config — 系统设置'
)

slide32_xml = slide32_xml.replace(
    'GLM 5',
    'cc · Esc'
)
slide32_xml = slide32_xml.replace(
    'Agent 和代码能力强，但目前请求较慢、不太稳定',
    'cc — 系统配置快速切换（cc -switch）  ·  双 Esc — 中断 AI 输出'
)

# Right column: 实用技巧
slide32_xml = slide32_xml.replace(
    '🌍 国外模型',
    '💡 使用技巧'
)

slide32_xml = slide32_xml.replace(
    'Claude Opus 4.7',
    'Skills 调用'
)
slide32_xml = slide32_xml.replace(
    '工程 & Agent 能力最强。缺点：价格较高，有封号风险',
    '通过 /skill-name 直接调用内置 Skills'
)

slide32_xml = slide32_xml.replace(
    'GPT 5.5',
    '记忆系统'
)
slide32_xml = slide32_xml.replace(
    '自主规划与自我纠错强，思维发散。缺点：上下文较短（无 1M）',
    '~/.claude/memory/ 持久化记忆，跨会话保持上下文'
)

slide32_xml = slide32_xml.replace(
    '推荐方案：不对比，都要 → QShot',
    '推荐方案：善用 / 命令 · 配置自动化'
)
slide32_xml = slide32_xml.replace(
    '集成多个 AI 对话框，同一个问题同步问给多个 AI，直接对比输出内容',
    '结合 Skills + Memory + 自动化配置，最大化 Claude Code 效率'
)

write(SLIDES / "slide32.xml", slide32_xml)
print("  slide32: commands content updated")

# 5c: New slide33 - Claude Code 内置 Skills
slide33_xml = read(SLIDES / "slide33.xml")

slide33_xml = slide33_xml.replace(
    'VibeCoding',
    'Claude Code 内置 Skills'
)

slide33_xml = slide33_xml.replace(
    'VibeCoding 是一种 AI 辅助编程的新范式，强调"随状态走、随感觉流"，',
    'Claude Code 内置丰富的 Skill，涵盖办公、开发、内容创作等场景：'
)
slide33_xml = slide33_xml.replace(
    '用自然语言指导 AI 生成代码，极大降低编程门槛。',
    '通过 /skill-name 直接调用，或让 AI 自动匹配触发。'
)

# Left column: 办公类
slide33_xml = slide33_xml.replace(
    '✅ 优点',
    '📋 办公类 Skills'
)
slide33_xml = slide33_xml.replace(
    '开发效率极高，快速迭代',
    '/xlsx — 电子表格创建与编辑'
)
slide33_xml = slide33_xml.replace(
    '降低编程入门门槛',
    '/pptx — PowerPoint 演示文稿制作'
)
slide33_xml = slide33_xml.replace(
    '适合原型验证和快速实现',
    '/docx — Word 文档生成与修改'
)
slide33_xml = slide33_xml.replace(
    '建议探索完成后固化为 Skill 使用',
    '/pdf — PDF 文件读取与处理'
)

# Right column: 开发类
slide33_xml = slide33_xml.replace(
    '⚠️ 注意事项',
    '🛠️ 开发 & 内容类'
)
slide33_xml = slide33_xml.replace(
    '需要明确约束，否则 AI 容易"放飞"',
    '/frontend-design — 前端界面设计开发'
)
slide33_xml = slide33_xml.replace(
    '生成的代码需人工 review',
    '/webapp-testing — Web 应用测试'
)
slide33_xml = slide33_xml.replace(
    '配合 Harness/规范使用更佳',
    '/claude-api — Claude API 集成开发'
)
slide33_xml = slide33_xml.replace(
    '自动化测试与质量保障',
    '/internal-comms — 内部沟通文档撰写'
)

write(SLIDES / "slide33.xml", slide33_xml)
print("  slide33: skills content updated")

# Step 6: Edit existing slides
print("\n=== Editing existing slides ===")

# 6a: Slide 2 (TOC) - update section titles and numbers
slide2_xml = read(SLIDES / "slide2.xml")
# Current TOC:
# 01 AI 进化史速览
# 02 实用干货分享
# 03 Web 自动化 & Agent
# 04 Superpowers 工作流
#
# New TOC:
# 01 AI 进化史速览
# 02 实用干货分享
# 03 实用命令 & 内置 Skills
# 04 Web 自动化 & Agent
# 05 Agent 约束(harness)

# Update the TOC items
# Find and replace "03" and "04" entries
slide2_xml = slide2_xml.replace(
    '03',
    '04'
)
# Need to be careful with this - it also changes "13" in 13/26 etc.
# Actually, the TOC has specific text "Web 自动化 & Agent" associated with 03
# and "Superpowers 工作流" associated with 04
# Let me check what the TOC text looks like in the XML
slide2_xml = slide2_xml.replace(
    'Web 自动化 & Agent',
    'TERM04_PLACEHOLDER'
)
slide2_xml = slide2_xml.replace(
    'Superpowers 工作流',
    'TERM05_PLACEHOLDER'
)
# Now add 03 entry
# The TOC has text items like "01", "02", "03", "04"
# I need to insert a "03 实用命令 & 内置 Skills" entry
# and shift 03→04, 04→05

# Actually, the simpler approach: modify the text content
# Let me replace the exact strings

# Restore placeholders
slide2_xml = slide2_xml.replace(
    'TERM04_PLACEHOLDER',
    'Web 自动化 & Agent'
)
slide2_xml = slide2_xml.replace(
    'TERM05_PLACEHOLDER',
    '05'
)

print("  slide2: TOC - complex, doing targeted modifications...")

# Let me use a different approach - read the full slide2 content and make precise changes
slide2_xml = read(SLIDES / "slide2.xml")

# The TOC has these text elements:
# "01" "AI 进化史速览" "大模型发展历程 & 对工作的影响"
# "02" "实用干货分享" "模型对比 · 工具推荐 · VibeCoding"
# "03" "Web 自动化 & Agent" "BrowserUse · Playwright · Agent Teams"
# "04" "Superpowers 工作流" "强制纪律的 AI 编程流程"

# New content:
# "01" "AI 进化史速览"
# "02" "实用干货分享"
# "03" "实用命令 & 内置 Skills"
# "04" "Web 自动化 & Agent"
# "05" "Agent 约束(harness)"

# Since the TOC has 4 entries, I need to add a 5th. But that's complex with XML.
# Option: Replace the TOC text to show all 5 sections
# Or simpler: modify the existing 4 entries

# Actually, adding a 5th entry to the TOC requires adding new XML shapes, which is complex.
# Let me modify the existing entries:
# - Keep 01, 02 unchanged
# - Change 03 text to "实用命令 & 内置 Skills" and its subtitle
# - Change 04 text to "Agent 约束(harness)" and its subtitle

# Wait, but order-wise, the user wants:
# 01 AI进化史速览
# 02 实用干货分享
# 03 实用命令 & 内置 Skills  ← NEW
# 04 Web 自动化 & Agent   ← old 03
# 05 Agent 约束(harness)   ← merged

# But with only 4 TOC entry slots, I can either:
# 1. Redesign the TOC to have 5 entries
# 2. Drop one entry

# Let me check the TOC layout in detail
print("  slide2: Need to handle TOC carefully...")

# Let me check the actual shape structure of the TOC to see if I can add a 5th entry

# For now, let me modify the existing entries and handle the TOC later
# The page numbers need updating first

# 6b: Slide 14 - Change "03" to "04"
slide14_xml = read(SLIDES / "slide14.xml")
slide14_xml = slide14_xml.replace(
    '<a:t>03</a:t>',
    '<a:t>04</a:t>'
)
write(SLIDES / "slide14.xml", slide14_xml)
print("  slide14: 03→04 updated")

# 6c: Slide 23 - Change "04" to "05" and "Superpowers 工作流" to "Agent 约束(harness)"
slide23_xml = read(SLIDES / "slide23.xml")
slide23_xml = slide23_xml.replace(
    '<a:t>04</a:t>',
    '<a:t>05</a:t>'
)
slide23_xml = slide23_xml.replace(
    'Superpowers 工作流',
    'Agent 约束(harness)'
)
slide23_xml = slide23_xml.replace(
    '强制纪律的 AI 编程流程',
    '记忆体系 · Superpowers · 强制工作流'
)
write(SLIDES / "slide23.xml", slide23_xml)
print("  slide23: 04→05, Superpowers→Agent 约束(harness) updated")

# Step 7: Update all page numbers
# The old total was 26. New total = 30 + 3 new slides = 33
# Wait, the old total was from the original. Let me check.
# Actually the slide says "X / 26" so total pages = 26 originally?
# No, wait. The original had 30 slides/page but said "X / 26".
# Hmm, that doesn't match. Let me check a few slides.

# Let me check what page numbers look like
print("\n=== Checking page numbers ===")
sample_pattern = r'(\d+) / 26'
for snum in [1, 2, 3, 6, 10, 14, 23, 28, 29]:
    path = SLIDES / f"slide{snum}.xml"
    if path.exists():
        content = read(path)
        matches = re.findall(sample_pattern, content)
        if matches:
            print(f"  slide{snum}.xml: page numbers found: {matches}")
        else:
            # Check for other patterns
            matches = re.findall(r'(\d+)\s*/\s*(\d+)', content)
            if matches:
                print(f"  slide{snum}.xml: number/number patterns: {matches}")

# ^ The page numbers might not be directly "X / 26" - they might have fonts etc.
# Actually, looking at the XML structure, the text is inside <a:t> tags
# Let me check more carefully
print("\n=== Finding page number text in slides ===")
for snum in range(1, 31):
    path = SLIDES / f"slide{snum}.xml"
    if path.exists():
        content = read(path)
        # Find text matching slide number patterns
        for m in re.finditer(r'<a:t>(\d+\s*/\s*\d+)</a:t>', content):
            print(f"  slide{snum}.xml: '{m.group(1)}'")

# I need to understand: was the original total 30 slides but showing "/ 26"?
# Let me check slide 30
slide30 = read(SLIDES / "slide30.xml")
print("\nSlide 30 content (looking for page numbers):")
for m in re.finditer(r'<a:t>([^<]*\d+[^<]*)</a:t>', slide30):
    print(f"  '{m.group(1)}'")

# From the original analysis, slide 2 showed "2 / 26" which makes no sense for 30 slides.
# It seems page numbers were used on SOME slides only.
# Let me check which slides have page numbers
print("\n=== Slides with page numbers ===")
for snum in range(1, 31):
    path = SLIDES / f"slide{snum}.xml"
    if path.exists():
        content = read(path)
        for m in re.finditer(r'<a:t>(\d+)\s*/\s*(\d+)</a:t>', content):
            print(f"  slide{snum}.xml: '{m.group(1)} / {m.group(2)}'")
