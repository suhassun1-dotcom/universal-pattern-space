from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# Copper colours
COPPER_DARK  = RGBColor(0x7A, 0x3B, 0x1E)
COPPER_MID   = RGBColor(0xB5, 0x62, 0x1E)
COPPER       = RGBColor(0xC8, 0x78, 0x2A)
COPPER_PALE  = RGBColor(0xED, 0xD5, 0xB8)
COPPER_GHOST = RGBColor(0xFB, 0xF3, 0xEC)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK    = RGBColor(0x2C, 0x15, 0x08)
TEXT_MID     = RGBColor(0x5C, 0x33, 0x20)
GREEN        = RGBColor(0x2E, 0x7D, 0x4F)
RED          = RGBColor(0x9B, 0x23, 0x35)
AMBER        = RGBColor(0xB8, 0x86, 0x0B)

def set_cell_bg(cell, colour: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), str(colour))
    tcPr.append(shd)

def set_cell_border(cell, sides=('top',), colour='C8782A', sz=12):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for side in sides:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:color'), colour)
        tcBorders.append(el)

def set_no_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        tblBorders.append(el)
    tblPr.append(tblBorders)

def add_run(para, text, bold=False, colour=None, size=None, italic=False):
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    if colour:
        run.font.color.rgb = colour
    if size:
        run.font.size = Pt(size)
    return run

doc = Document()

# ── Page margins ──
sec = doc.sections[0]
sec.page_width  = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin    = Cm(0)
sec.bottom_margin = Cm(0)
sec.left_margin   = Cm(0)
sec.right_margin  = Cm(0)

# ── Default style ──
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(8)
style.font.color.rgb = TEXT_DARK

# ════════════════════════════════════════════════
# HEADER TABLE (2 cols)
# ════════════════════════════════════════════════
hdr = doc.add_table(rows=1, cols=2)
hdr.alignment = WD_TABLE_ALIGNMENT.LEFT
set_no_border(hdr)
hdr.columns[0].width = Cm(12)
hdr.columns[1].width = Cm(9)

# Left cell
lc = hdr.cell(0, 0)
set_cell_bg(lc, COPPER_DARK)
lc.width = Cm(12)
lp1 = lc.paragraphs[0]
lp1.paragraph_format.space_before = Pt(8)
lp1.paragraph_format.space_after  = Pt(2)
lp1.paragraph_format.left_indent  = Cm(0.5)
add_run(lp1, 'Espiner — Bag Strength Test Development', bold=True, colour=WHITE, size=13)
lp2 = lc.add_paragraph()
lp2.paragraph_format.space_before = Pt(0)
lp2.paragraph_format.space_after  = Pt(8)
lp2.paragraph_format.left_indent  = Cm(0.5)
add_run(lp2, 'Bag Clamp Fixture  |  Design Session Progress Report', colour=COPPER_PALE, size=8)

# Right cell
rc = hdr.cell(0, 1)
set_cell_bg(rc, COPPER_DARK)
rc.width = Cm(9)
rc.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
rp1 = rc.paragraphs[0]
rp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rp1.paragraph_format.space_before = Pt(8)
rp1.paragraph_format.space_after  = Pt(1)
rp1.paragraph_format.right_indent = Cm(0.4)
add_run(rp1, 'PREPARED BY', colour=COPPER_PALE, size=6)
rp2 = rc.add_paragraph()
rp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rp2.paragraph_format.space_before = Pt(0)
rp2.paragraph_format.space_after  = Pt(3)
rp2.paragraph_format.right_indent = Cm(0.4)
add_run(rp2, 'Suhas Sundaresh K — simple design people', bold=True, colour=WHITE, size=7.5)
rp3 = rc.add_paragraph()
rp3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rp3.paragraph_format.space_before = Pt(0)
rp3.paragraph_format.space_after  = Pt(1)
rp3.paragraph_format.right_indent = Cm(0.4)
add_run(rp3, 'PERIOD  |  MACHINE', colour=COPPER_PALE, size=6)
rp4 = rc.add_paragraph()
rp4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rp4.paragraph_format.space_before = Pt(0)
rp4.paragraph_format.space_after  = Pt(8)
rp4.paragraph_format.right_indent = Cm(0.4)
add_run(rp4, '16–24 March 2026  |  Mark-10 F755 EM', bold=True, colour=WHITE, size=7.5)

# ── Confidential bar ──
conf_tbl = doc.add_table(rows=1, cols=1)
set_no_border(conf_tbl)
conf_cell = conf_tbl.cell(0, 0)
set_cell_bg(conf_cell, COPPER_MID)
cp = conf_cell.paragraphs[0]
cp.paragraph_format.space_before = Pt(2)
cp.paragraph_format.space_after  = Pt(2)
cp.paragraph_format.left_indent  = Cm(0.5)
add_run(cp, 'CONFIDENTIAL  ·  ISO 13485:2016  ·  All materials SS 316 throughout  ·  Full material certs required',
        bold=True, colour=WHITE, size=6)

# ── Summary box ──
sum_tbl = doc.add_table(rows=1, cols=1)
set_no_border(sum_tbl)
sc = sum_tbl.cell(0, 0)
set_cell_bg(sc, COPPER_GHOST)
set_cell_border(sc, sides=('left',), colour='C8782A', sz=18)
sc.paragraphs[0].paragraph_format.space_before = Pt(5)
sc.paragraphs[0].paragraph_format.space_after  = Pt(0)
sc.paragraphs[0].paragraph_format.left_indent  = Cm(0.5)
sc.paragraphs[0].paragraph_format.right_indent = Cm(0.4)
sp = sc.paragraphs[0]
add_run(sp, 'A full design session was completed for the mechanical clamping fixture holding a ', colour=TEXT_MID, size=7.5)
add_run(sp, 'ripstop nylon conical bag', bold=True, colour=COPPER_DARK, size=7.5)
add_run(sp, ' during destructive compression testing at ', colour=TEXT_MID, size=7.5)
add_run(sp, '590 N', bold=True, colour=COPPER_DARK, size=7.5)
add_run(sp, '. All component geometries are defined, key design decisions are locked, and structural calculations are verified. The design is ready for Fusion 360 CAD finalisation — one primary measurement task remains before machining release.', colour=TEXT_MID, size=7.5)
sc.add_paragraph().paragraph_format.space_after = Pt(4)

# ════════════════════════════════════════════════
# BODY — 2 COLUMNS
# ════════════════════════════════════════════════
body = doc.add_table(rows=1, cols=2)
set_no_border(body)
body.columns[0].width = Cm(11.2)
body.columns[1].width = Cm(9.8)
bl = body.cell(0, 0)
br = body.cell(0, 1)
bl.width = Cm(11.2)
br.width = Cm(9.8)

def sec_title(cell, text, first=False):
    p = cell.add_paragraph() if not first else cell.paragraphs[0]
    if first:
        p.clear()
    else:
        p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.4)
    r = p.add_run(text.upper())
    r.bold = True
    r.font.size = Pt(7)
    r.font.color.rgb = COPPER_MID
    # bottom border on paragraph
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), 'C8782A')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

# ── LEFT: Section title ──
sec_title(bl, 'Accomplishments This Session', first=True)

accomplishments = [
    ('Bag geometry confirmed', ' — Clamp plane Ø67.8 mm, upper ref Ø73.2 mm, taper ~8.1° half-angle'),
    ('Clamping principle finalised', ' — Conical SS surfaces + 1 mm Shore 20A silicone liners. Dual grip: threaded rod (torque-controlled) + self-energising wedge at 8.1° self-locking boundary'),
    ('All 7 component dimensions defined', ' — Bottom plate, ×4 standoffs, top plate, split collar, support ring, silicone liners, closing rods'),
    ('Structural calculations verified', ' — All safety factors confirmed (see table)'),
    ('5 documents produced', ' — ENG-CLJIG-001 Rev E, weekly report, CAD dimensions ref, ring cross-sections, this summary'),
    ('9 hours effort', ' — 5 hrs CAD + 4 hrs design development'),
]
for bold_text, rest in accomplishments:
    p = bl.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Cm(0.7)
    add_run(p, bold_text, bold=True, colour=COPPER_DARK, size=7.5)
    add_run(p, rest, colour=TEXT_DARK, size=7.5)

# ── LEFT: Calculations table ──
sec_title(bl, 'Structural Calculations Verified')

calcs = [
    ('Standoff stiffness', 'k = 2,897 N/mm → δ = 0.19 mm at 560 N', '—', '✓'),
    ('Standoff fatigue',   '14.1 MPa bending vs 120 MPa endurance',   '8.5×', '✓'),
    ('Ring wall (min)',    '3.4 mm wall, thick-walled cylinder',        '9.5×', '✓'),
    ('Friction grip',     '697 N grip > 590 N test load',              '>1×',  '✓'),
    ('Dovetail root stub','16 mm — adequate',                           '—',   '✓'),
    ('Assembly clearance','2.5 mm each side — elastomer through bore',  '—',   '✓'),
]

ct = bl.add_table(rows=1+len(calcs), cols=4)
set_no_border(ct)
ct.columns[0].width = Cm(3.2)
ct.columns[1].width = Cm(5.5)
ct.columns[2].width = Cm(1.2)
ct.columns[3].width = Cm(1.0)

# Header
for i, h in enumerate(['CHECK', 'RESULT', 'SF', '']):
    c = ct.cell(0, i)
    set_cell_bg(c, COPPER)
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Cm(0.15)
    add_run(p, h, bold=True, colour=WHITE, size=6.5)

for row_i, (label, result, sf, status) in enumerate(calcs):
    bg = COPPER_GHOST if row_i % 2 == 1 else WHITE
    row = ct.rows[row_i + 1]
    for ci in range(4): set_cell_bg(row.cells[ci], bg)
    vals = [label, result, sf, status]
    colours = [COPPER_DARK, TEXT_DARK, TEXT_DARK, GREEN]
    bolds = [True, False, False, True]
    for ci, (val, col, bld) in enumerate(zip(vals, colours, bolds)):
        p = row.cells[ci].paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        p.paragraph_format.left_indent  = Cm(0.15)
        add_run(p, val, bold=bld, colour=col, size=7)

# ══════════════════════════════════
# RIGHT COLUMN
# ══════════════════════════════════
sec_title(br, 'Key Design Decisions Locked', first=True)

decisions = [
    ('Dovetail Retention', 'Through male dovetail rail (60°, 12 mm). Collar slides into plate — shoulders resist uplift at bag failure. Zero retention fasteners.'),
    ('Leadscrew Closing', 'Half A = fixed jaw. Half B driven by ×2 M8 rods. Guided by dovetail walls — cannot skew regardless of tightening sequence.'),
    ('Half A Adjustability', 'M6×0.75 fine-pitch screw + slotted lock bolts. Bore centreline alignable with machine axis. Position recorded in DHF per session.'),
    ('Bore Surface Finish', 'Ra 3.2–6.3 µm — deliberate roughness for silicone RTV mechanical key. As-machined, no finishing pass — reduces cost.'),
    ('Base Plate Fixing', '5-bolt pattern: 4× #10-32 UNF + 1× ½-20 UNC. Countersinks on top face. No Mark-10 modification required.'),
    ('Elastomer Cone — Rejected', 'Cone on TP to stabilise ring rejected. Ring already stabilised by clamped bag wall. Revisit after first test data only.'),
]

dt = br.add_table(rows=3, cols=2)
set_no_border(dt)
dt.columns[0].width = Cm(4.7)
dt.columns[1].width = Cm(4.7)

for idx, (title, body_text) in enumerate(decisions):
    row_i, col_i = divmod(idx, 2)
    c = dt.cell(row_i, col_i)
    set_cell_bg(c, COPPER_GHOST)
    set_cell_border(c, sides=('top',), colour='D4956A', sz=12)
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(1)
    p.paragraph_format.left_indent  = Cm(0.2)
    p.paragraph_format.right_indent = Cm(0.2)
    add_run(p, title, bold=True, colour=COPPER_DARK, size=7)
    p2 = c.add_paragraph()
    p2.paragraph_format.space_before = Pt(1)
    p2.paragraph_format.space_after  = Pt(4)
    p2.paragraph_format.left_indent  = Cm(0.2)
    p2.paragraph_format.right_indent = Cm(0.2)
    add_run(p2, body_text, colour=TEXT_MID, size=6.8)

# ── RIGHT: Pending table ──
sec_title(br, 'Pending Before Machining Release')

pending = [
    ('Bag taper angle',       'Measure 5+ bags at 4+ heights — all dims depend on this', 'BLOCKER', RED),
    ('Top plate dims',        'Confirm rectangle dimensions in Fusion 360',               'CAD',     AMBER),
    ('Half A fixing holes',   'Confirm positions — clear of dovetail & rod holes',        'CAD',     AMBER),
    ('M8 rod hole positions', 'Min 12 mm clearance above/below bore edges',               'CAD',     AMBER),
    ('Standoff bolt circle',  'Maximise spread within plate footprint',                   'CAD',     AMBER),
    ('Clamping torque',       'Establish from first test — suggested 5 Nm per rod',       'Test',    AMBER),
    ('Rod & fastener lengths','Calculate from assembled CAD model',                        'CAD',     AMBER),
]

pt = br.add_table(rows=1+len(pending), cols=3)
set_no_border(pt)
pt.columns[0].width = Cm(2.8)
pt.columns[1].width = Cm(5.2)
pt.columns[2].width = Cm(1.4)

for i, h in enumerate(['ITEM', 'ACTION REQUIRED', 'PRIORITY']):
    c = pt.cell(0, i)
    set_cell_bg(c, COPPER)
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Cm(0.15)
    add_run(p, h, bold=True, colour=WHITE, size=6.5)

for row_i, (item, action, priority, pcol) in enumerate(pending):
    bg = COPPER_GHOST if row_i % 2 == 1 else WHITE
    row = pt.rows[row_i + 1]
    for ci in range(3): set_cell_bg(row.cells[ci], bg)
    for ci, (val, col, bld) in enumerate(zip([item, action, priority],
                                              [COPPER_DARK, TEXT_DARK, pcol],
                                              [True, False, True])):
        p = row.cells[ci].paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        p.paragraph_format.left_indent  = Cm(0.15)
        add_run(p, val, bold=bld, colour=col, size=6.8)

# ════════════════════════════════════════════════
# BANNER
# ════════════════════════════════════════════════
ban = doc.add_table(rows=1, cols=2)
set_no_border(ban)
ban.columns[0].width = Cm(2.8)
ban.columns[1].width = Cm(18.2)

bc1 = ban.cell(0, 0)
set_cell_bg(bc1, COPPER_DARK)
bp1 = bc1.paragraphs[0]
bp1.paragraph_format.space_before = Pt(6)
bp1.paragraph_format.space_after  = Pt(6)
bp1.paragraph_format.left_indent  = Cm(0.4)
add_run(bp1, 'IMMEDIATE\nNEXT STEP', bold=True, colour=COPPER_PALE, size=6.5)

bc2 = ban.cell(0, 1)
set_cell_bg(bc2, COPPER_MID)
bp2 = bc2.paragraphs[0]
bp2.paragraph_format.space_before = Pt(6)
bp2.paragraph_format.space_after  = Pt(6)
bp2.paragraph_format.left_indent  = Cm(0.4)
bp2.paragraph_format.right_indent = Cm(0.4)
add_run(bp2, 'Measure 5+ bag samples at 4+ heights each', bold=True, colour=COPPER_PALE, size=7.5)
add_run(bp2, ' to confirm taper angle. This single task unblocks all remaining bore and ring dimensions and enables full machining release. Current 8.1° is derived from 2 measurements on 1 bag only — insufficient for production.', colour=WHITE, size=7.5)

# ════════════════════════════════════════════════
# FOOTER
# ════════════════════════════════════════════════
ft = doc.add_table(rows=1, cols=2)
set_no_border(ft)
ft.columns[0].width = Cm(13)
ft.columns[1].width = Cm(8)

fc1 = ft.cell(0, 0)
set_cell_bg(fc1, COPPER_DARK)
fp1 = fc1.paragraphs[0]
fp1.paragraph_format.space_before = Pt(3)
fp1.paragraph_format.space_after  = Pt(3)
fp1.paragraph_format.left_indent  = Cm(0.5)
add_run(fp1, 'Espiner — Bag Strength Test Development  |  Bag Clamp Fixture  |  16–24 March 2026', colour=COPPER_PALE, size=6)

fc2 = ft.cell(0, 1)
set_cell_bg(fc2, COPPER_DARK)
fp2 = fc2.paragraphs[0]
fp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
fp2.paragraph_format.space_before = Pt(3)
fp2.paragraph_format.space_after  = Pt(3)
fp2.paragraph_format.right_indent = Cm(0.4)
add_run(fp2, 'ENG-CLJIG-001 Rev E  |  simple design people  |  CONFIDENTIAL', bold=True, colour=COPPER_PALE, size=6)

doc.save('espiner-progress-report.docx')
print("Done.")
