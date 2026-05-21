#!/usr/bin/env python3
"""
Professional PowerPoint: AR BML Update | 21 Mei 2026
SMOOT + SWAP ENERGI brand colors - Blue/Navy theme
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

# ── Slide dimensions (16:9 widescreen) ──────────────────────
W = Inches(13.33)
H = Inches(7.5)

# ── Color Palette ────────────────────────────────────────────
NAVY       = RGBColor(0x00, 0x1F, 0x5B)   # Deep Navy
BLUE       = RGBColor(0x00, 0x5B, 0xB5)   # Primary Blue
BLUE_MID   = RGBColor(0x17, 0x7A, 0xDF)   # Mid Blue
BLUE_LT    = RGBColor(0xE8, 0xF2, 0xFF)   # Light Blue bg
CYAN       = RGBColor(0x00, 0xC2, 0xDE)   # Cyan Accent
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
BLACK      = RGBColor(0x1A, 0x1A, 0x2E)
LGRAY      = RGBColor(0xF4, 0xF7, 0xFA)
MGRAY      = RGBColor(0xB2, 0xBE, 0xC9)
DGRAY      = RGBColor(0x52, 0x64, 0x77)
TBL_HDR    = RGBColor(0x00, 0x3A, 0x8C)   # Table header dark blue
TBL_ALT    = RGBColor(0xD0, 0xE5, 0xF8)   # Table alt row
TBL_TOT    = RGBColor(0x00, 0x1F, 0x5B)   # Table total row
GREEN      = RGBColor(0x1E, 0x8B, 0x4C)
RED        = RGBColor(0xC0, 0x39, 0x2B)
ORANGE     = RGBColor(0xD3, 0x84, 0x0A)
GOLD       = RGBColor(0xF5, 0xC5, 0x18)
SWAP_CYAN  = RGBColor(0x00, 0xB9, 0xD8)   # SWAP brand accent
SMOOT_BLUE = RGBColor(0x00, 0x4A, 0xAD)   # SMOOT brand blue

# ── Helpers ──────────────────────────────────────────────────

def rp(n):
    return f"Rp {int(n):,}".replace(",", ".")


def rect(slide, x, y, w, h, fc, lc=None, lw=0):
    shp = slide.shapes.add_shape(1, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fc
    if lc:
        shp.line.color.rgb = lc
        shp.line.width = Pt(lw) if lw else Pt(0.5)
    else:
        shp.line.fill.background()
    return shp


def txt(slide, text, x, y, w, h,
        size=11, bold=False, italic=False, color=None,
        align=PP_ALIGN.LEFT, wrap=True, font="Segoe UI"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    return tb


def header(slide, title, subtitle=None, h=Inches(1.25)):
    rect(slide, 0, 0, W, h, NAVY)
    rect(slide, 0, h - Inches(0.07), W, Inches(0.07), CYAN)
    txt(slide, title, Inches(0.45), Inches(0.12), Inches(10), Inches(0.75),
        size=22, bold=True, color=WHITE)
    if subtitle:
        txt(slide, subtitle, Inches(0.45), Inches(0.82), Inches(12), Inches(0.38),
            size=11, color=RGBColor(0xAD, 0xD8, 0xFF))


def footer(slide, pg=None):
    text = "Confidential  |  AR BML Update – 21 Mei 2026"
    if pg:
        text += f"  |  {pg}"
    rect(slide, 0, H - Inches(0.32), W, Inches(0.32), NAVY)
    txt(slide, text, Inches(0.3), H - Inches(0.28), Inches(12.7), Inches(0.28),
        size=8, color=MGRAY, align=PP_ALIGN.CENTER)


def set_bg(slide, color=BLUE_LT):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color


def kpi_card(slide, x, y, w, h, label, value, sub, bg, val_color=GOLD):
    rect(slide, x, y, w, h, bg)
    rect(slide, x, y, Inches(0.1), h, val_color)
    txt(slide, label, x + Inches(0.18), y + Inches(0.15), w - Inches(0.25), Inches(0.42),
        size=10, bold=True, color=WHITE)
    txt(slide, value, x + Inches(0.18), y + Inches(0.55), w - Inches(0.25), Inches(0.8),
        size=19, bold=True, color=val_color)
    txt(slide, sub, x + Inches(0.18), y + Inches(1.28), w - Inches(0.25), Inches(0.42),
        size=8, italic=True, color=RGBColor(0xCF, 0xE8, 0xFF))


def add_table(slide, data, x, y, w, h,
              col_w=None, hdr_bg=None, font_sz=8.5, total_last=False):
    if hdr_bg is None:
        hdr_bg = TBL_HDR
    rows, cols = len(data), len(data[0])
    tbl = slide.shapes.add_table(rows, cols, x, y, w, h).table

    if col_w:
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = cw

    for r, row in enumerate(data):
        is_hdr = r == 0
        is_tot = total_last and r == rows - 1
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            if is_hdr:
                fc, tc, bold = hdr_bg, WHITE, True
            elif is_tot:
                fc, tc, bold = TBL_TOT, WHITE, True
            elif r % 2 == 1:
                fc, tc, bold = WHITE, BLACK, False
            else:
                fc, tc, bold = TBL_ALT, BLACK, False

            cell.fill.solid()
            cell.fill.fore_color.rgb = fc

            p = cell.text_frame.paragraphs[0]
            s = str(val)
            # Right-align currency/number columns (not first col, not header)
            if not is_hdr and c > 0 and (s.startswith("Rp") or
                    s.replace(".", "").replace(",", "").lstrip("-").isdigit()):
                p.alignment = PP_ALIGN.RIGHT
            else:
                p.alignment = PP_ALIGN.LEFT

            run = p.add_run()
            run.text = s
            run.font.name = "Segoe UI"
            run.font.size = Pt(font_sz)
            run.font.bold = bold
            run.font.color.rgb = tc

    return tbl


def section_title(slide, text, y=Inches(1.35)):
    rect(slide, Inches(0.4), y, Inches(12.5), Inches(0.4), BLUE)
    txt(slide, text, Inches(0.55), y + Inches(0.05), Inches(12), Inches(0.32),
        size=12, bold=True, color=WHITE)


# ═══════════════════════════════════════════════════════════════
# SLIDE 1 – TITLE
# ═══════════════════════════════════════════════════════════════
def slide_title(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, NAVY)

    # Decorative right panel
    rect(sl, Inches(9.8),  0, Inches(3.53), H, BLUE)
    rect(sl, Inches(9.95), 0, Inches(0.12), H, CYAN)

    # Top accent stripe
    rect(sl, 0, 0, W, Inches(0.07), CYAN)

    # Diagonal accent shapes
    rect(sl, Inches(9.8), Inches(4.8), Inches(3.53), Inches(2.7), BLUE_MID)
    rect(sl, 0, H - Inches(0.07), W, Inches(0.07), GOLD)

    # SMOOT badge
    rect(sl, Inches(0.5), Inches(1.2), Inches(3.2), Inches(0.65), SMOOT_BLUE)
    txt(sl, "SMOOT", Inches(0.55), Inches(1.26), Inches(3.1), Inches(0.55),
        size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # SWAP ENERGI badge
    rect(sl, Inches(4.0), Inches(1.2), Inches(3.8), Inches(0.65), SWAP_CYAN)
    txt(sl, "SWAP ENERGI", Inches(4.05), Inches(1.26), Inches(3.7), Inches(0.55),
        size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # Main title
    txt(sl, "UPDATE AR BML", Inches(0.5), Inches(2.25), Inches(9), Inches(1.4),
        size=54, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Cyan underline
    rect(sl, Inches(0.5), Inches(3.6), Inches(5.5), Inches(0.07), CYAN)

    # Subtitle
    txt(sl, "Laporan Accounts Receivable kepada BML",
        Inches(0.5), Inches(3.75), Inches(9), Inches(0.48),
        size=16, color=RGBColor(0xAD, 0xD8, 0xFF))
    txt(sl, "Per 21 Mei 2026",
        Inches(0.5), Inches(4.28), Inches(5), Inches(0.42),
        size=14, bold=True, color=GOLD)

    # Total hutang highlight box
    rect(sl, Inches(0.5), Inches(5.15), Inches(6.0), Inches(1.7), BLUE)
    rect(sl, Inches(0.5), Inches(5.15), Inches(0.14), Inches(1.7), GOLD)
    txt(sl, "TOTAL TAGIHAN AR BML",
        Inches(0.72), Inches(5.28), Inches(5.6), Inches(0.42),
        size=10, bold=True, color=RGBColor(0xCF, 0xE8, 0xFF))
    txt(sl, rp(3_626_894_700),
        Inches(0.72), Inches(5.68), Inches(5.6), Inches(0.85),
        size=30, bold=True, color=GOLD)
    txt(sl, "SMOOT Motor  +  SWAP Energi  +  SMOOT After Sales",
        Inches(0.72), Inches(6.5), Inches(5.6), Inches(0.3),
        size=8, color=RGBColor(0xAD, 0xD8, 0xFF))

    # Right panel decorative text
    txt(sl, "AR", Inches(10.0), Inches(1.5), Inches(3.2), Inches(2.0),
        size=96, bold=True, color=RGBColor(0x00, 0x3A, 0x8C), align=PP_ALIGN.CENTER)
    txt(sl, "BML", Inches(10.0), Inches(3.2), Inches(3.2), Inches(1.2),
        size=48, bold=True, color=BLUE_MID, align=PP_ALIGN.CENTER)
    txt(sl, "2 0 2 6", Inches(10.0), Inches(4.4), Inches(3.2), Inches(0.6),
        size=22, bold=True, color=CYAN, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════
# SLIDE 2 – EXECUTIVE SUMMARY
# ═══════════════════════════════════════════════════════════════
def slide_summary(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "RINGKASAN EKSEKUTIF", "AR BML Update  |  21 Mei 2026")
    footer(sl, "1 / 12")

    # KPI cards
    cards = [
        ("INVOICE AR BML",        rp(3_626_894_700), "3 Entitas: Motor · Energi · After Sales", NAVY,     GOLD),
        ("TAGIHAN CUT-OFF APR'26", rp(1_976_747_940), "Percepatan & End Contract + De Sultan",   BLUE,     CYAN),
        ("OUTSTANDING SWAP-BML",   rp(3_688_669_071), "BBG De Sultan + BBG Tempur + Sewa Baterai", BLUE_MID, WHITE),
    ]
    for i, (lbl, val, sub, bg, vc) in enumerate(cards):
        kpi_card(sl, Inches(0.42 + i * 4.3), Inches(1.38), Inches(4.1), Inches(1.85), lbl, val, sub, bg, vc)

    # Summary table
    section_title(sl, "RINCIAN INVOICE AR BML (Total Hutang BML: Rp 3.626.894.700)", y=Inches(3.45))

    data = [
        ["Entitas",         "No Invoice / Ref",    "Keterangan",                           "Total Invoice"],
        ["SMOOT MOTOR",     "SO-2025/11-0216",     "90 Unit Smoot Zuzu OTR (PO-2511-0010)", rp(1_476_000_000)],
        ["SWAP ENERGI",     "SI-2026/04-0064 etc", "150 Unit LFP Battery Pack 51.2V 16AH",  rp(180_000_000)],
        ["SMOOT AFTER SALES","Pool Grab + SC + Maint", "Pool Grab / SC / Maintenance / Parts", rp(1_970_894_700)],
        ["TOTAL HUTANG BML","",                    "",                                      rp(3_626_894_700)],
    ]
    cw = [Inches(2.5), Inches(2.5), Inches(4.5), Inches(2.9)]
    add_table(sl, data, Inches(0.42), Inches(3.92), Inches(12.46), Inches(2.95),
              col_w=cw, font_sz=9.5, total_last=True)


# ═══════════════════════════════════════════════════════════════
# SLIDE 3 – INVOICE AR: SMOOT MOTOR
# ═══════════════════════════════════════════════════════════════
def slide_smoot_motor(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "INVOICE AR BML – SMOOT MOTOR", "Penjualan Unit Kendaraan Listrik ke BML")
    footer(sl, "2 / 12")

    # Company badge
    rect(sl, Inches(0.42), Inches(1.38), Inches(4.0), Inches(0.55), SMOOT_BLUE)
    txt(sl, "SMOOT MOTOR", Inches(0.5), Inches(1.44), Inches(3.9), Inches(0.44),
        size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Summary box
    rect(sl, Inches(4.8), Inches(1.38), Inches(4.5), Inches(1.4), NAVY)
    rect(sl, Inches(4.8), Inches(1.38), Inches(0.12), Inches(1.4), GOLD)
    txt(sl, "TOTAL TAGIHAN", Inches(5.0), Inches(1.48), Inches(4.2), Inches(0.4),
        size=11, bold=True, color=RGBColor(0xCF, 0xE8, 0xFF))
    txt(sl, rp(1_476_000_000), Inches(5.0), Inches(1.85), Inches(4.2), Inches(0.75),
        size=26, bold=True, color=GOLD)

    # Info box
    rect(sl, Inches(9.6), Inches(1.38), Inches(3.3), Inches(1.4), BLUE)
    txt(sl, "90 Unit Smoot Zuzu OTR\nElectric Motorcycle\nPO-2511-0010",
        Inches(9.75), Inches(1.48), Inches(3.05), Inches(1.2),
        size=11, bold=False, color=WHITE)

    section_title(sl, "DETAIL INVOICE", y=Inches(3.0))

    data = [
        ["No. Sales Order", "No. PO",      "Deskripsi",              "Qty",    "Total Invoice"],
        ["SO-2025/11-0216", "PO-2511-0010","Smoot Zuzu OTR (Unit)", "90 Unit", rp(1_476_000_000)],
        ["TOTAL",           "",             "",                       "",        rp(1_476_000_000)],
    ]
    cw = [Inches(2.5), Inches(2.2), Inches(3.8), Inches(1.5), Inches(2.4)]
    add_table(sl, data, Inches(0.42), Inches(3.47), Inches(12.46), Inches(1.6),
              col_w=cw, font_sz=10, total_last=True)

    # Note
    rect(sl, Inches(0.42), Inches(5.2), Inches(12.46), Inches(1.2), RGBColor(0xFF, 0xF3, 0xCD))
    rect(sl, Inches(0.42), Inches(5.2), Inches(0.1), Inches(1.2), ORANGE)
    txt(sl, "CATATAN:",
        Inches(0.6), Inches(5.28), Inches(11.8), Inches(0.35), size=9, bold=True, color=ORANGE)
    txt(sl, "Invoice ini merupakan tagihan atas penjualan 90 unit Smoot Zuzu OTR kepada BML.\n"
        "Status tagihan masih Outstanding (Accounts Receivable) dan menunggu pembayaran dari BML.",
        Inches(0.6), Inches(5.62), Inches(11.8), Inches(0.7), size=9, color=BLACK)


# ═══════════════════════════════════════════════════════════════
# SLIDE 4 – INVOICE AR: SWAP ENERGI
# ═══════════════════════════════════════════════════════════════
def slide_swap_energi(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "INVOICE AR BML – SWAP ENERGI", "Penjualan LFP Battery Pack ke BML")
    footer(sl, "3 / 12")

    # Company badge
    rect(sl, Inches(0.42), Inches(1.38), Inches(4.5), Inches(0.55), SWAP_CYAN)
    txt(sl, "SWAP ENERGI", Inches(0.5), Inches(1.44), Inches(4.4), Inches(0.44),
        size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # KPI boxes
    kpis = [
        ("Total Invoice", rp(180_000_000), NAVY, GOLD),
        ("Total Unit",    "150 Unit", BLUE, CYAN),
        ("Jenis Baterai", "LFP 51.2V 16AH", BLUE_MID, WHITE),
    ]
    for i, (lbl, val, bg, vc) in enumerate(kpis):
        x = Inches(5.0 + i * 2.8)
        rect(sl, x, Inches(1.38), Inches(2.6), Inches(1.1), bg)
        txt(sl, lbl, x + Inches(0.12), Inches(1.46), Inches(2.4), Inches(0.32),
            size=9, bold=True, color=RGBColor(0xCF, 0xE8, 0xFF))
        txt(sl, val, x + Inches(0.12), Inches(1.75), Inches(2.4), Inches(0.6),
            size=16, bold=True, color=vc)

    section_title(sl, "DETAIL INVOICE BATTERY PACK", y=Inches(2.72))

    data = [
        ["No. Sales Invoice", "No. PO",        "Deskripsi",                            "Qty",     "Total Invoice"],
        ["SI-2026/04-0064",   "PO-2603-0014",  "LFP Battery Pack Type A & B 51.2V 16AH","75 Unit", rp(90_000_000)],
        ["SI-2026/04-0131",   "PO-260/3-0015", "LFP Battery Pack Type B 51.2V 16AH",   "63 Unit", rp(75_600_000)],
        ["SI-2026/05-0035",   "PO-260/3-0015", "LFP Battery Pack Type A 51.2V 16AH",   "12 Unit", rp(14_400_000)],
        ["TOTAL",             "",               "",                                      "150 Unit", rp(180_000_000)],
    ]
    cw = [Inches(2.2), Inches(2.0), Inches(4.2), Inches(1.6), Inches(2.3)]
    add_table(sl, data, Inches(0.42), Inches(3.18), Inches(12.46), Inches(2.9),
              col_w=cw, font_sz=10, total_last=True)

    # Spec note
    rect(sl, Inches(0.42), Inches(6.25), Inches(12.46), Inches(0.9), RGBColor(0xE3, 0xF2, 0xFF))
    rect(sl, Inches(0.42), Inches(6.25), Inches(0.1), Inches(0.9), BLUE)
    txt(sl, "SPESIFIKASI: LFP Battery Pack 51.2V 16AH (Lithium Iron Phosphate)  |  "
        "Type A & Type B  |  Digunakan untuk armada Smoot Tempur & De Sultan",
        Inches(0.6), Inches(6.35), Inches(11.8), Inches(0.65), size=9, color=NAVY)


# ═══════════════════════════════════════════════════════════════
# SLIDE 5 – INVOICE AR: SMOOT AFTER SALES
# ═══════════════════════════════════════════════════════════════
def slide_aftersales(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "INVOICE AR BML – SMOOT AFTER SALES", "Layanan Purna Jual, Maintenance & Sparepart")
    footer(sl, "4 / 12")

    # KPI
    kpi_card(sl, Inches(0.42), Inches(1.38), Inches(4.0), Inches(1.7),
             "TOTAL AFTER SALES", rp(1_970_894_700), "5 Kategori Layanan", NAVY, GOLD)

    # Breakdown bars (visual)
    breakdown = [
        ("Biaya Maintenance",    1_135_300_000, BLUE),
        ("Pool Grab",              399_613_660, BLUE_MID),
        ("Service Center",         288_450_040, SWAP_CYAN),
        ("Project End Contract",   110_072_000, NAVY),
        ("Sparepart",               37_459_000, DGRAY),
    ]
    bar_x, bar_y = Inches(4.7), Inches(1.38)
    total_as = 1_970_894_700
    for i, (lbl, amt, color) in enumerate(breakdown):
        pct = amt / total_as
        bar_w = Inches(7.5 * pct)
        rect(sl, bar_x, bar_y + Inches(i * 0.32), bar_w, Inches(0.28), color)
        txt(sl, f"{lbl}  –  {rp(amt)}",
            bar_x + Inches(0.1), bar_y + Inches(i * 0.32 + 0.04), Inches(7.4), Inches(0.24),
            size=8.5, bold=False, color=WHITE if pct > 0.1 else BLACK)

    section_title(sl, "DETAIL INVOICE AFTER SALES", y=Inches(3.2))

    data = [
        ["Kategori Layanan",      "Keterangan",                           "Total Invoice"],
        ["Pool Grab",             "Pengelolaan pool unit Grab",            rp(399_613_660)],
        ["Service Center",        "Biaya operasional service center",      rp(288_450_040)],
        ["Biaya Maintenance",     "Biaya pemeliharaan armada",             rp(1_135_300_000)],
        ["Project End Contract",  "Biaya penyelesaian kontrak project",    rp(110_072_000)],
        ["Sparepart",             "Suku cadang & komponen",                rp(37_459_000)],
        ["TOTAL AFTER SALES",     "",                                      rp(1_970_894_700)],
    ]
    cw = [Inches(3.2), Inches(6.0), Inches(3.2)]
    add_table(sl, data, Inches(0.42), Inches(3.65), Inches(12.46), Inches(3.2),
              col_w=cw, font_sz=10, total_last=True)


# ═══════════════════════════════════════════════════════════════
# SLIDE 6 – TAGIHAN CUT-OFF APRIL 2026: OVERVIEW
# ═══════════════════════════════════════════════════════════════
def slide_cutoff_overview(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "TAGIHAN CUT-OFF APRIL 2026 – OVERVIEW", "Data Update Tagihan | Status per 1 Mei 2026")
    footer(sl, "5 / 12")

    # Status legend
    for i, (lbl, col, desc) in enumerate([
        ("SIAP BAYAR",    GREEN,  "Tagihan terverifikasi, siap dibayarkan"),
        ("CREDITED",      BLUE,   "Faktur Pajak sudah dikreditkan"),
        ("CANCELED",      RED,    "Faktur Pajak dibatalkan / tidak diakui"),
    ]):
        x = Inches(0.42 + i * 4.3)
        rect(sl, x, Inches(1.38), Inches(4.1), Inches(0.42), col)
        txt(sl, f"{lbl}  –  {desc}", x + Inches(0.12), Inches(1.44), Inches(3.9), Inches(0.32),
            size=9, bold=True, color=WHITE)

    section_title(sl, "RINGKASAN 4 KATEGORI TAGIHAN CUT-OFF", y=Inches(1.98))

    cat_data = [
        ["Kategori Tagihan",                  "Jumlah\nInvoice", "Amount (- PPh)",      "Total Invoice",       "Status"],
        ["Percepatan End Contract Tempur",    "8 inv",           rp(789_482_640),        rp(800_458_740),       "Siap Bayar"],
        ["End Contract Tempur",               "7 inv",           rp(17_527_200),         rp(17_848_800),        "Siap Bayar"],
        ["Rental De Sultan (Grab)",           "1 inv",           rp(1_137_567_600),      rp(1_158_440_400),     "Siap Bayar"],
        ["Tagihan IDLE (Credited – No.1-5)",  "5 inv",           rp(969_303_060),        "-",                   "Proses Cek"],
        ["Tagihan IDLE (Canceled – No.6-11)", "6 inv",           rp(1_410_409_290),      "-",                   "TIDAK DIAKUI"],
        ["TOTAL SIAP BAYAR",                  "",                rp(1_944_577_440),      rp(1_976_747_940),     ""],
    ]
    cw = [Inches(3.5), Inches(1.2), Inches(2.5), Inches(2.5), Inches(2.6)]
    add_table(sl, cat_data, Inches(0.42), Inches(2.44), Inches(12.46), Inches(4.0),
              col_w=cw, font_sz=9.5, total_last=True)

    # Note
    rect(sl, Inches(0.42), Inches(6.58), Inches(12.46), Inches(0.55), RGBColor(0xFF, 0xF3, 0xCD))
    rect(sl, Inches(0.42), Inches(6.58), Inches(0.1), Inches(0.55), ORANGE)
    txt(sl, "CATATAN: Tagihan IDLE No.6-11 senilai Rp 1.410.409.290 berstatus CANCELED dan tidak diakui karena tidak sesuai MoM yang berlaku dari hasil meeting.",
        Inches(0.6), Inches(6.65), Inches(11.8), Inches(0.42), size=8.5, color=BLACK)


# ═══════════════════════════════════════════════════════════════
# SLIDE 7 – PERCEPATAN END CONTRACT TEMPUR
# ═══════════════════════════════════════════════════════════════
def slide_percepatan(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "PERCEPATAN END CONTRACT – UNIT TEMPUR",
           "Sewa Smoot Tempur ex Grab | Cut-Off April 2026")
    footer(sl, "6 / 12")

    kpi_card(sl, Inches(0.42), Inches(1.38), Inches(3.8), Inches(1.5),
             "TOTAL INVOICE", rp(800_458_740), "8 Invoices | Status: Siap Bayar", NAVY, GOLD)
    kpi_card(sl, Inches(4.42), Inches(1.38), Inches(3.8), Inches(1.5),
             "AMOUNT (Net PPh)", rp(789_482_640), "Setelah potongan PPh", BLUE, CYAN)

    rect(sl, Inches(8.62), Inches(1.38), Inches(4.3), Inches(1.5), GREEN)
    txt(sl, "STATUS: SIAP BAYAR", Inches(8.72), Inches(1.48), Inches(4.1), Inches(0.4),
        size=10, bold=True, color=WHITE)
    txt(sl, "Semua Faktur Pajak\nBerstatus CREDITED\nUpdate: 1 Mei 2026",
        Inches(8.72), Inches(1.85), Inches(4.1), Inches(0.9), size=9.5, color=WHITE)

    section_title(sl, "DETAIL 8 INVOICE PERCEPATAN END CONTRACT", y=Inches(3.05))

    data = [
        ["No", "No Invoice",    "Tgl Invoice",  "Keterangan",                              "Amt (- PPh)",      "Total Invoice"],
        ["1",  "INV-2512-0080", "01 Dec 2025",  "120 unit Tempur ex Grab Surabaya",         rp(89_543_500),     rp(87_676_680)],
        ["2",  "INV-2601-0019", "01 Jan 2026",  "128 unit Tempur ex Grab JKT & Surabaya",   rp(88_800_120),     rp(90_429_480)],
        ["3",  "INV-2601-0133", "29 Jan 2026",  "50 unit Tempur ex Grab Jogja (29-31 Jan)",  rp(3_678_750),      rp(3_746_250)],
        ["4",  "INV-2602-0037", "01 Feb 2026",  "170 unit Tempur ex Grab Sub & Jogja Feb",   rp(112_099_960),    rp(114_156_840)],
        ["5",  "INV-2602-0103", "13 Feb 2026",  "100 unit Tempur ex Grab Surabaya 13-28 Feb",rp(38_368_000),     rp(39_072_000)],
        ["6",  "INV-2603-0021", "01 Mar 2026",  "270 unit Tempur ex Grab Sub & Jogja Maret", rp(198_448_670),    rp(202_089_930)],
        ["7",  "INV-2603-0093", "03 Mar 2026",  "47 unit Tempur ex Grab Sub 03-31 Maret",    rp(32_684_740),     rp(33_284_460)],
        ["8",  "INV-2604-0023", "01 Apr 2026",  "317 unit Tempur ex Grab Surabaya April",    rp(225_858_900),    rp(230_003_100)],
        ["",   "TOTAL",         "",             "",                                           rp(789_482_640),    rp(800_458_740)],
    ]
    cw = [Inches(0.4), Inches(1.6), Inches(1.2), Inches(4.5), Inches(2.2), Inches(2.2)]
    add_table(sl, data, Inches(0.42), Inches(3.5), Inches(12.46), Inches(3.65),
              col_w=cw, font_sz=8, total_last=True)


# ═══════════════════════════════════════════════════════════════
# SLIDE 8 – END CONTRACT TEMPUR + DE SULTAN
# ═══════════════════════════════════════════════════════════════
def slide_endcontract_desultan(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "END CONTRACT TEMPUR & RENTAL DE SULTAN",
           "Tagihan Sewa Unit Grab | Cut-Off April 2026")
    footer(sl, "7 / 12")

    # ── End Contract Tempur ──
    section_title(sl, "END CONTRACT GRAB – UNIT TEMPUR  |  Total Invoice: Rp 17.848.800", y=Inches(1.38))

    data_ec = [
        ["No", "No Invoice",    "Tgl Invoice",  "Keterangan",                             "Total Invoice"],
        ["1",  "INV-2509-0032", "01 Sep 2025",  "8 unit Tempur ex Grab – Sep 2025",        rp(1_776_000)],
        ["2",  "INV-2511-0023", "01 Nov 2025",  "8 unit Tempur ex Grab – Nov 2025",        rp(2_664_000)],
        ["3",  "INV-2511-0175", "28 Nov 2025",  "8 unit Tempur ex Grab – Okt 2025",        rp(2_752_800)],
        ["4",  "INV-2512-0022", "01 Dec 2025",  "8 unit Tempur ex Grab – Des 2025",        rp(2_752_800)],
        ["5",  "INV-2602-0022", "01 Feb 2026",  "8 unit Tempur JKT – 01-28 Feb 2026",      rp(2_486_400)],
        ["6",  "INV-2603-0020", "01 Mar 2026",  "8 unit Tempur JKT – 01-31 Mar 2026",      rp(2_752_800)],
        ["7",  "INV-2604-0021", "01 Apr 2026",  "8 unit Tempur JKT – 01-30 Apr 2026",      rp(2_664_000)],
        ["",   "TOTAL",         "",             "",                                          rp(17_848_800)],
    ]
    cw2 = [Inches(0.4), Inches(1.6), Inches(1.2), Inches(5.5), Inches(2.5)]
    add_table(sl, data_ec, Inches(0.42), Inches(1.82), Inches(10.5), Inches(2.98),
              col_w=cw2, font_sz=8, total_last=True)

    # Status badge end contract
    rect(sl, Inches(11.1), Inches(1.82), Inches(1.78), Inches(2.98), GREEN)
    txt(sl, "SIAP\nBAYAR\n\nCREDITED",
        Inches(11.15), Inches(2.3), Inches(1.6), Inches(2.0),
        size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # ── De Sultan ──
    section_title(sl, "RENTAL DE SULTAN GRAB  |  Total Invoice: Rp 1.158.440.400", y=Inches(5.0))

    data_ds = [
        ["No Invoice",     "Tgl Invoice",  "Keterangan",                               "Amt (- PPh)",      "Total Invoice"],
        ["INV-2604-0024",  "01 Apr 2026",  "1.338 unit De Sultan – 01-30 April 2026",  rp(1_137_567_600),  rp(1_158_440_400)],
        ["TOTAL",          "",             "",                                           rp(1_137_567_600),  rp(1_158_440_400)],
    ]
    cw3 = [Inches(2.0), Inches(1.4), Inches(5.0), Inches(2.2), Inches(2.2)]
    add_table(sl, data_ds, Inches(0.42), Inches(5.45), Inches(12.46), Inches(1.55),
              col_w=cw3, font_sz=9, total_last=True)


# ═══════════════════════════════════════════════════════════════
# SLIDE 9 – TAGIHAN IDLE
# ═══════════════════════════════════════════════════════════════
def slide_idle(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "TAGIHAN IDLE & OVER SLA – UNIT TEMPUR",
           "Pembayaran Sewa Grab Unit Mangkrak | Cut-Off April 2026")
    footer(sl, "8 / 12")

    # Status summary
    kpi_card(sl, Inches(0.42), Inches(1.38), Inches(4.0), Inches(1.35),
             "CREDITED (No.1-5)", rp(969_303_060), "Masih proses pengecekan Aftersales", BLUE, CYAN)
    kpi_card(sl, Inches(4.62), Inches(1.38), Inches(4.0), Inches(1.35),
             "CANCELED (No.6-11)", rp(1_410_409_290), "TIDAK DIAKUI – Tidak sesuai MoM", RED, WHITE)

    rect(sl, Inches(8.82), Inches(1.38), Inches(4.1), Inches(1.35), ORANGE)
    txt(sl, "TOTAL DITAGIHKAN", Inches(8.92), Inches(1.48), Inches(3.9), Inches(0.38),
        size=10, bold=True, color=WHITE)
    txt(sl, rp(2_379_712_350), Inches(8.92), Inches(1.83), Inches(3.9), Inches(0.65),
        size=20, bold=True, color=WHITE)
    txt(sl, "(11 Invoices)", Inches(8.92), Inches(2.45), Inches(3.9), Inches(0.22),
        size=9, color=WHITE, italic=True)

    section_title(sl, "DETAIL 11 INVOICE TAGIHAN IDLE", y=Inches(2.9))

    idle_data = [
        ["No", "No Invoice",    "Periode",  "Keterangan",                              "Total Invoice", "FP Status"],
        ["1",  "INV-2504-0085", "Apr 2025", "141 JKT + 88 Bali – Mar 2025",           rp(155_444_400), "CREDITED"],
        ["2",  "INV-2505-0110", "May 2025", "127 JKT + 85 Bali – Apr 2025",           rp(148_823_250), "CREDITED"],
        ["3",  "INV-2506-0106", "Jun 2025", "168 JKT + 106 Bali – Mei 2025",          rp(186_013_800), "CREDITED"],
        ["4",  "INV-2507-0139", "Jul 2025", "158 JKT + 99 Bali – Jun 2025",           rp(178_140_570), "CREDITED"],
        ["5",  "INV-2601-0176", "Jan 2026", "168 JKT + 106 Bali – Jul 2025 (Cancel)", rp(300_881_040), "CREDITED"],
        ["6",  "INV-2508-0176", "Aug 2026", "212 JKT + 124 Bali – Agu 2025",          rp(235_502_040), "CANCELED ✗"],
        ["7",  "INV-2509-0170", "Sep 2025", "261 JKT + 154 Bali – Sep 2025",          rp(281_518_200), "CANCELED ✗"],
        ["8",  "INV-2510-0148", "Oct 2025", "265 JKT + 164 Bali – Okt 2025",          rp(300_881_040), "CANCELED ✗"],
        ["9",  "INV-2511-0151", "Nov 2025", "216 JKT + 116 Bali – Nov 2025",          rp(224_974_800), "CANCELED ✗"],
        ["10", "INV-2512-0148", "Dec 2025", "213 JKT + 111 Bali – Des 2025",          rp(226_796_310), "CANCELED ✗"],
        ["11", "INV-2601-0176", "Jan 2026", "IDLE Januari 2026",                       rp(140_736_900), "CANCELED ✗"],
        ["",   "TOTAL",         "",         "",                                          rp(2_379_712_350), ""],
    ]
    cw = [Inches(0.4), Inches(1.6), Inches(0.8), Inches(3.5), Inches(2.2), Inches(1.6)]
    add_table(sl, idle_data, Inches(0.42), Inches(3.35), Inches(10.22), Inches(3.8),
              col_w=cw, font_sz=7.5, total_last=True)

    # Legend / note
    rect(sl, Inches(10.85), Inches(3.35), Inches(2.05), Inches(3.8), LGRAY)
    txt(sl, "KETERANGAN STATUS:", Inches(10.92), Inches(3.45), Inches(1.85), Inches(0.35),
        size=8, bold=True, color=NAVY)
    for i, (lbl, col, desc) in enumerate([
        ("CREDITED",   BLUE,   "FP dikreditkan"),
        ("CANCELED ✗", RED,    "FP dibatalkan,\ntidak diakui"),
    ]):
        rect(sl, Inches(10.92), Inches(3.88 + i * 1.2), Inches(1.75), Inches(1.0), col)
        txt(sl, f"{lbl}\n{desc}",
            Inches(11.0), Inches(3.95 + i * 1.2), Inches(1.6), Inches(0.85),
            size=8, bold=False, color=WHITE)


# ═══════════════════════════════════════════════════════════════
# SLIDE 10 – OUTSTANDING BML: OVERVIEW
# ═══════════════════════════════════════════════════════════════
def slide_outstanding_overview(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "TAGIHAN OUTSTANDING BML – SWAP ENERGI",
           "Tagihan Belum Terbayar | Jun 2025 – April 2026")
    footer(sl, "9 / 12")

    total_outstanding = 1_635_350_001 + 907_898_970 + 1_145_420_100
    kpi_card(sl, Inches(0.42), Inches(1.38), Inches(5.0), Inches(1.55),
             "TOTAL OUTSTANDING BML", rp(total_outstanding), "3 Kategori | Jun 2025 – Apr 2026", NAVY, GOLD)

    categories = [
        ("BBG De Sultan\nRp 10.000/unit",  rp(1_635_350_001), "21 Invoices", BLUE,     GOLD),
        ("BBG Tempur\nRp 11.000/unit",     rp(907_898_970),   "9 Invoices",  BLUE_MID, CYAN),
        ("Sewa Baterai\nRp 2.500/unit",    rp(1_145_420_100), "7 Invoices",  SWAP_CYAN, WHITE),
    ]
    for i, (lbl, val, sub, bg, vc) in enumerate(categories):
        x = Inches(5.65 + i * 2.6)
        rect(sl, x, Inches(1.38), Inches(2.4), Inches(1.55), bg)
        rect(sl, x, Inches(1.38), Inches(0.1), Inches(1.55), vc)
        txt(sl, lbl, x + Inches(0.18), Inches(1.46), Inches(2.1), Inches(0.5),
            size=9, bold=True, color=WHITE)
        txt(sl, val, x + Inches(0.18), Inches(1.93), Inches(2.1), Inches(0.72),
            size=14, bold=True, color=vc)
        txt(sl, sub, x + Inches(0.18), Inches(2.6), Inches(2.1), Inches(0.28),
            size=8, italic=True, color=RGBColor(0xCF, 0xE8, 0xFF))

    section_title(sl, "RINGKASAN TAGIHAN OUTSTANDING PER KATEGORI", y=Inches(3.12))

    data = [
        ["Kategori",                "Tarif",          "Periode",               "Jumlah Inv", "Total Invoice",    "Total Payment"],
        ["BBG De Sultan",           "Rp 10.000/unit", "Jun 2025 – Apr 2026",   "21 inv",     rp(1_635_350_001),  rp(1_635_350_001)],
        ["BBG Tempur",              "Rp 11.000/unit", "Sep 2025 – Apr 2026",   "9 inv",      rp(907_898_970),    rp(907_898_970)],
        ["Sewa Baterai Tempur",     "Rp 2.500/unit",  "Sep 2025 – Mar 2026",   "7 inv",      rp(1_145_420_100),  rp(1_124_781_900)],
        ["TOTAL OUTSTANDING",       "",               "",                      "",           rp(3_688_669_071),  rp(3_668_030_871)],
    ]
    cw = [Inches(2.5), Inches(1.5), Inches(2.5), Inches(1.2), Inches(2.3), Inches(2.3)]
    add_table(sl, data, Inches(0.42), Inches(3.57), Inches(12.46), Inches(2.85),
              col_w=cw, font_sz=9.5, total_last=True)

    rect(sl, Inches(0.42), Inches(6.55), Inches(12.46), Inches(0.58), RGBColor(0xE3, 0xF2, 0xFF))
    rect(sl, Inches(0.42), Inches(6.55), Inches(0.1), Inches(0.58), BLUE)
    txt(sl, "CATATAN: Tarif BBG berdasarkan kontrak PO 700 / PO 100 / PO 353 / PO 1000. "
        "Sewa Baterai 1.947 unit, PPh 23 dipotong dari Total Payment. "
        "1 invoice Sewa Baterai Feb 2026 belum diterima.",
        Inches(0.6), Inches(6.63), Inches(11.8), Inches(0.44), size=8.5, color=NAVY)


# ═══════════════════════════════════════════════════════════════
# SLIDE 11 – BBG DE SULTAN OUTSTANDING
# ═══════════════════════════════════════════════════════════════
def slide_bbg_desultan(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "OUTSTANDING: BBG DE SULTAN – Rp 10.000/unit",
           "Deposit Battery De Sultan | Jun 2025 – April 2026")
    footer(sl, "10 / 12")

    kpi_card(sl, Inches(0.42), Inches(1.38), Inches(3.6), Inches(1.22),
             "TOTAL INVOICE BBG DE SULTAN", rp(1_635_350_001), "21 Invoices | Jun 2025 – Apr 2026", NAVY, GOLD)
    rect(sl, Inches(4.22), Inches(1.38), Inches(3.5), Inches(1.22), BLUE)
    txt(sl, "Unit De Sultan", Inches(4.32), Inches(1.46), Inches(3.3), Inches(0.35), size=9, bold=True, color=RGBColor(0xCF, 0xE8, 0xFF))
    txt(sl, "100 – 1.338 unit\n(PO 100, 353, 700, 1000)", Inches(4.32), Inches(1.78), Inches(3.3), Inches(0.74), size=12, bold=True, color=CYAN)
    rect(sl, Inches(7.92), Inches(1.38), Inches(5.0), Inches(1.22), SWAP_CYAN)
    txt(sl, "Periode Tagihan", Inches(8.02), Inches(1.46), Inches(4.8), Inches(0.35), size=9, bold=True, color=NAVY)
    txt(sl, "Juni 2025  →  April 2026\nBBG (Battery Berlangganan Grab)", Inches(8.02), Inches(1.78), Inches(4.8), Inches(0.74), size=12, bold=True, color=NAVY)

    section_title(sl, "DETAIL 21 INVOICE BBG DE SULTAN (Rp 10.000/unit/hari × 365 hari / 2× Rp 5.000/730 hari)", y=Inches(2.78))

    data = [
        ["Periode",      "No Invoice",    "DPP",               "PPN",              "Total Invoice",    "Keterangan Unit"],
        ["Jun 2025",     "INV-2506-0174", rp(17_135_135),      rp(1_884_865),       rp(19_020_000),    "100 unit (PO 700) – Mei"],
        ["Jul 2025",     "INV-2507-0097", rp(13_963_964),      rp(1_536_036),       rp(15_500_000),    "100 unit (PO 700) – Juli"],
        ["Aug 2025",     "INV-2508-0018", rp(64_653_153),      rp(7_111_847),       rp(71_765_000),    "463 unit (PO 700) – Mei-Jun"],
        ["Sep 2025",     "INV-2509-0027", rp(62_567_568),      rp(6_882_432),       rp(69_450_000),    "463 unit (PO 700) – Mei-Jun"],
        ["Jul-Aug 2025", "INV-2509-0141", rp(57_265_765),      rp(6_299_235),       rp(63_565_000),    "237 unit (PO 700) – Jun"],
        ["26-31 Aug '25","INV-2509-0143", rp(1_801_802),       rp(198_198),         rp(2_000_000),     "100 unit (PO 100) – Agu"],
        ["Sep 2025",     "INV-2509-0142", rp(32_027_027),      rp(3_522_973),       rp(35_550_000),    "237 unit (PO 700) – Jun"],
        ["Sep 2025",     "INV-2509-0144", rp(13_513_514),      rp(1_486_486),       rp(15_000_000),    "100 unit (PO 100) – Agu"],
        ["Oct 2025",     "INV-2510-0020", rp(139_639_640),     rp(15_360_360),      rp(155_000_000),   "1000 unit (PO 1000) – Jan-Agu"],
        ["Nov 2025",     "INV-2511-0017", rp(135_135_135),     rp(14_864_865),      rp(150_000_000),   "1000 unit (PO 1000) – Jan-Agu"],
        ["Nov 2025",     "INV-2511-0155", rp(17_702_703),      rp(1_947_297),       rp(19_650_000),    "131 unit (PO 353) – Okt"],
        ["Nov 2025",     "INV-2511-0156", rp(3_972_973),       rp(437_027),         rp(4_410_000),     "30 unit (PO 353) – Okt"],
        ["Nov 2025",     "INV-2511-0157", rp(7_576_577),       rp(833_423),         rp(8_410_000),     "129 unit (PO 353) – Nov"],
        ["Des 2025",     "INV-2512-0019", rp(180_135_135),     rp(19_814_865),      rp(199_950_000),   "1290 unit (PO 1000+353) – Jan-Nov"],
        ["Des 2025",     "INV-2512-0112", rp(2_567_568),       rp(282_432),         rp(2_850_000),     "19 unit (PO 353) – Nov"],
        ["Jan 2026",     "INV-2601-0016", rp(182_788_288),     rp(20_106_712),      rp(202_895_000),   "1309 unit (PO 1000+353) – Jan-Nov"],
        ["Jan 2026",     "INV-2601-0073", rp(4_049_550),       rp(445_450),         rp(4_495_000),     "31 unit (PO 353) – Des"],
        ["Feb 2026",     "INV-2602-0019", rp(169_009_009),     rp(18_590_991),      rp(187_600_000),   "1340 unit (PO 1000+353) – 01-28 Feb"],
        ["Feb 2026",     "INV-2602-0111", rp(135_135),         rp(14_865),          rp(150_000),       "2 unit (PO 353) – 14-28 Feb"],
        ["Mar 2026",     "INV-2603-0017", rp(186_837_838),     rp(20_552_162),      rp(207_390_000),   "1328 unit (PO 1000+353) – Mar"],
        ["Apr 2026",     "INV-2604-0060", rp(180_810_811),     rp(19_889_189),      rp(200_700_000),   "1338 unit – Apr 2026"],
        ["TOTAL",        "",              rp(1_473_288_290),   rp(162_061_711),     rp(1_635_350_001), ""],
    ]
    cw = [Inches(1.1), Inches(1.8), Inches(1.8), Inches(1.5), Inches(2.0), Inches(4.0)]
    add_table(sl, data, Inches(0.42), Inches(3.2), Inches(12.46), Inches(3.97),
              col_w=cw, font_sz=7.0, total_last=True)


# ═══════════════════════════════════════════════════════════════
# SLIDE 12 – BBG TEMPUR + SEWA BATERAI
# ═══════════════════════════════════════════════════════════════
def slide_bbg_tempur_sewa(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, BLUE_LT)
    header(sl, "OUTSTANDING: BBG TEMPUR & SEWA BATERAI",
           "Deposit BBG Rp 11.000/unit | Sewa Baterai Rp 2.500/unit/hari")
    footer(sl, "11 / 12")

    # ── BBG Tempur ──
    section_title(sl, "BBG TEMPUR Rp 11.000/unit  |  Total Invoice: Rp 907.898.970  |  9 Invoices", y=Inches(1.38))

    data_tempur = [
        ["Periode",     "No Invoice",    "DPP",           "PPN",          "Total Invoice",  "Keterangan"],
        ["Sep-Nov 2025","INV-2511-0026", rp(26_180_000),  rp(2_879_800),  rp(29_059_800),   "60 unit BBG – Sep-Nov 2025"],
        ["Des 2025",    "INV-2512-0020", rp(31_856_000),  rp(3_504_160),  rp(35_360_160),   "100 unit BBG – 1-31 Des 2025"],
        ["Jan 2026",    "INV-2601-0017", rp(34_100_000),  rp(3_751_000),  rp(37_851_000),   "100 unit BBG – 1-31 Jan 2026"],
        ["Jan 2026",    "INV-2601-0134", rp(39_402_000),  rp(4_334_220),  rp(43_736_220),   "199 unit BBG – 1-31 Jan 2026"],
        ["Feb 2026",    "INV-2602-0020", rp(92_092_000),  rp(10_130_120), rp(102_222_120),  "299 unit BBG – 1-28 Feb 2026"],
        ["Feb 2026",    "INV-2602-0112", rp(64_977_000),  rp(7_147_470),  rp(72_124_470),   "389 unit BBG – 8-23 Feb 2026"],
        ["Mar 2026",    "INV-2603-0018", rp(234_608_000), rp(25_806_880), rp(260_414_880),  "688 unit BBG – 1-31 Mar 2026"],
        ["Mar 2026",    "INV-2603-0152", rp(14_542_000),  rp(1_599_620),  rp(16_141_620),   "161 unit BBG – Mar 2026 (final)"],
        ["Apr 2026",    "INV-2604-0052", rp(280_170_000), rp(30_818_700), rp(310_988_700),  "849 unit BBG – 1-30 Apr 2026"],
        ["TOTAL",       "",              rp(817_927_000), rp(89_971_970), rp(907_898_970),  ""],
    ]
    cw2 = [Inches(1.2), Inches(1.8), Inches(1.7), Inches(1.5), Inches(2.0), Inches(3.8)]
    add_table(sl, data_tempur, Inches(0.42), Inches(1.82), Inches(12.0), Inches(3.15),
              col_w=cw2, font_sz=7.5, total_last=True)

    # ── Sewa Baterai ──
    section_title(sl, "SEWA BATERAI Rp 2.500/unit/hari  |  Total Invoice: Rp 1.145.420.100  |  7 Invoices  |  1.947 unit", y=Inches(5.12))

    data_sewa = [
        ["Periode",    "No Invoice",    "Total Invoice",  "PPh 23",      "Total Payment",  "Keterangan"],
        ["Sep 2025",   "INV-2509-0102", rp(162_087_750),  rp(2_920_500), rp(159_167_250),  "1.947 unit – Sep 2025"],
        ["Okt 2025",   "INV-2510-0138", rp(167_490_675),  rp(3_017_850), rp(164_472_825),  "1.947 unit – Okt 2025"],
        ["Nov 2025",   "INV-2511-0134", rp(162_087_750),  rp(2_920_500), rp(159_167_250),  "1.947 unit – Nov 2025"],
        ["Des 2025",   "INV-2512-0124", rp(167_490_675),  rp(3_017_850), rp(164_472_825),  "1.947 unit – Des 2025"],
        ["Jan 2026",   "INV-2601-0104", rp(167_490_675),  rp(3_017_850), rp(164_472_825),  "1.947 unit – Jan 2026"],
        ["Feb 2026",   "Blm Diterima",  rp(151_281_900),  rp(2_725_800), rp(148_556_100),  "1.947 unit – Feb 2026"],
        ["Mar 2026",   "INV-2603-0151", rp(167_490_675),  rp(3_017_850), rp(164_472_825),  "1.947 unit – Mar 2026"],
        ["TOTAL",      "",              rp(1_145_420_100), rp(20_638_200),rp(1_124_781_900),""],
    ]
    cw3 = [Inches(1.2), Inches(1.8), Inches(2.0), Inches(1.5), Inches(2.0), Inches(3.8)]
    add_table(sl, data_sewa, Inches(0.42), Inches(5.55), Inches(12.3), Inches(1.62),
              col_w=cw3, font_sz=7.5, total_last=True)


# ═══════════════════════════════════════════════════════════════
# SLIDE 13 – GRAND TOTAL & NEXT STEPS
# ═══════════════════════════════════════════════════════════════
def slide_grand_total(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(sl, NAVY)

    # Decorative right
    rect(sl, Inches(9.5), 0, Inches(3.83), H, BLUE)
    rect(sl, Inches(9.65), 0, Inches(0.12), H, CYAN)
    rect(sl, 0, 0, W, Inches(0.07), CYAN)
    rect(sl, 0, H - Inches(0.07), W, Inches(0.07), GOLD)

    # Header
    txt(sl, "RINGKASAN TOTAL TAGIHAN BML", Inches(0.45), Inches(0.18), Inches(9), Inches(0.75),
        size=24, bold=True, color=WHITE)
    txt(sl, "AR BML Update  |  Per 21 Mei 2026", Inches(0.45), Inches(0.88), Inches(9), Inches(0.38),
        size=12, color=RGBColor(0xAD, 0xD8, 0xFF))
    rect(sl, 0, Inches(1.32), Inches(9.4), Inches(0.06), CYAN)

    # Grand total boxes
    grand_items = [
        ("INVOICE AR BML\n(Total Hutang BML)",       rp(3_626_894_700), "SMOOT Motor + SWAP Energi\n+ SMOOT After Sales", GOLD),
        ("TAGIHAN CUT-OFF APR'26\n(Siap Dibayarkan)", rp(1_976_747_940), "End Contract + De Sultan\n(Excludes IDLE Canceled)", CYAN),
        ("OUTSTANDING SWAP-BML",                      rp(3_688_669_071), "BBG De Sultan + BBG Tempur\n+ Sewa Baterai", RGBColor(0xFF, 0xFF, 0xFF)),
    ]

    for i, (lbl, val, desc, vc) in enumerate(grand_items):
        y_pos = Inches(1.52 + i * 1.7)
        rect(sl, Inches(0.42), y_pos, Inches(9.0), Inches(1.58), BLUE)
        rect(sl, Inches(0.42), y_pos, Inches(0.14), Inches(1.58), vc)
        txt(sl, lbl, Inches(0.65), y_pos + Inches(0.1), Inches(4.0), Inches(0.65),
            size=11, bold=True, color=RGBColor(0xCF, 0xE8, 0xFF))
        txt(sl, val, Inches(0.65), y_pos + Inches(0.7), Inches(5.0), Inches(0.75),
            size=26, bold=True, color=vc)
        txt(sl, desc, Inches(5.9), y_pos + Inches(0.2), Inches(3.3), Inches(1.2),
            size=9, italic=True, color=RGBColor(0xAD, 0xD8, 0xFF))

    # Right panel – next steps
    rect(sl, Inches(9.8), Inches(1.52), Inches(3.3), Inches(4.7), NAVY)
    txt(sl, "FOLLOW UP", Inches(9.9), Inches(1.62), Inches(3.1), Inches(0.45),
        size=14, bold=True, color=CYAN)
    rect(sl, Inches(9.9), Inches(2.1), Inches(3.1), Inches(0.05), CYAN)

    steps = [
        "Pembayaran Invoice AR BML\nsenilai Rp 3,6 M segera\ndikonfirmasi jadwal bayar",
        "Verifikasi Tagihan Cut-Off\nRp 1,97 M – semua invoice\nberstatus Siap Bayar",
        "Penyelesaian Outstanding\nSWAP Rp 3,69 M – proses\nreconciliation & payment",
        "Tagihan IDLE Rp 1,41 M\n(Canceled) – MoM review\nbersama tim Aftersales",
    ]
    for i, s in enumerate(steps):
        rect(sl, Inches(9.9), Inches(2.25 + i * 1.0), Inches(0.35), Inches(0.35),
             GOLD if i < 2 else ORANGE)
        txt(sl, str(i + 1), Inches(9.92), Inches(2.27 + i * 1.0), Inches(0.3), Inches(0.3),
            size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        txt(sl, s, Inches(10.35), Inches(2.25 + i * 1.0), Inches(2.6), Inches(0.9),
            size=8, color=WHITE)

    # Bottom total
    rect(sl, Inches(0.42), Inches(6.55), Inches(9.0), Inches(0.65), GOLD)
    txt(sl, f"GRAND TOTAL SEMUA TAGIHAN:  {rp(3_626_894_700 + 1_976_747_940 + 3_688_669_071)}",
        Inches(0.55), Inches(6.63), Inches(8.7), Inches(0.5),
        size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════
def main():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    slide_title(prs)
    slide_summary(prs)
    slide_smoot_motor(prs)
    slide_swap_energi(prs)
    slide_aftersales(prs)
    slide_cutoff_overview(prs)
    slide_percepatan(prs)
    slide_endcontract_desultan(prs)
    slide_idle(prs)
    slide_outstanding_overview(prs)
    slide_bbg_desultan(prs)
    slide_bbg_tempur_sewa(prs)
    slide_grand_total(prs)

    out = "/home/user/Claude-PPT/AR_BML_UPDATE_21052026.pptx"
    prs.save(out)
    print(f"✓  Saved: {out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
