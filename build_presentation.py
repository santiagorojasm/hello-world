#!/usr/bin/env python3
"""
Cube Ventures - McKinsey-level Investor Presentation
Reporte de Inversionistas: Segundo Semestre 2025
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.dml import MSO_THEME_COLOR
import copy

# ─── Brand Colors ───────────────────────────────────────────────────────────
NAVY        = RGBColor(0x0D, 0x1F, 0x3C)   # Deep navy #0D1F3C
TEAL        = RGBColor(0x00, 0xB3, 0x8A)   # Cube green #00B38A
GOLD        = RGBColor(0xF5, 0xA6, 0x23)   # Warm gold #F5A623
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY  = RGBColor(0xF4, 0xF6, 0xF9)
MID_GRAY    = RGBColor(0xA0, 0xAA, 0xBB)
DARK_TEXT   = RGBColor(0x1A, 0x28, 0x44)
LIGHT_TEAL  = RGBColor(0xE0, 0xF7, 0xF1)
LIGHT_GOLD  = RGBColor(0xFE, 0xF3, 0xDC)
RED_ACCENT  = RGBColor(0xE5, 0x4B, 0x4B)

FONT = "Montserrat"

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]  # Completely blank

# ─── Helpers ────────────────────────────────────────────────────────────────

def add_rect(slide, x, y, w, h, fill=None, line_color=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.line.width = line_width
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width if line_width else Pt(1)
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, x, y, w, h,
             font_size=14, bold=False, color=DARK_TEXT,
             align=PP_ALIGN.LEFT, font_name=FONT,
             italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_multiline(slide, lines, x, y, w, h,
                  font_size=11, bold_first=False, color=DARK_TEXT,
                  line_spacing=1.2, align=PP_ALIGN.LEFT, font_name=FONT,
                  bold=False):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    from pptx.util import Pt as P
    from pptx.oxml.ns import qn
    from lxml import etree

    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = P(font_size)
        run.font.color.rgb = color
        run.font.bold = bold or (bold_first and i == 0)
    return txBox


def slide_bg(slide, color=WHITE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_kpi_box(slide, x, y, w, h, value, label, value_color=TEAL, bg=LIGHT_GRAY,
                value_size=32, label_size=9):
    add_rect(slide, x, y, w, h, fill=bg)
    add_text(slide, value, x, y+0.1, w, h*0.55,
             font_size=value_size, bold=True, color=value_color,
             align=PP_ALIGN.CENTER)
    add_text(slide, label, x, y+h*0.52, w, h*0.45,
             font_size=label_size, bold=False, color=DARK_TEXT,
             align=PP_ALIGN.CENTER)


def add_section_header(slide, title, subtitle=None):
    add_rect(slide, 0, 0, 13.33, 1.1, fill=NAVY)
    add_text(slide, title, 0.5, 0.1, 12, 0.55,
             font_size=22, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_text(slide, subtitle, 0.5, 0.62, 11, 0.42,
                 font_size=11, bold=False, color=TEAL, align=PP_ALIGN.LEFT)


def add_footnote(slide, text, y=7.2):
    add_text(slide, text, 0.4, y, 12.5, 0.25,
             font_size=7, bold=False, color=MID_GRAY, align=PP_ALIGN.LEFT)


def divider_line(slide, y, color=TEAL, x=0.4, w=12.53):
    line = slide.shapes.add_connector(1, Inches(x), Inches(y),
                                       Inches(x+w), Inches(y))
    line.line.color.rgb = color
    line.line.width = Pt(1.5)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 1 – COVER
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, NAVY)

# Full-bleed green accent bar (left)
add_rect(slide, 0, 0, 0.35, 7.5, fill=TEAL)

# Gold accent bottom strip
add_rect(slide, 0, 6.8, 13.33, 0.7, fill=GOLD)

# Company name top-right watermark
add_text(slide, "CUBE VENTURES", 9.5, 0.3, 3.5, 0.5,
         font_size=11, bold=True, color=RGBColor(0xFF,0xFF,0xFF),
         align=PP_ALIGN.RIGHT)

# Main title
add_text(slide, "REPORTE DE\nINVERSIONISTAS", 1.0, 1.6, 9, 1.8,
         font_size=52, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

# Subtitle
add_text(slide, "SEGUNDO SEMESTRE 2025", 1.0, 3.6, 9, 0.5,
         font_size=18, bold=False, color=TEAL, align=PP_ALIGN.LEFT)

divider_line(slide, 4.22, color=TEAL, x=1.0, w=8)

# Key stat preview
add_text(slide, "Portafolio activo · Valoración potencial 2026 · Primer Exit · Estrategia de liquidez",
         1.0, 4.35, 11, 0.4,
         font_size=10, bold=False, color=MID_GRAY, align=PP_ALIGN.LEFT)

# Date & confidential
add_text(slide, "Mayo 2026", 1.0, 5.1, 4, 0.35,
         font_size=12, bold=False, color=WHITE)
add_text(slide, "CONFIDENCIAL", 8.5, 5.1, 4.5, 0.35,
         font_size=11, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)

# Bottom bar text
add_text(slide, "Meridian 75  ·  Exit Accelerator  ·  Cube Academy  ·  Aceleradora",
         0.5, 6.88, 12.33, 0.4,
         font_size=9, bold=False, color=NAVY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 2 – AGENDA
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_rect(slide, 0, 0, 13.33, 1.1, fill=NAVY)
add_text(slide, "Agenda", 0.5, 0.18, 12, 0.7,
         font_size=28, bold=True, color=WHITE)

items = [
    ("01", "Contexto Global de Venture Capital 2025"),
    ("02", "Evolución y Valoración del Portafolio"),
    ("03", "Primer Exit del Portafolio: Patiya"),
    ("04", "Empresas Destacadas: Trii · Geacco · Comanager"),
    ("05", "Estrategia de Liquidez y Exits 2026"),
    ("06", "Análisis Integral del Portafolio"),
    ("07", "Ecosistema y Tendencias Sectoriales"),
    ("08", "Perspectivas y Compromisos 2026"),
]

for i, (num, title) in enumerate(items):
    row = i % 4
    col = i // 4
    x = 0.5 + col * 6.5
    y = 1.35 + row * 1.45
    add_rect(slide, x, y, 6.1, 1.2, fill=LIGHT_GRAY)
    add_rect(slide, x, y, 0.55, 1.2, fill=TEAL)
    add_text(slide, num, x+0.02, y+0.3, 0.55, 0.55,
             font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, title, x+0.65, y+0.28, 5.3, 0.65,
             font_size=12, bold=False, color=DARK_TEXT, align=PP_ALIGN.LEFT)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 3 – MENSAJE DEL MANAGING DIRECTOR
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Mensaje del Managing Director",
                   "Visión estratégica · Cierre 2025")

# Quote block
add_rect(slide, 0.5, 1.3, 0.08, 5.6, fill=TEAL)
add_text(slide,
    '"En un ciclo exigente para el Venture Capital, los mejores retornos se '
    'construyen con disciplina, foco y capacidad de adaptación. Nuestro compromiso '
    'es seguir trabajando con transparencia, rigor y ambición para proteger y '
    'maximizar el valor del portafolio."',
    0.8, 1.35, 7.5, 1.5,
    font_size=12, bold=False, color=DARK_TEXT, italic=True)

add_text(slide, "Santiago Rojas Montoya · Managing Director & Cofundador",
         0.8, 2.8, 7.5, 0.35,
         font_size=10, bold=True, color=TEAL)

divider_line(slide, 3.3, color=LIGHT_GRAY, x=0.5, w=12.3)

# 4 callout boxes
callouts = [
    (TEAL,  "PRIMER EXIT", "Patiya: primera validación de liquidez del portafolio"),
    (GOLD,  "TRII", "US$567M en AUM · EBITDA positivo · Expansión a México"),
    (NAVY,  "GEACCO", "89% de crecimiento · US$1.17M · Cruce de umbral US$1M ARR"),
    (RGBColor(0x6C,0x63,0xFF), "ESTRATEGIA", "Exit Accelerator activo · 5 exits priorizados"),
]
for i, (col, title, body) in enumerate(callouts):
    x = 0.5 + i * 3.18
    add_rect(slide, x, 3.55, 3.0, 2.4, fill=col)
    add_text(slide, title, x+0.15, 3.65, 2.7, 0.45,
             font_size=11, bold=True, color=WHITE)
    add_text(slide, body, x+0.15, 4.2, 2.7, 1.6,
             font_size=9.5, bold=False, color=WHITE)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 4 – CONTEXTO GLOBAL VC 2025
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Contexto Global de Venture Capital 2025",
                   "El capital global se concentra · América Latina madura · Colombia redefine su base")

# Big numbers row
kpis = [
    ("US$425B", "Inversión VC Global 2025\n(Crunchbase)"),
    ("+63%", "Crecimiento EE.UU.\nvs. 2024"),
    ("40–50%", "Capital global\ncapturado por IA"),
    ("US$4.1B", "Inversión VC\nAmérica Latina 2025"),
    ("+13.8%", "Crecimiento LatAm\nvs. 2024"),
    ("US$224M", "Colombia 2025\n(62 rondas)"),
]
for i, (val, lbl) in enumerate(kpis):
    col = i % 3
    row = i // 3
    x = 0.35 + col * 4.32
    y = 1.25 + row * 2.65
    bg = LIGHT_TEAL if i < 3 else LIGHT_GRAY
    vc = TEAL if i < 3 else NAVY
    add_kpi_box(slide, x, y, 4.0, 2.3, val, lbl,
                value_color=vc, bg=bg, value_size=30, label_size=9)

# Insight callout
add_rect(slide, 0.35, 6.55, 12.5, 0.7, fill=NAVY)
add_text(slide,
    "INSIGHT: El capital global se polarizó — EE.UU. concentró megamegarondas en IA mientras "
    "LatAm mostró madurez sin profundidad. Colombia contrajo por ausencia de megarondas, "
    "no por deterioro estructural del ecosistema.",
    0.55, 6.6, 12.1, 0.58,
    font_size=9, bold=False, color=WHITE)

add_footnote(slide, "Fuentes: Crunchbase 2026 · KPMG Colombia Tech Report 2026 · Cuántico VP – Startuplinks LatAm VC Report 2026")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 5 – LATAM VC: PAÍSES
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "América Latina: Mapa de Inversión 2025",
                   "Brasil y México concentran el 73% del capital · Colombia y Chile lideran en eficiencia relativa")

# Country bars
countries = [
    ("Brasil",    2032, 15.6,  True),
    ("México",    980,  23.6,  True),
    ("Chile",     249,  None,  True),
    ("Colombia",  224, -36.7,  True),
    ("Argentina", 172, -58.8,  False),
    ("Uruguay",   40,  100.0,  True),
]
max_val = 2032
bar_start_x = 3.2
bar_max_w = 8.5
y_start = 1.3

for i, (country, val, pct, positive) in enumerate(countries):
    y = y_start + i * 0.95
    bar_w = (val / max_val) * bar_max_w
    bar_color = TEAL if positive else RED_ACCENT
    if country == "Colombia":
        bar_color = GOLD
    add_text(slide, country, 0.3, y+0.15, 2.8, 0.55,
             font_size=12, bold=(country in ["Brasil","Colombia"]), color=DARK_TEXT)
    add_rect(slide, bar_start_x, y+0.12, bar_w, 0.52, fill=bar_color)
    add_text(slide, f"US${val:,}M", bar_start_x + bar_w + 0.1, y+0.15, 2.0, 0.45,
             font_size=11, bold=True, color=DARK_TEXT)
    if pct is not None:
        pct_color = TEAL if pct > 0 else RED_ACCENT
        pct_str = f"+{pct:.1f}%" if pct > 0 else f"{pct:.1f}%"
        add_text(slide, pct_str, bar_start_x + bar_w + 1.6, y+0.15, 1.5, 0.45,
                 font_size=10, bold=True, color=pct_color)

# Colombia callout
add_rect(slide, 0.3, 7.0, 12.5, 0.28, fill=LIGHT_TEAL)
add_text(slide,
    "Colombia: La contracción refleja ausencia de megarondas, no deterioro. "
    "Fintech lidera con 59% del capital (US$505M, ticket promedio US$11.2M). "
    "Nuevos fondos locales consolidan el pipeline: Meridian 75, Alive Ventures, EWA Capital, LatinLeap.",
    0.5, 7.02, 12.0, 0.24, font_size=8, bold=False, color=DARK_TEXT)

add_footnote(slide, "Fuente: Cuántico VP – Startuplinks · Latin America VC Report 2026")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 6 – PORTAFOLIO: SNAPSHOT 2025
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Portafolio Cube Ventures: Snapshot al Cierre 2025",
                   "Valor concentrado en Tier 1 · Base depurada · 23 startups activas")

# Large KPIs
main_kpis = [
    ("US$916,881", "Valor Total del Portafolio\nCierre 2025", NAVY, WHITE),
    ("US$1,444,802", "Valoración Potencial 2026\n(Escenario Ponderado)", TEAL, WHITE),
    ("+57.6%", "Crecimiento Potencial\n2025 → 2026", GOLD, NAVY),
]
for i, (val, lbl, bg, tc) in enumerate(main_kpis):
    x = 0.35 + i * 4.32
    add_rect(slide, x, 1.2, 4.0, 1.8, fill=bg)
    add_text(slide, val, x+0.1, 1.3, 3.8, 0.9,
             font_size=26, bold=True, color=tc, align=PP_ALIGN.CENTER)
    add_text(slide, lbl, x+0.1, 2.18, 3.8, 0.72,
             font_size=9, bold=False, color=tc, align=PP_ALIGN.CENTER)

# Tier breakdown
tiers = [
    ("Tier 1", "US$767,834", "83.7%", "Motor de valor\n7 empresas activas", TEAL),
    ("Tier 2", "US$117,301", "12.8%", "Recomposición\n8 empresas activas", NAVY),
    ("Tier 3", "US$31,746",  "3.5%",  "Depuración natural\n6 activas · 76% cerradas", MID_GRAY),
]
for i, (tier, val, pct, desc, col) in enumerate(tiers):
    x = 0.35 + i * 4.32
    add_rect(slide, x, 3.2, 4.0, 2.6, fill=LIGHT_GRAY)
    add_rect(slide, x, 3.2, 4.0, 0.35, fill=col)
    add_text(slide, tier, x+0.12, 3.22, 3.76, 0.3,
             font_size=11, bold=True, color=WHITE)
    add_text(slide, val, x+0.12, 3.65, 3.76, 0.7,
             font_size=22, bold=True, color=col, align=PP_ALIGN.CENTER)
    add_text(slide, pct, x+0.12, 4.3, 3.76, 0.4,
             font_size=14, bold=True, color=DARK_TEXT, align=PP_ALIGN.CENTER)
    add_text(slide, desc, x+0.12, 4.75, 3.76, 0.9,
             font_size=9, bold=False, color=DARK_TEXT, align=PP_ALIGN.CENTER)

# Portfolio stats strip
add_rect(slide, 0.35, 6.0, 12.5, 0.8, fill=NAVY)
stats = [
    ("111", "Compañías históricas"),
    ("23", "Startups activas (20.7%)"),
    ("84", "Cierres naturales"),
    ("3", "En pausa"),
    ("1", "Exit ejecutado"),
    ("5", "Exits priorizados 2026"),
]
for i, (n, lbl) in enumerate(stats):
    x = 0.5 + i * 2.06
    add_text(slide, n,   x, 6.05, 1.8, 0.38,
             font_size=18, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    add_text(slide, lbl, x, 6.42, 1.8, 0.3,
             font_size=7.5, bold=False, color=WHITE, align=PP_ALIGN.CENTER)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 7 – EVOLUCIÓN DEL PORTAFOLIO
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Evolución del Portafolio 2024 → 2025",
                   "Variación por Tier · Dinámica interna · Reclasificaciones")

# Year-over-year comparison table
add_rect(slide, 0.35, 1.25, 12.5, 0.4, fill=NAVY)
for j, header in enumerate(["Tier", "Cierre 2024", "Cierre 2025", "Variación $", "Variación %", "Clave"]):
    x = [0.4, 1.7, 3.5, 5.3, 7.1, 8.7][j]
    w = [1.2, 1.7, 1.7, 1.7, 1.5, 4.0][j]
    add_text(slide, header, x, 1.28, w, 0.35,
             font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

rows = [
    ("Tier 1", "$784,500", "$767,834", "-$16,666", "-2.1%",
     "Reclasificación de Tu Techo Solar, Chak y Transgenia → Tier 2", LIGHT_TEAL, TEAL),
    ("Tier 2", "$64,838",  "$117,301", "+$52,463", "+80.9%",
     "Incorporación de 3 empresas desde Tier 1 + 1 exit ejecutado", LIGHT_GOLD, GOLD),
    ("Tier 3", "$92,626",  "$31,746",  "-$60,880", "-65.7%",
     "Cierres adicionales: 52% → 76% cerradas (depuración natural)", LIGHT_GRAY, RED_ACCENT),
    ("TOTAL",  "$941,965", "$916,881", "-$25,084", "-2.7%",
     "Contracción menor refleja solidez del núcleo Tier 1", NAVY, WHITE),
]

for i, (tier, v24, v25, delta, pct, note, bg, tc) in enumerate(rows):
    y = 1.65 + i * 1.05
    add_rect(slide, 0.35, y, 12.5, 1.0, fill=bg)
    vals = [tier, v24, v25, delta, pct]
    xs = [0.4, 1.7, 3.5, 5.3, 7.1]
    ws = [1.2, 1.7, 1.7, 1.7, 1.5]
    for j, (val, x, w) in enumerate(zip(vals, xs, ws)):
        is_bold = (tier == "TOTAL") or j == 0
        fc = tc if tier == "TOTAL" else DARK_TEXT
        if j == 3:
            fc = TEAL if "+" in val else RED_ACCENT
        if j == 4:
            fc = TEAL if "+" in val else RED_ACCENT
        add_text(slide, val, x, y+0.22, w, 0.55,
                 font_size=10 if tier != "TOTAL" else 11,
                 bold=is_bold, color=fc, align=PP_ALIGN.CENTER)
    add_text(slide, note, 8.7, y+0.22, 4.0, 0.55,
             font_size=8, bold=False,
             color=WHITE if tier == "TOTAL" else DARK_TEXT)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 8 – PRIMER EXIT: PATIYA
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Primer Exit del Portafolio: Patiya (Solfint S.A.S.)",
                   "Validación de la tesis · Liquidez en mercado difícil · Relaciones profundas con founders")

# Hero callout
add_rect(slide, 0.35, 1.25, 12.5, 1.6, fill=TEAL)
add_text(slide, "HITO: PRIMER EVENTO DE LIQUIDEZ DE CUBE VENTURES",
         0.6, 1.3, 12.0, 0.45,
         font_size=14, bold=True, color=WHITE)
add_text(slide,
    "La recompra de participación por parte del emprendedor Juan David Palacios es la "
    "primera validación concreta de que incluso en un mercado difícil, es posible construir "
    "oportunidades de salida. Confirma que Cube genera relaciones profundas con sus fundadores.",
    0.6, 1.75, 12.0, 0.95,
    font_size=10, bold=False, color=WHITE)

# Timeline
add_text(slide, "CRONOLOGÍA DEL EXIT", 0.5, 3.1, 5, 0.35,
         font_size=11, bold=True, color=DARK_TEXT)

timeline = [
    ("H2 2022",    "Ingresa a Cohorte 4 de Cube Ventures · Cede 5.5% accionario"),
    ("2025",       "Cambio estratégico · Alejamiento del modelo startup tradicional"),
    ("Abr–Jun 25", "Acuerdo de recompra COP $12M · 12 cuotas mensuales de $1M"),
    ("May 2026",   "Proceso prejurídico exitoso · Abono adicional COP $3M · Total pagado: $6M"),
    ("Jun 2026",   "Saldo pendiente COP $6M pactado en 2 meses · Seguimiento activo"),
]
for i, (period, event) in enumerate(timeline):
    y = 3.55 + i * 0.62
    add_rect(slide, 0.4, y, 2.1, 0.52, fill=NAVY if i < 4 else GOLD)
    tc = WHITE if i < 4 else NAVY
    add_text(slide, period, 0.42, y+0.1, 2.06, 0.35,
             font_size=9, bold=True, color=tc, align=PP_ALIGN.CENTER)
    add_rect(slide, 2.5, y+0.22, 0.08, 0.08, fill=TEAL)
    add_text(slide, event, 2.75, y+0.08, 10.0, 0.4,
             font_size=9, bold=False, color=DARK_TEXT)

# Key insight
add_rect(slide, 0.35, 7.0, 12.5, 0.32, fill=NAVY)
add_text(slide,
    "IMPLICACIÓN ESTRATÉGICA: Este exit abre el modelo de recompra como mecanismo de liquidez "
    "adicional para el portafolio, complementando los procesos de M&A y venta secundaria.",
    0.5, 7.03, 12.1, 0.27, font_size=8.5, bold=False, color=WHITE)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 9 – SPOTLIGHT: TRII
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_rect(slide, 0, 0, 13.33, 1.1, fill=NAVY)
add_text(slide, "Spotlight: Trii", 0.5, 0.1, 7, 0.55,
         font_size=22, bold=True, color=WHITE)
add_text(slide, "Plataforma de inversión bursátil digital · Colombia · Perú · Chile · Próximo: México",
         0.5, 0.63, 12, 0.38, font_size=11, bold=False, color=TEAL)

# Y Combinator badge
add_rect(slide, 10.5, 1.25, 2.5, 0.6, fill=GOLD)
add_text(slide, "Y COMBINATOR '21", 10.55, 1.3, 2.4, 0.5,
         font_size=9, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

# Metric grid
metrics = [
    ("US$567M",   "AUM Q1 2026",     "(vs US$100M en 2023\ncuando invertimos)"),
    ("300K+",     "Cuentas Fondeadas","3 países operando"),
    ("US$2.1M",   "Revenue Q1 2026",  "(vs US$1.38M en 2025)"),
    ("EBITDA+",   "Rentabilidad",     "Alcanzada Q2 2025\nSostenida en 2026"),
    ("+71%",      "Crecimiento AUM",  "2024 → 2025"),
    ("México",    "Próxima Expansión","Mayor mercado hispano\nde inversión minorista"),
]
for i, (val, lbl, sub) in enumerate(metrics):
    col = i % 3
    row = i // 2
    x = 0.35 + col * 4.32
    y = 2.0 + (i // 3) * 2.3
    bg = LIGHT_TEAL if i % 2 == 0 else LIGHT_GRAY
    add_rect(slide, x, y, 4.0, 2.1, fill=bg)
    add_text(slide, val, x+0.1, y+0.15, 3.8, 0.8,
             font_size=28, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    add_text(slide, lbl, x+0.1, y+0.95, 3.8, 0.38,
             font_size=10, bold=True, color=DARK_TEXT, align=PP_ALIGN.CENTER)
    add_text(slide, sub, x+0.1, y+1.35, 3.8, 0.65,
             font_size=8, bold=False, color=MID_GRAY, align=PP_ALIGN.CENTER)

# Investment note
add_rect(slide, 0.35, 6.6, 12.5, 0.65, fill=NAVY)
add_text(slide,
    "Nuestra inversión: US$5,000 · Feb 2023 · 0.02% de participación  "
    "│  ESTRATEGIA: Preservar exposición al crecimiento · Evaluar oportunidad de secundario · "
    "Expansión México 2026 abre nueva dimensión de escala",
    0.55, 6.65, 12.1, 0.55, font_size=8.5, bold=False, color=WHITE)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 10 – SPOTLIGHT: GEACCO
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_rect(slide, 0, 0, 13.33, 1.1, fill=TEAL)
add_text(slide, "Spotlight: Geacco  |  Desempeño ALTO (3.92/4.0)",
         0.5, 0.1, 10, 0.55, font_size=22, bold=True, color=WHITE)
add_text(slide,
    "Gestión y automatización de bienes en leasing · SaaS B2B · Clientes Tier-1 · Expansión México 2026",
    0.5, 0.63, 12, 0.38, font_size=11, bold=False, color=LIGHT_TEAL)

# Revenue growth visual
add_rect(slide, 0.35, 1.2, 5.8, 4.8, fill=NAVY)
add_text(slide, "CRECIMIENTO DE INGRESOS", 0.5, 1.3, 5.5, 0.4,
         font_size=10, bold=True, color=WHITE)
# Bar 2024
add_rect(slide, 1.2, 4.3, 1.5, 1.4, fill=MID_GRAY)
add_text(slide, "$618K", 1.1, 3.85, 1.7, 0.4,
         font_size=13, bold=True, color=MID_GRAY, align=PP_ALIGN.CENTER)
add_text(slide, "2024", 1.1, 5.72, 1.7, 0.3,
         font_size=10, bold=False, color=WHITE, align=PP_ALIGN.CENTER)
# Bar 2025
add_rect(slide, 3.1, 2.9, 1.5, 2.8, fill=TEAL)
add_text(slide, "$1.17M", 3.0, 2.5, 1.7, 0.4,
         font_size=13, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
add_text(slide, "2025", 3.0, 5.72, 1.7, 0.3,
         font_size=10, bold=False, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "+89.47%", 0.45, 3.9, 5.5, 0.55,
         font_size=24, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
add_text(slide, "#1 en facturación del portafolio",
         0.45, 5.3, 5.5, 0.35, font_size=9, bold=False,
         color=MID_GRAY, align=PP_ALIGN.CENTER)

# Right side facts
facts = [
    ("CLIENTES", "Itaú · Davivienda · Banco de Occidente\nBanco Finandina · Tangelo · IRIS"),
    ("HITO 2025", "Primera empresa del portafolio\nen cruzar US$1M ARR"),
    ("2026", "Expansión a México · S2 2026\nSistema financiero más grande de LatAm"),
    ("EXIT SCORE", "Candidata #1 a exit vía M&A\ncon actores financieros regionales"),
]
for i, (title, body) in enumerate(facts):
    x = 6.5
    y = 1.2 + i * 1.2
    add_rect(slide, x, y, 6.5, 1.1, fill=LIGHT_TEAL if i % 2 == 0 else LIGHT_GRAY)
    add_text(slide, title, x+0.15, y+0.08, 2.0, 0.38,
             font_size=9, bold=True, color=TEAL)
    add_text(slide, body, x+0.15, y+0.45, 6.2, 0.58,
             font_size=9.5, bold=False, color=DARK_TEXT)

add_rect(slide, 0.35, 6.2, 12.5, 0.6, fill=GOLD)
add_text(slide,
    "ESTRATEGIA EXIT: M&A estratégico con actores financieros regionales  │  "
    "Venta secundaria total de la participación  │  Consolidación con jugadores fintech/paytech",
    0.55, 6.28, 12.0, 0.48, font_size=9, bold=False, color=NAVY)

add_footnote(slide, "Nuestra inversión: 5.5% vía inversión en especie · US$9,263 · Cohorte 3 (S1 2022)")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 11 – TIER 1: DESEMPEÑO COMPARATIVO
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Tier 1: Análisis Integral de Desempeño",
                   "7 empresas · Evaluación por 7 criterios VC · Fuente: Endeavor & Glisco Partners metodología")

# Performance table
add_rect(slide, 0.35, 1.25, 12.5, 0.4, fill=NAVY)
headers = ["Empresa", "Score", "Categoría", "Revenue 2025", "Crecimiento YoY", "Prioridad Exit"]
xs = [0.4, 2.1, 3.3, 5.0, 7.2, 9.4]
ws = [1.6, 1.1, 1.6, 2.1, 2.1, 3.6]
for j, (h, x, w) in enumerate(zip(headers, xs, ws)):
    add_text(slide, h, x, 1.28, w, 0.35,
             font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

companies = [
    ("Geacco",     "3.92", "ALTO",       "$1,170,933", "+89.47%", "#1 - M&A / Secundario",    LIGHT_TEAL, TEAL),
    ("Bmotik",     "3.25", "MEDIO-ALTO", "$394,465",   "+56.74%", "#2 - Adquisición / Capital",LIGHT_TEAL, NAVY),
    ("Esuite",     "2.85", "MEDIO",      "N/D",        "–",       "Monitoreo",                 LIGHT_GRAY, DARK_TEXT),
    ("Contactiza", "2.42", "MEDIO-BAJO", "Activa",     "–",       "Monitoreo",                 LIGHT_GRAY, DARK_TEXT),
    ("Asisvisa",   "2.35", "MEDIO-BAJO", "$126,400",   "-7.45%",  "#3 - Recompra founders",    LIGHT_GOLD, GOLD),
    ("Comanager",  "2.20", "MEDIO-BAJO", "Post-pivote","–",       "#4 - Post-ronda US$400K",   LIGHT_GOLD, GOLD),
    ("Igo",        "2.10", "MEDIO-BAJO", "Recurrente", "–",       "Monitoreo",                 LIGHT_GRAY, DARK_TEXT),
]
for i, (name, score, cat, rev, grw, exit_str, bg, tc) in enumerate(companies):
    y = 1.65 + i * 0.75
    add_rect(slide, 0.35, y, 12.5, 0.72, fill=bg)
    vals = [name, score, cat, rev, grw, exit_str]
    for j, (val, x, w) in enumerate(zip(vals, xs, ws)):
        fc = tc if j in [0, 1, 2] else DARK_TEXT
        if j == 4:
            fc = TEAL if "+" in val else (RED_ACCENT if "-" in val else DARK_TEXT)
        add_text(slide, val, x, y+0.15, w, 0.45,
                 font_size=9, bold=(j == 0), color=fc,
                 align=PP_ALIGN.LEFT if j in [0, 5] else PP_ALIGN.CENTER)

add_footnote(slide, "Score 1–4: 1=insuficiente · 4=sobresaliente · Criterios: Tracción comercial, Mercado, Innovación IA, Equipo, Eficiencia capital, Estrategia exit, Escalabilidad")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 12 – ESTRATEGIA DE LIQUIDEZ
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Estrategia de Liquidez y Exits 2026",
                   "5 activos priorizados · 3 mecanismos de salida · Exit Accelerator activo")

# Title insight
add_rect(slide, 0.35, 1.2, 12.5, 0.55, fill=LIGHT_TEAL)
add_text(slide,
    "En América Latina, la liquidez no ocurre sola — se construye. "
    "Con el vehículo entrando en su etapa final de maduración en 2027, los próximos 18 meses son críticos.",
    0.55, 1.25, 12.1, 0.44, font_size=10, bold=False, color=NAVY)

# Exit cards
exits = [
    ("Trii",      "Preservación + Oportunista",
     "• Participación minoritaria (0.02%)\n• Evaluar secundario si hay interés\n• Mantener hasta evento mayor valorización",
     "PRIORIDAD: ALTA", TEAL),
    ("Geacco",    "M&A Estratégico",
     "• Comprador estratégico regional\n• Venta secundaria total\n• Consolidación con fintech/paytech\n• Expansión México ↑ valoración",
     "PRIORIDAD: ALTA", NAVY),
    ("Comanager", "Post-Ronda → Secundario",
     "• Ronda SAFE US$400K en curso\n• Cap post-money US$5M\n• Esperar cierre y validación pivote\n• Coinversión con nuevos LPs",
     "PRIORIDAD: MEDIA", GOLD),
    ("Bmotik",    "Adquisición Total",
     "• Constellation (1,000+ adquisiciones)\n• Interés activo por >1 año\n• Umbral: US$800K–1M ARR\n• Trayectoria: ~US$550K ARR 2025",
     "PRIORIDAD: ALTA", RGBColor(0x6C,0x63,0xFF)),
    ("Asisvisa",  "Recompra / Legaltech",
     "• Recompra por founders\n• Boutique legaltech regional\n• Procesos ágiles de monetización",
     "PRIORIDAD: MEDIA", MID_GRAY),
]
for i, (name, mech, detail, prio, col) in enumerate(exits):
    col_idx = i % 3
    row_idx = i // 3
    x = 0.35 + col_idx * 4.32
    y = 2.0 + row_idx * 2.5
    add_rect(slide, x, y, 4.0, 2.35, fill=LIGHT_GRAY)
    add_rect(slide, x, y, 4.0, 0.35, fill=col)
    add_text(slide, f"{name}  ·  {mech}", x+0.1, y+0.05, 3.8, 0.28,
             font_size=9.5, bold=True, color=WHITE)
    add_text(slide, detail, x+0.1, y+0.42, 3.8, 1.4,
             font_size=8.5, bold=False, color=DARK_TEXT)
    prio_col = TEAL if "ALTA" in prio else GOLD
    add_text(slide, prio, x+0.1, y+1.95, 3.8, 0.3,
             font_size=8, bold=True, color=prio_col)

add_footnote(slide, "Exit Accelerator: +1,068 organizaciones identificadas · 120 contactados · 22 reuniones · 5 oportunidades avanzadas (caso Legalítica)")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 13 – VALORACIÓN POTENCIAL 2026
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Valoración Potencial del Portafolio: Escenarios 2026",
                   "Análisis de 3 escenarios · Ponderación: Conservador 50% · Optimista 25% · Pesimista 25%")

# Scenario comparison table
add_rect(slide, 0.35, 1.25, 12.5, 0.4, fill=NAVY)
cols_h = ["Escenario", "Peso", "Tier 1", "Tier 2", "Tier 3", "TOTAL"]
xs_t = [0.4, 2.1, 3.3, 5.5, 7.7, 9.9]
ws_t = [1.6, 1.1, 2.1, 2.1, 2.1, 3.0]
for h, x, w in zip(cols_h, xs_t, ws_t):
    add_text(slide, h, x, 1.28, w, 0.35,
             font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

scenarios = [
    ("OPTIMISTA",    "25%", "$1,850,000", "$72,248",  "$59,837",  "$1,982,085", LIGHT_TEAL, TEAL),
    ("CONSERVADOR",  "50%", "$1,300,000", "$55,390",  "$49,277",  "$1,404,667", LIGHT_GRAY, NAVY),
    ("PESIMISTA",    "25%", "$950,000",   "$9,633",   "$28,158",  "$987,791",   LIGHT_GOLD, RED_ACCENT),
    ("PONDERADO★",   "100%","—",          "—",        "—",        "$1,444,802", NAVY, WHITE),
]
for i, (esc, peso, t1, t2, t3, total, bg, tc) in enumerate(scenarios):
    y = 1.65 + i * 0.78
    add_rect(slide, 0.35, y, 12.5, 0.75, fill=bg)
    for j, (val, x, w) in enumerate(zip([esc, peso, t1, t2, t3, total], xs_t, ws_t)):
        is_total = j == 5
        fc = tc if bg == NAVY else (tc if j == 0 else DARK_TEXT)
        if is_total:
            fc = tc
        add_text(slide, val, x, y+0.15, w, 0.48,
                 font_size=10 if not is_total else 12,
                 bold=is_total or j == 0, color=fc,
                 align=PP_ALIGN.CENTER)

# Growth visual
add_rect(slide, 0.35, 4.7, 5.8, 1.9, fill=NAVY)
add_text(slide, "TRAYECTORIA DE VALOR", 0.5, 4.8, 5.5, 0.35,
         font_size=10, bold=True, color=WHITE)
add_text(slide, "$916,881", 0.5, 5.2, 2.5, 0.55,
         font_size=22, bold=True, color=MID_GRAY, align=PP_ALIGN.CENTER)
add_text(slide, "Hoy", 0.5, 5.75, 2.5, 0.3,
         font_size=9, color=MID_GRAY, align=PP_ALIGN.CENTER)
add_text(slide, "→", 3.0, 5.25, 0.6, 0.5,
         font_size=22, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
add_text(slide, "$1,444,802", 3.6, 5.2, 2.3, 0.55,
         font_size=22, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
add_text(slide, "Potencial 2026", 3.6, 5.75, 2.3, 0.3,
         font_size=9, color=TEAL, align=PP_ALIGN.CENTER)

# Right side: what drives value
add_rect(slide, 6.5, 4.7, 6.3, 1.9, fill=LIGHT_TEAL)
add_text(slide, "+57.6% DE CRECIMIENTO POTENCIAL REQUIERE:", 6.6, 4.78, 6.0, 0.38,
         font_size=9.5, bold=True, color=NAVY)
drivers = ["✓ Eventos de exit (M&A / Secundario) en Geacco o Bmotik",
           "✓ Geacco alcanza valoración ≥ US$3M (en curso con 2x revenue)",
           "✓ Bmotik cruza US$800K–1M ARR → Constellation deal",
           "✓ Trii mantiene momentum de AUM hacia ronda / IPO"]
for k, d in enumerate(drivers):
    add_text(slide, d, 6.65, 5.18 + k * 0.32, 6.0, 0.3,
             font_size=8.5, bold=False, color=DARK_TEXT)

add_footnote(slide, "Las cifras son estimaciones prospectivas basadas en desempeño observado y supuestos de mercado. No constituyen garantía de resultados.")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 14 – ECOSISTEMA DEL PORTAFOLIO
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Ecosistema del Portafolio: Análisis Sectorial 2025",
                   "23 startups activas · IA como factor diferenciador transversal · Fintech lidera el valor")

sectors = [
    ("Fintech",     4, "17%", "Vertical con mayor tracción. Concentra 59% del capital Colombia. Candidata #1 a exits.",    TEAL,  "$505M cap Colombia"),
    ("Healthtech",  3, "13%", "Alta estabilidad. Demanda institucional creciente. US$71.7B global 2025.",                  NAVY,  "Sector #3 Colombia"),
    ("Edtech",      2, "9%",  "Mercado amplio y demanda estructural. Diferenciación clave para 2026.",                     GOLD,  "Crecimiento sostenido"),
    ("Legaltech",   2, "9%",  "IA transforma procesos jurídicos. Brechas sin resolver. Alta adopción temprana.",            RGBColor(0x6C,0x63,0xFF), "IA aceleradora"),
    ("Climatech",   2, "9%",  "Emergente estratégico. Capital institucional en crecimiento. ESG momentum.",                RGBColor(0x2E,0xCC,0x71), "Monitor LP impact"),
    ("Biztech/SaaS",2, "9%",  "B2B predecible. SaaS Enterprise lidera Colombia con 27% de startups activas.",             RGBColor(0xE6,0x7E,0x22), "#1 Colombia 2026"),
]
for i, (sector, n, pct, insight, col, badge) in enumerate(sectors):
    col_idx = i % 2
    row_idx = i // 2
    x = 0.35 + col_idx * 6.5
    y = 1.3 + row_idx * 1.9
    add_rect(slide, x, y, 6.2, 1.75, fill=LIGHT_GRAY)
    add_rect(slide, x, y, 0.5, 1.75, fill=col)
    add_text(slide, str(n), x+0.02, y+0.5, 0.48, 0.55,
             font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, sector, x+0.65, y+0.1, 3.0, 0.45,
             font_size=14, bold=True, color=DARK_TEXT)
    add_text(slide, f"{n} startups · {pct} del portafolio activo",
             x+0.65, y+0.52, 3.5, 0.3,
             font_size=8, bold=False, color=MID_GRAY)
    add_text(slide, insight, x+0.65, y+0.85, 5.4, 0.75,
             font_size=8.5, bold=False, color=DARK_TEXT)
    add_rect(slide, x+4.8, y+0.08, 1.25, 0.32, fill=col)
    add_text(slide, badge, x+4.82, y+0.1, 1.21, 0.28,
             font_size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# AI callout bar
add_rect(slide, 0.35, 7.0, 12.5, 0.32, fill=NAVY)
add_text(slide,
    "IA TRANSVERSAL: Geacco +89% revenue con IA operativa · Legalítica y Healthtech en adopción · "
    "Ausencia de IA fue factor determinante en cierres durante 2025.",
    0.55, 7.03, 12.1, 0.27, font_size=8.5, bold=False, color=WHITE)

add_footnote(slide, "Colombia Tech Report 2026 (KPMG) · Latin America VC Report 2026 (Cuántico VP) · Crunchbase 2026")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 15 – CIERRES: CONTEXTO Y CONCLUSIONES
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Cierres 2025: Contexto, Causas y Conclusiones",
                   "76% de cierres refleja madurez del portafolio, no anomalía · Base activa con calidad superior")

# Stats bar
add_rect(slide, 0.35, 1.25, 12.5, 0.9, fill=LIGHT_GRAY)
stats_c = [
    ("84", "Startups cerradas"),
    ("3",  "En pausa"),
    ("23", "Activas (20.7%)"),
    ("76%","Tasa de cierre\nvs. 52% H1 2025"),
    ("1:10","Seed → Serie A\n(Colombia promedio)"),
    ("26m","Mediana entre\nrondas Colombia"),
]
for i, (n, lbl) in enumerate(stats_c):
    x = 0.5 + i * 2.06
    vc = TEAL if i in [2] else (RED_ACCENT if i == 3 else NAVY)
    add_text(slide, n,   x, 1.28, 1.9, 0.45,
             font_size=22, bold=True, color=vc, align=PP_ALIGN.CENTER)
    add_text(slide, lbl, x, 1.72, 1.9, 0.38,
             font_size=7.5, bold=False, color=DARK_TEXT, align=PP_ALIGN.CENTER)

# 3 causas
add_text(slide, "FACTORES CLAVE DE CIERRE", 0.5, 2.3, 7, 0.35,
         font_size=11, bold=True, color=DARK_TEXT)

causas = [
    ("1", "Modelos poco sostenibles",
     "Propuestas de valor no adaptadas a demanda real. Inversores exigen unit "
     "economics sólidos y eficiencia de capital. Sin demostrar esto antes de agotar "
     "el runway → sin acceso a rondas de seguimiento.", RED_ACCENT),
    ("2", "Equipo fundador",
     "Ausencia de ingresos para founders fue factor determinante en Tier 3. "
     "Presión del mercado llevó al cierre antes de explorar pivotes o reestructuración. "
     "Dependencia de levantamiento de capital sin revenue propio.", GOLD),
    ("3", "Intensificación competitiva por IA",
     "IA redefinió el estándar competitivo en Legaltech, Healthtech y Biztech. "
     "Captó 40-50% del capital en EE.UU. Startups sin IA adoptada perdieron "
     "diferenciación y viabilidad.", NAVY),
]
for i, (num, title, body, col) in enumerate(causas):
    y = 2.75 + i * 1.45
    add_rect(slide, 0.35, y, 8.5, 1.35, fill=LIGHT_GRAY)
    add_rect(slide, 0.35, y, 0.45, 1.35, fill=col)
    add_text(slide, num, 0.35, y+0.38, 0.45, 0.5,
             font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, title, 0.9, y+0.1, 3.5, 0.4,
             font_size=10.5, bold=True, color=col)
    add_text(slide, body,  0.9, y+0.5, 7.8, 0.78,
             font_size=8.5, bold=False, color=DARK_TEXT)

# Right: positive reframe
add_rect(slide, 9.2, 2.75, 3.9, 4.35, fill=NAVY)
add_text(slide, "LECTURA POSITIVA", 9.35, 2.85, 3.6, 0.38,
         font_size=10, bold=True, color=TEAL)
points = [
    "Base activa de calidad superior: 23 startups con ingresos sostenidos y modelos validados",
    "Portafolio más preciso: energía concentrada en compañías con real potencial de retorno",
    "Tasa coherente con benchmarks: Colombia 1 de 10 seed→Serie A es la norma global",
    "Tier 1 no se deterioró: núcleo de valor ($767K) sólido y en crecimiento",
    "Aprendizaje acumulado: criterios de selección y acompañamiento continúan mejorando",
]
for k, pt in enumerate(points):
    add_text(slide, f"→ {pt}", 9.35, 3.3 + k * 0.7, 3.6, 0.62,
             font_size=8.5, bold=False, color=WHITE)

add_footnote(slide, "Colombia Tech Report 2025-2026 · Latin America VC Report 2026 · Crunchbase 2026")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 16 – CUBE VENTURES: PLATAFORMA INTEGRAL
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Cube Ventures: Evolución hacia Plataforma de Creación de Valor",
                   "Más que una aceleradora · 4 líneas complementarias · Sinergias para el portafolio")

# Intro
add_text(slide,
    "Cube Ventures evoluciona de aceleradora tradicional a plataforma integral. "
    "Las 4 líneas son complementarias y crean un ciclo virtuoso: "
    "la aceleradora genera portafolio, Meridian 75 lo capitaliza, Exit Accelerator lo liquida, "
    "Cube Academy atrae capital inteligente.",
    0.5, 1.2, 12.5, 0.65, font_size=10.5, bold=False, color=DARK_TEXT)

# 4 pillars
pillars = [
    ("ACELERADORA", "Origen del portafolio",
     "• 111 compañías históricas\n• 5 cohortes (2020–2024)\n"
     "• Community emprendedora\n• Pipeline para Meridian 75",
     NAVY, "FUNDAMENTO"),
    ("MERIDIAN 75", "Brazo de Venture Capital",
     "• Inversión con mayor capacidad\n• Rondas activas del portafolio\n"
     "• Alternativas para LPs en follow-ons\n• Expansión del deal flow",
     TEAL, "CAPITAL"),
    ("EXIT ACCELERATOR", "Liquidez activa",
     "• +1,068 potenciales compradores\n• Pipeline M&A estructurado\n"
     "• SPVs / Continuation vehicles\n• Exits priorizados en 5 activos",
     GOLD, "LIQUIDEZ"),
    ("CUBE ACADEMY", "Capital inteligente",
     "• Alianza Babson College + Cornell\n• Formación de family offices\n"
     "• Educación de nuevos LPs\n• Red de capital institucional",
     RGBColor(0x6C,0x63,0xFF), "COMUNIDAD"),
]
for i, (name, sub, detail, col, badge) in enumerate(pillars):
    x = 0.35 + i * 3.22
    add_rect(slide, x, 2.0, 3.0, 4.8, fill=col)
    add_text(slide, name, x+0.12, 2.1, 2.76, 0.45,
             font_size=11, bold=True, color=WHITE)
    add_text(slide, sub, x+0.12, 2.55, 2.76, 0.35,
             font_size=8.5, bold=False, color=WHITE, italic=True)
    divider_line(slide, 3.0, color=WHITE, x=x+0.1, w=2.8)
    add_text(slide, detail, x+0.12, 3.1, 2.76, 3.3,
             font_size=9, bold=False, color=WHITE)
    add_rect(slide, x+0.6, 6.45, 1.8, 0.3, fill=WHITE)
    add_text(slide, badge, x+0.62, 6.47, 1.76, 0.26,
             font_size=8, bold=True, color=col, align=PP_ALIGN.CENTER)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 17 – PERSPECTIVAS Y COMPROMISOS 2026
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, WHITE)
add_section_header(slide, "Perspectivas y Compromisos 2026",
                   "5 prioridades estratégicas · Norte definido · 18 meses decisivos")

add_text(slide,
    "2026 tiene un norte definido: acompañar a las compañías de mayor representación hacia "
    "eventos de liquidez, sostener el valor del portafolio mientras ese proceso madura y "
    "construir capacidad operativa para actuar con mayor escala y precisión.",
    0.5, 1.2, 12.5, 0.6, font_size=10.5, bold=False, color=DARK_TEXT)

priorities = [
    ("01", "EXITS Y LIQUIDEZ",
     "Ejecutar exits en Geacco (M&A regional) y Bmotik (Constellation, umbral US$800K-1M ARR). "
     "Opcionalidad en Trii. Recompra Asisvisa. Exit Accelerator como motor.",
     TEAL, True),
    ("02", "VALORIZACIÓN DEL NÚCLEO",
     "Sostener la trayectoria de Geacco hacia US$2.5M+ revenue 2026. Apoyar Bmotik en "
     "capital de crecimiento. Maximizar valor antes de salida.",
     NAVY, True),
    ("03", "COMANAGER: RONDA Y PIVOTE",
     "Acompañar cierre de ronda US$400K (SAFE/nota convertible, US$5M cap). "
     "Validar Entity Connect como plataforma antifraude. Evaluar secundario post-ronda.",
     GOLD, False),
    ("04", "ADOPCIÓN DE IA EN PORTAFOLIO",
     "Acelerar integración de IA en Legaltech y Biztech. Replicar el caso Geacco: "
     "IA como diferenciador de propuesta de valor, no como feature.",
     RGBColor(0x6C,0x63,0xFF), False),
    ("05", "MERIDIAN 75 Y NUEVOS LPs",
     "Fortalecer Meridian 75 para follow-ons en Tier 1. Cube Academy como canal de "
     "captación de capital inteligente. Alianzas Babson + Cornell.",
     RGBColor(0x2E,0xCC,0x71), False),
]
for i, (num, title, body, col, priority) in enumerate(priorities):
    col_idx = i % 3
    row_idx = i // 3
    x = 0.35 + col_idx * 4.32
    y = 2.0 + row_idx * 2.3
    if i >= 3:
        x = 0.35 + (i-3) * 6.5
        w = 6.2
    else:
        w = 4.0
    add_rect(slide, x, y, w, 2.1, fill=LIGHT_GRAY)
    add_rect(slide, x, y, w, 0.38, fill=col)
    add_text(slide, f"{num}  {title}", x+0.1, y+0.06, w-0.2, 0.3,
             font_size=9.5, bold=True, color=WHITE)
    add_text(slide, body, x+0.1, y+0.48, w-0.2, 1.52,
             font_size=9, bold=False, color=DARK_TEXT)
    if priority:
        add_rect(slide, x+w-1.1, y+1.72, 0.95, 0.25, fill=TEAL)
        add_text(slide, "PRIORITARIO", x+w-1.08, y+1.74, 0.91, 0.21,
                 font_size=6.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_footnote(slide, "Cube Ventures · Reporte de Inversionistas H2 2025 · Confidencial")


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 18 – CLOSING
# ═══════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, NAVY)

add_rect(slide, 0, 0, 0.35, 7.5, fill=TEAL)
add_rect(slide, 0, 6.8, 13.33, 0.7, fill=GOLD)

add_text(slide, "Gracias por su\nconfianza y paciencia.",
         1.0, 0.8, 10, 2.0,
         font_size=44, bold=True, color=WHITE)

divider_line(slide, 3.1, color=TEAL, x=1.0, w=8)

add_text(slide,
    "Este ha sido un ciclo exigente para el Venture Capital. Creemos que los mejores "
    "retornos se construyen en períodos donde hay disciplina, foco y capacidad de "
    "adaptación. Seguimos trabajando con transparencia, rigor y ambición.",
    1.0, 3.3, 10.5, 1.2, font_size=11, bold=False, color=MID_GRAY, italic=True)

add_text(slide, "Santiago Rojas Montoya", 1.0, 4.7, 8, 0.45,
         font_size=16, bold=True, color=WHITE)
add_text(slide, "Managing Director & Cofundador · Cube Ventures",
         1.0, 5.15, 8, 0.35, font_size=10, bold=False, color=TEAL)

# Contact strip
add_text(slide, "sr@cube.ventures", 1.0, 5.7, 5, 0.35,
         font_size=10, bold=False, color=GOLD)

add_text(slide, "Meridian 75  ·  Exit Accelerator  ·  Cube Academy  ·  Aceleradora",
         0.5, 6.88, 12.33, 0.4,
         font_size=9, bold=False, color=NAVY, align=PP_ALIGN.CENTER)

add_text(slide, "CUBE VENTURES", 9.5, 0.25, 3.5, 0.45,
         font_size=16, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
add_text(slide, "REPORTE H2 2025 · CONFIDENCIAL",
         9.5, 0.7, 3.5, 0.3,
         font_size=7.5, bold=False, color=MID_GRAY, align=PP_ALIGN.RIGHT)

# ─── Save ────────────────────────────────────────────────────────────────
output_path = "/home/user/hello-world/Cube_Ventures_Investor_Report_H2_2025.pptx"
prs.save(output_path)
print(f"✓ Saved: {output_path}")
print(f"  Slides: {len(prs.slides)}")
