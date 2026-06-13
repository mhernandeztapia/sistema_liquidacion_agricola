"""
generador_pdf.py Generará la liquidación de sueldo en formato PDF.
Utiliza ReportLab (Platypus) para un documento profesional.
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from datetime import datetime


# Paleta de colores corporativos
VERDE_OSCURO  = colors.HexColor("#1B4332")
VERDE_MEDIO   = colors.HexColor("#2D6A4F")
VERDE_CLARO   = colors.HexColor("#52B788")
GRIS_CLARO    = colors.HexColor("#F1F5F2")
GRIS_TEXTO    = colors.HexColor("#6C757D")
ROJO          = colors.HexColor("#DC3545")
AMARILLO      = colors.HexColor("#FFC107")


def _fmt(monto: int) -> str:
    """Formatea un entero como moneda chilena: $1.234.567"""
    return f"${monto:,.0f}".replace(",", ".")


def generar_pdf_liquidacion(datos: dict, resumen: dict, ruta_salida: str):
    doc = SimpleDocTemplate(
        ruta_salida,
        pagesize=letter,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    # Estilos personalizados
    s_titulo = ParagraphStyle("titulo",
        fontSize=20, textColor=colors.white, alignment=TA_CENTER,
        fontName="Helvetica-Bold", spaceAfter=4)
    s_subtitulo = ParagraphStyle("subtitulo",
        fontSize=11, textColor=VERDE_CLARO, alignment=TA_CENTER,
        fontName="Helvetica", spaceAfter=0)
    s_seccion = ParagraphStyle("seccion",
        fontSize=11, textColor=VERDE_OSCURO, fontName="Helvetica-Bold",
        spaceBefore=12, spaceAfter=4)
    s_nota = ParagraphStyle("nota",
        fontSize=8, textColor=GRIS_TEXTO, alignment=TA_CENTER,
        fontName="Helvetica-Oblique")

    story = []

    # ─── ENCABEZADO ───────────────────────────────────────────
    encabezado_data = [
        [Paragraph("🌿  SERVICARG", s_titulo)],
        [Paragraph("Liquidación de Sueldo", s_subtitulo)],
        [Paragraph(f"{datos['mes']} {datos['anio']}", s_subtitulo)],
    ]
    enc_table = Table(encabezado_data, colWidths=["100%"])
    enc_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), VERDE_OSCURO),
        ("ROUNDEDCORNERS", [6, 6, 6, 6]),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(enc_table)
    story.append(Spacer(1, 0.4*cm))

    # Fecha de emisión
    fecha_emision = Paragraph(
        f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        s_nota
    )
    story.append(fecha_emision)
    story.append(Spacer(1, 0.5*cm))

    # ─── DATOS DEL TRABAJADOR ─────────────────────────────────
    story.append(Paragraph("DATOS DEL TRABAJADOR", s_seccion))
    story.append(HRFlowable(width="100%", thickness=1, color=VERDE_MEDIO, spaceAfter=6))

    info_data = [
        ["Nombre",          datos["nombre"],    "RUT",       datos["rut"] or "—"],
        ["Empresa/Faena",   datos["empresa"] or "—", "Labor", datos["labor"]],
        ["Sueldo Base",     _fmt(datos["sueldo_base"]),
         "Días trabajados", str(datos["dias_trabajados"])],
    ]
    col_w = [3.5*cm, 7*cm, 3.5*cm, 5.5*cm]
    info_table = Table(info_data, colWidths=col_w, hAlign="LEFT")
    info_table.setStyle(TableStyle([
        ("FONTNAME",    (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("FONTNAME",    (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",    (2, 0), (2, -1), "Helvetica-Bold"),
        ("BACKGROUND",  (0, 0), (-1, -1), GRIS_CLARO),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [GRIS_CLARO, colors.white]),
        ("GRID",        (0, 0), (-1, -1), 0.5, colors.HexColor("#DEE2E6")),
        ("TOPPADDING",  (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 0.4*cm))

    # ─── HABERES ──────────────────────────────────────────────
    story.append(Paragraph("HABERES", s_seccion))
    story.append(HRFlowable(width="100%", thickness=1, color=VERDE_MEDIO, spaceAfter=6))

    haberes_rows = [
        ["Concepto", "Imponible", "No Imponible", "Total"],
    ]
    haberes_rows.append([
        "Sueldo Base (proporcional)",
        _fmt(resumen["sueldo_base_proporcional"]), "—",
        _fmt(resumen["sueldo_base_proporcional"])
    ])
    if resumen["valor_horas_extra"]:
        haberes_rows.append([
            f"Horas Extras ({datos['horas_extra']} hrs × $5.000)",
            _fmt(resumen["valor_horas_extra"]), "—",
            _fmt(resumen["valor_horas_extra"])
        ])
    if resumen["bono_colacion"]:
        haberes_rows.append([
            "Bono Colación", "—",
            _fmt(resumen["bono_colacion"]),
            _fmt(resumen["bono_colacion"])
        ])
    if resumen["bono_movilizacion"]:
        haberes_rows.append([
            "Bono Movilización", "—",
            _fmt(resumen["bono_movilizacion"]),
            _fmt(resumen["bono_movilizacion"])
        ])
    if resumen["bono_asistencia"]:
        haberes_rows.append([
            "Bono Asistencia Perfecta", "—",
            _fmt(resumen["bono_asistencia"]),
            _fmt(resumen["bono_asistencia"])
        ])
    if resumen["asig_familiar"]:
        haberes_rows.append([
            "Asignación Familiar", "—",
            _fmt(resumen["asig_familiar"]),
            _fmt(resumen["asig_familiar"])
        ])

    # Fila de total haberes
    haberes_rows.append([
        "TOTAL HABERES",
        _fmt(resumen["total_imponible"]),
        _fmt(resumen["total_haberes"] - resumen["total_imponible"]),
        _fmt(resumen["total_haberes"])
    ])

    hab_col_w = [8*cm, 3*cm, 3.5*cm, 3.5*cm]
    hab_table = Table(haberes_rows, colWidths=hab_col_w, hAlign="LEFT")
    hab_table.setStyle(TableStyle([
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME",    (0, -1),(-1,-1), "Helvetica-Bold"),
        ("FONTNAME",    (0, 1), (-1, -2),"Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("BACKGROUND",  (0, 0), (-1, 0), VERDE_MEDIO),
        ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
        ("BACKGROUND",  (0, -1),(-1, -1), VERDE_CLARO),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, GRIS_CLARO]),
        ("ALIGN",       (1, 0), (-1, -1), "RIGHT"),
        ("GRID",        (0, 0), (-1, -1), 0.5, colors.HexColor("#DEE2E6")),
        ("TOPPADDING",  (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0),(-1,-1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(hab_table)
    story.append(Spacer(1, 0.4*cm))

    # ─── DESCUENTOS ───────────────────────────────────────────
    story.append(Paragraph("DESCUENTOS", s_seccion))
    story.append(HRFlowable(width="100%", thickness=1, color=ROJO, spaceAfter=6))

    desc_rows = [["Concepto", "Tasa", "Monto"]]
    desc_rows.append([
        "AFP",
        f"{resumen['tasa_afp']*100:.2f}%",
        _fmt(resumen["desc_afp"])
    ])
    desc_rows.append([
        "Salud (Isapre/Fonasa)",
        f"{resumen['tasa_salud']*100:.2f}%",
        _fmt(resumen["desc_salud"])
    ])
    desc_rows.append([
        "Seguro de Cesantía",
        f"{resumen['tasa_cesantia']*100:.2f}%",
        _fmt(resumen["desc_cesantia"])
    ])
    if resumen["desc_prestamo"]:
        desc_rows.append(["Préstamo Empresa",  "—", _fmt(resumen["desc_prestamo"])])
    if resumen["desc_adelanto"]:
        desc_rows.append(["Adelanto de Sueldo","—", _fmt(resumen["desc_adelanto"])])
    if resumen["desc_otros"]:
        desc_rows.append(["Otros Descuentos",  "—", _fmt(resumen["desc_otros"])])
    desc_rows.append(["TOTAL DESCUENTOS", "", _fmt(resumen["total_descuentos"])])

    desc_col_w = [9*cm, 3*cm, 3.5*cm]  # Solo 3 columnas
    desc_table = Table(desc_rows, colWidths=desc_col_w, hAlign="LEFT")
    desc_table.setStyle(TableStyle([
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME",    (0, -1),(-1,-1),"Helvetica-Bold"),
        ("FONTNAME",    (0, 1), (-1,-2),"Helvetica"),
        ("FONTSIZE",    (0, 0), (-1,-1), 9),
        ("BACKGROUND",  (0, 0), (-1, 0), ROJO),
        ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
        ("BACKGROUND",  (0,-1), (-1,-1), colors.HexColor("#FFEBEE")),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[colors.white, GRIS_CLARO]),
        ("ALIGN",       (1, 0), (-1,-1),"RIGHT"),
        ("GRID",        (0, 0), (-1,-1), 0.5, colors.HexColor("#DEE2E6")),
        ("TOPPADDING",  (0, 0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING", (0, 0), (-1,-1), 6),
    ]))
    story.append(desc_table)
    story.append(Spacer(1, 0.5*cm))

    # ─── LÍQUIDO A PAGAR ──────────────────────────────────────
    liquido_data = [[
        Paragraph("LÍQUIDO A PAGAR", ParagraphStyle(
            "liq_lbl", fontSize=14, textColor=colors.white,
            fontName="Helvetica-Bold", alignment=TA_LEFT)),
        Paragraph(_fmt(resumen["liquido_a_pagar"]), ParagraphStyle(
            "liq_monto", fontSize=16, textColor=AMARILLO,
            fontName="Helvetica-Bold", alignment=TA_RIGHT))
    ]]
    liq_table = Table(liquido_data, colWidths=[9*cm, 9.5*cm])
    liq_table.setStyle(TableStyle([
        ("BACKGROUND",     (0, 0), (-1,-1), VERDE_OSCURO),
        ("TOPPADDING",     (0, 0), (-1,-1), 12),
        ("BOTTOMPADDING",  (0, 0), (-1,-1), 12),
        ("LEFTPADDING",    (0, 0), (-1,-1), 14),
        ("RIGHTPADDING",   (0, 0), (-1,-1), 14),
        ("ROUNDEDCORNERS", [6, 6, 6, 6]),
    ]))
    story.append(liq_table)
    story.append(Spacer(1, 0.8*cm))

    # ─── FIRMA ────────────────────────────────────────────────
    firma_data = [
        ["_________________________", "   ", "_________________________"],
        ["Trabajador", "   ", "Empleador / Empresa"],
        [datos["nombre"], "   ", datos.get("empresa", "Servicarg")],
    ]
    firma_table = Table(firma_data, colWidths=[7*cm, 2*cm, 7*cm], hAlign="CENTER")
    firma_table.setStyle(TableStyle([
        ("ALIGN",        (0, 0), (-1,-1), "CENTER"),
        ("FONTSIZE",     (0, 0), (-1,-1), 8),
        ("FONTNAME",     (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTNAME",     (0, 2), (-1, 2), "Helvetica"),
        ("TEXTCOLOR",    (0, 1), (-1, 1), VERDE_OSCURO),
        ("TEXTCOLOR",    (0, 2), (-1, 2), GRIS_TEXTO),
        ("TOPPADDING",   (0, 0), (-1,-1), 3),
    ]))
    story.append(firma_table)

    # ─── PIE DE PÁGINA ────────────────────────────────────────
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GRIS_TEXTO))
    story.append(Paragraph(
        "Documento generado por Sistema de Liquidación SERVICARG  •  "
        f"Emitido el {datetime.now().strftime('%d/%m/%Y')}",
        s_nota
    ))

    doc.build(story)
