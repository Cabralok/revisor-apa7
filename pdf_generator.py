from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from io import BytesIO
import base64
import os

# Colores Profes al Rescate
AZUL_NEON = colors.HexColor('#1E90FF')
AZUL_OSCURO = colors.HexColor('#0A1628')
GRIS_OSCURO = colors.HexColor('#2D3748')
GRIS_MEDIO = colors.HexColor('#718096')
GRIS_CLARO = colors.HexColor('#F7FAFC')
VERDE = colors.HexColor('#38A169')
ROJO = colors.HexColor('#E53E3E')
AMARILLO = colors.HexColor('#D69E2E')
BLANCO = colors.white

LOGO_PATH = os.path.join(os.path.dirname(__file__), 'logo.png')

def generar_pdf(datos):
    buffer = BytesIO()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2.5*cm,
        bottomMargin=2*cm
    )

    estilos = getSampleStyleSheet()
    
    # Estilos personalizados
    estilo_titulo = ParagraphStyle(
        'Titulo',
        parent=estilos['Normal'],
        fontSize=20,
        fontName='Helvetica-Bold',
        textColor=AZUL_OSCURO,
        spaceAfter=4,
        alignment=TA_LEFT
    )
    estilo_subtitulo = ParagraphStyle(
        'Subtitulo',
        parent=estilos['Normal'],
        fontSize=11,
        fontName='Helvetica',
        textColor=GRIS_MEDIO,
        spaceAfter=20,
        alignment=TA_LEFT
    )
    estilo_dimension = ParagraphStyle(
        'Dimension',
        parent=estilos['Normal'],
        fontSize=12,
        fontName='Helvetica-Bold',
        textColor=AZUL_OSCURO,
        spaceBefore=16,
        spaceAfter=8,
        borderPad=6,
    )
    estilo_item = ParagraphStyle(
        'Item',
        parent=estilos['Normal'],
        fontSize=9,
        fontName='Helvetica',
        textColor=GRIS_OSCURO,
        spaceAfter=2,
        leftIndent=10
    )
    estilo_sugerencia = ParagraphStyle(
        'Sugerencia',
        parent=estilos['Normal'],
        fontSize=8,
        fontName='Helvetica-Oblique',
        textColor=GRIS_MEDIO,
        spaceAfter=6,
        leftIndent=22
    )
    estilo_puntaje_global = ParagraphStyle(
        'PuntajeGlobal',
        parent=estilos['Normal'],
        fontSize=36,
        fontName='Helvetica-Bold',
        textColor=AZUL_NEON,
        alignment=TA_CENTER,
        spaceAfter=4
    )
    estilo_label_puntaje = ParagraphStyle(
        'LabelPuntaje',
        parent=estilos['Normal'],
        fontSize=10,
        fontName='Helvetica',
        textColor=GRIS_MEDIO,
        alignment=TA_CENTER,
        spaceAfter=20
    )
    estilo_footer = ParagraphStyle(
        'Footer',
        parent=estilos['Normal'],
        fontSize=8,
        fontName='Helvetica',
        textColor=GRIS_MEDIO,
        alignment=TA_CENTER
    )

    elementos = []

    # ---- ENCABEZADO ----
    es_comparacion = "version1" in datos

    # Logo + título en tabla
    logo_celda = ""
    if os.path.exists(LOGO_PATH):
        from reportlab.platypus import Image
        logo = Image(LOGO_PATH, width=1.8*cm, height=1.8*cm)
        logo_celda = logo
    else:
        logo_celda = Paragraph("🎓", estilo_titulo)

    if es_comparacion:
        titulo_texto = "Comparación de Versiones APA 7"
        subtitulo_texto = f"{datos['version1']['nombre']}  →  {datos['version2']['nombre']}"
    else:
        titulo_texto = "Informe de Revisión APA 7"
        subtitulo_texto = datos.get("nombre_archivo", "Documento analizado")

    tabla_header = Table(
        [[logo_celda, [
            Paragraph(titulo_texto, estilo_titulo),
            Paragraph(subtitulo_texto, estilo_subtitulo)
        ]]],
        colWidths=[2.2*cm, None]
    )
    tabla_header.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (0, 0), 0),
        ('RIGHTPADDING', (0, 0), (0, 0), 8),
    ]))
    elementos.append(tabla_header)
    elementos.append(HRFlowable(width="100%", thickness=2, color=AZUL_NEON, spaceAfter=16))

    # ---- PUNTAJE GLOBAL ----
    if es_comparacion:
        p1 = datos['version1']['puntaje']
        p2 = datos['version2']['puntaje']
        diff = datos['diferencia_puntaje']
        signo = "+" if diff >= 0 else ""
        color_diff = VERDE if diff >= 0 else ROJO

        tabla_puntajes = Table(
            [[
                Paragraph(f"{p1}/100", estilo_puntaje_global),
                Paragraph("→", ParagraphStyle('arrow', fontSize=24, fontName='Helvetica', textColor=GRIS_MEDIO, alignment=TA_CENTER)),
                Paragraph(f"{p2}/100", estilo_puntaje_global),
                Paragraph(f"{signo}{diff}", ParagraphStyle('diff', fontSize=28, fontName='Helvetica-Bold', textColor=color_diff, alignment=TA_CENTER))
            ],
            [
                Paragraph("Entrega 1", estilo_label_puntaje),
                Paragraph("", estilo_label_puntaje),
                Paragraph("Entrega 2", estilo_label_puntaje),
                Paragraph("Diferencia", estilo_label_puntaje),
            ]],
            colWidths=[4*cm, 2*cm, 4*cm, 4*cm]
        )
        tabla_puntajes.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND', (0, 0), (-1, -1), GRIS_CLARO),
            ('ROWBACKGROUNDS', (0, 0), (-1, -1), [GRIS_CLARO, GRIS_CLARO]),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#E2E8F0')),
            ('ROUNDEDCORNERS', [8]),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        elementos.append(tabla_puntajes)
    else:
        puntaje = datos.get("puntaje", 0)
        color_puntaje = VERDE if puntaje >= 80 else (AMARILLO if puntaje >= 60 else ROJO)
        estilo_puntaje_global.textColor = color_puntaje
        
        tabla_puntaje = Table(
            [[Paragraph(f"{puntaje}/100", estilo_puntaje_global)],
             [Paragraph("Puntaje de Cumplimiento APA 7", estilo_label_puntaje)]],
            colWidths=[None]
        )
        tabla_puntaje.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('BACKGROUND', (0, 0), (-1, -1), GRIS_CLARO),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#E2E8F0')),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))
        elementos.append(tabla_puntaje)

    elementos.append(Spacer(1, 20))

    # ---- CONTENIDO POR DIMENSIÓN ----
    if es_comparacion:
        elementos += _generar_comparacion(datos, estilo_dimension, estilo_item, estilo_sugerencia)
    else:
        elementos += _generar_dimensiones(datos, estilo_dimension, estilo_item, estilo_sugerencia)

    # ---- FOOTER ----
    elementos.append(Spacer(1, 20))
    elementos.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#E2E8F0'), spaceBefore=8))
    elementos.append(Paragraph("Profes al Rescate • Revisión APA 7ª edición • @profes_al_rescate", estilo_footer))

    doc.build(elementos)
    return buffer.getvalue()


def _icono_estado(estado):
    if estado == "ok":
        return "✅"
    elif estado == "error":
        return "❌"
    else:
        return "⚠️"


def _color_estado(estado):
    if estado == "ok":
        return VERDE
    elif estado == "error":
        return ROJO
    else:
        return AMARILLO


def _generar_dimensiones(datos, estilo_dim, estilo_item, estilo_sugerencia):
    elementos = []
    for dim in datos.get("dimensiones", []):
        puntaje_dim = dim.get("puntaje", 0)
        color_p = VERDE if puntaje_dim >= 80 else (AMARILLO if puntaje_dim >= 60 else ROJO)
        
        # Encabezado de dimensión
        tabla_dim = Table(
            [[
                Paragraph(dim["nombre"], estilo_dim),
                Paragraph(f"{puntaje_dim}/100", ParagraphStyle('pd', fontSize=14, fontName='Helvetica-Bold', textColor=color_p, alignment=TA_RIGHT, spaceBefore=16))
            ]],
            colWidths=[None, 2.5*cm]
        )
        tabla_dim.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
            ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.HexColor('#E2E8F0')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elementos.append(tabla_dim)

        for it in dim.get("items", []):
            icono = _icono_estado(it["estado"])
            color = _color_estado(it["estado"])
            
            estilo_item_col = ParagraphStyle(
                f'item_{it["estado"]}',
                parent=estilo_item,
                textColor=color
            )
            
            texto = f"{icono}  <b>{it['titulo']}</b>: {it['detalle']}"
            elementos.append(Paragraph(texto, estilo_item_col))
            
            if it.get("sugerencia"):
                elementos.append(Paragraph(f"→ {it['sugerencia']}", estilo_sugerencia))

        elementos.append(Spacer(1, 8))
    
    return elementos


def _generar_comparacion(datos, estilo_dim, estilo_item, estilo_sugerencia):
    elementos = []
    
    estilo_corregido = ParagraphStyle('corr', parent=estilo_item, textColor=VERDE)
    estilo_persiste = ParagraphStyle('pers', parent=estilo_item, textColor=ROJO)
    estilo_regresion = ParagraphStyle('regr', parent=estilo_item, textColor=AMARILLO)
    
    for dim in datos.get("dimensiones", []):
        p1 = dim.get("puntaje_v1", 0)
        p2 = dim.get("puntaje_v2", 0)
        diff = p2 - p1
        signo = "+" if diff >= 0 else ""
        color_diff = VERDE if diff >= 0 else ROJO

        tabla_dim = Table(
            [[
                Paragraph(dim["nombre"], estilo_dim),
                Paragraph(f"{p1}→{p2} ({signo}{diff})", ParagraphStyle('pd2', fontSize=11, fontName='Helvetica-Bold', textColor=color_diff, alignment=TA_RIGHT, spaceBefore=16))
            ]],
            colWidths=[None, 3.5*cm]
        )
        tabla_dim.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
            ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.HexColor('#E2E8F0')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elementos.append(tabla_dim)

        corregidos = dim.get("corregidos", [])
        persisten = dim.get("persisten", [])
        regresiones = dim.get("regresiones", [])

        if corregidos:
            elementos.append(Paragraph("<b>✅ Corregidos</b>", ParagraphStyle('cat', parent=estilo_item, textColor=VERDE, fontName='Helvetica-Bold', spaceBefore=4)))
            for c in corregidos:
                elementos.append(Paragraph(f"  ✅ {c['titulo']}: {c['detalle']}", estilo_corregido))

        if persisten:
            elementos.append(Paragraph("<b>❌ Persisten</b>", ParagraphStyle('cat2', parent=estilo_item, textColor=ROJO, fontName='Helvetica-Bold', spaceBefore=4)))
            for p in persisten:
                elementos.append(Paragraph(f"  ❌ {p['titulo']}: {p['detalle']}", estilo_persiste))
                if p.get("sugerencia"):
                    elementos.append(Paragraph(f"  → {p['sugerencia']}", estilo_sugerencia))

        if regresiones:
            elementos.append(Paragraph("<b>⚠️ Errores nuevos (regresiones)</b>", ParagraphStyle('cat3', parent=estilo_item, textColor=AMARILLO, fontName='Helvetica-Bold', spaceBefore=4)))
            for r in regresiones:
                elementos.append(Paragraph(f"  ⚠️ {r['titulo']}: {r['detalle']}", estilo_regresion))
                if r.get("sugerencia"):
                    elementos.append(Paragraph(f"  → {r['sugerencia']}", estilo_sugerencia))

        elementos.append(Spacer(1, 8))

    return elementos
