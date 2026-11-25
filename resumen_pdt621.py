import os
import re
import subprocess
import sys
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def asegurar_librerias():
    """Verifica e instala automáticamente pdfplumber y openpyxl si no están disponibles."""
    faltantes = []
    for paquete in ("pdfplumber", "openpyxl"):
        try:
            __import__(paquete)
        except ImportError:
            faltantes.append(paquete)

    if not faltantes:
        return

    ventana = tk.Tk()
    ventana.title("Instalando componentes...")
    ventana.geometry("420x130")
    ventana.resizable(False, False)
    tk.Label(
        ventana,
        text="Instalando librerías necesarias por primera vez:\n" + ", ".join(faltantes) + "\n\nPor favor espera unos segundos...",
        font=("Segoe UI", 10),
        pady=20
    ).pack()
    ventana.update()

    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *faltantes])
        ventana.destroy()
    except Exception as e:
        ventana.destroy()
        root_err = tk.Tk()
        root_err.withdraw()
        messagebox.showerror(
            "Error de instalación",
            f"No se pudieron instalar las librerías ({', '.join(faltantes)}).\n"
            f"Verifica tu conexión a internet.\n\nDetalle: {e}"
        )
        root_err.destroy()
        sys.exit(1)


asegurar_librerias()

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule


# Lista completa y ordenada de los casilleros relevantes del PDT 621
CASILLEROS_ORDEN = [
    "100", "101", "102", "103", "154", "155", "160", "161", "162", "163",
    "106", "127", "105", "109", "112", "131",
    "107", "108", "156", "157", "110", "111", "113",
    "114", "115", "116", "117", "119", "120", "122", "178",
    "172", "173",
    "340", "341", "182",
    "301", "312", "380", "315", "894", "336",
    "140", "353", "302", "145", "351", "303", "184", "352", "304",
    "171", "168", "164", "179", "176", "165",
    "326", "327", "347", "305", "328",
    "681", "683", "682", "185", "342", "317",
    "187", "343", "319", "188", "344", "324"
]

DESCRIPCIONES_CASILLEROS = {
    "100": "Ventas Netas Gravadas 18% (Base)",
    "101": "Ventas Netas Gravadas 18% (IGV)",
    "102": "Descuentos / Devoluciones Ventas 18% (Base)",
    "103": "Descuentos / Devoluciones Ventas 18% (IGV)",
    "154": "Ventas Netas Ley 31556 - 10.5% (Base)",
    "155": "Ventas Netas Ley 31556 - 10.5% (IGV)",
    "160": "Ventas Bienes Ley 27037 (Base)",
    "161": "Ventas Bienes Ley 27037 (IGV)",
    "162": "Descuentos Ley 27037 (Base)",
    "163": "Descuentos Ley 27037 (IGV)",
    "106": "Exportaciones Facturadas",
    "127": "Exportaciones Embarcadas",
    "105": "Ventas No Gravadas (Con efecto en ratio)",
    "109": "Ventas No Gravadas (Sin efecto en ratio)",
    "112": "Otras Ventas (Art. 6 Regl.)",
    "131": "Total IGV Ventas",
    "107": "Compras Nac. Gravadas Exclusivas 18% (Base)",
    "108": "Compras Nac. Gravadas Exclusivas 18% (IGV)",
    "156": "Compras Nac. Ley 31556 - 10.5% (Base)",
    "157": "Compras Nac. Ley 31556 - 10.5% (IGV)",
    "110": "Compras Nac. Mixtas Grav/No Grav (Base)",
    "111": "Compras Nac. Mixtas Grav/No Grav (IGV)",
    "113": "Compras Nac. Destinadas a No Gravadas",
    "114": "Compras Imp. Gravadas Exclusivas 18% (Base)",
    "115": "Compras Imp. Gravadas Exclusivas 18% (IGV)",
    "116": "Compras Imp. Mixtas Grav/No Grav (Base)",
    "117": "Compras Imp. Mixtas Grav/No Grav (IGV)",
    "119": "Compras Imp. Destinadas a No Gravadas",
    "120": "Compras Internas No Gravadas",
    "122": "Compras Importadas No Gravadas",
    "178": "Total Crédito Fiscal Compras",
    "172": "Crédito Fiscal Especial",
    "173": "Coeficiente Prorrata IGV",
    "340": "IVAP Base",
    "341": "IVAP Tributo",
    "182": "Otros Créditos IVAP",
    "301": "Ingresos Netos Renta (Base)",
    "312": "Impuesto a la Renta Calculado",
    "380": "Coeficiente Renta",
    "315": "Porcentaje Renta",
    "894": "Coeficiente/Porcentaje Aplicado Cas. 312",
    "336": "Pagos a Cuenta en Exceso Ejercicio",
    "140": "IGV Impuesto Resultante / Saldo a Favor",
    "353": "IVAP Impuesto Resultante",
    "302": "Renta Impuesto Resultante / Saldo a Favor",
    "145": "IGV Saldo a Favor Período Anterior",
    "351": "IVAP Saldo a Favor Período Anterior",
    "303": "Renta Saldo a Favor Período Anterior",
    "184": "IGV Tributo a Pagar o Saldo a Favor",
    "352": "IVAP Tributo a Pagar",
    "304": "Renta Tributo a Pagar o Saldo a Favor",
    "171": "Percepciones Declaradas en el Período",
    "168": "Saldo Percepciones Períodos Anteriores",
    "164": "Saldo de Percepciones No Aplicadas",
    "179": "Retenciones Declaradas en el Período",
    "176": "Saldo Retenciones Períodos Anteriores",
    "165": "Saldo de Retenciones No Aplicadas",
    "326": "Retenciones Tercera Período",
    "327": "Retenciones Tercera Períodos Anteriores",
    "347": "Compensación Saldo Exportador (IGV)",
    "305": "Compensación Saldo Exportador (Renta)",
    "328": "Impuesto Temporal Activos Netos (ITAN)",
    "681": "Sub Total IGV",
    "683": "Sub Total IVAP",
    "682": "Sub Total Renta",
    "185": "Pagos Previos IGV",
    "342": "Pagos Previos IVAP",
    "317": "Pagos Previos Renta",
    "187": "Interés Moratorio IGV",
    "343": "Interés Moratorio IVAP",
    "319": "Interés Moratorio Renta",
    "188": "Total Deuda Tributaria IGV",
    "344": "Total Deuda Tributaria IVAP",
    "324": "Total Deuda Tributaria Renta"
}

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

SET_CASILLEROS = set(CASILLEROS_ORDEN)

RE_IMPORTE = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")
RE_PERIODO = re.compile(
    r"\bPer[ií]odo(?:\s+tributario)?\s*[:\-]?\s*"
    r"(?:(20\d{2})[ /-]?(\d{2})|(\d{2})[/-](20\d{2}))(?![\d/-])",
    re.IGNORECASE,
)

# Tolerancia vertical estricta (4.5 pt): en PDFs compactos de 2 páginas las filas están separadas
# por 10.0 pt, mientras que dentro de una misma fila código y valor tienen diferencia ~0.0 pt.
TOLERANCIA_Y = 4.5


def leer_periodo(texto):
    periodos = set()
    for anio, mes, mes_primero, anio_final in RE_PERIODO.findall(texto):
        mes = mes or mes_primero
        if not 1 <= int(mes) <= 12:
            raise ValueError("Mes inválido en el período del PDF.")
        periodos.add((anio or anio_final) + mes)
    if len(periodos) != 1:
        raise ValueError("No se encontró un período único válido.")
    return periodos.pop()


def fecha_pdt(valor):
    try:
        return datetime.strptime(valor, "%d/%m/%Y").date()
    except (TypeError, ValueError):
        return None


def indice_recomendado(opciones):
    """Sugiere la más reciente solo si las versiones se pueden distinguir."""
    fechas = [fecha_pdt(opcion["fecha_presentacion"]) for opcion in opciones]
    if (not fechas or None in fechas or len(set(fechas)) != len(fechas)
            or any(not opcion["orden"] or opcion["tipo_declaracion"] == "N/D" for opcion in opciones)):
        return None
    return fechas.index(max(fechas))


def es_constancia_o_no_pdt(texto_plano):
    """Detecta si el PDF es una Constancia de Presentación/Pago (NPS, 1662, etc.) en lugar de la Declaración PDT 621."""
    upper = texto_plano.upper()
    marcadores_constancia = [
        "CONSTANCIA DE PRESENTACI",
        "CONSTANCIA DE PAGO",
        "FORMULARIO 1662",
        "BOLETA DE PAGO",
        "NUMERO DE OPERACION SUNAT",
        "NÚMERO DE OPERACIÓN SUNAT",
        "RESUMEN DE TRANSACCIONES",
        "DETALLE DE TRIBUTOS PAGADOS"
    ]
    for m in marcadores_constancia:
        if m in upper:
            return True

    tiene_ventas = "IGV VENTAS" in upper
    tiene_compras = "IGV COMPRAS" in upper
    if not (tiene_ventas and tiene_compras):
        return True

    return False


def extraer_datos_pagina(pagina):
    """
    Extrae casilleros y el indicador de IGV Justo usando coordenadas (x0, top)
    con tolerancia vertical estricta (4.5 pt) para soportar tanto PDFs de 2 páginas
    (Aplicativo Desktop compacto) como de 3 páginas (Portal SUNAT).
    """
    words = pagina.extract_words(keep_blank_chars=False, use_text_flow=False)
    resultados = {}
    igv_justo_pag = None

    codigos_objs = []
    numeros_objs = []

    for w in words:
        txt = w["text"].strip()
        if txt in SET_CASILLEROS:
            codigos_objs.append(w)
        if RE_IMPORTE.fullmatch(txt):
            numeros_objs.append(w)

    # 1. Extraer casilleros en la misma franja horizontal estricta
    for cod_w in codigos_objs:
        cod = cod_w["text"].strip()
        cx1 = cod_w["x1"]
        cy_mid = (cod_w["top"] + cod_w["bottom"]) / 2.0

        limite_x_der = float("inf")
        for otro_w in codigos_objs:
            if otro_w is cod_w:
                continue
            oy_mid = (otro_w["top"] + otro_w["bottom"]) / 2.0
            if abs(oy_mid - cy_mid) <= TOLERANCIA_Y and otro_w["x0"] > cx1:
                if otro_w["x0"] < limite_x_der:
                    limite_x_der = otro_w["x0"]

        candidatos = []
        for num_w in numeros_objs:
            if num_w is cod_w:
                continue
            ny_mid = (num_w["top"] + num_w["bottom"]) / 2.0
            if abs(ny_mid - cy_mid) <= TOLERANCIA_Y:
                if cx1 - 1.0 <= num_w["x0"] < limite_x_der - 1.0:
                    candidatos.append(num_w)

        if candidatos:
            candidatos.sort(key=lambda item: item["x0"])
            val_txt = candidatos[0]["text"].strip().replace(",", "")
            try:
                val_num = float(val_txt)
                if cod not in resultados or resultados[cod] == 0.0:
                    resultados[cod] = val_num
            except ValueError:
                pass

    # 2. Detectar respuesta de IGV Justo por coordenadas en la misma fila de "(Si/No)?"
    for w in words:
        if "(Si/No)?" in w["text"] or "Justo)" in w["text"]:
            jy_mid = (w["top"] + w["bottom"]) / 2.0
            jx1 = w["x1"]
            respuestas = []
            for rw in words:
                ry_mid = (rw["top"] + rw["bottom"]) / 2.0
                rtxt = rw["text"].strip().upper()
                if abs(ry_mid - jy_mid) <= TOLERANCIA_Y and rw["x0"] > jx1:
                    if rtxt in ("SI", "SÍ", "NO"):
                        respuestas.append((rw["x0"], "SI" if rtxt.startswith("S") else "NO"))
            if respuestas:
                respuestas.sort(key=lambda x: x[0])
                igv_justo_pag = respuestas[-1][1]
                break

    return resultados, igv_justo_pag


def extraer_datos_pdt(ruta_pdf):
    """Extrae metadatos y casilleros de un PDF de Declaración PDT 621. Ignora constancias."""
    paginas_texto = []
    valores_totales = {}
    igv_justo_final = None

    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text(layout=True) or ""
            paginas_texto.append(texto)

        plano = " ".join(" ".join(paginas_texto).split())

        if es_constancia_o_no_pdt(plano):
            return {
                "ignorado": True,
                "archivo": os.path.basename(ruta_pdf),
                "motivo": "Constancia de presentación o documento distinto al detalle PDT 621"
            }

        rucs = set(re.findall(r"\bRUC\s*[:\-]?\s*((?:10|15|17|20)\d{9})\b", plano, re.I))
        if len(rucs) != 1:
            return {"ignorado": True, "archivo": os.path.basename(ruta_pdf), "motivo": "Sin RUC único legible"}

        try:
            periodo = leer_periodo(plano)
        except ValueError as e:
            return {"ignorado": True, "archivo": os.path.basename(ruta_pdf), "motivo": str(e)}

        for pagina in pdf.pages:
            vals_pag, justo_pag = extraer_datos_pagina(pagina)
            for k, v in vals_pag.items():
                if k not in valores_totales or valores_totales[k] == 0.0:
                    valores_totales[k] = v
            if justo_pag is not None:
                igv_justo_final = justo_pag

    if igv_justo_final is None:
        m_justo = re.search(r"pr[oó]rroga\s+del\s+pago\s+del\s+IGV.*?\?\s*(S[ií]|No)\b", plano, re.I)
        igv_justo_final = ("SI" if m_justo.group(1).lower().startswith("s") else "NO") if m_justo else "NO"

    # El Aplicativo Desktop parte la razón social en dos líneas: "INKA TOYS Período 202601 CORPORACION S.A.C."
    razon = re.search(r"Raz[oó]n\s+Social\s*[:\-]?\s+(.+?)\s+(?:Per[ií]odo\s*[:\-]?\s*\d{6}\s+([^\d]{1,80}?)\s+(?=N[uú]mero\s+de\s+Orden)"
                      r"|Per[ií]odo|N[uú]mero\s+de\s+Orden)", plano, re.I)
    orden = re.search(r"(?:N[uú]mero|N[.°ºo]+|Nro\.?)\s*(?:de\s+)?orden\s*[:#\-]?\s*(\d+)\b", plano, re.I)
    # SUNAT puede imprimir el valor entre las dos líneas de la etiqueta.
    encabezado = re.split(r"\bIGV\s+(?:CUENTA\s+PROPIA|VENTAS)\b", paginas_texto[0],
                          maxsplit=1, flags=re.I)[0]
    tipo_dec = re.search(r"\bTipo\s+de\s+(?:Declaraci[oó]n\s*)?[:\-]?\s*"
                         r"(Original|Rectificatoria|Sustitutoria)\b", encabezado, re.I)
    fecha_pres = re.search(r"\bFecha\s+de\s+(?:Presentaci[oó]n\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})"
                           r"|(\d{2}/\d{2}/\d{4})\s+Presentaci[oó]n)\b", encabezado, re.I)
    fecha = fecha_pdt(next(g for g in fecha_pres.groups() if g)) if fecha_pres else None

    # Las casillas en blanco en el PDT valen 0.0
    valores_completos = {cod: valores_totales.get(cod, 0.0) for cod in CASILLEROS_ORDEN}

    return {
        "ignorado": False,
        "archivo": os.path.basename(ruta_pdf),
        "ruc": rucs.pop(),
        "periodo": periodo,
        "razon_social": re.sub(r"\b0?621\b", "", " ".join(filter(None, razon.groups()))).strip() if razon else "CONTRIBUYENTE SIN NOMBRE",
        "orden": orden.group(1) if orden else None,
        "tipo_declaracion": tipo_dec.group(1).capitalize() if tipo_dec else "N/D",
        "fecha_presentacion": fecha.strftime("%d/%m/%Y") if fecha else "N/D",
        "igv_justo": igv_justo_final,
        "valores": valores_completos,
        "leidos_no_cero": sum(1 for v in valores_completos.values() if v != 0.0)
    }


def generar_excel_empresa(ruc, razon_social, anio, registros_mes, ruta_salida):
    """Genera el libro Excel con hoja RESUMEN separando 18% de Ley 31556 (10.5%) y hojas PDT 01..12."""
    wb = openpyxl.Workbook()
    ws_res = wb.active
    ws_res.title = "RESUMEN"
    ws_res.views.sheetView[0].showGridLines = True

    fill_navy = PatternFill("solid", fgColor="0B2545")

    fill_ventas_hdr = PatternFill("solid", fgColor="D9E8F5")
    fill_ventas_sub = PatternFill("solid", fgColor="EBF3FA")
    fill_ventas_ley = PatternFill("solid", fgColor="D6EAF8")
    fill_ventas_cell = PatternFill("solid", fgColor="F5F9FD")

    fill_compras_hdr = PatternFill("solid", fgColor="FDE49E")
    fill_compras_sub = PatternFill("solid", fgColor="FEF0C7")
    fill_compras_ley = PatternFill("solid", fgColor="FAD7A0")
    fill_compras_cell = PatternFill("solid", fgColor="FFF9E6")

    fill_imp_hdr = PatternFill("solid", fgColor="F5CBA7")
    fill_imp_sub = PatternFill("solid", fgColor="FAE5D3")
    fill_imp_cell = PatternFill("solid", fgColor="FDF2E9")

    fill_audit_hdr = PatternFill("solid", fgColor="E8DAEF")
    fill_audit_cell = PatternFill("solid", fgColor="F5EEF8")

    fill_sunat_hdr = PatternFill("solid", fgColor="FADBD8")
    fill_sunat_cell = PatternFill("solid", fgColor="FDEDEC")

    fill_perc_hdr = PatternFill("solid", fgColor="D4EFDF")
    fill_perc_cell = PatternFill("solid", fgColor="EAFAF1")

    font_title = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    font_hdr_white = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_hdr_dark = Font(name="Calibri", size=9, bold=True, color="0B2545")
    font_mes = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_num = Font(name="Calibri", size=10, bold=False, color="111111")
    font_total = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    thin_side = Side(border_style="thin", color="8C9BA5")
    med_navy = Side(border_style="medium", color="0B2545")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_hdr = Border(left=thin_side, right=thin_side, top=med_navy, bottom=thin_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    num_fmt = "#,##0.00"

    # Crear las 12 hojas mensuales (PDT 01 .. PDT 12)
    mapa_filas_casillero = {}
    for idx_cod, cod in enumerate(CASILLEROS_ORDEN, start=5):
        mapa_filas_casillero[cod] = idx_cod

    for mes_num in range(1, 13):
        nombre_hoja = f"PDT {mes_num:02d}"
        ws_m = wb.create_sheet(title=nombre_hoja)
        dato_mes = registros_mes.get(mes_num)

        ws_m["A1"] = "CAMPO / CASILLERO"
        ws_m["B1"] = "DESCRIPCIÓN SUNAT"
        ws_m["C1"] = "VALOR IMPORTADO"
        ws_m["D1"] = "ESTADO"
        ws_m["E1"] = "ARCHIVO PDF ORIGEN"

        ws_m["E2"] = dato_mes["archivo"] if dato_mes else "Sin PDT cargado"
        ws_m["E3"] = f"Orden: {dato_mes['orden'] or 'N/D'} | {dato_mes['tipo_declaracion']}" if dato_mes else "-"
        ws_m["E4"] = f"Presentado: {dato_mes['fecha_presentacion']}" if dato_mes else "-"

        for col_i in range(1, 6):
            c = ws_m.cell(row=1, column=col_i)
            c.fill = fill_navy
            c.font = font_hdr_white
            c.alignment = align_center
            c.border = border_cell

        ws_m["A2"], ws_m["B2"], ws_m["C2"] = "RUC", "Número de RUC", dato_mes["ruc"] if dato_mes else ruc
        ws_m["A3"], ws_m["B3"], ws_m["C3"] = "PERIODO", "Período Tributario", dato_mes["periodo"] if dato_mes else f"{anio}{mes_num:02d}"
        ws_m["A4"], ws_m["B4"], ws_m["C4"] = "IGV_JUSTO", "Acogimiento Prórr. IGV Justo", dato_mes["igv_justo"] if dato_mes else "NO"
        ws_m["D4"] = "PDT Cargado" if dato_mes else "Sin PDF (0.00)"

        for r_meta in range(2, 5):
            ws_m.cell(row=r_meta, column=1).font = Font(bold=True)
            ws_m.cell(row=r_meta, column=3).alignment = align_center
            for col_i in range(1, 5):
                ws_m.cell(row=r_meta, column=col_i).border = border_cell

        for cod in CASILLEROS_ORDEN:
            r_idx = mapa_filas_casillero[cod]
            val = dato_mes["valores"].get(cod, 0.0) if dato_mes else 0.0
            estado = "Importado" if dato_mes else "Sin PDF"

            c1 = ws_m.cell(row=r_idx, column=1, value=f"Casillero {cod}")
            c2 = ws_m.cell(row=r_idx, column=2, value=DESCRIPCIONES_CASILLEROS.get(cod, ""))
            c3 = ws_m.cell(row=r_idx, column=3, value=val)
            c4 = ws_m.cell(row=r_idx, column=4, value=estado)

            c3.number_format = num_fmt
            c3.alignment = align_right
            c4.alignment = align_center
            for cell_x in (c1, c2, c3, c4):
                cell_x.border = border_cell

        ws_m.column_dimensions["A"].width = 16
        ws_m.column_dimensions["B"].width = 45
        ws_m.column_dimensions["C"].width = 18
        ws_m.column_dimensions["D"].width = 16
        ws_m.column_dimensions["E"].width = 48

    # =========================================================================
    # HOJA RESUMEN UNIVERSAL CON SEPARACIÓN 18% vs LEY 31556 (10.5%)
    # Columnas A hasta AR (44 columnas). Compras separadas en Nacionales (J:T) e Importadas (U:AB).
    # =========================================================================
    ws_res.merge_cells("A1:AR1")
    ws_res["A1"] = f"{razon_social} (RUC: {ruc}) - RESUMEN TRIBUTARIO {anio}"
    ws_res["A1"].fill = fill_navy
    ws_res["A1"].font = font_title
    ws_res["A1"].alignment = align_center
    ws_res.row_dimensions[1].height = 26

    # Nivel 1 (Fila 2): Grandes Bloques
    ws_res.merge_cells("A2:A4")
    ws_res["A2"] = "MESES"

    ws_res.merge_cells("B2:I2")
    ws_res["B2"] = "VENTAS (IGV)"

    ws_res.merge_cells("J2:T2")
    ws_res["J2"] = "COMPRAS NACIONALES"

    ws_res.merge_cells("U2:AB2")
    ws_res["U2"] = "COMPRAS IMPORTADAS"

    ws_res.merge_cells("AC2:AC4")
    ws_res["AC2"] = "TOTAL COMPRAS"

    ws_res.merge_cells("AD2:AG2")
    ws_res["AD2"] = "RENTA Y AUDITORÍA BASE"

    ws_res.merge_cells("AH2:AJ2")
    ws_res["AH2"] = "DETERMINACIÓN SUNAT"

    ws_res.merge_cells("AK2:AM2")
    ws_res["AK2"] = "PERCEPCIONES Y RETENCIONES"

    ws_res.merge_cells("AN2:AR2")
    ws_res["AN2"] = "DECLARACIÓN UTILIZADA"

    # Nivel 2 (Fila 3): Sub-bloques separados por tasa
    ws_res.merge_cells("B3:D3")
    ws_res["B3"] = "GRAVADAS NETAS 18%"
    ws_res.merge_cells("E3:G3")
    ws_res["E3"] = "LEY 31556 (10.5%)"
    ws_res["H3"] = "NO GRAV / EXP"
    ws_res.merge_cells("I3:I4")
    ws_res["I3"] = "TOTAL VENTAS"

    ws_res.merge_cells("J3:L3")
    ws_res["J3"] = "GRAVADAS 18%"
    ws_res.merge_cells("M3:O3")
    ws_res["M3"] = "LEY 31556 (10.5%)"
    ws_res.merge_cells("P3:R3")
    ws_res["P3"] = "MIXTAS (GRAV Y NO GRAV)"
    ws_res["S3"] = "NO GRAVADAS"
    ws_res.merge_cells("T3:T4")
    ws_res["T3"] = "TOTAL NACIONALES"

    ws_res.merge_cells("U3:W3")
    ws_res["U3"] = "GRAVADAS 18%"
    ws_res.merge_cells("X3:Z3")
    ws_res["X3"] = "MIXTAS (GRAV Y NO GRAV)"
    ws_res["AA3"] = "NO GRAVADAS"
    ws_res.merge_cells("AB3:AB4")
    ws_res["AB3"] = "TOTAL IMPORTADAS"

    ws_res["AD3"] = "INGRESOS NETOS"
    ws_res["AE3"] = "DIF. IGV vs RENTA"
    ws_res["AF3"] = "COEF / %"
    ws_res["AG3"] = "IGV JUSTO"

    ws_res["AH3"] = "IGV RESULT."
    ws_res["AI3"] = "IR RESULT."
    ws_res["AJ3"] = "PAGOS REALIZ."

    ws_res["AK3"] = "PERCEP. PERÍODO"
    ws_res["AL3"] = "PERCEP. NO APLIC."
    ws_res["AM3"] = "RETENC. NO APLIC."

    ws_res.merge_cells("AN3:AN4")
    ws_res["AN3"] = "ESTADO"
    ws_res.merge_cells("AO3:AO4")
    ws_res["AO3"] = "PDF DE ORIGEN"

    for col, titulo in (("AP", "TIPO"), ("AQ", "N.º DE ORDEN"), ("AR", "PRESENTACIÓN")):
        ws_res.merge_cells(f"{col}3:{col}4")
        ws_res[f"{col}3"] = titulo

    # Nivel 3 (Fila 4): Detalle de columnas
    sub_headers = {
        "B4": "SUBTOTAL", "C4": "IGV 18%", "D4": "TOTAL",
        "E4": "SUBTOTAL", "F4": "IGV 10.5%", "G4": "TOTAL",
        "H4": "SUBTOTAL",
        "J4": "SUBTOTAL", "K4": "IGV 18%", "L4": "TOTAL",
        "M4": "SUBTOTAL", "N4": "IGV 10.5%", "O4": "TOTAL",
        "P4": "SUBTOTAL", "Q4": "IGV", "R4": "TOTAL",
        "S4": "SUBTOTAL",
        "U4": "SUBTOTAL", "V4": "IGV 18%", "W4": "TOTAL",
        "X4": "SUBTOTAL", "Y4": "IGV", "Z4": "TOTAL",
        "AA4": "SUBTOTAL",
        "AD4": "CAS. 301", "AE4": "ALERTA", "AF4": "APLICADO", "AG4": "PRÓRR.",
        "AH4": "CAS. 184", "AI4": "CAS. 304", "AJ4": "MANUAL",
        "AK4": "CAS. 171", "AL4": "CAS. 164", "AM4": "CAS. 165"
    }
    for coord, texto in sub_headers.items():
        ws_res[coord] = texto

    for r in range(2, 5):
        ws_res.row_dimensions[r].height = 21
        for c_idx in range(1, 45):
            cell = ws_res.cell(row=r, column=c_idx)
            cell.border = border_hdr
            cell.alignment = align_center
            if c_idx == 1 or c_idx >= 40:
                cell.fill = fill_navy
                cell.font = font_hdr_white
            elif 2 <= c_idx <= 9:
                if r >= 3 and 5 <= c_idx <= 7:
                    cell.fill = fill_ventas_ley
                else:
                    cell.fill = fill_ventas_hdr if r == 2 else fill_ventas_sub
                cell.font = font_hdr_dark
            elif 10 <= c_idx <= 20 or c_idx == 29:
                if r >= 3 and 13 <= c_idx <= 15:
                    cell.fill = fill_compras_ley
                else:
                    cell.fill = fill_compras_hdr if r == 2 else fill_compras_sub
                cell.font = font_hdr_dark
            elif 21 <= c_idx <= 28:
                cell.fill = fill_imp_hdr if r == 2 else fill_imp_sub
                cell.font = font_hdr_dark
            elif 30 <= c_idx <= 33:
                cell.fill = fill_audit_hdr
                cell.font = font_hdr_dark
            elif 34 <= c_idx <= 36:
                cell.fill = fill_sunat_hdr
                cell.font = font_hdr_dark
            else:
                cell.fill = fill_perc_hdr
                cell.font = font_hdr_dark

    def ref_cas(hoja, cod):
        return f"{hoja}!C{mapa_filas_casillero[cod]}"

    for mes_num in range(1, 13):
        r = mes_num + 4
        h = f"'PDT {mes_num:02d}'"
        ws_res.row_dimensions[r].height = 19

        c_mes = ws_res.cell(row=r, column=1, value=MESES_ES[mes_num])
        c_mes.fill = fill_navy
        c_mes.font = font_mes
        c_mes.alignment = align_left
        c_mes.border = border_cell

        dato = registros_mes.get(mes_num)

        # VENTAS 18% (B, C, D)
        ws_res[f"B{r}"] = f"={ref_cas(h, '100')}-{ref_cas(h, '102')}"
        ws_res[f"C{r}"] = f"={ref_cas(h, '101')}-{ref_cas(h, '103')}"
        ws_res[f"D{r}"] = f"=B{r}+C{r}"

        # VENTAS LEY 31556 - 10.5% (E, F, G)
        ws_res[f"E{r}"] = f"={ref_cas(h, '154')}"
        ws_res[f"F{r}"] = f"={ref_cas(h, '155')}"
        ws_res[f"G{r}"] = f"=E{r}+F{r}"

        # VENTAS NO GRAVADAS / EXPORTACIONES (H)
        ws_res[f"H{r}"] = f"={ref_cas(h, '105')}+{ref_cas(h, '106')}+{ref_cas(h, '109')}+{ref_cas(h, '112')}"

        # TOTAL VENTAS (I)
        ws_res[f"I{r}"] = f"=D{r}+G{r}+H{r}"

        # COMPRAS NACIONALES 18% (J, K, L)
        ws_res[f"J{r}"] = f"={ref_cas(h, '107')}"
        ws_res[f"K{r}"] = f"={ref_cas(h, '108')}"
        ws_res[f"L{r}"] = f"=J{r}+K{r}"

        # COMPRAS NACIONALES LEY 31556 - 10.5% (M, N, O)
        ws_res[f"M{r}"] = f"={ref_cas(h, '156')}"
        ws_res[f"N{r}"] = f"={ref_cas(h, '157')}"
        ws_res[f"O{r}"] = f"=M{r}+N{r}"

        # COMPRAS NACIONALES MIXTAS (P, Q, R)
        ws_res[f"P{r}"] = f"={ref_cas(h, '110')}"
        ws_res[f"Q{r}"] = f"={ref_cas(h, '111')}"
        ws_res[f"R{r}"] = f"=P{r}+Q{r}"

        # COMPRAS NACIONALES NO GRAVADAS (S): destinadas a no gravadas + internas no gravadas
        ws_res[f"S{r}"] = f"={ref_cas(h, '113')}+{ref_cas(h, '120')}"

        # TOTAL NACIONALES (T)
        ws_res[f"T{r}"] = f"=L{r}+O{r}+R{r}+S{r}"

        # COMPRAS IMPORTADAS 18% (U, V, W)
        ws_res[f"U{r}"] = f"={ref_cas(h, '114')}"
        ws_res[f"V{r}"] = f"={ref_cas(h, '115')}"
        ws_res[f"W{r}"] = f"=U{r}+V{r}"

        # COMPRAS IMPORTADAS MIXTAS (X, Y, Z)
        ws_res[f"X{r}"] = f"={ref_cas(h, '116')}"
        ws_res[f"Y{r}"] = f"={ref_cas(h, '117')}"
        ws_res[f"Z{r}"] = f"=X{r}+Y{r}"

        # COMPRAS IMPORTADAS NO GRAVADAS (AA): destinadas a no gravadas + importadas no gravadas
        ws_res[f"AA{r}"] = f"={ref_cas(h, '119')}+{ref_cas(h, '122')}"

        # TOTAL IMPORTADAS (AB) y TOTAL COMPRAS (AC)
        ws_res[f"AB{r}"] = f"=W{r}+Z{r}+AA{r}"
        ws_res[f"AC{r}"] = f"=T{r}+AB{r}"

        # RENTA Y AUDITORÍA (AD, AE, AF, AG)
        ws_res[f"AD{r}"] = f"={ref_cas(h, '301')}"
        ws_res[f"AE{r}"] = f"=(B{r}+E{r}+H{r})-AD{r}"
        ws_res[f"AF{r}"] = f"=IF({ref_cas(h, '894')}>0,{ref_cas(h, '894')},IF({ref_cas(h, '380')}>0,{ref_cas(h, '380')},{ref_cas(h, '315')}/100))"
        ws_res[f"AG{r}"] = f"={h}!C4"

        # SUNAT (AH, AI, AJ)
        ws_res[f"AH{r}"] = f"={ref_cas(h, '184')}"
        ws_res[f"AI{r}"] = f"={ref_cas(h, '304')}"
        ws_res[f"AJ{r}"] = None

        # PERCEPCIONES Y RETENCIONES (AK, AL, AM)
        ws_res[f"AK{r}"] = f"={ref_cas(h, '171')}"
        ws_res[f"AL{r}"] = f"={ref_cas(h, '164')}"
        ws_res[f"AM{r}"] = f"={ref_cas(h, '165')}"

        # Identificación de la declaración utilizada (AN a AR)
        if dato:
            ws_res[f"AN{r}"] = "OK"
            ws_res[f"AO{r}"] = dato["archivo"]
            ws_res[f"AP{r}"] = dato["tipo_declaracion"]
            ws_res[f"AQ{r}"] = dato["orden"] or "N/D"
            ws_res[f"AR{r}"] = fecha_pdt(dato["fecha_presentacion"]) or "N/D"
        else:
            ws_res[f"AN{r}"] = "Sin PDF"
            ws_res[f"AO{r}"] = "-"

        for c_idx in range(2, 45):
            cell = ws_res.cell(row=r, column=c_idx)
            cell.border = border_cell
            cell.font = font_num

            if c_idx == 32:
                cell.number_format = "0.0000"
                cell.alignment = align_right
            elif c_idx == 44:
                cell.number_format = "dd/mm/yyyy"
                cell.alignment = align_center
            elif c_idx in (33, 40, 42, 43):
                cell.alignment = align_center
            elif c_idx == 41:
                cell.alignment = align_left
            else:
                cell.number_format = num_fmt
                cell.alignment = align_right

            if 2 <= c_idx <= 9:
                cell.fill = fill_ventas_cell
            elif 10 <= c_idx <= 20 or c_idx == 29:
                cell.fill = fill_compras_cell
            elif 21 <= c_idx <= 28:
                cell.fill = fill_imp_cell
            elif 30 <= c_idx <= 33:
                cell.fill = fill_audit_cell
            elif 34 <= c_idx <= 36:
                cell.fill = fill_sunat_cell
            elif 37 <= c_idx <= 39:
                cell.fill = fill_perc_cell
            else:
                cell.fill = PatternFill("solid", fgColor="F2F4F4")

    # Fila 17: TOTAL ANUAL
    ws_res.row_dimensions[17].height = 22
    c_tot = ws_res.cell(row=17, column=1, value="TOTAL ANUAL")
    c_tot.fill = fill_navy
    c_tot.font = font_total
    c_tot.alignment = align_left
    c_tot.border = border_cell

    for c_idx in range(2, 45):
        col_let = get_column_letter(c_idx)
        cell = ws_res.cell(row=17, column=c_idx)
        cell.fill = fill_navy
        cell.font = font_total
        cell.border = border_cell

        if c_idx in (32, 33) or c_idx >= 41:
            cell.value = "-"
            cell.alignment = align_center
        elif c_idx == 40:
            cell.value = f"{len(registros_mes)}/12 PDFs"
            cell.alignment = align_center
        else:
            cell.value = f"=SUM({col_let}5:{col_let}16)"
            cell.number_format = num_fmt
            cell.alignment = align_right

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006", bold=True)
    ws_res.conditional_formatting.add(
        "AE5:AE16",
        FormulaRule(formula=["ROUND(AE5,2)<>0"], stopIfTrue=True, fill=red_fill, font=red_font)
    )

    anchos = {
        "A": 14,
        "B": 13, "C": 12, "D": 13, "E": 13, "F": 12, "G": 13, "H": 13, "I": 14,
        "J": 13, "K": 12, "L": 13, "M": 13, "N": 12, "O": 13, "P": 13, "Q": 11, "R": 13, "S": 13, "T": 14,
        "U": 13, "V": 12, "W": 13, "X": 13, "Y": 11, "Z": 13, "AA": 13, "AB": 14, "AC": 14,
        "AD": 14, "AE": 14, "AF": 10, "AG": 10,
        "AH": 13, "AI": 13, "AJ": 13,
        "AK": 13, "AL": 13, "AM": 13,
        "AN": 11, "AO": 42, "AP": 18, "AQ": 18, "AR": 18
    }
    for col_l, w in anchos.items():
        ws_res.column_dimensions[col_l].width = w

    ws_res.freeze_panes = "B5"
    wb.save(ruta_salida)


class AppGeneradorPDT:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador Universal PDT 621 -> Excel")
        self.root.geometry("700x460")
        self.root.resizable(False, False)

        self.carpeta_origen = tk.StringVar()
        self.carpeta_destino = tk.StringVar()

        self._construir_ui()

    def _construir_ui(self):
        pad = {"padx": 14, "pady": 8}

        lbl_titulo = tk.Label(
            self.root,
            text="Consolidador Automático de PDT 621 (SUNAT)",
            font=("Segoe UI", 13, "bold"),
            fg="#0B2545"
        )
        lbl_titulo.pack(pady=(14, 4))

        lbl_sub = tk.Label(
            self.root,
            text="Selecciona la carpeta con tus PDFs (ignora constancias y agrupa por RUC y Año)",
            font=("Segoe UI", 9),
            fg="#444444"
        )
        lbl_sub.pack(pady=(0, 10))

        frm_origen = tk.Frame(self.root)
        frm_origen.pack(fill="x", **pad)
        tk.Label(frm_origen, text="Carpeta con PDTs (PDF):", width=22, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Entry(frm_origen, textvariable=self.carpeta_origen, width=50).pack(side="left", padx=6)
        tk.Button(frm_origen, text="Explorar...", command=self.seleccionar_origen).pack(side="left")

        frm_dest = tk.Frame(self.root)
        frm_dest.pack(fill="x", **pad)
        tk.Label(frm_dest, text="Carpeta de salida Excel:", width=22, anchor="w", font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Entry(frm_dest, textvariable=self.carpeta_destino, width=50).pack(side="left", padx=6)
        tk.Button(frm_dest, text="Explorar...", command=self.seleccionar_destino).pack(side="left")

        btn_procesar = tk.Button(
            self.root,
            text="GENERAR EXCEL(S) RESUMEN",
            font=("Segoe UI", 10, "bold"),
            bg="#0B2545",
            fg="white",
            height=2,
            width=32,
            command=self.procesar
        )
        btn_procesar.pack(pady=12)

        tk.Label(
            self.root,
            text="La U es la U",
            font=("Segoe UI", 8),
            fg="#707070"
        ).pack(side="bottom", anchor="e", padx=14, pady=(0, 4))

        self.log_box = tk.Text(self.root, height=13, width=82, font=("Consolas", 9), state="disabled", bg="#F8F9FA")
        self.log_box.pack(fill="both", expand=True, padx=14, pady=(0, 4))

    def log(self, mensaje):
        self.log_box.config(state="normal")
        self.log_box.insert("end", mensaje + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        self.root.update_idletasks()

    def seleccionar_origen(self):
        ruta = filedialog.askdirectory(title="Selecciona la carpeta que contiene los PDFs del PDT 621")
        if ruta:
            self.carpeta_origen.set(ruta)
            if not self.carpeta_destino.get():
                self.carpeta_destino.set(ruta)

    def seleccionar_destino(self):
        ruta = filedialog.askdirectory(title="Selecciona dónde guardar los archivos Excel")
        if ruta:
            self.carpeta_destino.set(ruta)

    def elegir_declaracion(self, opciones):
        dialogo = tk.Toplevel(self.root)
        dialogo.title("Elegir declaración del mes (Original / Rectificatoria)")
        dialogo.transient(self.root)
        dato = opciones[0]
        ttk.Label(dialogo, text=f"RUC {dato['ruc']}   ·   Período {dato['periodo']}",
                  padding=(12, 10)).pack(anchor="w")
        sugerido = indice_recomendado(opciones)
        aviso = ("Se ha seleccionado la presentación más reciente. Confírmala o elige otra."
                 if sugerido is not None else
                 "Hay fechas repetidas o datos incompletos. Elige la declaración que deseas usar.")
        ttk.Label(dialogo, text=aviso,
                  padding=(12, 0)).pack(anchor="w")
        tabla = ttk.Treeview(
            dialogo,
            columns=("pdf", "tipo", "orden", "fecha"),
            show="headings",
            selectmode="browse",
            height=min(10, len(opciones))
        )
        tabla.heading("pdf", text="Archivo PDF")
        tabla.heading("tipo", text="Tipo")
        tabla.heading("orden", text="N.º Orden")
        tabla.heading("fecha", text="Presentación")
        tabla.column("pdf", width=320)
        tabla.column("tipo", width=110)
        tabla.column("orden", width=110)
        tabla.column("fecha", width=100)
        for i, candidato in enumerate(opciones):
            tabla.insert(
                "", "end", iid=str(i),
                values=(
                    candidato["archivo"],
                    candidato["tipo_declaracion"],
                    candidato["orden"] or "N/D",
                    candidato["fecha_presentacion"]
                )
            )
        tabla.pack(fill="both", expand=True, padx=12, pady=10)
        barra = ttk.Frame(dialogo, padding=(12, 0, 12, 10))
        barra.pack(fill="x")
        seleccionado = None

        def aceptar():
            nonlocal seleccionado
            if tabla.selection():
                seleccionado = opciones[int(tabla.selection()[0])]
                dialogo.destroy()

        boton = ttk.Button(barra, text="Usar PDF seleccionado", command=aceptar, state="disabled")
        boton.pack(side="right")
        tabla.bind("<<TreeviewSelect>>", lambda evento: boton.configure(state="normal"))
        if sugerido is not None:
            fila = str(sugerido)
            tabla.selection_set(fila)
            tabla.focus(fila)
            tabla.see(fila)
            boton.configure(state="normal")
        ttk.Button(barra, text="Cancelar lote", command=dialogo.destroy).pack(side="left")
        dialogo.wait_visibility()
        dialogo.grab_set()
        tabla.focus_set()
        self.root.wait_window(dialogo)
        return seleccionado

    def procesar(self):
        origen = self.carpeta_origen.get().strip()
        destino = self.carpeta_destino.get().strip()

        if not origen or not os.path.isdir(origen):
            messagebox.showwarning("Atención", "Selecciona una carpeta de origen válida.")
            return
        if not destino or not os.path.isdir(destino):
            messagebox.showwarning("Atención", "Selecciona una carpeta de salida válida.")
            return

        pdfs = [os.path.join(origen, f) for f in os.listdir(origen) if f.lower().endswith(".pdf")]
        if not pdfs:
            messagebox.showinfo("Sin archivos", "No se encontraron archivos PDF en la carpeta seleccionada.")
            return

        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

        self.log(f"Escaneando {len(pdfs)} archivo(s) PDF en la carpeta...")

        candidatos = {}
        for ruta_pdf in sorted(pdfs):
            try:
                datos = extraer_datos_pdt(ruta_pdf)
            except Exception as error:
                self.log(f"[ERROR LECTURA] {os.path.basename(ruta_pdf)}: {error}")
                continue

            if datos.get("ignorado"):
                self.log(f"[IGNORADO] {datos['archivo']} -> {datos['motivo']}")
                continue

            clave = (datos["ruc"], datos["periodo"])
            candidatos.setdefault(clave, []).append(datos)
            self.log(
                f"[PDT OK] {datos['archivo']} | RUC {datos['ruc']} | Período {datos['periodo']} ({datos['tipo_declaracion']})"
            )

        agrupado = {}
        nombres_ruc = {}
        for (ruc, periodo), opciones in sorted(candidatos.items()):
            if len(opciones) == 1:
                elegido = opciones[0]
            else:
                elegido = self.elegir_declaracion(opciones)
                if elegido is None:
                    self.log("Proceso cancelado por el usuario.")
                    return
                self.log(f"[ELEGIDO] {periodo} -> {elegido['archivo']} (Orden {elegido['orden'] or 'N/D'})")

            agrupado.setdefault((ruc, periodo[:4]), {})[int(periodo[4:])] = elegido
            if elegido["razon_social"] != "CONTRIBUYENTE SIN NOMBRE":
                nombres_ruc[ruc] = elegido["razon_social"]

        if not agrupado:
            messagebox.showwarning("Sin declaraciones", "No se encontró ningún detalle de PDT 621 válido (solo constancias u otros PDFs).")
            return

        generados = 0
        for (ruc, anio), meses_dict in agrupado.items():
            razon = nombres_ruc.get(ruc, f"RUC_{ruc}")
            razon_limpia = re.sub(r'[\\/*?:"<>|]', "", razon).strip()
            nombre_archivo = f"RESUMEN_PDT621_{anio}_{ruc}_{razon_limpia[:30]}.xlsx"
            ruta_xlsx = os.path.join(destino, nombre_archivo)

            generar_excel_empresa(ruc, razon, anio, meses_dict, ruta_xlsx)
            generados += 1
            self.log(f"\n>>> EXCEL GENERADO: {nombre_archivo}")

            for m_num, d in sorted(meses_dict.items()):
                v = d["valores"]
                base_igv = (v["100"] - v["102"]) + v["154"] + v["105"] + v["106"] + v["109"] + v["112"]
                base_ir = v["301"]
                if round(base_igv - base_ir, 2) != 0:
                    self.log(
                        f"    [ALERTA {MESES_ES[m_num]}] Descuadre Base IGV ({base_igv:,.2f}) vs Renta ({base_ir:,.2f}) = {base_igv - base_ir:,.2f}"
                    )

        messagebox.showinfo("Proceso Completado", f"Se generaron {generados} archivo(s) Excel correctamente.")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppGeneradorPDT(root)
    root.mainloop()
