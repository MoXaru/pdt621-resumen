"""Pruebas con datos ficticios (no hay PDFs reales en el repositorio)."""
import os
import tempfile
import unittest
from unittest import mock

import openpyxl

import resumen_pdt621 as app


def palabra(texto, x0, top):
    return {"text": texto, "x0": x0, "x1": x0 + 6 * len(texto), "top": top, "bottom": top + 8}


class PaginaFalsa:
    def __init__(self, texto, palabras):
        self.texto = texto
        self.palabras = palabras

    def extract_text(self, layout=True):
        return self.texto

    def extract_words(self, **kwargs):
        return self.palabras


class PdfFalso:
    def __init__(self, paginas):
        self.pages = paginas

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


# Encabezado como lo imprime el Aplicativo Desktop: la razón social se parte en dos líneas.
TEXTO_DEMO = """
621 RUC 20123456789
Razón Social EMPRESA DEMO        Período 202601
             COMERCIAL S.A.C.
Número de Orden 1000000001   Fecha de 16/02/2026
                             Presentación
Tipo de Declaración Original
IGV CUENTA PROPIA
IGV VENTAS
IGV COMPRAS
"""

PALABRAS_DEMO = [
    # Ventas netas 18 %: 100 base / 101 IGV
    palabra("100", 300, 100), palabra("1,000.00", 330, 100),
    palabra("101", 420, 100), palabra("180.00", 450, 100),
    # Compras importadas 18 %: 114 base / 115 IGV
    palabra("114", 300, 200), palabra("500.00", 330, 200),
    palabra("115", 420, 200), palabra("90.00", 450, 200),
    # Compras nacionales 18 % en cero
    palabra("107", 300, 190), palabra("0.00", 330, 190),
    # Renta
    palabra("301", 300, 300), palabra("1000.00", 330, 300),
    # Pregunta del IGV Justo
    palabra("(Si/No)?", 300, 400), palabra("Si", 360, 400),
]


def leer_demo():
    pdf = PdfFalso([PaginaFalsa(TEXTO_DEMO, PALABRAS_DEMO)])
    with mock.patch.object(app.pdfplumber, "open", return_value=pdf):
        return app.extraer_datos_pdt("PDT621-01-2026.pdf")


class PruebasLectura(unittest.TestCase):
    def test_periodo(self):
        self.assertEqual(app.leer_periodo("Período 202603"), "202603")
        self.assertEqual(app.leer_periodo("Periodo tributario: 03/2026"), "202603")
        with self.assertRaises(ValueError):
            app.leer_periodo("Período 202613")

    def test_constancia_se_ignora(self):
        self.assertTrue(app.es_constancia_o_no_pdt("CONSTANCIA DE PRESENTACIÓN IGV VENTAS IGV COMPRAS"))
        self.assertFalse(app.es_constancia_o_no_pdt("IGV VENTAS ... IGV COMPRAS"))

    def test_datos_del_encabezado(self):
        d = leer_demo()
        self.assertFalse(d["ignorado"])
        self.assertEqual(d["ruc"], "20123456789")
        self.assertEqual(d["periodo"], "202601")
        self.assertEqual(d["razon_social"], "EMPRESA DEMO COMERCIAL S.A.C.")
        self.assertEqual(d["orden"], "1000000001")
        self.assertEqual(d["tipo_declaracion"], "Original")
        self.assertEqual(d["fecha_presentacion"], "16/02/2026")
        self.assertEqual(d["igv_justo"], "SI")

    def test_casilleros(self):
        v = leer_demo()["valores"]
        self.assertEqual(v["100"], 1000.0)
        self.assertEqual(v["101"], 180.0)
        self.assertEqual(v["114"], 500.0)
        self.assertEqual(v["115"], 90.0)
        self.assertEqual(v["107"], 0.0)
        self.assertEqual(v["156"], 0.0)  # casillero en blanco = 0


class PruebasExcel(unittest.TestCase):
    def test_compras_nacionales_e_importadas_separadas(self):
        dato = leer_demo()
        with tempfile.TemporaryDirectory() as carpeta:
            ruta = os.path.join(carpeta, "demo.xlsx")
            app.generar_excel_empresa("20123456789", dato["razon_social"], "2026", {1: dato}, ruta)
            wb = openpyxl.load_workbook(ruta)
        ws = wb["RESUMEN"]
        fila = app.CASILLEROS_ORDEN.index("114") + 5
        self.assertEqual(ws["J2"].value, "COMPRAS NACIONALES")
        self.assertEqual(ws["U2"].value, "COMPRAS IMPORTADAS")
        self.assertEqual(ws["U5"].value, f"='PDT 01'!C{fila}")
        self.assertNotIn(f"C{fila}", ws["J5"].value)
        self.assertEqual(ws["AC5"].value, "=T5+AB5")
        self.assertEqual(wb["PDT 01"][f"C{fila}"].value, 500.0)


if __name__ == "__main__":
    unittest.main()
