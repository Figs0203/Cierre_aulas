"""
Pruebas unitarias del parser de Notas (Interactiva).

Cubren casos atípicos de estructura por curso:
- Archivos SIN columna 'Sección' (p. ej. cursos CMI).
- Columnas de estado/semáforo con rellenos ámbar/dorado que NO deben marcar a
  los estudiantes como excluidos.
- Reconciliación de la exclusión real (naranja) en columnas de identidad.
"""

from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill

from parsing.notas_parser import parse_notas_archivo


def _build_notas(path: Path, headers, rows, fills=None):
    """Construye un archivo de Notas sintético.

    Args:
        path: Ruta destino.
        headers: Encabezados (fila 1).
        rows: Filas de datos.
        fills: Dict opcional {(row_idx_1based, col_idx_1based): "RRGGBB"}.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calificaciones"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    if fills:
        for (r, c), color in fills.items():
            ws.cell(r, c).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    wb.save(path)
    wb.close()


class TestNotasParserEstructurasAtipicas(unittest.TestCase):

    def test_notas_sin_columna_seccion(self):
        """Sin columna 'Sección', la sección debe quedar en None (no leer una nota)."""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_CMI205.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Calificación Módulo 1", "Calificación Módulo 2", "Calificación Módulo 3",
                "Promedio curso completo", "", "N. Semáforo de secuencia",
            ]
            rows = [
                ["1000000001", "jperez", "Perez Castro", "Juan Camilo", 4.5, 4.0, 4.8, 4.4, None, "Verde"],
                ["1000000002", "mgomez", "Gomez Toro", "Maria Alejandra", 5.0, 4.8, 5.0, 4.9, None, "Verde"],
            ]
            _build_notas(path, headers, rows)

            estudiantes, secciones = parse_notas_archivo(path)
            self.assertEqual(len(estudiantes), 2)
            for es in estudiantes:
                self.assertIsNone(es.seccion, "No debe interpretarse una calificación como sección")
            # Los módulos sí se detectan
            self.assertEqual(set(estudiantes[0].calificaciones_modulos.keys()), {"modulo_1", "modulo_2", "modulo_3"})
            # 'Promedio curso completo' se reconoce como nota final oficial
            self.assertIsNotNone(estudiantes[0].nota_final, "Debe detectar 'Promedio curso completo' como nota final")
            self.assertEqual(estudiantes[0].nota_final.value, 4.4)
            self.assertFalse(any(es.is_orange_fill for es in estudiantes))

    def test_semaforo_ambar_no_marca_excluidos(self):
        """Un relleno ámbar/dorado en la columna de semáforo NO debe excluir al estudiante."""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_CMI205.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Calificación Módulo 1", "Calificación Módulo 2", "Calificación Módulo 3",
                "Promedio curso completo", "", "N. Semáforo de secuencia",
            ]
            rows = [
                [f"100000000{i}", f"user{i}", f"Apellido {i}", f"Nombre {i}", 4.5, 4.0, 4.8, 4.4, None, "Amarillo"]
                for i in range(1, 6)
            ]
            # Relleno ámbar FFC000 en la última columna (semáforo) de TODAS las filas
            fills = {(r, 10): "FFC000" for r in range(2, 7)}
            _build_notas(path, headers, rows, fills=fills)

            estudiantes, _ = parse_notas_archivo(path)
            self.assertEqual(len(estudiantes), 5)
            self.assertEqual(
                sum(1 for es in estudiantes if es.is_orange_fill), 0,
                "El semáforo ámbar no debe marcar al estudiante como excluido",
            )

    def test_naranja_en_identidad_si_marca_excluido(self):
        """El naranja de exclusión en columnas de identidad SÍ debe detectarse."""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_CMI205.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Calificación Módulo 1", "Calificación Módulo 2", "Calificación Módulo 3",
                "Promedio curso completo", "", "N. Semáforo de secuencia",
            ]
            rows = [
                ["1000000001", "user1", "Uno", "Estudiante", 4.5, 4.0, 4.8, 4.4, None, "Amarillo"],
                ["1000000002", "user2", "Dos", "Excluido", 2.0, 1.5, 2.0, 2.0, None, "Rojo"],
            ]
            fills = {
                # Semáforo ámbar para ambos (no debe excluir)
                (2, 10): "FFC000",
                (3, 10): "FFC000",
                # Naranja real sobre la identidad de la fila 3
                (3, 1): "FF8C00", (3, 2): "FF8C00", (3, 3): "FF8C00", (3, 4): "FF8C00",
            }
            _build_notas(path, headers, rows, fills=fills)

            estudiantes, _ = parse_notas_archivo(path)
            excluidos = [es for es in estudiantes if es.is_orange_fill]
            self.assertEqual(len(excluidos), 1)
            self.assertEqual(excluidos[0].first_name.value, "Excluido")


    def test_selecciona_hoja_con_notas_entre_varias(self):
        """Si hay varias hojas, debe elegir la que tiene calificaciones, no la del listado general.

        Caso real CMI: 'Calificaciones' solo tiene identidad + sección (otro curso) y
        'CURSO REALIZADO' tiene módulos y promedio. Debe usar 'CURSO REALIZADO'.
        """
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "CMI120_ECONOMIA_2.xlsx"
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "Calificaciones"
            ws1.append(["OrgDefinedId", "Username", "Last Name", "First Name", "Sección"])
            ws1.append(["9000000001", "otro1", "Apellido", "Nombre", "FI0171 CLASE 4164 E. FINANZAS"])
            ws2 = wb.create_sheet("CURSO REALIZADO")
            ws2.append([
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Calificación Módulo 1", "Calificación Módulo 2", "Calificación Módulo 3",
                "Promedio curso completo", "", "N.º", "Semáforo de secuencia",
            ])
            ws2.append(["1000000001", "juan.perez1", "Perez Castro", "Juan Camilo", 5.0, 4.5, 4.8, 4.8, None, 1, "Verde"])
            wb.save(path)
            wb.close()

            from parsing.notas_parser import _select_notas_sheet

            wb2 = openpyxl.load_workbook(path)
            self.assertEqual(_select_notas_sheet(wb2), "CURSO REALIZADO")
            wb2.close()

            estudiantes, _ = parse_notas_archivo(path)
            self.assertEqual(len(estudiantes), 1)
            self.assertEqual(estudiantes[0].first_name.value, "Juan Camilo")
            self.assertIsNotNone(estudiantes[0].nota_final)
            self.assertEqual(set(estudiantes[0].calificaciones_modulos.keys()), {"modulo_1", "modulo_2", "modulo_3"})


    def test_detecta_modulo_sin_palabra_modulo(self):
        """Regresión CMI121: 'Cuestionario 1' (sin la palabra 'módulo') debe ser módulo 1.

        Además, la columna de identidad puede llamarse 'Documento' en lugar de
        'OrgDefinedId'.
        """
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "CMI121_DERECHO_2.xlsx"
            headers = [
                "Documento", "Username", "Apellidos", "Nombres", "Sección",
                "Cuestionario 1: estrategia y herramientas de búsqueda",
                "Cuestionario módulo 2: evaluación y organización de la información",
                "Cuestionario módulo 3: uso ético de la información",
                "Nota final", "", "N.º", "Semáforo de secuencia",
            ]
            rows = [["1000000001", "user1", "Apellido", "Nombre", "DR1004-1914-DERECHO",
                     5.0, 3.8, 4.9, 4.4, None, 1, "Verde"]]
            _build_notas(path, headers, rows)

            estudiantes, _ = parse_notas_archivo(path)
            es = estudiantes[0]
            # Identidad por columna 'Documento'
            self.assertEqual(es.org_defined_id.value, "1000000001")
            # Módulo 1 detectado aunque el encabezado no diga 'módulo'
            self.assertEqual(
                {k: v.value for k, v in es.calificaciones_modulos.items()},
                {"modulo_1": 5.0, "modulo_2": 3.8, "modulo_3": 4.9},
            )
            self.assertEqual(es.nota_final.value, 4.4)


class TestRespaldoPosicionalNotas(unittest.TestCase):
    """Si los encabezados de notas no se reconocen, se infieren por posición.

    Regla: de izquierda a derecha, las columnas con notas son módulo 1..N y la
    última es la nota final. El respaldo valida los DATOS de cada columna para no
    confundir identidad, fechas, porcentajes ni estados con notas.
    """

    def test_infiere_modulos_y_final_con_encabezados_atipicos(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_ABBUEI999.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name", "Sección",
                "Puntaje A", "Puntaje B", "Puntaje C", "Total",
            ]
            rows = [
                ["1000000001", "user1", "Apellido", "Nombre", "ABBUEI999-001",
                 4.5, 4.0, 4.2, 4.3],
                ["1000000002", "user2", "Apellido", "Nombre", "ABBUEI999-001",
                 3.8, 4.1, 3.9, 4.0],
            ]
            _build_notas(path, headers, rows)

            diagnostico = {}
            estudiantes, _ = parse_notas_archivo(path, diagnostics=diagnostico)
            es = estudiantes[0]
            self.assertEqual(
                {k: v.value for k, v in es.calificaciones_modulos.items()},
                {"modulo_1": 4.5, "modulo_2": 4.0, "modulo_3": 4.2},
            )
            self.assertEqual(es.nota_final.value, 4.3)
            self.assertEqual(diagnostico["metodo_modulos"], "posicional")
            self.assertEqual(diagnostico["nota_final"]["columna"], 9)
            self.assertEqual(
                {k: v["columna"] for k, v in diagnostico["modulos"].items()},
                {"1": 6, "2": 7, "3": 8},
            )
            self.assertTrue(diagnostico["avisos"])

    def test_no_confunde_identidad_ni_fechas_con_notas(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_ABBUEI999.xlsx"
            headers = [
                "Documento", "Username", "Apellidos", "Nombres", "Sección",
                "Fecha de matrícula", "Nota A", "Nota B", "Nota C",
            ]
            rows = [
                ["1000000001", "user1", "Apellido", "Nombre", "ABBUEI999-001",
                 datetime(2026, 3, 1), 4.5, 4.0, 4.2],
                ["1000000002", "user2", "Apellido", "Nombre", "ABBUEI999-001",
                 datetime(2026, 3, 2), 3.8, 4.1, 3.9],
            ]
            _build_notas(path, headers, rows)

            diagnostico = {}
            estudiantes, _ = parse_notas_archivo(path, diagnostics=diagnostico)
            es = estudiantes[0]
            # Columnas 7 y 8 = módulos 1 y 2; columna 9 = nota final.
            self.assertEqual(
                {k: v.value for k, v in es.calificaciones_modulos.items()},
                {"modulo_1": 4.5, "modulo_2": 4.0},
            )
            self.assertEqual(es.nota_final.value, 4.2)
            self.assertEqual(diagnostico["nota_final"]["columna"], 9)
            # El documento (1000000001) y la fecha NUNCA se toman como notas.
            self.assertNotIn(1, [v["columna"] for v in diagnostico["modulos"].values()])
            self.assertNotIn(6, [v["columna"] for v in diagnostico["modulos"].values()])

    def test_una_sola_columna_numerica_es_la_nota_final(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_ABBUEI999.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Calificación X",
            ]
            rows = [["1000000001", "user1", "Apellido", "Nombre", 4.7]]
            _build_notas(path, headers, rows)

            diagnostico = {}
            estudiantes, _ = parse_notas_archivo(path, diagnostics=diagnostico)
            es = estudiantes[0]
            self.assertEqual(es.calificaciones_modulos, {})
            self.assertEqual(es.nota_final.value, 4.7)
            self.assertEqual(diagnostico["metodo_modulos"], "no_detectado")
            self.assertEqual(diagnostico["metodo_nota_final"], "posicional")

    def test_respeta_los_encabezados_cuando_si_se_reconocen(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_ABBUEI999.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Calificación Módulo 1", "Calificación Módulo 2",
                "Promedio curso completo",
            ]
            rows = [["1000000001", "user1", "Apellido", "Nombre", 4.5, 4.0, 4.2]]
            _build_notas(path, headers, rows)

            diagnostico = {}
            estudiantes, _ = parse_notas_archivo(path, diagnostics=diagnostico)
            self.assertEqual(diagnostico["metodo_modulos"], "encabezados")
            self.assertEqual(diagnostico["nota_final"]["columna"], 7)
            self.assertEqual(
                {k: v.value for k, v in estudiantes[0].calificaciones_modulos.items()},
                {"modulo_1": 4.5, "modulo_2": 4.0},
            )

    def test_completa_por_posicion_los_modulos_que_faltan(self):
        """Si solo algunos módulos tienen encabezado, se completan los faltantes."""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_ABBUEI999.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Calificación Módulo 1", "Puntaje B", "Puntaje C", "Total",
            ]
            rows = [
                ["1000000001", "user1", "Apellido", "Nombre", 4.5, 4.0, 4.2, 4.3],
                ["1000000002", "user2", "Apellido", "Nombre", 3.8, 4.1, 3.9, 4.0],
            ]
            _build_notas(path, headers, rows)

            diagnostico = {}
            estudiantes, _ = parse_notas_archivo(path, diagnostics=diagnostico)
            es = estudiantes[0]
            self.assertEqual(
                {k: v.value for k, v in es.calificaciones_modulos.items()},
                {"modulo_1": 4.5, "modulo_2": 4.0, "modulo_3": 4.2},
            )
            self.assertEqual(es.nota_final.value, 4.3)
            self.assertEqual(diagnostico["metodo_modulos"], "mixto")
            self.assertEqual(diagnostico["nota_final"]["columna"], 8)

    def test_no_usa_una_columna_simbolica_como_nota_final_por_seguridad(self):
        """Una columna 'Aprobado/No aplica' puede ser un estado, no una nota."""
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_ABBUEI999.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Nota A", "Nota B", "Resultado",
            ]
            rows = [
                ["1000000001", "user1", "Apellido", "Nombre", 4.5, 4.0, "Aprobado"],
                ["1000000002", "user2", "Apellido", "Nombre", 3.8, 4.1, "No aplica"],
            ]
            _build_notas(path, headers, rows)

            diagnostico = {}
            estudiantes, _ = parse_notas_archivo(path, diagnostics=diagnostico)
            es = estudiantes[0]
            # La última columna NUMÉRICA es la nota final; la simbólica no se usa.
            self.assertEqual(
                {k: v.value for k, v in es.calificaciones_modulos.items()},
                {"modulo_1": 4.5},
            )
            self.assertEqual(es.nota_final.value, 4.0)
            self.assertTrue(
                any("simbólica" in aviso for aviso in diagnostico["avisos"]),
                diagnostico["avisos"],
            )

    def test_descarta_columnas_con_formato_de_porcentaje(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Notas_ABBUEI999.xlsx"
            headers = [
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Asistencia", "Nota A", "Nota B",
            ]
            rows = [
                ["1000000001", "user1", "Apellido", "Nombre", 0.8, 4.5, 4.0],
                ["1000000002", "user2", "Apellido", "Nombre", 0.9, 3.8, 4.1],
            ]
            _build_notas(path, headers, rows)

            wb = openpyxl.load_workbook(path)
            ws = wb["Calificaciones"]
            for r in (2, 3):
                ws.cell(r, 5).number_format = "0%"
            wb.save(path)
            wb.close()

            estudiantes, _ = parse_notas_archivo(path)
            es = estudiantes[0]
            self.assertEqual(
                {k: v.value for k, v in es.calificaciones_modulos.items()},
                {"modulo_1": 4.5},
            )
            self.assertEqual(es.nota_final.value, 4.0)

    def test_selecciona_hoja_con_notas_aunque_los_titulos_sean_atipicos(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "CMI999.xlsx"
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "Listado"
            ws1.append(["OrgDefinedId", "Username", "Last Name", "First Name"])
            ws1.append(["9000000001", "otro", "Apellido", "Nombre"])

            ws2 = wb.create_sheet("CURSO REALIZADO")
            ws2.append([
                "OrgDefinedId", "Username", "Last Name", "First Name",
                "Puntaje A", "Puntaje B", "Puntaje C",
            ])
            ws2.append(["1000000001", "user1", "Apellido", "Nombre", 5.0, 4.5, 4.8])
            wb.save(path)
            wb.close()

            from parsing.notas_parser import _select_notas_sheet

            wb2 = openpyxl.load_workbook(path)
            self.assertEqual(_select_notas_sheet(wb2), "CURSO REALIZADO")
            wb2.close()

            estudiantes, _ = parse_notas_archivo(path)
            self.assertEqual(len(estudiantes), 1)
            self.assertEqual(
                {k: v.value for k, v in estudiantes[0].calificaciones_modulos.items()},
                {"modulo_1": 5.0, "modulo_2": 4.5},
            )
            self.assertEqual(estudiantes[0].nota_final.value, 4.8)


if __name__ == "__main__":
    unittest.main()
