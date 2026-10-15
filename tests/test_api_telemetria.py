from datetime import datetime
import os
import unittest
from unittest import mock

from fastapi.testclient import TestClient
import pytz

import api_telemetria as api

AHORA = pytz.timezone("America/Mexico_City").localize(datetime(2026, 10, 5, 11, 30, 7))
FILAS = [
    {"Unidad": "1939", "Estatus": "DETENIDO CONFIRMADO", "Minutos Detenido": 1434.6,
     "Latitud": 19.631638, "Longitud": "-99.17491", "Ubicación": "Avenida Independencia"},
    {"Unidad": "2241", "Estatus": "DETENIDO SAMSARA", "Minutos Detenido": 22,
     "Latitud": 20.618138, "Longitud": -100.402748, "Ubicación": "Querétaro"},
    {"Unidad": "2002", "Estatus": "TRAFICO LENTO", "Minutos Trafico": 12,
     "Latitud": 19.3, "Longitud": -90.7, "Ubicación": "Champotón"},
    {"Unidad": "1944", "Estatus": "RUTA", "Latitud": "", "Longitud": "", "Ubicación": "Arco Norte"},
]


class ApiTelemetriaTests(unittest.TestCase):
    def setUp(self):
        api._cache.clear()
        api._fallos.clear()
        api._bloqueados.clear()
        self.env = mock.patch.dict(os.environ, {"TELEMETRIA_API_KEY": "clave-prueba"})
        self.env.start()
        self.cliente = TestClient(api.app)

    def tearDown(self):
        self.env.stop()

    def test_respuesta_con_formato_simple(self):
        respuesta = api.construir_respuesta_estatus("EC-05", FILAS, AHORA)
        self.assertEqual(respuesta["fecha"], "2026-10-05")
        self.assertEqual(respuesta["hora"], "11:30:07")
        self.assertEqual(respuesta["total_unidades"], 4)
        self.assertEqual(respuesta["resumen"], {"detenido": 2, "ruta": 2})
        self.assertEqual(respuesta["unidades"][0], {
            "unidad": "1939", "estatus": "DETENIDO", "tiempo_minutos": 1435,
            "coordenadas": {"latitud": 19.631638, "longitud": -99.17491},
            "ubicacion": "Avenida Independencia", "fecha_gps": None,
        })
        # TRAFICO LENTO avanza: el cliente lo recibe como RUTA, sin tiempo.
        self.assertEqual(respuesta["unidades"][2]["estatus"], "RUTA")
        self.assertIsNone(respuesta["unidades"][2]["tiempo_minutos"])
        self.assertIsNone(respuesta["unidades"][3]["tiempo_minutos"])
        self.assertIsNone(respuesta["unidades"][3]["coordenadas"])

    def test_cliente_solo_recibe_detenido_o_ruta(self):
        casos = {
            "DETENIDO CONFIRMADO": "DETENIDO", "DETENIDO SAMSARA": "DETENIDO",
            "DETENIDO LOGITRACK": "DETENIDO", "REVISAR": "DETENIDO", "RETEN": "DETENIDO",
            "TRAFICO LENTO": "RUTA", "RUTA": "RUTA",
        }
        for interno, esperado in casos.items():
            self.assertEqual(api.estatus_api(interno), esperado, interno)
        revisar = api.unidad_a_json({"Unidad": "9", "Estatus": "REVISAR", "Minutos Detenido": 3})
        self.assertEqual((revisar["estatus"], revisar["tiempo_minutos"]), ("DETENIDO", 3))

    def test_fecha_gps_es_la_de_la_lectura_de_las_coordenadas(self):
        samsara = AHORA.replace(minute=28, second=41)
        logitrack = AHORA.replace(minute=29, second=5)
        normal = api.unidad_a_json({
            "Unidad": "1", "Estatus": "RUTA", "Estatus Samsara Actual": "RUTA",
            "Fecha GPS": samsara, "Fecha Logitrack": logitrack,
        })
        self.assertEqual(normal["fecha_gps"], "2026-10-05 11:28:41")
        # Rescatada: Samsara sin GPS vigente, coordenadas y fecha de Logitrack.
        rescatada = api.unidad_a_json({
            "Unidad": "2", "Estatus": "DETENIDO LOGITRACK", "Estatus Samsara Actual": "GPS VIEJO",
            "Fecha GPS": AHORA.replace(hour=6), "Fecha Logitrack": logitrack,
        })
        self.assertEqual(rescatada["fecha_gps"], "2026-10-05 11:29:05")

    def test_refresco_alineado_a_marcas_del_reloj(self):
        # 2026-10-05 12:00:00 en Mexico (UTC-6) = 18:00:00 UTC.
        mediodia = pytz.utc.localize(datetime(2026, 10, 5, 18, 0, 0)).timestamp()
        self.assertEqual(api.segundos_para_siguiente_marca(180, mediodia - 100), 100)
        self.assertEqual(api.segundos_para_siguiente_marca(180, mediodia + 1), 179)
        self.assertEqual(api.segundos_para_siguiente_marca(180, mediodia + 1800 - 30), 30)

    def test_refresco_calcula_todos_los_centros_en_paralelo(self):
        calculados = []

        def calcular(centro):
            calculados.append(centro)
            return api.construir_respuesta_estatus(centro, FILAS, AHORA)

        with mock.patch.object(api, "calcular_estatus", side_effect=calcular), \
                mock.patch.object(api, "centros_samsara", return_value=["EC-04", "EC-05", "EC-10"]):
            api.refrescar_centros(["TODOS"], 4)
        self.assertEqual(sorted(calculados), ["EC-04", "EC-05", "EC-10"])
        self.assertEqual(sorted(api._cache), ["ec-04", "ec-05", "ec-10"])

    UNIDADES = [
        {"unidad": "1939", "division": "EC-05", "flota": {"id": 40, "nombre": "MeLi Dedicado SD05"},
         "placas": "58BA9Y", "activa": True},
        {"unidad": "1786", "division": "EC-10", "flota": {"id": 70, "nombre": "Avl EC-10"},
         "placas": "70AT5K", "activa": True},
        {"unidad": "1500", "division": "AF", "flota": None, "placas": "", "activa": False},
    ]

    def _consultar_unidades(self, **params):
        with mock.patch.object(api.catalogo_unidades, "unidades_con_cache", return_value=self.UNIDADES):
            return self.cliente.get(
                "/telemetria/consultarUnidades", params=params, headers={"X-API-Key": "clave-prueba"}
            )

    def test_catalogo_de_unidades_sin_bajas_por_defecto(self):
        respuesta = self._consultar_unidades()
        self.assertEqual(respuesta.status_code, 200)
        self.assertEqual([u["unidad"] for u in respuesta.json()["unidades"]], ["1939", "1786"])
        con_bajas = self._consultar_unidades(incluir_bajas="true").json()
        self.assertEqual(con_bajas["total_unidades"], 3)

    def test_catalogo_de_unidades_filtra_por_division_unidad_y_flota(self):
        self.assertEqual(self._consultar_unidades(division="ec-10").json()["unidades"][0]["unidad"], "1786")
        self.assertEqual(self._consultar_unidades(unidad="1939, 1786").json()["total_unidades"], 2)
        self.assertEqual(self._consultar_unidades(flota=40).json()["unidades"][0]["placas"], "58BA9Y")

    def test_catalogo_de_unidades_requiere_clave_y_oculta_errores_de_sql(self):
        self.assertEqual(self.cliente.get("/telemetria/consultarUnidades").status_code, 401)
        with mock.patch.object(
            api.catalogo_unidades, "unidades_con_cache", side_effect=RuntimeError("Login failed for user x")
        ):
            respuesta = self.cliente.get(
                "/telemetria/consultarUnidades", headers={"X-API-Key": "clave-prueba"}
            )
        self.assertEqual(respuesta.status_code, 502)
        self.assertNotIn("Login", respuesta.text)

    def test_busca_reporte_por_perfil_o_nombre(self):
        config = {"reportes": [{"nombre": "Reporte EC-05", "perfil_filtros": "EC-05",
                                "telemetria_logitrack": {"activo": True}}]}
        for centro in ["EC-05", "ec-05", "Reporte EC-05"]:
            self.assertIsNotNone(api.buscar_reporte(config, centro), centro)
        self.assertIsNone(api.buscar_reporte(config, "EC-99"))

    def test_reportes_solo_excel_no_se_usan_en_la_api(self):
        # Caso real: "Reporte Sayer" y "Reporte EC-02" comparten perfil EC-02 y no cruzan Logitrack.
        config = {"reportes": [
            {"nombre": "Reporte Sayer", "perfil_filtros": "EC-02"},
            {"nombre": "Reporte EC-02", "perfil_filtros": "EC-02"},
        ]}
        self.assertIsNone(api.buscar_reporte(config, "EC-02"))
        with mock.patch.object(api, "centros_samsara", return_value=["EC-02", "EC-10"]):
            reporte, nombre = api.resolver_reporte(config, "ec-02")
        self.assertEqual(nombre, "EC-02")
        self.assertEqual(reporte["etiquetas"], ["EC-02"])
        self.assertEqual(reporte["tipo_filtro_etiqueta"], "parentTagIds")
        self.assertEqual(reporte["filtros"]["distancia_cercania_geocerca_metros"], 400)
        self.assertEqual(reporte["filtros"]["etiquetas_geocercas_excluidas"], ["EC-02"])
        self.assertNotIn("telemetria_logitrack", reporte)

    def test_centro_que_no_es_ec_no_consulta_samsara(self):
        with mock.patch.object(api, "centros_samsara") as samsara:
            reporte, _ = api.resolver_reporte({"reportes": []}, "Sayer Full")
        self.assertIsNone(reporte)
        samsara.assert_not_called()

    def _consultar_centros(self, texto):
        def calcular(centro):
            return api.construir_respuesta_estatus(centro, FILAS, AHORA)

        with mock.patch.object(api, "calcular_estatus", side_effect=calcular), \
                mock.patch.object(api, "centros_samsara", return_value=["EC-04", "EC-05", "EC-10"]):
            return self.cliente.get(
                "/telemetria/consultarEstatus", params={"centro_operativo": texto},
                headers={"X-API-Key": "clave-prueba"},
            )

    def test_varios_centros_regresan_lista(self):
        respuesta = self._consultar_centros("EC-05, ec-10,EC-05")
        self.assertEqual(respuesta.status_code, 200)
        self.assertEqual([x["centro_operativo"] for x in respuesta.json()], ["EC-05", "EC-10"])

    def test_todos_regresa_todos_los_centros_de_samsara(self):
        respuesta = self._consultar_centros("todos")
        self.assertEqual([x["centro_operativo"] for x in respuesta.json()], ["EC-04", "EC-05", "EC-10"])

    def test_un_centro_sigue_regresando_objeto(self):
        self.assertEqual(self._consultar_centros("EC-10").json()["centro_operativo"], "EC-10")

    def test_centro_inexistente_en_lista_regresa_404(self):
        respuesta = self._consultar_centros("EC-05,EC-99")
        self.assertEqual(respuesta.status_code, 404)
        self.assertIn("EC-99", respuesta.json()["detail"])

    def test_limite_de_centros_por_consulta(self):
        respuesta = self._consultar_centros(",".join(f"EC-{n:02d}" for n in range(11)))
        self.assertEqual(respuesta.status_code, 422)

    def test_endpoint_requiere_api_key(self):
        self.assertEqual(self.cliente.get("/telemetria/consultarEstatus").status_code, 401)
        respuesta = self.cliente.get(
            "/telemetria/consultarEstatus", headers={"X-API-Key": "otra"}
        )
        self.assertEqual(respuesta.status_code, 401)

    def test_endpoint_sin_clave_configurada_no_responde_datos(self):
        with mock.patch.dict(os.environ, {"TELEMETRIA_API_KEY": "", "TELEMETRIA_API_KEYS": ""}):
            respuesta = self.cliente.get(
                "/telemetria/consultarEstatus", headers={"X-API-Key": ""}
            )
        self.assertEqual(respuesta.status_code, 503)

    def test_endpoint_responde_y_usa_cache(self):
        llamadas = []

        def calcular(centro):
            llamadas.append(centro)
            return api.construir_respuesta_estatus("EC-05", FILAS, AHORA)

        with mock.patch.object(api, "calcular_estatus", side_effect=calcular):
            for ruta in ["/telemetria/consultarEstatus", "/telemetria/consultarEstatus/"]:
                respuesta = self.cliente.get(
                    ruta, params={"centro_operativo": "EC-05"},
                    headers={"X-API-Key": "clave-prueba"},
                )
                self.assertEqual(respuesta.status_code, 200)
                self.assertEqual(respuesta.json()["resumen"]["detenido"], 2)
        self.assertEqual(llamadas, ["EC-05"])

    def test_refresco_en_segundo_plano_deja_cache_listo(self):
        detener = api.threading.Event()
        llamadas = []

        def calcular(centro):
            llamadas.append(centro)
            detener.set()  # una sola vuelta
            return api.construir_respuesta_estatus("EC-05", FILAS, AHORA)

        with mock.patch.object(api, "calcular_estatus", side_effect=calcular):
            api.refrescar_periodicamente(["EC-05"], 60, detener)
            respuesta = self.cliente.get(
                "/telemetria/consultarEstatus", headers={"X-API-Key": "clave-prueba"}
            )
        self.assertEqual(llamadas, ["EC-05"])  # la peticion uso el cache del refresco
        self.assertEqual(respuesta.json()["total_unidades"], 4)
        self.assertIn("ec-05", self.cliente.get("/salud").json()["segundos_desde_ultimo_calculo"])

    def test_refresco_fallido_no_detiene_el_ciclo(self):
        detener = api.threading.Event()
        intentos = []

        def calcular(centro):
            intentos.append(centro)
            if len(intentos) == 2:
                detener.set()
            raise RuntimeError("Samsara caido")

        with mock.patch.object(api, "calcular_estatus", side_effect=calcular), \
                mock.patch.object(detener, "wait", return_value=False):
            api.refrescar_periodicamente(["EC-05"], 60, detener)
        self.assertEqual(len(intentos), 2)

    def test_cache_vencido_recalcula_en_la_peticion(self):
        api._cache["ec-05"] = (api.time.monotonic() - 10_000, {"viejo": True})
        with mock.patch.object(
            api, "calcular_estatus",
            return_value=api.construir_respuesta_estatus("EC-05", FILAS, AHORA),
        ):
            respuesta = self.cliente.get(
                "/telemetria/consultarEstatus", headers={"X-API-Key": "clave-prueba"}
            )
        self.assertNotIn("viejo", respuesta.json())

    def _consultar(self, clave):
        with mock.patch.object(
            api, "calcular_estatus",
            return_value=api.construir_respuesta_estatus("EC-05", FILAS, AHORA),
        ):
            return self.cliente.get("/telemetria/consultarEstatus", headers={"X-API-Key": clave})

    def test_una_clave_por_consumidor_y_registro_de_quien_consulta(self):
        with mock.patch.dict(os.environ, {
            "TELEMETRIA_API_KEY": "", "TELEMETRIA_API_KEYS": "erp:clave-erp, portal:clave-portal",
        }):
            self.assertEqual(api.claves_configuradas(), {"erp": "clave-erp", "portal": "clave-portal"})
            with self.assertLogs(level="INFO") as registro:
                self.assertEqual(self._consultar("clave-portal").status_code, 200)
            self.assertTrue(any("por portal desde" in linea for linea in registro.output))
            self.assertEqual(self._consultar("clave-prueba").status_code, 401)

    def test_bloquea_ip_tras_intentos_fallidos(self):
        with mock.patch.dict(os.environ, {"TELEMETRIA_MAX_INTENTOS_FALLIDOS": "3"}):
            codigos = [self._consultar("mala").status_code for _ in range(3)]
            self.assertEqual(codigos, [401, 401, 401])
            bloqueada = self._consultar("clave-prueba")  # ni la clave correcta pasa
            self.assertEqual(bloqueada.status_code, 429)
            self.assertIn("Retry-After", bloqueada.headers)
            api._bloqueados.clear()
            self.assertEqual(self._consultar("clave-prueba").status_code, 200)

    def test_encabezados_de_seguridad_y_docs_apagados(self):
        respuesta = self._consultar("clave-prueba")
        self.assertEqual(respuesta.headers["Cache-Control"], "no-store")
        self.assertEqual(respuesta.headers["X-Content-Type-Options"], "nosniff")
        self.assertEqual(self.cliente.get("/docs").status_code, 404)
        self.assertEqual(self.cliente.get("/openapi.json").status_code, 404)

    def test_centro_operativo_demasiado_largo_se_rechaza(self):
        respuesta = self.cliente.get(
            "/telemetria/consultarEstatus", params={"centro_operativo": "x" * 200},
            headers={"X-API-Key": "clave-prueba"},
        )
        self.assertEqual(respuesta.status_code, 422)

    def test_centro_desconocido_regresa_404(self):
        with mock.patch.object(
            api.envio, "cargar_configuracion",
            return_value={"reportes": [{"nombre": "Reporte EC-05", "perfil_filtros": "EC-05"}]},
        ), mock.patch.object(api, "centros_samsara", return_value=["EC-05", "EC-10"]):
            respuesta = self.cliente.get(
                "/telemetria/consultarEstatus", params={"centro_operativo": "EC-99"},
                headers={"X-API-Key": "clave-prueba"},
            )
        self.assertEqual(respuesta.status_code, 404)

    def test_falla_de_samsara_regresa_502_sin_detalle_interno(self):
        with mock.patch.object(api, "calcular_estatus", side_effect=RuntimeError("token=secreto")):
            respuesta = self.cliente.get(
                "/telemetria/consultarEstatus", headers={"X-API-Key": "clave-prueba"}
            )
        self.assertEqual(respuesta.status_code, 502)
        self.assertNotIn("secreto", respuesta.text)


if __name__ == "__main__":
    unittest.main()
