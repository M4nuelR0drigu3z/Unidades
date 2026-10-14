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
        self.assertEqual(
            respuesta["resumen"],
            {"detenido": 2, "revisar": 0, "trafico_lento": 1, "reten": 0, "ruta": 1},
        )
        self.assertEqual(respuesta["unidades"][0], {
            "unidad": "1939", "estatus": "DETENIDO", "tiempo_minutos": 1435,
            "coordenadas": {"latitud": 19.631638, "longitud": -99.17491},
            "ubicacion": "Avenida Independencia",
        })
        self.assertEqual(respuesta["unidades"][2]["tiempo_minutos"], 12)
        self.assertIsNone(respuesta["unidades"][3]["tiempo_minutos"])
        self.assertIsNone(respuesta["unidades"][3]["coordenadas"])

    def test_busca_reporte_por_perfil_o_nombre(self):
        config = {"reportes": [{"nombre": "Reporte EC-05", "perfil_filtros": "EC-05"}]}
        for centro in ["EC-05", "ec-05", "Reporte EC-05"]:
            self.assertIsNotNone(api.buscar_reporte(config, centro), centro)
        self.assertIsNone(api.buscar_reporte(config, "EC-99"))

    def test_endpoint_requiere_api_key(self):
        self.assertEqual(self.cliente.get("/telemetria/consultarEstatus").status_code, 401)
        respuesta = self.cliente.get(
            "/telemetria/consultarEstatus", headers={"X-API-Key": "otra"}
        )
        self.assertEqual(respuesta.status_code, 401)

    def test_endpoint_sin_clave_configurada_no_responde_datos(self):
        with mock.patch.dict(os.environ, {"TELEMETRIA_API_KEY": ""}):
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

    def test_centro_desconocido_regresa_404(self):
        with mock.patch.object(
            api.envio, "cargar_configuracion",
            return_value={"reportes": [{"nombre": "Reporte EC-05", "perfil_filtros": "EC-05"}]},
        ):
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
