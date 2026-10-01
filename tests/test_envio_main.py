from datetime import datetime
from io import BytesIO
import os
from pathlib import Path
import unittest
from unittest import mock

from openpyxl import load_workbook
import pytz

import EnvioMain as envio


TZ_MX = pytz.timezone("America/Mexico_City")
AHORA = TZ_MX.localize(datetime(2026, 9, 7, 14, 0))


def unidad_samsara(nombre="2234", mph=0, minuto=0, lat=25.70, lon=-100.34):
    return {
        "Unidad": nombre,
        "Estatus": "RUTA",
        "Fecha GPS": TZ_MX.localize(datetime(2026, 9, 7, 14, minuto)) if minuto is not None else None,
        "GpsActual": {"latitude": lat, "longitude": lon, "speedMilesPerHour": mph},
    }


def lectura_logitrack(nombre="2234", kmh=0, hora="2026-09-07 13:58:00", lat=25.70, lon=-100.34):
    return {"unit_name": nombre, "datetime": hora, "lat": lat, "lon": lon, "speed": kmh, "engine_ign": 1}


def cruzar(samsara, logitrack, settings=None):
    return envio.enriquecer_con_logitrack([samsara], logitrack, AHORA, settings or {})[0]


class EnvioMainTests(unittest.TestCase):
    def test_normaliza_sufijo_logitrack_tdr(self):
        self.assertEqual(envio.normalizar_nombre_unidad("2234-TDR"), "2234")
        self.assertEqual(envio.normalizar_nombre_unidad("002197-TDR"), "2197")

    def test_normaliza_cualquier_combinacion_de_tdr(self):
        # Formatos reales vistos en units/last_status y variantes posibles.
        for nombre in ["2234 - TDR", "2234 -TDR", "2234tdr", "2234 TDR", "TDR-2234",
                       "tdr 02234", "2234_tdr", " 2234 ", "2234-TDR-"]:
            self.assertEqual(envio.normalizar_nombre_unidad(nombre), "2234", nombre)

    def test_normaliza_conserva_nombres_no_numericos_y_vacios(self):
        self.assertEqual(envio.normalizar_nombre_unidad("DC01"), "dc01")
        self.assertEqual(envio.normalizar_nombre_unidad("TDR"), "tdr")
        self.assertEqual(envio.normalizar_nombre_unidad(""), "")
        self.assertEqual(envio.normalizar_nombre_unidad(None), "")

    def test_vincula_nombre_logitrack_con_espacios_y_guion(self):
        fila = cruzar(unidad_samsara("2234"), [lectura_logitrack("2234 - TDR")])
        self.assertEqual(fila["Logitrack Encontrado"], "SI")

    def test_duplicados_logitrack_conservan_lectura_mas_reciente(self):
        fila = cruzar(unidad_samsara("1866"), [
            lectura_logitrack("1866", kmh=0, hora="2025-04-04 13:14:48", lat=19.0, lon=-99.0),
            lectura_logitrack("1866", kmh=0, hora="2026-09-07 13:58:00"),
        ])
        self.assertEqual(fila["Logitrack Fresco"], "SI")
        self.assertEqual(fila["Doble Comprobación Actual"], "SI")

    def test_ambas_detenidas_cerca_es_doble_comprobacion(self):
        fila = cruzar(unidad_samsara(mph=0), [lectura_logitrack(kmh=0, lat=25.7001, lon=-100.3401)])
        self.assertEqual(fila["Estatus"], "DETENIDO")
        self.assertEqual(fila["Doble Comprobación Actual"], "SI")
        self.assertEqual(fila["Ubicación Coincide"], "SI")
        self.assertEqual(fila["Fuente Confirmación"], "SAMSARA + LOGITRACK")

    def test_deriva_gps_de_2_kmh_cuenta_como_detenida(self):
        # Caso real 2001/2195: Samsara marca ~1.3 mph estando detenida.
        fila = cruzar(unidad_samsara(mph=1.3), [lectura_logitrack(kmh=0)])
        self.assertEqual(fila["Doble Comprobación Actual"], "SI")

    def test_detenidas_a_mas_de_150_m_no_son_doble_comprobacion(self):
        # 333 m en 2 min es posible: la unidad se movio entre lecturas.
        fila = cruzar(unidad_samsara(mph=0), [lectura_logitrack(kmh=0, lat=25.703)])
        self.assertEqual(fila["Ubicación Coincide"], "NO")
        self.assertEqual(fila["Doble Comprobación Actual"], "NO")
        self.assertEqual(fila["Estatus"], "DETENIDO")
        self.assertIn("cambió de posición", fila["Motivo Decisión"])

    def test_deriva_de_velocidad_no_amplia_tolerancia_de_detenidas(self):
        # Caso real 2193: 1.2 mph de deriva y 10 min de desfase a 494 m.
        fila = cruzar(
            unidad_samsara(mph=1.2, minuto=0),
            [lectura_logitrack(kmh=0, hora="2026-09-07 13:50:00", lat=25.7044)],
        )
        self.assertEqual(fila["Distancia Esperada Metros"], 150)
        self.assertEqual(fila["Doble Comprobación Actual"], "NO")

    def test_detenidas_a_distancia_imposible_se_revisan(self):
        fila = cruzar(unidad_samsara(mph=0), [lectura_logitrack(kmh=0, lat=25.80)])
        self.assertEqual(fila["Estatus"], "REVISAR")
        self.assertIn("Ubicación no coincide", fila["Motivo Decisión"])

    def test_en_ruta_la_tolerancia_crece_con_desfase_y_velocidad(self):
        # 100 km/h y 2 min de desfase: ~3.3 km recorridos entre lecturas.
        fila = cruzar(
            unidad_samsara(mph=62, minuto=0),
            [lectura_logitrack(kmh=100, hora="2026-09-07 13:58:00", lat=25.73)],
        )
        self.assertEqual(fila["Ubicación Coincide"], "SI")
        self.assertEqual(fila["Estatus"], "RUTA")
        self.assertEqual(fila["Fuente Confirmación"], "SAMSARA + LOGITRACK")
        self.assertEqual(fila["Motivo Decisión"], "Ambas telemetrías reportan movimiento")

    def test_en_ruta_demasiado_lejos_para_el_desfase_se_revisa(self):
        fila = cruzar(
            unidad_samsara(mph=62, minuto=0),
            [lectura_logitrack(kmh=100, hora="2026-09-07 13:58:00", lat=25.90)],
        )
        self.assertEqual(fila["Ubicación Coincide"], "NO")
        self.assertEqual(fila["Estatus"], "REVISAR")

    def test_samsara_mas_reciente_detenida_va_a_historial(self):
        # Caso real 2146: Samsara 0 km/h, Logitrack 104 km/h 2.5 min antes.
        fila = cruzar(
            unidad_samsara(mph=0, minuto=0),
            [lectura_logitrack(kmh=104, hora="2026-09-07 13:57:30", lat=25.72)],
        )
        self.assertEqual(fila["Estatus"], "DETENIDO")
        self.assertEqual(fila["Doble Comprobación Actual"], "NO")
        self.assertIn("Samsara (lectura más reciente)", fila["Motivo Decisión"])

    def test_samsara_mas_reciente_en_movimiento_gana(self):
        fila = cruzar(unidad_samsara(mph=40, minuto=0), [lectura_logitrack(kmh=0, hora="2026-09-07 13:58:00")])
        self.assertEqual(fila["Estatus"], "RUTA")
        self.assertIn("Logitrack aún reporta detención", fila["Motivo Decisión"])

    def test_logitrack_mas_reciente_detenida_gana(self):
        fila = cruzar(
            unidad_samsara(mph=40, minuto=0),
            [lectura_logitrack(kmh=0, hora="2026-09-07 14:03:00")],
        )
        self.assertEqual(fila["Estatus"], "DETENIDO LOGITRACK")
        self.assertEqual(fila["Fuente Confirmación"], "LOGITRACK")

    def test_logitrack_mas_reciente_en_movimiento_gana(self):
        fila = cruzar(
            unidad_samsara(mph=0, minuto=0),
            [lectura_logitrack(kmh=60, hora="2026-09-07 14:03:00")],
        )
        self.assertEqual(fila["Estatus"], "RUTA")
        self.assertEqual(fila["Fuente Confirmación"], "LOGITRACK")

    def test_logitrack_sin_velocidad_no_cuenta_como_detenido(self):
        lectura = lectura_logitrack()
        del lectura["speed"]
        fila = cruzar(unidad_samsara(mph=0), [lectura])
        self.assertEqual(fila["Estatus Logitrack"], "SIN VELOCIDAD")
        self.assertEqual(fila["Doble Comprobación Actual"], "NO")
        self.assertEqual(fila["Estatus"], "DETENIDO")

    def test_sin_coincidencia_conserva_samsara(self):
        fila = cruzar(unidad_samsara(mph=40), [lectura_logitrack("9999")])
        self.assertEqual(fila["Logitrack Encontrado"], "NO")
        self.assertEqual(fila["Estatus"], "RUTA")
        self.assertEqual(fila["Fuente Confirmación"], "SAMSARA ACTUAL")

    def test_logitrack_caido_mantiene_reporte_con_samsara(self):
        fila = cruzar(unidad_samsara(mph=0), [])
        self.assertEqual(fila["Estatus"], "DETENIDO")
        self.assertEqual(fila["Logitrack Encontrado"], "NO")

    def test_rescata_gps_samsara_viejo_con_logitrack_vigente(self):
        omitidas = [
            {"Unidad": "2234", "Motivo": "GPS VIEJO", "Detalle": "Antiguedad: 300 minutos"},
            {"Unidad": "2197", "Motivo": "GPS VIEJO", "Detalle": "Antiguedad: 300 minutos"},
            {"Unidad": "2257", "Motivo": "PATIO/GEOCERCA EXCLUIDA"},
        ]
        logitrack = [
            lectura_logitrack("2234 - TDR", kmh=0),
            lectura_logitrack("2197", kmh=0, hora="2026-09-07 10:00:00"),
            lectura_logitrack("2257", kmh=0),
        ]
        rescatadas, restantes = envio.rescatar_con_logitrack(omitidas, logitrack, AHORA, {}, {})
        self.assertEqual([x["Unidad"] for x in rescatadas], ["2234"])
        self.assertEqual(rescatadas[0]["Estatus"], "DETENIDO LOGITRACK")
        self.assertEqual(rescatadas[0]["Fuente Confirmación"], "LOGITRACK")
        self.assertEqual([x["Unidad"] for x in restantes], ["2197", "2257"])

    def test_no_rescata_unidad_que_logitrack_ubica_en_patio(self):
        patio = {"1": {"nombre": "Patio Tultitlan", "geofence": {"polygon": {"vertices": [
            {"latitude": 25.69, "longitude": -100.35}, {"latitude": 25.71, "longitude": -100.35},
            {"latitude": 25.71, "longitude": -100.33}, {"latitude": 25.69, "longitude": -100.33},
        ]}}}}
        omitidas = [{"Unidad": "2234", "Motivo": "GPS VIEJO", "Detalle": "Antiguedad: 300 minutos"}]
        rescatadas, restantes = envio.rescatar_con_logitrack(
            omitidas, [lectura_logitrack("2234", kmh=0)], AHORA, {}, patio
        )
        self.assertEqual(rescatadas, [])
        self.assertIn("Patio Tultitlan", restantes[0]["Detalle"])

    def test_distancia_a_geocerca_poligono_y_circulo(self):
        cuadro = {"polygon": {"vertices": [
            {"latitude": 19.60, "longitude": -99.20}, {"latitude": 19.61, "longitude": -99.20},
            {"latitude": 19.61, "longitude": -99.19}, {"latitude": 19.60, "longitude": -99.19},
        ]}}
        self.assertEqual(envio.distancia_a_geocerca(19.605, -99.195, cuadro), 0.0)
        # 0.002 grados al norte del borde superior: ~221 m.
        self.assertAlmostEqual(envio.distancia_a_geocerca(19.612, -99.195, cuadro), 221, delta=3)
        circulo = {"circle": {"latitude": 19.60, "longitude": -99.20, "radiusMeters": 100}}
        self.assertAlmostEqual(envio.distancia_a_geocerca(19.603, -99.20, circulo), 233, delta=3)

    def test_omite_detenidas_cerca_de_patio_y_conserva_en_ruta(self):
        patio = {"1": {"nombre": "MXXEM2", "geofence": {"circle": {
            "latitude": 19.60, "longitude": -99.20, "radiusMeters": 100}}}}
        datos = [
            {"Unidad": "2258", "Estatus": "DETENIDO CONFIRMADO", "Latitud": 19.6018, "Longitud": -99.20},
            {"Unidad": "2007", "Estatus": "RUTA", "Latitud": 19.6018, "Longitud": -99.20},
            {"Unidad": "2243", "Estatus": "DETENIDO CONFIRMADO", "Latitud": 19.609, "Longitud": -99.20},
            {"Unidad": "2014", "Estatus": "DETENIDO LOGITRACK", "Latitud": "19.6015", "Longitud": "-99.20"},
        ]
        conservadas, omitidas = envio.omitir_detenidas_cerca_de_geocerca(datos, patio, 400)
        self.assertEqual([x["Unidad"] for x in conservadas], ["2007", "2243"])
        self.assertEqual([x["Unidad"] for x in omitidas], ["2258", "2014"])
        self.assertEqual(omitidas[0]["Motivo"], "CERCA DE GEOCERCA")
        self.assertIn("MXXEM2", omitidas[0]["Detalle"])

    def test_geometrias_incluyen_excluidas_por_etiqueta_y_especiales_por_id(self):
        direcciones = [
            {"id": "1", "name": "MXXEM2", "tags": [{"id": "4363967"}], "geofence": {"circle": {}}},
            {"id": "257477773", "name": "TDR MEX1", "tags": [{"id": "4357031"}], "geofence": {"polygon": {}}},
            {"id": "3", "name": "Arco Norte", "tags": [{"id": "999"}], "geofence": {}},
        ]
        geometrias = envio.obtener_geometrias_geocercas(None, ["4363967"], ["257477773"], direcciones)
        self.assertEqual(sorted(g["nombre"] for g in geometrias.values()), ["MXXEM2", "TDR MEX1"])
        self.assertEqual(geometrias["257477773"]["motivo"], "GEOCERCA ESPECIAL")
        self.assertEqual(geometrias["1"]["motivo"], "PATIO/GEOCERCA EXCLUIDA")

    def test_excluye_por_coordenadas_aunque_samsara_no_reporte_geocerca(self):
        # Caso real 2148: dentro de TDR MEX1 con gps.address vacio.
        mex1 = {"257477773": {"nombre": "TDR MEX1", "motivo": "GEOCERCA ESPECIAL", "geofence": {
            "circle": {"latitude": 19.6256, "longitude": -99.1648, "radiusMeters": 80}}}}
        datos = [
            {"Unidad": "2148", "Estatus": "DETENIDO CONFIRMADO", "Latitud": 19.625611, "Longitud": -99.164793},
            {"Unidad": "2009", "Estatus": "DETENIDO CONFIRMADO", "Latitud": 20.0, "Longitud": -98.9},
            {"Unidad": "9", "Estatus": "RUTA", "Latitud": "", "Longitud": ""},
        ]
        conservadas, omitidas = envio.excluir_por_coordenadas_en_geocerca(datos, mex1)
        self.assertEqual([x["Unidad"] for x in conservadas], ["2009", "9"])
        self.assertEqual(omitidas[0]["Motivo"], "GEOCERCA ESPECIAL")
        self.assertIn("TDR MEX1", omitidas[0]["Detalle"])

    def test_sin_geometrias_validadas_no_rescata(self):
        omitidas = [{"Unidad": "2234", "Motivo": "GPS VIEJO"}]
        rescatadas, restantes = envio.rescatar_con_logitrack(
            omitidas, [lectura_logitrack("2234", kmh=0)], AHORA, {}, None
        )
        self.assertEqual(rescatadas, [])
        self.assertEqual(len(restantes), 1)

    def test_punto_en_geocerca_circulo_y_poligono(self):
        circulo = {"circle": {"latitude": 19.62, "longitude": -99.16, "radiusMeters": 200}}
        self.assertTrue(envio.punto_en_geocerca(19.6205, -99.1605, circulo))
        self.assertFalse(envio.punto_en_geocerca(19.63, -99.16, circulo))
        poligono = {"polygon": {"vertices": [
            {"latitude": 0, "longitude": 0}, {"latitude": 0, "longitude": 1},
            {"latitude": 1, "longitude": 1}, {"latitude": 1, "longitude": 0},
        ]}}
        self.assertTrue(envio.punto_en_geocerca(0.5, 0.5, poligono))
        self.assertFalse(envio.punto_en_geocerca(1.5, 0.5, poligono))
        self.assertFalse(envio.punto_en_geocerca(None, 0.5, poligono))

    def test_historial_sin_doble_lectura_queda_detenido_samsara(self):
        datos = [{"Estatus": "DETENIDO", "Minutos Detenido": 30,
                  "Doble Comprobación Actual": "NO", "Logitrack Fresco": "SI"}]
        resultado = envio.finalizar_comprobacion_telemetria(datos, {})
        self.assertEqual(resultado[0]["Estatus"], "DETENIDO SAMSARA")

    def test_historial_corto_queda_revisar_con_minimo_configurado(self):
        datos = [{"Estatus": "DETENIDO", "Minutos Detenido": 7, "Doble Comprobación Actual": "SI"}]
        resultado = envio.finalizar_comprobacion_telemetria(datos, {"detencion_minima_minutos": 10})
        self.assertEqual(resultado[0]["Estatus"], "REVISAR")
        self.assertIn("10 minutos", resultado[0]["Motivo Decisión"])

    def test_finalizar_no_toca_estatus_decididos_por_cruce(self):
        datos = [{"Estatus": "DETENIDO LOGITRACK"}, {"Estatus": "REVISAR", "Motivo Decisión": "x"}]
        resultado = envio.finalizar_comprobacion_telemetria(datos, {})
        self.assertEqual([x["Estatus"] for x in resultado], ["DETENIDO LOGITRACK", "REVISAR"])
        self.assertEqual(resultado[1]["Motivo Decisión"], "x")

    def test_resumen_telemetria_para_consola(self):
        filas = [
            cruzar(unidad_samsara("2234", mph=0), [lectura_logitrack("2234 - TDR", kmh=0, lat=25.7002)]),
            cruzar(unidad_samsara("2240", mph=0), []),
        ]
        texto = "\n".join(envio.resumen_telemetria(filas))
        self.assertIn("Samsara vs Logitrack", texto)
        self.assertIn("Vinculadas: 1/2", texto)
        self.assertIn("promedio 22 m", texto)
        self.assertIn("Sin coincidencia: 2240", texto)

    def test_resumen_telemetria_avisa_retraso_de_logitrack(self):
        filas = [cruzar(unidad_samsara(mph=0), [lectura_logitrack(kmh=0, hora="2026-09-07 13:23:00")])]
        self.assertIn("Logitrack con retraso: mediana 37 min", "\n".join(envio.resumen_telemetria(filas)))
        filas = [cruzar(unidad_samsara(mph=0), [lectura_logitrack(kmh=0)])]
        self.assertNotIn("retraso", "\n".join(envio.resumen_telemetria(filas)))

    def test_google_chat_muestra_solo_detenido_sin_detalle_de_telemetria(self):
        filas = [
            {"Unidad": "1", "Estatus": "DETENIDO CONFIRMADO", "Tiempo Detenido": "20 min", "Logitrack Encontrado": "SI"},
            {"Unidad": "2", "Estatus": "DETENIDO SAMSARA", "Tiempo Detenido": "8 min"},
            {"Unidad": "3", "Estatus": "DETENIDO LOGITRACK"},
            {"Unidad": "4", "Estatus": "RUTA"},
        ]
        texto = envio.construir_reporte_google(filas, AHORA, "EC-05")
        self.assertIn("Detenido: 3", texto)
        self.assertNotIn("CONFIRMADO", texto.upper())
        self.assertNotIn("SAMSARA", texto.upper())
        self.assertNotIn("LOGITRACK", texto.upper())
        self.assertIn("20 min", texto)
        # Lo confirmado por ambas telemetrias sigue apareciendo primero.
        self.assertLess(texto.index("1            |"), texto.index("2            |"))

    def test_chat_prueba_usa_webhook_de_pruebas(self):
        reporte = {"nombre": "Reporte EC-05", "entregas": [
            {"canal": "google_chat", "activo": True, "webhook_env": "GOOGLE_CHAT_WEBHOOK_URL"},
            {"canal": "correo", "activo": True},
        ]}
        enviados = []
        original = envio.enviar_google_chat
        envio.enviar_google_chat = lambda texto, url: enviados.append((texto, url))
        try:
            with mock.patch.dict(os.environ, {
                "GOOGLE_CHAT_WEBHOOK_URL": "https://produccion",
                "GOOGLE_CHAT_WEBHOOK_URL_PRUEBAS": "https://pruebas",
            }):
                envio.entregar_reporte(reporte, [], [], AHORA, False, chat_prueba=True)
        finally:
            envio.enviar_google_chat = original
        self.assertEqual(len(enviados), 1)
        self.assertEqual(enviados[0][1], "https://pruebas")
        self.assertIn("PRUEBA", enviados[0][0])

    def test_logitrack_prepara_candidato_con_doble_senal(self):
        tz = pytz.timezone("America/Mexico_City")
        now = tz.localize(datetime(2026, 9, 7, 14, 0))
        datos = [{
            "Unidad": "2234",
            "Estatus": "RUTA",
            "GpsActual": {
                "latitude": 25.70,
                "longitude": -100.34,
                "speedMilesPerHour": 0,
            },
        }]
        logitrack = [{
            "unit_name": "2234-TDR",
            "datetime": "2026-09-07 13:55:00",
            "lat": "25.7001",
            "lon": "-100.3401",
            "speed": 0,
            "engine_ign": 0,
        }]

        resultado = envio.enriquecer_con_logitrack(datos, logitrack, now, {})

        self.assertEqual(resultado[0]["Estatus"], "DETENIDO")
        self.assertEqual(resultado[0]["Logitrack Encontrado"], "SI")
        self.assertEqual(resultado[0]["Estatus Logitrack"], "DETENIDO")
        self.assertEqual(resultado[0]["Doble Comprobación Actual"], "SI")

    def test_logitrack_viejo_no_cuenta_como_doble_comprobacion(self):
        tz = pytz.timezone("America/Mexico_City")
        now = tz.localize(datetime(2026, 9, 7, 14, 0))
        datos = [{
            "Unidad": "2197",
            "Estatus": "DETENIDO",
            "GpsActual": {
                "latitude": 25.70,
                "longitude": -100.34,
                "speedMilesPerHour": 0,
            },
        }]
        logitrack = [{
            "unit_name": "2197-TDR",
            "datetime": "2026-09-07 10:00:00",
            "lat": "25.70",
            "lon": "-100.34",
            "speed": 0,
        }]

        resultado = envio.enriquecer_con_logitrack(datos, logitrack, now, {})

        self.assertEqual(resultado[0]["Logitrack Fresco"], "NO")
        self.assertEqual(resultado[0]["Estatus Logitrack"], "DESACTUALIZADO")
        self.assertEqual(resultado[0]["Doble Comprobación Actual"], "NO")

    def test_finaliza_detenido_confirmado_solo_con_historial_y_doble_senal(self):
        datos = [{
            "Estatus": "DETENIDO",
            "Minutos Detenido": 18,
            "Doble Comprobación Actual": "SI",
            "Logitrack Fresco": "SI",
        }]

        resultado = envio.finalizar_comprobacion_telemetria(datos, {})

        self.assertEqual(resultado[0]["Estatus"], "DETENIDO CONFIRMADO")
        self.assertEqual(
            resultado[0]["Fuente Confirmación"],
            "HISTORICO SAMSARA + LOGITRACK",
        )

    def test_sin_historial_no_confirma_aunque_ambas_lecturas_estan_detenidas(self):
        datos = [{
            "Estatus": "DETENIDO",
            "Minutos Detenido": None,
            "Doble Comprobación Actual": "SI",
            "Logitrack Fresco": "SI",
        }]

        resultado = envio.finalizar_comprobacion_telemetria(datos, {})

        self.assertEqual(resultado[0]["Estatus"], "REVISAR")

    def test_reporte_ec02_usa_solo_las_seis_etiquetas_sayer_directas(self):
        config = envio.cargar_configuracion(Path("config/reportes.json"))
        reporte = next(x for x in config["reportes"] if x["nombre"] == "Reporte EC-02")
        self.assertEqual(reporte["tipo_filtro_etiqueta"], "tagIds")
        self.assertEqual(
            reporte["etiquetas"],
            [
                "Sayer Full",
                "Sayer Patios y T.",
                "Sayer Pipas",
                "Sayer Sencillo",
                "Sayer Thorton",
                "Sayer Vuelteros",
            ],
        )

    def test_construye_jerarquia_de_padres_e_hijos(self):
        tags = [
            {"id": "1", "nombre": "Padre", "parentTagId": ""},
            {"id": "2", "nombre": "Hijo", "parentTagId": "1"},
            {"id": "3", "nombre": "Nieto", "parentTagId": "2"},
        ]
        catalogo = envio.construir_jerarquia_etiquetas(tags)
        self.assertEqual(catalogo["totalEtiquetas"], 3)
        self.assertEqual(catalogo["totalEtiquetasPadre"], 2)
        self.assertEqual(catalogo["jerarquia"][0]["hijos"][0]["nombre"], "Hijo")
        self.assertEqual(
            catalogo["jerarquia"][0]["hijos"][0]["hijos"][0]["nombre"], "Nieto"
        )

    def test_perfil_ec02_sobrescribe_exclusion_global(self):
        base = {"excluir_unidades_en_geocerca": True, "geocercas_especiales_ids": ["1"]}
        perfiles = {
            "EC-02": {
                "excluir_unidades_en_geocerca": False,
                "geocercas_especiales_ids": [],
            }
        }
        filtros = envio.combinar_filtros(base, {"perfil_filtros": "EC-02"}, perfiles)
        self.assertFalse(filtros["excluir_unidades_en_geocerca"])
        self.assertEqual(filtros["geocercas_especiales_ids"], [])

    def test_busqueda_remota_por_nombre_usa_external_id_y_cache(self):
        class Response:
            status_code = 200

            def raise_for_status(self):
                return None

            def json(self):
                return {"data": {"id": "55", "name": "Sayer Full"}}

        class Session:
            def __init__(self):
                self.urls = []

            def get(self, url, timeout):
                self.urls.append(url)
                return Response()

        session, cache = Session(), {}
        primero = envio.resolver_etiquetas_samsara(session, ["Sayer Full"], cache)
        segundo = envio.resolver_etiquetas_samsara(session, ["sayer full"], cache)
        self.assertEqual(primero, {"Sayer Full": "55"})
        self.assertEqual(segundo, {"sayer full": "55"})
        self.assertEqual(len(session.urls), 1)
        self.assertIn("samsara.name:Sayer%20Full", session.urls[0])

    def test_paginacion_conserva_parametros_y_cursor(self):
        class Response:
            status_code = 200
            url = "https://api.samsara.com/tags"

            def __init__(self, payload):
                self.payload = payload

            def raise_for_status(self):
                return None

            def json(self):
                return self.payload

        class Session:
            def __init__(self):
                self.calls = []

            def get(self, url, params, timeout):
                self.calls.append((url, params, timeout))
                if len(self.calls) == 1:
                    return Response({"data": [{"id": "1"}], "pagination": {"hasNextPage": True, "endCursor": "abc"}})
                return Response({"data": [{"id": "2"}], "pagination": {"hasNextPage": False}})

        session = Session()
        data = envio.obtener_paginas(session, "/tags", {"limit": 10})
        self.assertEqual([x["id"] for x in data], ["1", "2"])
        self.assertEqual(session.calls[0][1], {"limit": 10})
        self.assertEqual(session.calls[1][1], {"limit": 10, "after": "abc"})

    def test_resuelve_etiquetas_sin_importar_mayusculas_o_acentos(self):
        tags = [{"id": "10", "name": "Sayer Full"}, {"id": "20", "name": "Sáyer Pipas"}]
        resultado = envio.resolver_etiquetas(["sayer full", "SAYER PIPAS"], tags)
        self.assertEqual(resultado, {"sayer full": "10", "SAYER PIPAS": "20"})

    def test_etiqueta_faltante_falla_claramente(self):
        with self.assertRaisesRegex(envio.ConfiguracionError, "no encontradas"):
            envio.resolver_etiquetas(["No existe"], [])

    def test_geocerca_es_configurable_por_reporte(self):
        tz = pytz.timezone("America/Mexico_City")
        now = tz.localize(datetime(2026, 8, 19, 12, 0))
        vehicles = [{"id": "1", "name": "2239", "gps": {
            "time": now.isoformat(), "speedMilesPerHour": 0, "isEcuSpeed": True,
            "address": {"id": "99", "name": "Patio"},
        }}]
        incluidos, omitidos = envio.procesar_vehiculos(vehicles, now, {"gps_max_minutos": 60}, {"99": "Patio"})
        self.assertEqual(incluidos, [])
        self.assertEqual(omitidos[0]["Motivo"], "PATIO/GEOCERCA EXCLUIDA")
        self.assertIn("Coordenadas", omitidos[0])
        incluidos, omitidos = envio.procesar_vehiculos(vehicles, now, {"gps_max_minutos": 60}, {})
        self.assertEqual(omitidos, [])
        self.assertEqual(incluidos[0]["Estatus"], "DETENIDO")

    def test_incluir_todas_conserva_gps_viejo_y_speed_cero_sin_ecu(self):
        tz = pytz.timezone("America/Mexico_City")
        now = tz.localize(datetime(2026, 8, 19, 12, 0))
        vehicles = [{"id": "1", "name": "100", "gps": {
            "time": "2026-08-18T10:00:00-06:00", "speedMilesPerHour": 0,
            "isEcuSpeed": False, "latitude": 19.1, "longitude": -99.2,
            "reverseGeo": {"formattedLocation": "Ubicacion anterior"},
        }}]
        filtros = {
            "gps_max_minutos": 60,
            "excluir_speed_cero_sin_ecu": True,
            "incluir_todas_las_unidades": True,
        }
        incluidos, omitidos = envio.procesar_vehiculos(vehicles, now, filtros, {})
        self.assertEqual(len(incluidos), 1)
        self.assertEqual(omitidos, [])
        self.assertEqual(incluidos[0]["Coordenadas"], "19.1,-99.2")

    def test_operador_samsara_usa_asignacion_mas_reciente(self):
        class Session:
            def get(self, url, params, timeout):
                class Response:
                    status_code = 200
                    url = "https://api.samsara.com/fleet/driver-vehicle-assignments"

                    def raise_for_status(self):
                        return None

                    def json(self):
                        return {"data": [
                            {"startTime": "2026-08-20T10:00:00Z", "isPassenger": False,
                             "vehicle": {"id": "1"}, "driver": {"id": "10", "name": "Anterior"}},
                            {"startTime": "2026-08-20T11:00:00Z", "isPassenger": False,
                             "vehicle": {"id": "1"}, "driver": {"id": "20", "name": "Actual"}},
                        ], "pagination": {"hasNextPage": False}}
                return Response()

        now = datetime(2026, 8, 20, 12, 0, tzinfo=pytz.UTC)
        datos = [{"Unidad": "100", "SamsaraVehicleId": "1"}]
        resultado = envio.enriquecer_operadores_samsara(Session(), datos, now)
        self.assertEqual(resultado[0]["Operador"], "Actual")
        self.assertEqual(resultado[0]["ID Operador"], "20")

    def test_excel_contiene_tres_hojas(self):
        now = datetime(2026, 8, 19, 12, 0, tzinfo=pytz.UTC)
        datos = [{"Unidad": "100", "Estatus": "RUTA", "Fecha GPS": now}]
        omitidas = [{"Unidad": "200", "Motivo": "GPS VIEJO"}]
        contenido = {"columnas": ["Unidad", "Estatus", "Fecha GPS"], "incluir_omitidas_excel": True}
        archivo = envio.crear_excel_reporte("Prueba", datos, omitidas, now, contenido)
        wb = load_workbook(BytesIO(archivo), read_only=True, data_only=True)
        self.assertEqual(wb.sheetnames, ["Resumen", "Unidades", "Omitidas"])
        self.assertEqual(wb["Unidades"]["A6"].value, "100")
        self.assertEqual(wb["Omitidas"]["A2"].value, "200")
        wb.close()

    def test_excel_puede_concentrarse_en_una_sola_hoja(self):
        now = datetime(2026, 8, 20, 12, 0, tzinfo=pytz.UTC)
        datos = [
            {"Unidad": "100", "Estatus": "RUTA"},
            {"Unidad": "200", "Estatus": "DETENIDO"},
        ]
        contenido = {
            "columnas": ["Unidad", "Estatus"],
            "resumen_estados_excel": True,
            "incluir_resumen_excel": False,
            "incluir_omitidas_excel": False,
        }
        archivo = envio.crear_excel_reporte("EC-02", datos, [], now, contenido)
        wb = load_workbook(BytesIO(archivo), read_only=True, data_only=True)
        self.assertEqual(wb.sheetnames, ["Unidades"])
        self.assertEqual(wb["Unidades"]["A6"].value, "100")
        self.assertEqual(wb["Unidades"]["B3"].value, 2)
        self.assertEqual(wb["Unidades"]["D3"].value, 1)
        self.assertEqual(wb["Unidades"]["F3"].value, 1)
        wb.close()

    def test_contenido_ordena_detenidos_antes_de_ruta(self):
        datos = [
            {"Unidad": "2", "Estatus": "RUTA"},
            {"Unidad": "10", "Estatus": "DETENIDO"},
            {"Unidad": "1", "Estatus": "DETENIDO"},
        ]
        resultado = envio.filtrar_contenido(
            datos, {"orden_estados": ["DETENIDO", "RUTA"]}
        )
        self.assertEqual([x["Unidad"] for x in resultado], ["1", "10", "2"])

    def test_divide_chat_sin_perder_lineas(self):
        self.assertEqual(envio.dividir_mensaje("uno\ndos\ntres", 8), ["uno\ndos", "tres"])

    def test_google_chat_incluye_coordenadas(self):
        tz = pytz.timezone("America/Mexico_City")
        now = tz.localize(datetime(2026, 8, 19, 12, 0))
        texto = envio.construir_reporte_google(
            [{"Unidad": "100", "Estatus": "RUTA", "Coordenadas": "19.1,-99.2"}],
            now,
            "Prueba",
        )
        self.assertIn("COORDENADAS", texto)
        self.assertIn("19.1,-99.2", texto)

    def test_google_chat_resume_detenidos_confirmados(self):
        tz = pytz.timezone("America/Mexico_City")
        now = tz.localize(datetime(2026, 9, 7, 14, 0))
        texto = envio.construir_reporte_google(
            [{
                "Unidad": "2234",
                "Estatus": "DETENIDO CONFIRMADO",
                "Tiempo Detenido": "12 min",
            }],
            now,
            "EC-05",
        )
        self.assertIn("Detenido: 1", texto)
        self.assertIn("12 min", texto)

    def test_entregas_usan_activas_si_no_se_fuerza_canal(self):
        reporte = {
            "nombre": "EC-05",
            "entregas": [
                {"canal": "google_chat", "activo": True},
                {"canal": "correo", "activo": False},
            ],
        }
        entregas = envio.seleccionar_entregas(reporte)
        self.assertEqual([x["canal"] for x in entregas], ["google_chat"])

    def test_canal_forzado_selecciona_correo_aunque_este_inactivo(self):
        reporte = {
            "nombre": "EC-05",
            "entregas": [
                {"canal": "google_chat", "activo": True},
                {"canal": "correo", "activo": False},
            ],
        }
        entregas = envio.seleccionar_entregas(reporte, "correo")
        self.assertEqual([x["canal"] for x in entregas], ["correo"])

    def test_solo_puede_seleccionar_reporte_inactivo(self):
        config = {
            "reportes": [
                {"nombre": "Reporte EC-05", "activo": True},
                {"nombre": "Reporte EC-02", "activo": False},
            ]
        }
        seleccionados = envio.seleccionar_reportes(config, ["Reporte EC-02"])
        self.assertEqual([x["nombre"] for x in seleccionados], ["Reporte EC-02"])

    def test_sin_solo_conserva_unicamente_reportes_activos(self):
        config = {
            "reportes": [
                {"nombre": "Reporte EC-05", "activo": True},
                {"nombre": "Reporte EC-02", "activo": False},
            ]
        }
        seleccionados = envio.seleccionar_reportes(config, None)
        self.assertEqual([x["nombre"] for x in seleccionados], ["Reporte EC-05"])


if __name__ == "__main__":
    unittest.main()
