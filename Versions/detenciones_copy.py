import logging
import math
import time
import requests
from datetime import timedelta, timezone
from dateutil import parser as dp


# ----------------------------------------------------------------------
# CONFIGURACION DE DETENCIONES

STOP_SPEED_MPH = 1.0
STOP_RADIUS_METERS = 100

MIN_STOP_MINUTES = 5
RECENT_WINDOW_MINUTES = 15
ECU_FRESH_MINUTES = 5
ENGINE_FRESH_MINUTES = 30

# Si el inicio detenido queda cerca del inicio de la ventana,
# se consulta una ventana mayor para buscar el inicio real.
WINDOW_START_TOLERANCE_MINUTES = 10

# Ventanas progresivas para no consultar 24h desde el inicio.
VENTANAS_HORAS = [1, 3, 6, 12, 24]

# Umbrales para diferenciar trafico lento vs ruta
MIN_DISTANCIA_LINEAL_RUTA = 80          # metros
MIN_DISTANCIA_ACUMULADA_TRAFICO = 120   # metros


# ----------------------------------------------------------------------
# DISTANCIA ENTRE DOS COORDENADAS
def haversine_meters(lat1, lon1, lat2, lon2):
    radio_tierra = 6371000

    lat1 = math.radians(float(lat1))
    lon1 = math.radians(float(lon1))
    lat2 = math.radians(float(lat2))
    lon2 = math.radians(float(lon2))

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    )

    return 2 * radio_tierra * math.atan2(math.sqrt(a), math.sqrt(1 - a))


# ----------------------------------------------------------------------
# UTILIDADES DE TIEMPO
def parse_time(value):
    if not value:
        return None

    try:
        return dp.parse(value)
    except Exception:
        return None


def esta_pegado_al_inicio_ventana(detenido_desde, start_time, tolerancia_minutos=10):
    """
    Si detenido_desde está muy cerca del inicio de la ventana consultada,
    probablemente la unidad ya venía detenida desde antes.

    Ejemplo:
    - Ventana 1h inicia 13:30
    - detenido_desde 13:31
    - Esto NO significa que empezó a las 13:31.
      Significa que la ventana no alcanzó a ver más atrás.
    """

    if not detenido_desde or not start_time:
        return False

    try:
        detenido_dt = detenido_desde.astimezone(timezone.utc)
        start_dt = start_time.astimezone(timezone.utc)

        diferencia_min = abs((detenido_dt - start_dt).total_seconds()) / 60

        return diferencia_min <= tolerancia_minutos

    except Exception:
        return False


def filtrar_ultimos_minutos(gps_history, minutos=15):
    if not gps_history:
        return []

    puntos = [
        p for p in gps_history
        if p.get("time")
    ]

    if not puntos:
        return []

    puntos = sorted(puntos, key=lambda x: dp.parse(x["time"]))
    ultimo_dt = dp.parse(puntos[-1]["time"])
    limite = ultimo_dt - timedelta(minutes=minutos)

    return [
        p for p in puntos
        if dp.parse(p["time"]) >= limite
    ]


# ----------------------------------------------------------------------
# ENGINE STATE Y ECU SPEED
def normalizar_engine_state(valor):
    if valor is None:
        return ""

    if isinstance(valor, dict):
        valor = valor.get("value") or valor.get("state") or valor.get("engineState")

    texto = str(valor).strip()

    mapa = {
        "Off": "OFF",
        "On": "ON",
        "Idle": "IDLE",
        "off": "OFF",
        "on": "ON",
        "idle": "IDLE",
    }

    return mapa.get(texto, texto.upper())


def obtener_ultimo_engine_state(engine_states, referencia_time=None, max_minutos=30):
    if not engine_states:
        return ""

    puntos = [
        e for e in engine_states
        if e.get("time")
    ]

    if not puntos:
        return ""

    puntos = sorted(puntos, key=lambda x: dp.parse(x["time"]))

    ref_dt = parse_time(referencia_time) if referencia_time else None

    if not ref_dt:
        ultimo = puntos[-1]
        return normalizar_engine_state(
            ultimo.get("value") or ultimo.get("state") or ultimo.get("engineState")
        )

    candidatos = [
        e for e in puntos
        if dp.parse(e["time"]) <= ref_dt
    ]

    if not candidatos:
        return ""

    ultimo = candidatos[-1]
    ultimo_dt = dp.parse(ultimo["time"])

    diferencia_min = abs((ref_dt - ultimo_dt).total_seconds()) / 60

    if diferencia_min > max_minutos:
        return ""

    return normalizar_engine_state(
        ultimo.get("value") or ultimo.get("state") or ultimo.get("engineState")
    )


def obtener_ecu_speed_fresco(ecu_speed_history, referencia_time=None, max_minutos=5):
    if not ecu_speed_history:
        return None

    puntos = [
        e for e in ecu_speed_history
        if e.get("time") and e.get("value") is not None
    ]

    if not puntos:
        return None

    puntos = sorted(puntos, key=lambda x: dp.parse(x["time"]))

    ref_dt = parse_time(referencia_time) if referencia_time else None

    if not ref_dt:
        try:
            return float(puntos[-1].get("value"))
        except Exception:
            return None

    candidatos = [
        e for e in puntos
        if dp.parse(e["time"]) <= ref_dt
    ]

    if not candidatos:
        return None

    ultimo = candidatos[-1]
    ultimo_dt = dp.parse(ultimo["time"])

    diferencia_min = abs((ref_dt - ultimo_dt).total_seconds()) / 60

    if diferencia_min > max_minutos:
        return None

    try:
        return float(ultimo.get("value"))
    except Exception:
        return None


# ----------------------------------------------------------------------
# VALIDAR SI UN PUNTO HISTORICO SIGUE CERCA DEL GPS ACTUAL
def es_punto_detenido_contra_actual(punto, gps_actual):
    try:
        speed = float(punto.get("speedMilesPerHour") or 0)

        distancia = haversine_meters(
            punto.get("latitude"),
            punto.get("longitude"),
            gps_actual.get("latitude"),
            gps_actual.get("longitude"),
        )

        return speed <= STOP_SPEED_MPH and distancia <= STOP_RADIUS_METERS

    except Exception:
        return False


# ----------------------------------------------------------------------
# TRAMO FINAL DETENIDO
def calcular_tramo_final_detenido(gps_history, gps_actual, fecha_hora_reporte):
    """
    Valida primero el estado actual real.

    Si la unidad se movio antes en la ventana pero al final quedó quieta,
    debe marcar DETENIDO, no TRAFICO LENTO.
    """

    if not gps_history:
        return {
            "esta_detenido": False,
            "detenido_desde": None,
            "minutos_detenido": None,
            "motivo": "Sin historial GPS",
        }

    puntos = [
        p for p in gps_history
        if p.get("time")
        and p.get("latitude") is not None
        and p.get("longitude") is not None
    ]

    if not puntos:
        return {
            "esta_detenido": False,
            "detenido_desde": None,
            "minutos_detenido": None,
            "motivo": "Sin puntos GPS validos",
        }

    puntos = sorted(puntos, key=lambda x: dp.parse(x["time"]))

    detenido_desde = None

    for punto in reversed(puntos):
        try:
            speed = float(punto.get("speedMilesPerHour") or 0)

            distancia_actual = haversine_meters(
                punto.get("latitude"),
                punto.get("longitude"),
                gps_actual.get("latitude"),
                gps_actual.get("longitude"),
            )

            if speed <= STOP_SPEED_MPH and distancia_actual <= STOP_RADIUS_METERS:
                detenido_desde = punto.get("time")
            else:
                break

        except Exception:
            break

    if not detenido_desde:
        return {
            "esta_detenido": False,
            "detenido_desde": None,
            "minutos_detenido": None,
            "motivo": "No hay tramo final detenido",
        }

    detenido_desde_dt = dp.parse(detenido_desde).astimezone(fecha_hora_reporte.tzinfo)

    minutos = int((fecha_hora_reporte - detenido_desde_dt).total_seconds() // 60)
    minutos = max(minutos, 0)

    return {
        "esta_detenido": minutos >= MIN_STOP_MINUTES,
        "detenido_desde": detenido_desde_dt,
        "minutos_detenido": minutos,
        "motivo": "Tramo final detenido",
    }


# ----------------------------------------------------------------------
# DISTANCIA LINEAL ENTRE PRIMER Y ULTIMO PUNTO
def distancia_entre_primer_y_ultimo(gps_history):
    if not gps_history or len(gps_history) < 2:
        return 0

    puntos = sorted(gps_history, key=lambda x: dp.parse(x["time"]))

    primero = puntos[0]
    ultimo = puntos[-1]

    try:
        return haversine_meters(
            primero.get("latitude"),
            primero.get("longitude"),
            ultimo.get("latitude"),
            ultimo.get("longitude"),
        )
    except Exception:
        return 0


# ----------------------------------------------------------------------
# DISTANCIA ACUMULADA ENTRE TODOS LOS PUNTOS
def distancia_total_recorrida(gps_history):
    if not gps_history or len(gps_history) < 2:
        return 0

    puntos = sorted(gps_history, key=lambda x: dp.parse(x["time"]))

    total = 0

    for i in range(1, len(puntos)):
        p1 = puntos[i - 1]
        p2 = puntos[i]

        try:
            total += haversine_meters(
                p1.get("latitude"),
                p1.get("longitude"),
                p2.get("latitude"),
                p2.get("longitude"),
            )
        except Exception:
            continue

    return total


# ----------------------------------------------------------------------
# CLASIFICAR MOVIMIENTO POR HISTORIAL RECIENTE
def determinar_estado_por_historial(gps_history):
    """
    Esta funcion debe usarse solo con los ultimos 10-15 minutos.
    No debe usarse con toda la ventana historica.
    """

    if not gps_history:
        return "DETENIDO", 0, 0

    puntos = [
        p for p in gps_history
        if p.get("time")
        and p.get("latitude") is not None
        and p.get("longitude") is not None
    ]

    if len(puntos) < 2:
        return "DETENIDO", 0, 0

    distancia_lineal = distancia_entre_primer_y_ultimo(puntos)
    distancia_acumulada = distancia_total_recorrida(puntos)

    velocidades = []

    for p in puntos:
        try:
            velocidades.append(float(p.get("speedMilesPerHour") or 0))
        except Exception:
            continue

    max_speed = max(velocidades) if velocidades else 0

    if max_speed > 5 or distancia_lineal >= MIN_DISTANCIA_LINEAL_RUTA:
        return "RUTA", distancia_lineal, distancia_acumulada

    if distancia_acumulada >= MIN_DISTANCIA_ACUMULADA_TRAFICO:
        return "TRAFICO LENTO", distancia_lineal, distancia_acumulada

    return "DETENIDO", distancia_lineal, distancia_acumulada


# ----------------------------------------------------------------------
# CALCULAR DESDE CUANDO ESTA DETENIDO
def calcular_inicio_detenido(gps_history, gps_actual, fecha_hora_reporte):
    tramo_final = calcular_tramo_final_detenido(
        gps_history=gps_history,
        gps_actual=gps_actual,
        fecha_hora_reporte=fecha_hora_reporte,
    )

    return {
        "detenido_desde": tramo_final.get("detenido_desde"),
        "minutos_detenido": tramo_final.get("minutos_detenido"),
        "motivo": tramo_final.get("motivo"),
    }


# ----------------------------------------------------------------------
# CALCULAR DESDE CUANDO VIENE EN TRAFICO LENTO
def calcular_inicio_trafico_lento(gps_history, fecha_hora_reporte):
    """
    Calcula desde cuando viene en trafico lento usando solo una ventana reciente.
    """

    if not gps_history or len(gps_history) < 2:
        return {
            "trafico_desde": None,
            "minutos_trafico": None,
            "motivo": "Sin historial GPS suficiente",
        }

    puntos = sorted(gps_history, key=lambda x: dp.parse(x["time"]))

    trafico_desde = None

    for i in range(len(puntos) - 1, 0, -1):
        punto_actual = puntos[i]
        punto_anterior = puntos[i - 1]

        try:
            speed_actual = float(punto_actual.get("speedMilesPerHour") or 0)
            speed_anterior = float(punto_anterior.get("speedMilesPerHour") or 0)

            distancia = haversine_meters(
                punto_anterior.get("latitude"),
                punto_anterior.get("longitude"),
                punto_actual.get("latitude"),
                punto_actual.get("longitude"),
            )

        except Exception:
            break

        if speed_actual > 5 or speed_anterior > 5:
            break

        if distancia > MIN_DISTANCIA_LINEAL_RUTA:
            break

        trafico_desde = punto_anterior.get("time")

    if not trafico_desde:
        return {
            "trafico_desde": None,
            "minutos_trafico": None,
            "motivo": "No se encontro inicio de trafico lento",
        }

    trafico_desde_dt = dp.parse(trafico_desde).astimezone(fecha_hora_reporte.tzinfo)

    minutos = int((fecha_hora_reporte - trafico_desde_dt).total_seconds() // 60)

    return {
        "trafico_desde": trafico_desde_dt,
        "minutos_trafico": max(minutos, 0),
        "motivo": "Calculado por GPS reciente",
    }


# ----------------------------------------------------------------------
# CONSULTAR HISTORIAL EN SAMSARA
def consultar_historial_gps(token, vehicle_ids, start_time, end_time, timeout=60):
    url = "https://api.samsara.com/fleet/vehicles/stats/history"

    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
    }

    params = {
        "types": "gps,engineStates,ecuSpeedMph",
        "vehicleIds": ",".join(str(v) for v in vehicle_ids),
        "startTime": start_time.isoformat().replace("+00:00", "Z"),
        "endTime": end_time.isoformat().replace("+00:00", "Z"),
    }

    data_total = []
    after = None

    while True:
        params_consulta = dict(params)

        if after:
            params_consulta["after"] = after

        response = requests.get(
            url,
            headers=headers,
            params=params_consulta,
            timeout=timeout,
        )

        response.raise_for_status()

        payload = response.json()

        data_total.extend(payload.get("data", []))

        pagination = payload.get("pagination", {})

        if not pagination.get("hasNextPage"):
            break

        after = pagination.get("endCursor")

        if not after:
            break

        time.sleep(0.2)

    return data_total


# ----------------------------------------------------------------------
# FORMATEAR MINUTOS
def formatear_tiempo_detenido(minutos):
    if minutos is None:
        return "N/D"

    horas = minutos // 60
    mins = minutos % 60

    if horas <= 0:
        return f"{mins} min"

    return f"{horas}h {mins}m"


# ----------------------------------------------------------------------
# ENRIQUECER RESULTS CON ESTATUS REAL
def enriquecer_minutos_detenido(results, token, now_mx):
    """
    Procesa solamente unidades que inicialmente vienen como DETENIDO.

    Prioridad:
    1. Tramo final detenido.
    2. Si el inicio detenido está pegado al inicio de la ventana,
       buscar una ventana mayor para encontrar el inicio real.
    3. ECU speed fresco.
    4. Movimiento reciente de los ultimos 15 min.
    """

    candidatos = [
        row for row in results
        if str(row.get("Estatus", "")).strip() == "DETENIDO"
        and row.get("SamsaraVehicleId")
        and row.get("GpsActual")
    ]

    print(f"[Detenciones] Candidatos DETENIDO: {len(candidatos)}")

    for row in candidatos:
        row["_detencion_confirmada"] = False

        print(
            f"[Detenciones] Unidad {row.get('Unidad')} | "
            f"Estatus inicial: {row.get('Estatus')} | "
            f"ID Samsara: {row.get('SamsaraVehicleId')} | "
            f"GPS actual: {bool(row.get('GpsActual'))}"
        )

    if not candidatos:
        return results

    for horas in VENTANAS_HORAS:
        pendientes = [
            row for row in candidatos
            if not row.get("_detencion_confirmada")
            and str(row.get("Estatus", "")).strip() == "DETENIDO"
        ]

        if not pendientes:
            break

        start_time = (now_mx - timedelta(hours=horas)).astimezone(timezone.utc)
        end_time = now_mx.astimezone(timezone.utc)

        vehicle_ids = [row["SamsaraVehicleId"] for row in pendientes]

        try:
            historiales = consultar_historial_gps(
                token=token,
                vehicle_ids=vehicle_ids,
                start_time=start_time,
                end_time=end_time,
            )

            print(
                f"[Detenciones] Ventana {horas}h | "
                f"Vehiculos consultados: {len(vehicle_ids)} | "
                f"Historiales recibidos: {len(historiales)}"
            )

        except Exception as e:
            logging.exception(f"Error consultando historial GPS ventana {horas}h")
            print(f"[Detenciones][ERROR] Ventana {horas}h: {e}")
            continue

        historiales_por_id = {
            str(item.get("id")): item
            for item in historiales
        }

        for row in pendientes:
            vehicle_id = str(row.get("SamsaraVehicleId"))
            item_historial = historiales_por_id.get(vehicle_id, {})

            gps_history = item_historial.get("gps", [])
            engine_states = item_historial.get("engineStates", [])
            ecu_speed_history = item_historial.get("ecuSpeedMph", [])

            gps_actual = row.get("GpsActual") or {}
            gps_actual_time = gps_actual.get("time")

            engine_actual = obtener_ultimo_engine_state(
                engine_states=engine_states,
                referencia_time=gps_actual_time,
                max_minutos=ENGINE_FRESH_MINUTES,
            )

            ecu_actual = obtener_ecu_speed_fresco(
                ecu_speed_history=ecu_speed_history,
                referencia_time=gps_actual_time,
                max_minutos=ECU_FRESH_MINUTES,
            )

            row["Motor"] = engine_actual
            row["EcuSpeedActual"] = ecu_actual

            print(
                f"[Detenciones] Unidad {row.get('Unidad')} | "
                f"Motor={engine_actual or 'N/D'} | "
                f"ECU fresco={ecu_actual if ecu_actual is not None else 'N/D'}"
            )

            # 1. Primero validamos tramo final detenido.
            tramo_final = calcular_tramo_final_detenido(
                gps_history=gps_history,
                gps_actual=gps_actual,
                fecha_hora_reporte=now_mx,
            )

            print(
                f"[Detenciones] Unidad {row.get('Unidad')} | "
                f"Tramo final detenido={tramo_final.get('esta_detenido')} | "
                f"Desde={tramo_final.get('detenido_desde')} | "
                f"Min={tramo_final.get('minutos_detenido')} | "
                f"Motivo={tramo_final.get('motivo')}"
            )

            if tramo_final["esta_detenido"]:
                row["Estatus"] = "DETENIDO"
                row["Detenido Desde"] = tramo_final["detenido_desde"]
                row["Minutos Detenido"] = tramo_final["minutos_detenido"]
                row["Tiempo Detenido"] = formatear_tiempo_detenido(
                    tramo_final["minutos_detenido"]
                )
                row["Tiempo Trafico"] = ""
                row["Minutos Trafico"] = None
                row["Trafico Desde"] = None
                row["Ventana Detenido"] = f"{horas}h"

                pegado_inicio = esta_pegado_al_inicio_ventana(
                    detenido_desde=tramo_final["detenido_desde"],
                    start_time=start_time,
                    tolerancia_minutos=WINDOW_START_TOLERANCE_MINUTES,
                )

                print(
                    f"[Detenciones] Unidad {row.get('Unidad')} | "
                    f"Pegado inicio ventana={pegado_inicio} | "
                    f"Ventana={horas}h"
                )

                # Si está pegado al inicio y aún hay ventanas mayores,
                # no confirmamos. Seguimos buscando una ventana más amplia.
                if pegado_inicio and horas != VENTANAS_HORAS[-1]:
                    continue

                row["_detencion_confirmada"] = True
                continue

            # 2. Si ECU speed fresco marca movimiento claro, es RUTA.
            if ecu_actual is not None and ecu_actual > 5:
                row["Estatus"] = "RUTA"
                row["Tiempo Detenido"] = ""
                row["Minutos Detenido"] = None
                row["Detenido Desde"] = None
                row["Tiempo Trafico"] = ""
                row["Minutos Trafico"] = None
                row["Trafico Desde"] = None
                row["Ventana Detenido"] = f"{horas}h"
                row["_detencion_confirmada"] = True

                print(
                    f"[Detenciones] Unidad {row.get('Unidad')} | "
                    f"RUTA por ECU speed fresco > 5 mph"
                )
                continue

            # 3. Si no está detenido, analizamos solo últimos 15 minutos.
            gps_history_reciente = filtrar_ultimos_minutos(
                gps_history,
                minutos=RECENT_WINDOW_MINUTES,
            )

            estado_historial, distancia_lineal, distancia_acumulada = (
                determinar_estado_por_historial(gps_history_reciente)
            )

            print(
                f"[Detenciones] Unidad {row.get('Unidad')} | "
                f"Estado reciente={estado_historial} | "
                f"Puntos recientes={len(gps_history_reciente)} | "
                f"Distancia lineal={distancia_lineal:.1f} m | "
                f"Distancia acumulada={distancia_acumulada:.1f} m"
            )

            if estado_historial == "RUTA":
                row["Estatus"] = "RUTA"
                row["Tiempo Detenido"] = ""
                row["Minutos Detenido"] = None
                row["Detenido Desde"] = None
                row["Tiempo Trafico"] = ""
                row["Minutos Trafico"] = None
                row["Trafico Desde"] = None
                row["Ventana Detenido"] = f"{horas}h"
                row["_detencion_confirmada"] = True
                continue

            if estado_historial == "TRAFICO LENTO":
                calculo_trafico = calcular_inicio_trafico_lento(
                    gps_history=gps_history_reciente,
                    fecha_hora_reporte=now_mx,
                )

                row["Estatus"] = "TRAFICO LENTO"
                row["Tiempo Detenido"] = ""
                row["Minutos Detenido"] = None
                row["Detenido Desde"] = None
                row["Ventana Detenido"] = f"{horas}h"
                row["_detencion_confirmada"] = True

                if calculo_trafico["trafico_desde"] is not None:
                    row["Trafico Desde"] = calculo_trafico["trafico_desde"]
                    row["Minutos Trafico"] = calculo_trafico["minutos_trafico"]
                    row["Tiempo Trafico"] = formatear_tiempo_detenido(
                        calculo_trafico["minutos_trafico"]
                    )
                else:
                    row["Trafico Desde"] = None
                    row["Minutos Trafico"] = None
                    row["Tiempo Trafico"] = "N/D"

                continue

            # 4. Si el historial reciente tampoco muestra movimiento claro,
            # se mantiene como detenido.
            calculo = calcular_inicio_detenido(
                gps_history=gps_history,
                gps_actual=gps_actual,
                fecha_hora_reporte=now_mx,
            )

            print(
                f"[Detenciones] Unidad {row.get('Unidad')} | "
                f"Desde={calculo.get('detenido_desde')} | "
                f"Min={calculo.get('minutos_detenido')} | "
                f"Motivo={calculo.get('motivo')}"
            )

            if calculo["detenido_desde"] is not None:
                row["Estatus"] = "DETENIDO"
                row["Detenido Desde"] = calculo["detenido_desde"]
                row["Minutos Detenido"] = calculo["minutos_detenido"]
                row["Tiempo Detenido"] = formatear_tiempo_detenido(
                    calculo["minutos_detenido"]
                )
                row["Ventana Detenido"] = f"{horas}h"

                pegado_inicio = esta_pegado_al_inicio_ventana(
                    detenido_desde=calculo["detenido_desde"],
                    start_time=start_time,
                    tolerancia_minutos=WINDOW_START_TOLERANCE_MINUTES,
                )

                if pegado_inicio and horas != VENTANAS_HORAS[-1]:
                    continue

                row["_detencion_confirmada"] = True

    for row in candidatos:
        if str(row.get("Estatus", "")).strip() == "DETENIDO":
            if row.get("Tiempo Detenido") is None:
                row["Tiempo Detenido"] = "N/D"

    for row in results:
        row.pop("_detencion_confirmada", None)

    return results