"""Migra HistoricoKmSamsara a una tabla corregida y auditable.

El histórico original nunca se actualiza. Sin ``--aplicar`` el programa solo
simula el resultado y muestra los totales que produciría la migración.

Reglas de corrección
--------------------
1. ``DiferenciaKmCorregida`` se recalcula como lectura final menos lectura
   inicial. Se descartan valores negativos y saltos diarios mayores al límite.
2. ``KmNoRegistradosCorregidos`` solo se calcula entre registros consecutivos
   del mismo ``IdTracto`` y el mismo ``Origen``.
3. Un cambio GPS <-> Odómetro reinicia la referencia y produce cero kilómetros
   no registrados; nunca se restan contadores con bases distintas.
4. El original y el motivo de cada corrección se conservan en la tabla destino.

Ejemplos
--------
Simulación de solo lectura::

    python scripts/migrar_historico_km_samsara.py

Crear la tabla corregida::

    python scripts/migrar_historico_km_samsara.py --aplicar

Recrear una tabla corregida existente (destructivo solo para la tabla destino)::

    python scripts/migrar_historico_km_samsara.py --aplicar --reemplazar-destino
"""

from __future__ import annotations

import argparse
import re
import sys
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

import pyodbc
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).resolve().parent))
from comparar_kilometros_mensuales import conectar_sql, localizar_tabla_historica


REPO_ROOT = Path(__file__).resolve().parents[1]
NOMBRE_DESTINO_PREDETERMINADO = "HistoricoKmSamsaraCorregido"


@dataclass(frozen=True)
class TablaSql:
    base: str
    esquema: str
    tabla: str

    @property
    def nombre_completo(self) -> str:
        return ".".join(f"[{parte.replace(']', ']]')}]" for parte in (
            self.base,
            self.esquema,
            self.tabla,
        ))


def argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Crea un histórico corregido sin modificar HistoricoKmSamsara. "
            "Sin --aplicar funciona como simulación de solo lectura."
        )
    )
    parser.add_argument(
        "--aplicar",
        action="store_true",
        help="Crea y llena la tabla destino dentro de una transacción.",
    )
    parser.add_argument(
        "--reemplazar-destino",
        action="store_true",
        help=(
            "Elimina la tabla destino si ya existe. Solo funciona junto con "
            "--aplicar; nunca elimina la tabla original."
        ),
    )
    parser.add_argument(
        "--tabla-destino",
        default=NOMBRE_DESTINO_PREDETERMINADO,
        help=(
            "Nombre simple de la tabla destino, en la misma base y esquema de "
            "la tabla original."
        ),
    )
    parser.add_argument(
        "--max-km-dia",
        type=Decimal,
        default=Decimal("1500"),
        help="Máximo recorrido aceptable dentro de un registro diario.",
    )
    parser.add_argument(
        "--max-km-entre-lecturas",
        type=Decimal,
        default=Decimal("1500"),
        help=(
            "Máximo salto aceptable entre el final anterior y el inicio actual "
            "cuando la fuente es la misma."
        ),
    )
    return parser.parse_args()


def validar_argumentos(args: argparse.Namespace) -> None:
    if args.reemplazar_destino and not args.aplicar:
        raise SystemExit("--reemplazar-destino requiere también --aplicar.")
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", args.tabla_destino):
        raise SystemExit(
            "--tabla-destino solo puede contener letras, números y guion bajo."
        )
    if args.max_km_dia <= 0 or args.max_km_entre_lecturas <= 0:
        raise SystemExit("Los límites de kilómetros deben ser mayores que cero.")


def parsear_tabla(nombre_completo: str) -> TablaSql:
    coincidencia = re.fullmatch(
        r"\[([^]]+(?:]][^]]*)*)\]\.\[([^]]+(?:]][^]]*)*)\]\.\[([^]]+(?:]][^]]*)*)\]",
        nombre_completo,
    )
    if not coincidencia:
        raise RuntimeError(f"Nombre SQL inesperado: {nombre_completo}")
    partes = tuple(parte.replace("]]", "]") for parte in coincidencia.groups())
    return TablaSql(*partes)


def tabla_existe(cursor: pyodbc.Cursor, tabla: TablaSql) -> bool:
    consulta = (
        f"SELECT COUNT(*) FROM [{tabla.base.replace(']', ']]')}].sys.tables t "
        f"JOIN [{tabla.base.replace(']', ']]')}].sys.schemas s "
        "ON s.schema_id = t.schema_id "
        "WHERE s.name = ? AND t.name = ?"
    )
    return bool(cursor.execute(consulta, tabla.esquema, tabla.tabla).fetchone()[0])


def expresiones(max_dia: str, max_gap: str) -> dict[str, str]:
    delta_dia = (
        "(TRY_CONVERT(float, DistranciaMetrosFinal) - "
        " TRY_CONVERT(float, DistanciaMetrosInicial)) / 1000.0"
    )
    delta_gap = (
        "(TRY_CONVERT(float, DistanciaMetrosInicial) - "
        " TRY_CONVERT(float, FinalAnterior)) / 1000.0"
    )
    diferencia_corregida = f"""
        CASE
            WHEN TRY_CONVERT(float, DistanciaMetrosInicial) IS NULL
              OR TRY_CONVERT(float, DistranciaMetrosFinal) IS NULL THEN 0
            WHEN {delta_dia} < 0 THEN 0
            WHEN {delta_dia} > {max_dia} THEN 0
            ELSE {delta_dia}
        END
    """
    km_no_corregidos = f"""
        CASE
            WHEN FinalAnterior IS NULL THEN 0
            WHEN COALESCE(OrigenAnterior, '') <> COALESCE(Origen, '') THEN 0
            WHEN TRY_CONVERT(float, DistanciaMetrosInicial) IS NULL THEN 0
            WHEN {delta_gap} < 0 THEN 0
            WHEN {delta_gap} > {max_gap} THEN 0
            ELSE {delta_gap}
        END
    """
    motivo_diferencia = f"""
        CASE
            WHEN TRY_CONVERT(float, DistanciaMetrosInicial) IS NULL
              OR TRY_CONVERT(float, DistranciaMetrosFinal) IS NULL
                THEN 'LECTURA_DIARIA_INCOMPLETA'
            WHEN {delta_dia} < 0 THEN 'REINICIO_CONTADOR_DENTRO_DIA'
            WHEN {delta_dia} > {max_dia} THEN 'SALTO_DIARIO_ATIPICO'
            ELSE 'OK'
        END
    """
    motivo_gap = f"""
        CASE
            WHEN FinalAnterior IS NULL THEN 'SIN_LECTURA_ANTERIOR'
            WHEN COALESCE(OrigenAnterior, '') <> COALESCE(Origen, '')
                THEN 'CAMBIO_DE_FUENTE'
            WHEN TRY_CONVERT(float, DistanciaMetrosInicial) IS NULL
                THEN 'LECTURA_INICIAL_INCOMPLETA'
            WHEN {delta_gap} < 0 THEN 'REINICIO_CONTADOR_ENTRE_DIAS'
            WHEN {delta_gap} > {max_gap} THEN 'SALTO_ENTRE_DIAS_ATIPICO'
            ELSE 'OK'
        END
    """
    return {
        "diferencia": diferencia_corregida,
        "gap": km_no_corregidos,
        "motivo_diferencia": motivo_diferencia,
        "motivo_gap": motivo_gap,
    }


def cte_ordenado(origen: str) -> str:
    return f"""
        WITH Ordenado AS (
            SELECT
                s.*,
                LAG(s.DistranciaMetrosFinal) OVER (
                    PARTITION BY s.IdTracto
                    ORDER BY TRY_CONVERT(datetime2, s.FechaInicio),
                             TRY_CONVERT(datetime2, s.FechaFin), s.id
                ) AS FinalAnterior,
                LAG(s.FechaFin) OVER (
                    PARTITION BY s.IdTracto
                    ORDER BY TRY_CONVERT(datetime2, s.FechaInicio),
                             TRY_CONVERT(datetime2, s.FechaFin), s.id
                ) AS FechaFinalAnterior,
                LAG(s.Origen) OVER (
                    PARTITION BY s.IdTracto
                    ORDER BY TRY_CONVERT(datetime2, s.FechaInicio),
                             TRY_CONVERT(datetime2, s.FechaFin), s.id
                ) AS OrigenAnterior
            FROM {origen} s
        )
    """


def resumen_simulacion(
    cursor: pyodbc.Cursor,
    origen: TablaSql,
    max_km_dia: Decimal,
    max_km_gap: Decimal,
) -> dict[str, Any]:
    exp = expresiones(str(max_km_dia), str(max_km_gap))
    consulta = f"""
        {cte_ordenado(origen.nombre_completo)}
        SELECT
            COUNT(*) AS Registros,
            CAST(SUM(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0))
                AS decimal(20,2)) AS TotalOriginal,
            CAST(SUM(({exp['diferencia']}) + ({exp['gap']}))
                AS decimal(20,2)) AS TotalCorregido,
            SUM(CASE WHEN ({exp['motivo_diferencia']}) <> 'OK'
                     THEN 1 ELSE 0 END) AS DiferenciasCorregidas,
            SUM(CASE WHEN ({exp['motivo_gap']}) NOT IN ('OK','SIN_LECTURA_ANTERIOR')
                     THEN 1 ELSE 0 END) AS GapsCorregidos,
            SUM(CASE WHEN ({exp['motivo_gap']}) = 'CAMBIO_DE_FUENTE'
                     THEN 1 ELSE 0 END) AS CambiosFuente,
            SUM(CASE WHEN ({exp['motivo_gap']}) = 'SALTO_ENTRE_DIAS_ATIPICO'
                     THEN 1 ELSE 0 END) AS SaltosEntreDias,
            SUM(CASE WHEN ({exp['motivo_diferencia']}) = 'SALTO_DIARIO_ATIPICO'
                     THEN 1 ELSE 0 END) AS SaltosDiarios
        FROM Ordenado
    """
    fila = cursor.execute(consulta).fetchone()
    columnas = [descripcion[0] for descripcion in cursor.description]
    return dict(zip(columnas, fila))


def imprimir_resumen(resumen: dict[str, Any], titulo: str) -> None:
    print(f"\n{titulo}")
    print("-" * len(titulo))
    for clave, valor in resumen.items():
        print(f"{clave}: {valor}")


def crear_destino(
    cursor: pyodbc.Cursor,
    origen: TablaSql,
    destino: TablaSql,
    lote: str,
    max_km_dia: Decimal,
    max_km_gap: Decimal,
) -> None:
    exp = expresiones(str(max_km_dia), str(max_km_gap))
    cursor.execute(
        f"""
        SELECT TOP (0)
            s.*,
            CAST(NULL AS float) AS DiferenciaKmOriginal,
            CAST(NULL AS float) AS KmNoRegistradosOriginal,
            CAST(NULL AS float) AS DiferenciaKmCorregida,
            CAST(NULL AS float) AS KmNoRegistradosCorregidos,
            CAST(NULL AS float) AS KilometrosTotalesCorregidos,
            CAST(NULL AS varchar(40)) AS MotivoDiferencia,
            CAST(NULL AS varchar(40)) AS MotivoKmNoRegistrados,
            CAST(NULL AS varchar(12)) AS EstatusCorreccion,
            CAST(NULL AS datetime2) AS FechaMigracion,
            CAST(NULL AS uniqueidentifier) AS LoteMigracion
        INTO {destino.nombre_completo}
        FROM {origen.nombre_completo} s;
        """
    )

    columnas_originales = (
        "id, IdTracto, Unidad, DistanciaMetrosInicial, DistranciaMetrosFinal, "
        "DiferenciaKm, KmNoRegistrados, Origen, FechaInicio, FechaFin"
    )
    cursor.execute(
        f"""
        {cte_ordenado(origen.nombre_completo)}
        INSERT INTO {destino.nombre_completo} (
            {columnas_originales},
            DiferenciaKmOriginal, KmNoRegistradosOriginal,
            DiferenciaKmCorregida, KmNoRegistradosCorregidos,
            KilometrosTotalesCorregidos,
            MotivoDiferencia, MotivoKmNoRegistrados, EstatusCorreccion,
            FechaMigracion, LoteMigracion
        )
        SELECT
            {columnas_originales},
            TRY_CONVERT(float, DiferenciaKm),
            TRY_CONVERT(float, KmNoRegistrados),
            ({exp['diferencia']}),
            ({exp['gap']}),
            ({exp['diferencia']}) + ({exp['gap']}),
            ({exp['motivo_diferencia']}),
            ({exp['motivo_gap']}),
            CASE
                WHEN ({exp['motivo_diferencia']}) = 'OK'
                 AND ({exp['motivo_gap']}) IN ('OK','SIN_LECTURA_ANTERIOR')
                    THEN 'OK'
                ELSE 'CORREGIDO'
            END,
            SYSUTCDATETIME(),
            CONVERT(uniqueidentifier, ?)
        FROM Ordenado;
        """,
        lote,
    )

    nombre_indice = f"IX_{destino.tabla}_IdTracto_FechaInicio"
    cursor.execute(
        f"CREATE INDEX [{nombre_indice}] ON {destino.nombre_completo} "
        "(IdTracto, FechaInicio);"
    )


def validar_migracion(
    cursor: pyodbc.Cursor, origen: TablaSql, destino: TablaSql
) -> dict[str, Any]:
    fila = cursor.execute(
        f"""
        SELECT
            (SELECT COUNT(*) FROM {origen.nombre_completo}) AS FilasOrigen,
            (SELECT COUNT(*) FROM {destino.nombre_completo}) AS FilasDestino,
            (SELECT COUNT(*) FROM {destino.nombre_completo}
             WHERE KilometrosTotalesCorregidos IS NULL) AS TotalesNulos,
            (SELECT COUNT(*) FROM {destino.nombre_completo}
             WHERE EstatusCorreccion = 'CORREGIDO') AS FilasCorregidas,
            (SELECT CAST(SUM(KilometrosTotalesCorregidos) AS decimal(20,2))
             FROM {destino.nombre_completo}) AS TotalCorregido
        """
    ).fetchone()
    columnas = [descripcion[0] for descripcion in cursor.description]
    resultado = dict(zip(columnas, fila))
    if resultado["FilasOrigen"] != resultado["FilasDestino"]:
        raise RuntimeError("La cantidad de filas del origen y destino no coincide.")
    if resultado["TotalesNulos"] != 0:
        raise RuntimeError("La tabla destino contiene totales corregidos nulos.")
    return resultado


def main() -> None:
    args = argumentos()
    validar_argumentos(args)
    load_dotenv(REPO_ROOT / ".env")

    conexion = conectar_sql()
    conexion.autocommit = False
    try:
        cursor = conexion.cursor()
        origen = parsear_tabla(localizar_tabla_historica(conexion))
        destino = TablaSql(origen.base, origen.esquema, args.tabla_destino)

        if origen.nombre_completo.casefold() == destino.nombre_completo.casefold():
            raise SystemExit("La tabla destino no puede ser la tabla original.")

        print(f"Origen:  {origen.nombre_completo}")
        print(f"Destino: {destino.nombre_completo}")
        print(f"Máximo km/día: {args.max_km_dia}")
        print(f"Máximo km entre lecturas: {args.max_km_entre_lecturas}")

        simulacion = resumen_simulacion(
            cursor,
            origen,
            args.max_km_dia,
            args.max_km_entre_lecturas,
        )
        imprimir_resumen(simulacion, "SIMULACIÓN (sin modificar SQL)")

        if not args.aplicar:
            conexion.rollback()
            print("\nNo se realizaron cambios. Use --aplicar después de revisar el resumen.")
            return

        existe = tabla_existe(cursor, destino)
        if existe and not args.reemplazar_destino:
            raise RuntimeError(
                f"La tabla {destino.nombre_completo} ya existe. "
                "Use otro nombre o --reemplazar-destino."
            )
        if existe:
            print(f"Eliminando exclusivamente la tabla destino {destino.nombre_completo}...")
            cursor.execute(f"DROP TABLE {destino.nombre_completo};")

        lote = str(uuid.uuid4())
        print(f"Creando tabla corregida. Lote: {lote}")
        crear_destino(
            cursor,
            origen,
            destino,
            lote,
            args.max_km_dia,
            args.max_km_entre_lecturas,
        )
        validacion = validar_migracion(cursor, origen, destino)
        imprimir_resumen(validacion, "VALIDACIÓN DE LA MIGRACIÓN")
        conexion.commit()
        print("\nMigración confirmada. La tabla original no fue modificada.")
    except Exception:
        conexion.rollback()
        raise
    finally:
        conexion.close()


if __name__ == "__main__":
    main()
