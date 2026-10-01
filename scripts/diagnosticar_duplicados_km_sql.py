"""Diagnostico de solo lectura para HistoricoKmSamsara."""

from __future__ import annotations

import sys
from pathlib import Path

from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).resolve().parent))
from comparar_kilometros_mensuales import conectar_sql, localizar_tabla_historica


INICIO = "2026-01-01T00:00:00"
FIN = "2026-09-02T00:00:00"


def imprimir(titulo: str, columnas: list[str], filas: list[tuple]) -> None:
    print(f"\n## {titulo}")
    print(" | ".join(columnas))
    for fila in filas:
        print(" | ".join("" if valor is None else str(valor) for valor in fila))


def main() -> None:
    load_dotenv(Path(__file__).resolve().parents[1] / ".env")
    with conectar_sql() as conexion:
        tabla = localizar_tabla_historica(conexion)
        cursor = conexion.cursor()
        print(f"Tabla: {tabla}")

        parametros = (INICIO, FIN)
        resumen = cursor.execute(
            f"""
            WITH b AS (
                SELECT *
                FROM {tabla}
                WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
                  AND TRY_CONVERT(datetime2, FechaInicio) < ?
            ), exactos AS (
                SELECT COUNT(*) AS Veces
                FROM b
                GROUP BY IdTracto, Unidad, DistanciaMetrosInicial,
                         DistranciaMetrosFinal, DiferenciaKm, KmNoRegistrados,
                         Origen, FechaInicio, FechaFin
                HAVING COUNT(*) > 1
            ), diarios AS (
                SELECT COUNT(*) AS Veces
                FROM b
                GROUP BY IdTracto, CONVERT(date, TRY_CONVERT(datetime2, FechaInicio))
                HAVING COUNT(*) > 1
            )
            SELECT
                (SELECT COUNT(*) FROM b) AS Registros,
                (SELECT COUNT(*) FROM exactos) AS GruposDuplicadosExactos,
                (SELECT COALESCE(SUM(Veces - 1), 0) FROM exactos) AS FilasExactasSobrantes,
                (SELECT COUNT(*) FROM diarios) AS DiasConMasDeUnRegistro,
                (SELECT COALESCE(SUM(Veces - 1), 0) FROM diarios) AS FilasDiariasSobrantes,
                SUM(CASE WHEN ABS(COALESCE(KmNoRegistrados, 0)) > 1000 THEN 1 ELSE 0 END)
                    AS FilasKmNoRegistradosMayor1000,
                CAST(SUM(COALESCE(DiferenciaKm, 0)) AS decimal(18,2)) AS SumaDiferenciaKm,
                CAST(SUM(COALESCE(KmNoRegistrados, 0)) AS decimal(18,2)) AS SumaKmNoRegistrados,
                CAST(SUM(COALESCE(DiferenciaKm, 0) + COALESCE(KmNoRegistrados, 0))
                    AS decimal(18,2)) AS TotalKm
            FROM b
            """,
            *parametros,
        ).fetchall()
        imprimir(
            "Resumen",
            [
                "Registros", "Grupos exactos", "Filas exactas sobrantes",
                "Días repetidos", "Filas diarias sobrantes", "KmNoReg >1000",
                "Suma DiferenciaKm", "Suma KmNoRegistrados", "Total km",
            ],
            resumen,
        )

        totales = cursor.execute(
            f"""
            WITH b AS (
                SELECT *,
                    ROW_NUMBER() OVER (
                        PARTITION BY IdTracto, Unidad, DistanciaMetrosInicial,
                                     DistranciaMetrosFinal, DiferenciaKm,
                                     KmNoRegistrados, Origen, FechaInicio, FechaFin
                        ORDER BY id
                    ) AS rn_exacto,
                    ROW_NUMBER() OVER (
                        PARTITION BY IdTracto,
                                     CONVERT(date, TRY_CONVERT(datetime2, FechaInicio))
                        ORDER BY TRY_CONVERT(datetime2, FechaFin) DESC, id DESC
                    ) AS rn_dia
                FROM {tabla}
                WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
                  AND TRY_CONVERT(datetime2, FechaInicio) < ?
            )
            SELECT 'Original' AS Escenario,
                   COUNT(*) Registros,
                   CAST(SUM(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0))
                        AS decimal(18,2)) TotalKm
            FROM b
            UNION ALL
            SELECT 'Sin duplicados exactos', COUNT(*),
                   CAST(SUM(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0))
                        AS decimal(18,2))
            FROM b WHERE rn_exacto = 1
            UNION ALL
            SELECT 'Un registro por unidad/dia (ultimo)', COUNT(*),
                   CAST(SUM(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0))
                        AS decimal(18,2))
            FROM b WHERE rn_dia = 1
            """,
            *parametros,
        ).fetchall()
        imprimir("Impacto de depuración", ["Escenario", "Registros", "Total km"], totales)

        impacto_atipicos = cursor.execute(
            f"""
            SELECT
                SUM(CASE WHEN ABS(COALESCE(KmNoRegistrados,0)) > 1000
                         THEN 1 ELSE 0 END) AS FilasAtipicas,
                CAST(SUM(CASE WHEN ABS(COALESCE(KmNoRegistrados,0)) > 1000
                              THEN COALESCE(KmNoRegistrados,0) ELSE 0 END)
                    AS decimal(18,2)) AS KmNoRegistradosAtipicos,
                CAST(SUM(CASE WHEN ABS(COALESCE(DiferenciaKm,0)) > 1000
                              THEN COALESCE(DiferenciaKm,0) ELSE 0 END)
                    AS decimal(18,2)) AS DiferenciaKmAtipica,
                CAST(SUM(CASE WHEN ABS(COALESCE(KmNoRegistrados,0)) <= 1000
                                   AND ABS(COALESCE(DiferenciaKm,0)) <= 1000
                              THEN COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0)
                              ELSE 0 END)
                    AS decimal(18,2)) AS TotalSinSaltosMayores1000
            FROM {tabla}
            WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
              AND TRY_CONVERT(datetime2, FechaInicio) < ?
            """,
            *parametros,
        ).fetchall()
        imprimir(
            "Impacto de saltos mayores a 1,000 km",
            ["Filas atípicas KmNoReg", "KmNoReg atípicos", "DiferenciaKm atípica", "Total sin saltos >1000"],
            impacto_atipicos,
        )

        nombres_repetidos = cursor.execute(
            f"""
            SELECT TOP (20)
                Unidad,
                CONVERT(char(10), TRY_CONVERT(datetime2, FechaInicio), 120) Dia,
                COUNT(*) Registros,
                COUNT(DISTINCT CAST(IdTracto AS varchar(100))) IdTractoDistintos,
                STRING_AGG(CAST(IdTracto AS varchar(max)), ', ') IdTractos
            FROM {tabla}
            WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
              AND TRY_CONVERT(datetime2, FechaInicio) < ?
            GROUP BY Unidad, CONVERT(char(10), TRY_CONVERT(datetime2, FechaInicio), 120)
            HAVING COUNT(*) > 1
            ORDER BY COUNT(*) DESC, Unidad, Dia
            """,
            *parametros,
        ).fetchall()
        imprimir(
            "Mismo nombre de unidad repetido en el día",
            ["Unidad", "Día", "Registros", "IdTracto distintos", "IdTractos"],
            nombres_repetidos,
        )

        repetidos = cursor.execute(
            f"""
            SELECT TOP (20)
                Unidad,
                CONVERT(char(10), TRY_CONVERT(datetime2, FechaInicio), 120) AS Dia,
                COUNT(*) AS Registros,
                COUNT(DISTINCT CONCAT(
                    COALESCE(CONVERT(varchar(50), DistanciaMetrosInicial),'NULL'),'|',
                    COALESCE(CONVERT(varchar(50), DistranciaMetrosFinal),'NULL'),'|',
                    COALESCE(CONVERT(varchar(50), DiferenciaKm),'NULL'),'|',
                    COALESCE(CONVERT(varchar(50), KmNoRegistrados),'NULL'),'|',
                    COALESCE(Origen,'NULL'),'|',
                    COALESCE(CONVERT(varchar(30), FechaInicio,126),'NULL'),'|',
                    COALESCE(CONVERT(varchar(30), FechaFin,126),'NULL')
                )) AS Variantes,
                CAST(SUM(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0))
                    AS decimal(18,2)) AS KmSumados
            FROM {tabla}
            WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
              AND TRY_CONVERT(datetime2, FechaInicio) < ?
            GROUP BY Unidad, CONVERT(char(10), TRY_CONVERT(datetime2, FechaInicio), 120)
            HAVING COUNT(*) > 1
            ORDER BY COUNT(*) DESC, KmSumados DESC
            """,
            *parametros,
        ).fetchall()
        imprimir(
            "Mayores repeticiones unidad/día",
            ["Unidad", "Día", "Registros", "Variantes", "Km sumados"],
            repetidos,
        )

        anomalias = cursor.execute(
            f"""
            SELECT TOP (25)
                Unidad,
                CONVERT(char(19), TRY_CONVERT(datetime2, FechaInicio), 120) FechaInicio,
                CONVERT(char(19), TRY_CONVERT(datetime2, FechaFin), 120) FechaFin,
                CAST(DiferenciaKm AS decimal(18,2)) DiferenciaKm,
                CAST(KmNoRegistrados AS decimal(18,2)) KmNoRegistrados,
                CAST(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0)
                    AS decimal(18,2)) TotalFila,
                Origen
            FROM {tabla}
            WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
              AND TRY_CONVERT(datetime2, FechaInicio) < ?
            ORDER BY ABS(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0)) DESC
            """,
            *parametros,
        ).fetchall()
        imprimir(
            "Filas con mayor kilometraje absoluto",
            ["Unidad", "Inicio", "Fin", "DiferenciaKm", "KmNoRegistrados", "Total fila", "Origen"],
            anomalias,
        )

        saltos = cursor.execute(
            f"""
            WITH ordenado AS (
                SELECT
                    IdTracto, Unidad, FechaInicio, FechaFin,
                    DistanciaMetrosInicial, DistranciaMetrosFinal,
                    DiferenciaKm, KmNoRegistrados, Origen,
                    LAG(DistranciaMetrosFinal) OVER (
                        PARTITION BY IdTracto
                        ORDER BY TRY_CONVERT(datetime2, FechaInicio), id
                    ) AS FinalAnterior,
                    LAG(FechaFin) OVER (
                        PARTITION BY IdTracto
                        ORDER BY TRY_CONVERT(datetime2, FechaInicio), id
                    ) AS FechaFinalAnterior,
                    LAG(Origen) OVER (
                        PARTITION BY IdTracto
                        ORDER BY TRY_CONVERT(datetime2, FechaInicio), id
                    ) AS OrigenAnterior
                FROM {tabla}
                WHERE TRY_CONVERT(datetime2, FechaInicio) >= '2025-12-01'
                  AND TRY_CONVERT(datetime2, FechaInicio) < ?
            )
            SELECT TOP (20)
                Unidad,
                CONVERT(char(19), TRY_CONVERT(datetime2, FechaInicio), 120),
                CONVERT(char(19), TRY_CONVERT(datetime2, FechaFinalAnterior), 120),
                CAST(DistanciaMetrosInicial / 1000.0 AS decimal(18,2)) AS InicialKm,
                CAST(FinalAnterior / 1000.0 AS decimal(18,2)) AS FinalAnteriorKm,
                CAST((DistanciaMetrosInicial-FinalAnterior)/1000.0 AS decimal(18,2)) AS SaltoCalculado,
                CAST(KmNoRegistrados AS decimal(18,2)) AS KmNoRegistrados,
                OrigenAnterior,
                Origen
            FROM ordenado
            WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
              AND ABS(COALESCE(KmNoRegistrados,0)) > 1000
            ORDER BY ABS(KmNoRegistrados) DESC
            """,
            FIN,
            INICIO,
        ).fetchall()
        imprimir(
            "Origen de los mayores KmNoRegistrados",
            [
                "Unidad", "Inicio actual", "Fin anterior", "Inicial actual km",
                "Final anterior km", "Salto calculado", "KmNoRegistrados",
                "Origen anterior", "Origen actual",
            ],
            saltos,
        )

        foco = cursor.execute(
            f"""
            SELECT Unidad,
                   CONVERT(char(7), TRY_CONVERT(datetime2, FechaInicio), 120) Mes,
                   COUNT(*) Registros,
                   CAST(SUM(COALESCE(DiferenciaKm,0)) AS decimal(18,2)) DiferenciaKm,
                   CAST(SUM(COALESCE(KmNoRegistrados,0)) AS decimal(18,2)) KmNoRegistrados,
                   CAST(SUM(COALESCE(DiferenciaKm,0)+COALESCE(KmNoRegistrados,0))
                       AS decimal(18,2)) Total
            FROM {tabla}
            WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
              AND TRY_CONVERT(datetime2, FechaInicio) < ?
              AND Unidad IN ('2282','DC06','1725','1726')
            GROUP BY Unidad, CONVERT(char(7), TRY_CONVERT(datetime2, FechaInicio), 120)
            ORDER BY Unidad, Mes
            """,
            *parametros,
        ).fetchall()
        imprimir(
            "Unidades con mayores diferencias",
            ["Unidad", "Mes", "Registros", "DiferenciaKm", "KmNoRegistrados", "Total"],
            foco,
        )


if __name__ == "__main__":
    main()
