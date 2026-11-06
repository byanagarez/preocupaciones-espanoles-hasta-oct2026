"""Consulta los CSV preparados: ranking, diferencias por edad y evolución.

Ejemplos:
    python comparar.py
    python comparar.py --tipo personal --problema "La sanidad" --mes 2026-10
"""
import argparse
from pathlib import Path

import pandas as pd

raiz = Path(__file__).resolve().parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--datos", type=Path, default=raiz / "datos_limpios")
parser.add_argument("--tipo", choices=["espana", "personal"], default="espana")
parser.add_argument("--problema", default="La vivienda")
parser.add_argument("--mes", help="AAAA-MM. Por defecto, el último mes disponible.")
args = parser.parse_args()

datos = pd.read_csv(args.datos / "preocupaciones_total.csv", sep=";", decimal=",",
                    encoding="utf-8-sig", dtype={"estudio": str})
mes = args.mes or datos.mes.max()
if mes not in set(datos.mes):
    parser.error(f"Mes sin datos: {mes}. Disponibles: {', '.join(sorted(datos.mes.unique()))}")
if args.problema not in set(datos.problema):
    parser.error("Problema no encontrado. Copia su nombre exacto desde preocupaciones_total.csv.")

general = datos[datos.dimension.eq("general") & datos.tipo_problema.eq(args.tipo)]
ranking = general[general.mes.eq(mes) & general.es_tema.eq(True)].nlargest(10, "porcentaje")
print(f"\n10 temas más mencionados | {mes} | {args.tipo}")
print(ranking[["problema", "porcentaje"]].round(2).to_string(index=False, na_rep="sin dato"))

edad = datos[datos.dimension.eq("Edad") & datos.tipo_problema.eq(args.tipo)
             & datos.mes.eq(mes) & datos.problema.eq(args.problema) & ~datos.grupo.eq("N.C.")]
orden_edad = ["De 18 a 24 años", "De 25 a 34 años", "De 35 a 44 años", "De 45 a 54 años",
             "De 55 a 64 años", "De 65 a 74 años", "75 y más años"]
edad = edad.set_index("grupo").reindex(orden_edad)
print(f"\n{args.problema} por edad | {mes} | {args.tipo}")
print(edad[["porcentaje", "base_ponderada"]].round(2).to_string(na_rep="sin dato"))

# Se reindexa el calendario solo para MOSTRAR los huecos. No se rellena con ceros.
evolucion = general[general.problema.eq(args.problema)].set_index("mes")[["porcentaje"]]
calendario = pd.period_range(general.mes.min(), general.mes.max(), freq="M").astype(str)
evolucion = evolucion.reindex(calendario)
evolucion.index.name = "mes"
print(f"\nEvolución nacional de {args.problema} | {args.tipo}")
print(evolucion.round(2).to_string(na_rep="sin dato"))
print("\nUnidad: porcentaje de personas que mencionan el tema (escala 0-100).")
print("Los cambios son descriptivos y no constituyen pruebas de significación estadística.")
