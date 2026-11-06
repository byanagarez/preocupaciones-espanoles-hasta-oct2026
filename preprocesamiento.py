"""Prepara barómetros CIS tabulados; no necesita microdatos ni registro.

Uso (desde la carpeta del proyecto): python preprocesamiento.py
Dependencias: pandas y openpyxl. No modifica los Excel originales.
Los cruces son independientes: no permiten reconstruir individuos ni cruces conjuntos.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
import unicodedata
import warnings
from pathlib import Path

import openpyxl
import pandas as pd

MESES = dict(enumerate([
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
    "agosto", "septiembre", "octubre", "noviembre", "diciembre"
], 1))
ORDENES = {"primer problema": "primero", "segundo problema": "segundo",
           "tercer problema": "tercero", "total": "total"}
NO_SUSTANTIVAS = {"ninguno", "ninguna", "ninguno de ellos", "n.s.", "n.c."}
OMITIR = {"contenidos", "ficha tecnica", "tabla ponderaciones", "notas", "resultados"}


def identificar_dimension(cabecera):
    # En marzo, una pestaña del archivo absoluto se llama RV EG23 pero contiene
    # tamaño de municipio. La dimensión debe salir del título INTERIOR.
    reglas = [("sexo de", "Sexo"), ("edad de", "Edad"),
              ("estudios de", "Nivel de estudios"), ("tamano de municipio", "Tamaño de municipio"),
              ("recuerdo de voto", "RV EG23"), ("poblacion activa", "Ocupación"),
              ("identificacion subjetiva de clase", "Clase social"),
              ("autoubicacion ideologica", "Escala de ideología"), ("religiosidad", "Religiosidad")]
    matches = [nombre for texto, nombre in reglas if texto in cabecera]
    if len(matches) != 1:
        raise ValueError(f"Dimensión no reconocida inequívocamente: {cabecera}")
    return matches[0]


def normalizar(valor):
    """Normaliza solo para localizar etiquetas; no agrupa problemas distintos."""
    if valor is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


def numero(valor):
    if valor is None or str(valor).strip() in {"", "-", "–", "—", ".."}:
        return None
    if isinstance(valor, bool):
        raise ValueError("Se encontró un booleano en una celda numérica")
    if isinstance(valor, str):
        valor = valor.strip().replace("\u00a0", "").replace(",", ".")
    resultado = float(valor)
    if not math.isfinite(resultado):
        raise ValueError(f"Valor numérico no finito: {valor}")
    return resultado


def descubrir(carpeta):
    """Tolera sufijos de descarga (1); duplica solo si los bytes son idénticos."""
    encontrados, duplicados = {}, []
    patron = re.compile(r"^(\d+)-multi(_abs)?(?:\s*\(\d+\))?\.xlsx$", re.I)
    for ruta in sorted(carpeta.glob("*.xlsx")):
        m = patron.match(ruta.name)
        if not m:
            continue
        clave = (m[1], "absolutos" if m[2] else "porcentajes")
        digest = hashlib.sha256(ruta.read_bytes()).hexdigest()
        if clave in encontrados:
            anterior, anterior_hash = encontrados[clave]
            if anterior_hash != digest:
                raise ValueError(f"Copias diferentes del mismo estudio: {anterior.name} y {ruta.name}")
            duplicados.append(ruta.name)
        else:
            encontrados[clave] = (ruta, digest)
    if not encontrados:
        raise ValueError(f"No se encontraron archivos *-multi.xlsx en {carpeta}")
    for estudio, _ in encontrados:
        for tipo in ("porcentajes", "absolutos"):
            if (estudio, tipo) not in encontrados:
                raise ValueError(f"Falta el archivo de {tipo} del estudio {estudio}")
    return encontrados, duplicados


def leer(ruta, estudio, formato):
    """Lee todos los cruces publicados y conserva referencias de hoja y fila."""
    registros, inventario, notas = [], [], []
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Workbook contains no default style.*")
        libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        referencia = None
        for hoja in libro:
            if normalizar(hoja.title) == "notas":
                notas = [str(x) for r in hoja.values for x in r if x is not None]
            if normalizar(hoja.title) in OMITIR:
                continue
            filas = list(hoja.values)
            cabecera = " ".join(normalizar(x) for r in filas[:5] for x in r if x is not None)
            # Algunos estudios añaden notas/estimaciones de voto: no son cruces.
            if "cruce por" not in cabecera:
                continue
            dimension_real = identificar_dimension(cabecera)
            fecha = re.search(r"barometro de (\w+) (\d{4})", cabecera)
            codigo = re.search(r"\b(\d+)/\d+\b", cabecera)
            if not fecha or not codigo or codigo[1] != estudio:
                raise ValueError(f"{ruta.name}/{hoja.title}: fecha o estudio no reconocidos")
            mes_num = next((k for k, v in MESES.items() if v == fecha[1]), None)
            if mes_num is None:
                raise ValueError(f"Mes desconocido: {fecha[1]}")
            mes = f"{fecha[2]}-{mes_num:02d}"
            if referencia is not None and referencia != mes:
                raise ValueError(f"Meses diferentes dentro de {ruta.name}")
            referencia = mes
            marca = "porcentajes verticales" if formato == "porcentajes" else "valores absolutos"
            if marca not in cabecera:
                raise ValueError(f"{ruta.name}: el contenido no corresponde a {formato}")
            etiquetas = [normalizar(r[0]) for r in filas]
            comienzos = [i for i, t in enumerate(etiquetas) if t.startswith("pregunta ")]
            detectadas = set()
            for pos, inicio in enumerate(comienzos):
                fin = comienzos[pos + 1] if pos + 1 < len(comienzos) else len(filas)
                enunciado = " ".join(etiquetas[inicio + 1:min(inicio + 4, fin)])
                if "principal problema que existe actualmente en espana" in enunciado:
                    tipo = "espana"
                elif "personalmente, le afecta mas" in enunciado:
                    tipo = "personal"
                else:
                    continue
                if tipo in detectadas:
                    raise ValueError(f"Dos preguntas candidatas para {tipo} en {ruta.name}/{hoja.title}")
                detectadas.add(tipo)
                ordenes_vistos = set()
                for h in range(inicio + 1, fin - 1):
                    orden = ORDENES.get(etiquetas[h])
                    encabezado = filas[h + 1]
                    if orden is None or len(encabezado) < 2 or normalizar(encabezado[1]) != "total":
                        continue
                    if orden in ordenes_vistos:
                        raise ValueError(f"Bloque {orden} repetido: {ruta.name}/{hoja.title}/{tipo}")
                    ordenes_vistos.add(orden)
                    bases = [n for n in range(h + 2, fin) if etiquetas[n] == "(n)"]
                    if not bases:
                        raise ValueError(f"Falta base (N): {ruta.name}/{hoja.title}/{tipo}/{orden}")
                    n = bases[0]
                    grupos = [(j, str(g).strip()) for j, g in enumerate(encabezado) if j > 0 and g is not None and str(g).strip()]
                    if len({g for _, g in grupos}) != len(grupos):
                        raise ValueError(f"Grupos duplicados en {ruta.name}/{hoja.title}")
                    problemas = set()
                    for i in range(h + 2, n):
                        problema = filas[i][0]
                        if problema is None or not str(problema).strip():
                            continue
                        problema = str(problema).strip()
                        # Conservamos literalmente los nombres; cambios entre meses se auditan.
                        if problema in problemas:
                            raise ValueError(f"Problema repetido: {problema}")
                        problemas.add(problema)
                        for j, grupo in grupos:
                            valor = numero(filas[i][j])
                            base = numero(filas[n][j])
                            if base is None or base < 0:
                                raise ValueError(f"Base inválida: {ruta.name}/{hoja.title}/{n+1}")
                            if valor is not None and (valor < 0 or (formato == "porcentajes" and valor > 100.000001)):
                                raise ValueError(f"Valor fuera de rango en {ruta.name}/{hoja.title}/{i+1}")
                            dimension = "general" if grupo == "TOTAL" else dimension_real
                            # El total nacional aparece en cada cruce: lo tomamos una sola vez de Edad.
                            if grupo == "TOTAL" and dimension_real != "Edad":
                                continue
                            registros.append(dict(estudio=estudio, mes=mes, tipo_problema=tipo,
                                pregunta=str(filas[inicio][0]), orden_mencion=orden,
                                dimension=dimension, grupo=grupo, problema=problema,
                                es_tema=normalizar(problema) not in NO_SUSTANTIVAS,
                                valor_publicado=valor, base_publicada=base,
                                archivo=ruta.name, hoja=hoja.title, fila_excel=i+1))
                    if orden == "total":
                        inventario.append(dict(estudio=estudio, mes=mes, dimension=dimension_real, hoja=hoja.title,
                            tipo_problema=tipo, pregunta=str(filas[inicio][0]),
                            categorias=len(problemas), grupos=len(grupos)-1))
                if ordenes_vistos != set(ORDENES.values()):
                    raise ValueError(f"Bloques incompletos en {ruta.name}/{hoja.title}/{tipo}: {ordenes_vistos}")
            if detectadas != {"espana", "personal"}:
                raise ValueError(f"Faltan preguntas de preocupaciones en {ruta.name}/{hoja.title}")
        if referencia is None:
            raise ValueError(f"No hay cruces reconocidos en {ruta.name}")
    finally:
        libro.close()
    return pd.DataFrame(registros), inventario, notas


def ejecutar(entrada, salida):
    archivos, duplicados = descubrir(entrada)
    porcentajes, absolutos, inventario, fuentes, notas_fuente = [], [], [], [], []
    for (estudio, formato), (ruta, digest) in sorted(archivos.items()):
        df, inv, notas = leer(ruta, estudio, formato)
        (porcentajes if formato == "porcentajes" else absolutos).append(df)
        if formato == "porcentajes":
            inventario.extend(inv)
            notas_fuente.append(dict(estudio=estudio, mes=df.mes.iloc[0], notas=notas))
            print(f"OK {estudio}: {df.mes.iloc[0]} | {df.dimension.nunique()} dimensiones")
        fuentes.append(dict(archivo=ruta.name, sha256=digest, formato=formato, estudio=estudio,
                            mes=df.mes.iloc[0], bytes=ruta.stat().st_size))
    p, a = pd.concat(porcentajes, ignore_index=True), pd.concat(absolutos, ignore_index=True)
    claves = ["estudio", "mes", "tipo_problema", "orden_mencion", "dimension", "grupo", "problema"]
    if p.duplicated(claves).any() or a.duplicated(claves).any():
        raise ValueError("Claves repetidas en los datos extraídos")
    p = p.rename(columns={"valor_publicado": "porcentaje_publicado", "base_publicada": "base_ponderada"})
    a = a[claves + ["valor_publicado", "base_publicada", "archivo", "hoja", "fila_excel"]].rename(columns={
        "valor_publicado": "absoluto_publicado", "base_publicada": "base_absoluta_publicada",
        "archivo": "archivo_absolutos", "hoja": "hoja_absolutos", "fila_excel": "fila_excel_absolutos"})
    datos = p.merge(a, on=claves, how="outer", validate="one_to_one", indicator=True)
    if not datos["_merge"].eq("both").all():
        raise ValueError("Porcentajes y absolutos no contienen las mismas categorías/grupos")
    datos = datos.drop(columns="_merge")
    datos["porcentaje"] = datos.porcentaje_publicado.where(datos.base_ponderada > 0)
    datos["estado"] = "publicado"
    datos.loc[datos.porcentaje_publicado.isna(), "estado"] = "sin_valor_publicado"
    datos.loc[datos.base_ponderada.eq(0), "estado"] = "sin_base"
    datos = datos.sort_values(claves).reset_index(drop=True)
    total = datos[datos.orden_mencion.eq("total")].copy()
    edad = total[total.dimension.isin(["Edad", "general"])].copy()
    # Diferencias en puntos porcentuales SOLO entre meses consecutivos presentes.
    comparacion = total.sort_values("mes").copy()
    agrupacion = ["tipo_problema", "dimension", "grupo", "problema"]
    g = comparacion.groupby(agrupacion, sort=False)
    comparacion["mes_anterior_disponible"] = g["mes"].shift()
    comparacion["porcentaje_anterior"] = g["porcentaje"].shift()
    actual = pd.to_datetime(comparacion.mes + "-01")
    anterior = pd.to_datetime(comparacion.mes_anterior_disponible + "-01")
    comparacion["meses_entre_observaciones"] = (actual.dt.year-anterior.dt.year)*12 + actual.dt.month-anterior.dt.month
    comparacion["diferencia_pp"] = comparacion.porcentaje - comparacion.porcentaje_anterior
    comparacion["es_cambio_mensual"] = comparacion.meses_entre_observaciones.eq(1)
    # Auditoría de cobertura: una categoría ausente sigue siendo ausente, nunca cero.
    meses = sorted(total.mes.unique())
    cobertura = total.groupby(["tipo_problema", "dimension", "grupo", "problema"], as_index=False).agg(
        meses_presentes=("mes", lambda x: ", ".join(sorted(set(x)))), numero_meses=("mes", "nunique"))
    cobertura["presente_todos_meses"] = cobertura.numero_meses.eq(len(meses))
    ranking = total[total.es_tema & total.porcentaje.notna()].copy()
    ranking["puesto"] = ranking.groupby(["mes", "tipo_problema", "dimension", "grupo"])["porcentaje"].rank(method="min", ascending=False)
    ranking = ranking.sort_values(["mes", "tipo_problema", "dimension", "grupo", "puesto", "problema"])
    # Solo se exporta una vez superadas las comprobaciones de TODOS los estudios.
    salida.mkdir(parents=True, exist_ok=True)
    productos = {"preocupaciones_total.csv": total, "preocupaciones_edad.csv": edad,
        "detalle_orden_menciones.csv": datos, "comparacion_temporal.csv": comparacion,
        "ranking.csv": ranking, "cobertura_categorias.csv": cobertura,
        "inventario_estudios.csv": pd.DataFrame(inventario), "fuentes.csv": pd.DataFrame(fuentes)}
    for nombre, df in productos.items():
        df.to_csv(salida / nombre, sep=";", index=False, encoding="utf-8-sig", decimal=",")
    # JSON web en escala 0-100, con null para bases vacías; no contiene NaN.
    web_cols = ["estudio", "mes", "tipo_problema", "dimension", "grupo", "problema", "es_tema", "porcentaje", "base_ponderada", "estado"]
    payload = dict(fuente="CIS, tabulaciones verticales", escala_porcentaje="0-100",
        meses=meses, datos=json.loads(total[web_cols].to_json(orient="records", force_ascii=False, double_precision=10)))
    (salida / "preocupaciones_web.json").write_text(json.dumps(payload, ensure_ascii=False, allow_nan=False), encoding="utf-8")
    informe = dict(meses=meses, estudios=sorted(total.estudio.unique()),
        filas_total=len(total), filas_edad=len(edad), filas_detalle=len(datos),
        duplicados_identicos_ignorados=duplicados,
        meses_sin_archivo_por_ano={str(y): [f"{y}-{m:02d}" for m in range(1, max(int(x[5:]) for x in meses if x.startswith(str(y)))+1) if f"{y}-{m:02d}" not in meses] for y in sorted({int(x[:4]) for x in meses})},
        advertencias=["Los absolutos publicados están redondeados: no son microdatos ni bases brutas garantizadas.",
          "Los porcentajes ya están ponderados. No aplicar otra vez la tabla de ponderaciones.",
          "Una categoría ausente no equivale a cero. Revisar cobertura_categorias.csv.",
          "Las diferencias entre meses son descriptivas: no se calculan significaciones ni intervalos de confianza.",
          "No se promedian meses ni se combinan tramos de edad automáticamente.",
          "Los distintos cruces son marginales independientes: no cruzar Edad con Sexo usando estas tablas."])
    (salida / "control_calidad.json").write_text(json.dumps(informe, ensure_ascii=False, indent=2), encoding="utf-8")
    (salida / "notas_fuente.json").write_text(json.dumps(notas_fuente, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nCompletado: {len(meses)} meses, {len(total):,} registros TOTAL, {len(edad):,} de edad/general.")
    print("Meses:", ", ".join(meses))
    print("Salida:", salida)
    print("Revisa control_calidad.json y cobertura_categorias.csv antes de interpretar cambios.")


if __name__ == "__main__":
    raiz = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--entrada", type=Path, default=raiz / "datos_originales")
    parser.add_argument("--salida", type=Path, default=raiz / "datos_limpios")
    args = parser.parse_args()
    try:
        ejecutar(args.entrada, args.salida)
    except (ValueError, KeyError, IndexError, OSError) as exc:
        print(f"\nERROR: {exc}\nLa ejecución no ha finalizado. Corrige el error y repítela antes de utilizar sus resultados.", file=sys.stderr)
        sys.exit(1)
