"""
Entrega 7 - Figuras y tablas corregidas
"Cafeterías de especialidad, valor del suelo y valorización en barrios de Bogotá"
Haciendo Economía 2026-2, Universidad del Rosario.

Este script reemplaza al de la Entrega 5 (ENTREGA 5/Código/figuras_entrega5.py) e incorpora
la retroalimentación de las Entregas 5 y 6. Hace todo desde los archivos del repositorio:
  1. Lee la base RAW de la Entrega 4 (68 registros) y aplica la regla de depuración
     documentada en el informe -> 57 encuestas válidas.
  2. Lee el inventario de 25 cafeterías (Entrega 3a).
  3. Escribe las tablas de apoyo (depuración, cobertura por local, frecuencias,
     resumen por barrio, percepción por perfil y por nivel del suelo).
  4. Exporta los 25 locales como GeoJSON.
  5. Genera las 8 figuras del informe, incluido el MAPA (figura 1).

Uso, desde la raíz del repositorio:
    python "ENTREGA 7/Código/figuras_entrega7.py"

Requisitos: Python 3.9+, pandas, numpy, openpyxl, matplotlib (ver requirements.txt).
"""
from pathlib import Path
import json

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize
from matplotlib.lines import Line2D

# --------------------------------------------------------------------------- rutas
RAIZ = Path(__file__).resolve().parents[2]
RAW_XLSX = RAIZ / "ENTREGA 4" / "Base_Datos_RAW_Entrega4.xlsx"
HOJA_RAW = "1_Datos_RAW_Completos_68"
INVENTARIO_CSV = RAIZ / "ENTREGA 3" / "base_secundaria_cafeterias_3a-v3.csv"

SALIDA = RAIZ / "ENTREGA 7"
FIG = SALIDA / "Gráficos"
TAB = SALIDA / "Tablas"
DAT = SALIDA / "Datos"
for p in (FIG, TAB, DAT):
    p.mkdir(parents=True, exist_ok=True)

ESTADOS_VALIDOS = ["En muestra (14 locales)", "Adicional del marco (25 locales)"]

# --------------------------------------------------------------------------- estilo
AZUL, NARANJA, AQUA, VIOLETA = "#2a78d6", "#eb6834", "#1baf7a", "#8b5cf6"
GRIS_NOSE = "#b9b8b2"          # «no sabe / no puede comparar / no puede opinar»: siempre gris
TINTA, TINTA2, TINTA3 = "#0b0b0b", "#52514e", "#8a8984"
REJILLA = "#e6e5e0"
RAMPA_AZUL = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]
CMAP_SUELO = LinearSegmentedColormap.from_list("suelo", RAMPA_AZUL)
DIV_ROJO, DIV_ROJO_C = "#c0392b", "#e88a80"
DIV_AZUL, DIV_AZUL_C = "#1c5cab", "#86b6ef"
NEUTRO = "#dedcd6"

plt.rcParams.update({
    "font.family": ["Inter", "DejaVu Sans"], "font.size": 10,
    "axes.edgecolor": REJILLA, "axes.labelcolor": TINTA2,
    "axes.titlesize": 13, "axes.titleweight": "semibold", "axes.titlecolor": TINTA,
    "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
    "xtick.color": TINTA2, "ytick.color": TINTA2, "axes.grid": False,
    "figure.facecolor": "white", "savefig.dpi": 200, "savefig.bbox": "tight",
})

FUENTE_ENC = ("Fuente: encuesta de percepción. 68 envíos el 8 y 9 de nov. de 2026; "
              "las 57 respuestas válidas son del 8 de nov.")
FUENTE_SUELO = ("Valor del suelo: valor de referencia por barrio registrado en el inventario de la Entrega 3a "
                "(UAECD, 2016 y 2025).")
PERCEPCION = "Son percepciones declaradas por los encuestados, no mediciones de precios, arriendos o mudanzas."


def nota(fig, texto, y=-0.02):
    fig.text(0.01, y, texto, ha="left", va="top", fontsize=8, color=TINTA3, wrap=True)


def guardar(fig, nombre):
    fig.savefig(FIG / nombre)
    plt.close(fig)
    print(f"  figura -> {nombre}")


def corto(nombre):
    return (nombre.replace(" (Coffee Master 2026)", "").replace(" Coffee Roasters", "")
            .replace("Librería Café", "Librería").replace(" & Panadería Artesanal", "")
            .replace("Café Amor Perfecto", "Amor Perfecto").replace(" - Sede ", " - "))


# =========================================================================== 1. datos
print("1) Base RAW y depuración")
raw = pd.read_excel(RAW_XLSX, sheet_name=HOJA_RAW)
raw.columns = [c.strip() for c in raw.columns]
ren = {c: f"p{c[0]}" for c in raw.columns if c[:2] in {f"{i}." for i in range(1, 10)}}
ren["¿Acepta participar voluntariamente?"] = "consentimiento"
raw = raw.rename(columns=ren)
assert len(raw) == 68, "La base RAW debería tener 68 registros"


def categoria_depuracion(e):
    if e == ESTADOS_VALIDOS[0]:
        return "Válido: local de la muestra principal (14 locales)"
    if e == ESTADOS_VALIDOS[1]:
        return "Válido: local adicional del inventario de 25"
    if str(e).startswith("N/A"):
        return "Excluido: ruta pegada por error (no identifica un local)"
    return "Excluido: fuera de las 7 zonas o prueba inicial"


raw["depuracion"] = raw["estado_registro_raw"].map(categoria_depuracion)
raw["fecha"] = pd.to_datetime(raw["Marca temporal"]).dt.strftime("%d-%b")
enc = raw[raw["estado_registro_raw"].isin(ESTADOS_VALIDOS) & raw["latitud"].notna()].copy()
assert len(enc) == 57, f"Se esperaban 57 válidos y hay {len(enc)}"
assert enc["consentimiento"].str.startswith("Sí").all()

cols_pub = ["id_encuesta", "Marca temporal", "consentimiento", "estado_registro_raw", "id_cafe",
            "nombre_establecimiento_verificado", "zona_estudio", "barrio_verificado_catastro",
            "localidad", "cod_barrio", "latitud", "longitud", "valor_suelo_m2_2025_cop",
            "var_valor_suelo_16_25_pct", "precio_filtrado_ref_cop", "rating_gmaps"] + \
           [f"p{i}" for i in range(1, 10)]
enc[cols_pub].to_csv(DAT / "encuesta_limpia_57.csv", index=False, encoding="utf-8-sig")

inv = pd.read_csv(INVENTARIO_CSV, encoding="utf-8-sig")
assert len(inv) == 25
inv = inv.merge(enc.groupby("id_cafe").size().rename("n_encuestas"), on="id_cafe", how="left")
inv["n_encuestas"] = inv["n_encuestas"].fillna(0).astype(int)
plan14 = set(enc.loc[enc.estado_registro_raw == ESTADOS_VALIDOS[0], "id_cafe"])
inv["grupo_campo"] = np.select([inv["id_cafe"].isin(plan14), inv["n_encuestas"] > 0],
                               ["Plan (14 locales)", "Adicional con encuestas"], "Sin encuestas")
inv["nombre_corto"] = inv["nombre_establecimiento"].map(corto)
# BAR_0205 aparece con dos nombres y dos valores del suelo (8,90 y 8,75 M): se trata como
# un solo barrio, «El Nogal», con el promedio de ambos valores (ver README).
inv["barrio_std"] = np.where(inv["cod_barrio"] == "BAR_0205", "El Nogal", inv["barrio"])

# ----------------------------------------------------------------- etiquetas cortas
P = {
    "p1": {"Vivo aquí.": "Vive aquí", "Trabajo o estudio aquí.": "Trabaja o estudia aquí",
           "No vivo ni trabajo aquí, pero vengo con frecuencia.": "Visita con frecuencia",
           "Vengo de vez en cuando o es la primera vez que vengo.": "Visita ocasional / primera vez"},
    "p2": {"Sobre todo cafés, restaurantes modernos o tiendas de diseño.": "Cafés, restaurantes modernos o diseño",
           "Sobre todo negocios de siempre (tiendas de barrio, panaderías, talleres).": "Negocios de siempre",
           "Casi no he visto negocios nuevos.": "Casi ningún negocio nuevo",
           "No conozco el sector lo suficiente para saberlo.": "No conoce el sector"},
    "p5": {"Están mucho más caros.": "Mucho más caros", "Siguen más o menos igual.": "Más o menos igual",
           "Están más baratos.": "Más baratos",
           "No conocía el sector hace 3 años, no puedo comparar.": "No puede comparar"},
    "p6": {"Han subido mucho.": "Subieron mucho", "Siguen más o menos igual.": "Más o menos igual",
           "Han bajado.": "Bajaron", "No sé, no conozco lo suficiente el sector.": "No sabe"},
    "p7": {"Sí, muchos.": "Sí, muchos", "Sí, algunos.": "Sí, algunos",
           "No, casi todos siguen aquí.": "No, casi todos siguen",
           "No sé, no conozco lo suficiente el sector.": "No sabe"},
    "p9": {"Ha mejorado mucho.": "Mejoró mucho", "Ha mejorado un poco.": "Mejoró un poco",
           "Sigue igual, no veo diferencia.": "Sigue igual", "Ha empeorado un poco.": "Empeoró un poco",
           "Ha empeorado mucho.": "Empeoró mucho", "No lo conocía antes, no puedo opinar.": "No puede opinar"},
}
for q, m in P.items():
    enc[q + "_c"] = enc[q].map(m)
    assert enc[q + "_c"].notna().all(), f"Categoría no mapeada en {q}"

enc["perfil"] = np.where(enc["p1_c"].isin(["Vive aquí", "Trabaja o estudia aquí"]),
                         "Vive, trabaja o estudia aquí", "Visitante (frecuente u ocasional)")
CORTES = [0, 5.5e6, 7e6, np.inf]
NIVELES = ["menos de 5,5 M", "de 5,5 a 7 M", "más de 7 M"]
enc["nivel_suelo"] = pd.cut(enc["valor_suelo_m2_2025_cop"], CORTES, labels=NIVELES, right=False)

# Indicadores: (columna, valores que cuentan como "sí", valores que se excluyen de la base)
IND = {
    "P5 · Precios mucho más caros": ("p5_c", ["Mucho más caros"], ["No puede comparar"]),
    "P6 · Arriendos subieron mucho": ("p6_c", ["Subieron mucho"], ["No sabe"]),
    "P7 · Vio irse vecinos (muchos o algunos)": ("p7_c", ["Sí, muchos", "Sí, algunos"], ["No sabe"]),
    "P9 · El sector mejoró (mucho o un poco)": ("p9_c", ["Mejoró mucho", "Mejoró un poco"], ["No puede opinar"]),
}


def conteo(sub, col, si, fuera):
    base = sub[~sub[col].isin(fuera)]
    return int(base[col].isin(si).sum()), len(base), int(sub[col].isin(fuera).sum())


# =========================================================================== 2. tablas
print("2) Tablas de apoyo")
dep = (raw.groupby("depuracion").agg(n=("id_encuesta", "size"),
                                     fechas=("fecha", lambda s: ", ".join(f"{k}: {v}" for k, v in s.value_counts().items())))
       .reindex(["Válido: local de la muestra principal (14 locales)", "Válido: local adicional del inventario de 25",
                 "Excluido: fuera de las 7 zonas o prueba inicial",
                 "Excluido: ruta pegada por error (no identifica un local)"]).reset_index())
dep.loc[len(dep)] = ["Total recibido", int(dep["n"].sum()), ""]
dep.to_csv(TAB / "tabla_depuracion_68_a_57.csv", index=False, encoding="utf-8-sig")

inv.sort_values(["grupo_campo", "n_encuestas"], ascending=[True, False])[
    ["id_cafe", "nombre_establecimiento", "zona_estudio", "barrio", "cod_barrio", "grupo_campo", "n_encuestas"]
].to_csv(TAB / "tabla_cobertura_por_local.csv", index=False, encoding="utf-8-sig")

filas = []
for i in range(1, 10):
    for cat, n in enc[f"p{i}"].value_counts().items():
        filas.append({"pregunta": f"P{i}", "categoria": cat, "n": n, "pct_sobre_57": round(100 * n / 57, 1)})
pd.DataFrame(filas).to_csv(TAB / "frecuencias_p1_p9_n57.csv", index=False, encoding="utf-8-sig")

barrio = (inv.groupby("cod_barrio", as_index=False)
          .agg(barrio=("barrio_std", "first"), zona_estudio=("zona_estudio", "first"),
               cafes_inventario=("id_cafe", "size"),
               valor_suelo_m2_2025_cop=("valor_suelo_m2_2025_cop", "mean"),
               var_valor_suelo_16_25_pct=("var_valor_suelo_16_25_pct", "mean"),
               precio_filtrado_medio=("precio_filtrado_ref_cop", "mean"),
               encuestas_validas=("n_encuestas", "sum")))
barrio.round(1).to_csv(TAB / "resumen_por_barrio.csv", index=False, encoding="utf-8-sig")


def tabla_grupos(col_grupo, orden):
    out = []
    for g in orden:
        sub = enc[enc[col_grupo] == g]
        fila = {"grupo": g, "respuestas": len(sub)}
        for k, (c, si, fuera) in IND.items():
            a, b, nose = conteo(sub, c, si, fuera)
            fila[k] = f"{a} de {b}"
            fila[k + " · no puede responder"] = nose
        out.append(fila)
    return pd.DataFrame(out)


tabla_grupos("perfil", ["Vive, trabaja o estudia aquí", "Visitante (frecuente u ocasional)"]) \
    .to_csv(TAB / "percepcion_por_perfil.csv", index=False, encoding="utf-8-sig")
tabla_grupos("nivel_suelo", NIVELES).to_csv(TAB / "percepcion_por_nivel_suelo.csv", index=False, encoding="utf-8-sig")

# =========================================================================== 3. GeoJSON
geo = {"type": "FeatureCollection", "name": "cafeterias_inventario_25",
       "crs": {"type": "name", "properties": {"name": "urn:ogc:def:crs:OGC:1.3:CRS84"}}, "features": []}
for _, r in inv.iterrows():
    props = {k: (v.item() if hasattr(v, "item") else v) for k, v in r.items() if k not in ("latitud", "longitud")}
    geo["features"].append({"type": "Feature", "properties": props,
                            "geometry": {"type": "Point", "coordinates": [float(r["longitud"]), float(r["latitud"])]}})
(DAT / "cafeterias_inventario_25.geojson").write_text(json.dumps(geo, ensure_ascii=False, indent=1), encoding="utf-8")

# =========================================================================== 4. figuras
print("3) Figuras")

# ---- Figura 1: MAPA del inventario
fig, ax = plt.subplots(figsize=(9.5, 7.6))
norm = Normalize(inv["valor_suelo_m2_2025_cop"].min() / 1e6, inv["valor_suelo_m2_2025_cop"].max() / 1e6)
con, sin = inv[inv.n_encuestas > 0], inv[inv.n_encuestas == 0]
sc = ax.scatter(con["longitud"], con["latitud"], s=60 + 45 * con["n_encuestas"],
                c=con["valor_suelo_m2_2025_cop"] / 1e6, cmap=CMAP_SUELO, norm=norm,
                edgecolor="white", linewidth=1.5, zorder=3)
ax.scatter(sin["longitud"], sin["latitud"], s=60, facecolor="white",
           edgecolor=sin["valor_suelo_m2_2025_cop"].map(lambda v: CMAP_SUELO(norm(v / 1e6))), linewidth=2, zorder=3)
zc = inv.groupby("zona_estudio")[["longitud", "latitud"]].mean()
pos_txt = {"Usaquén": (-74.018, 4.703), "El Refugio y borde oriental": (-74.036, 4.682),
           "Chapinero central y Zona Rosa": (-74.036, 4.667), "Chapinero Alto y Pardo Rubio": (-74.036, 4.632),
           "Quinta Camacho y Teusaquillo": (-74.112, 4.672), "Salitre - Nicolás de Federmán": (-74.118, 4.628),
           "La Macarena": (-74.040, 4.606)}
for z, (tx, ty) in pos_txt.items():
    mz = inv[inv.zona_estudio == z]
    k = ((mz.longitud - zc.loc[z, "longitud"]) ** 2 + (mz.latitud - zc.loc[z, "latitud"]) ** 2).idxmin()
    ax.annotate(z, xy=(mz.loc[k, "longitud"], mz.loc[k, "latitud"]), xytext=(tx, ty), fontsize=8.5, color=TINTA2,
                ha="left", va="center", arrowprops=dict(arrowstyle="-", color=TINTA3, lw=0.7, shrinkA=2, shrinkB=4))
ax.set_aspect(1 / np.cos(np.deg2rad(4.65)))
ax.set_xlim(-74.135, -73.995)
ax.set_ylim(4.598, 4.712)
ax.set_xlabel("Longitud")
ax.set_ylabel("Latitud")
ax.grid(color=REJILLA, linewidth=0.6)
cb = fig.colorbar(sc, ax=ax, shrink=0.55, pad=0.02)
cb.set_label("Valor de referencia del suelo del barrio, 2025 (millones COP/m²)", color=TINTA2)
cb.outline.set_visible(False)
leg = [Line2D([], [], marker="o", ls="", ms=np.sqrt(60 + 45 * k), mfc="#9ec5f4", mec="white",
              label=f"{k} encuesta{'s' if k > 1 else ''}") for k in (1, 5, 9)]
leg.append(Line2D([], [], marker="o", ls="", ms=np.sqrt(60), mfc="white", mec="#6da7ec", mew=2, label="Sin encuestas"))
ax.legend(handles=leg, loc="upper left", frameon=False, fontsize=8.5, labelspacing=1.2,
          title="Tamaño = encuestas válidas", title_fontsize=8.5, borderpad=1.0)
ax.set_title("Los 25 cafés del inventario, coloreados según el valor del suelo de su barrio")
nota(fig, "Cada punto es un café; su color es el valor de referencia del suelo del barrio donde está (no el del predio "
          "del café).\nEl inventario son 25 cafés de especialidad seleccionados en 7 zonas: no es un censo de las "
          "cafeterías de Bogotá. " + FUENTE_SUELO)
guardar(fig, "fig01_mapa_inventario_valor_suelo.png")

# ---- Figura 2: PREGUNTA CENTRAL - cafés del inventario por barrio vs suelo y vs su variación
fig, axs = plt.subplots(1, 2, figsize=(12, 5.4), sharey=True, gridspec_kw={"wspace": 0.12})
b = barrio.copy()
b["v"] = b["valor_suelo_m2_2025_cop"] / 1e6
paneles = [(axs[0], "v", "Valor del suelo 2025 (millones COP/m²)", "A. Frente al valor del suelo en 2025"),
           (axs[1], "var_valor_suelo_16_25_pct", "Variación del valor del suelo 2016-2025 (%)",
            "B. Frente a la valorización 2016-2025")]
for ax, col, xl, tit in paneles:
    for n in (1, 2, 3):
        fila = b[b.cafes_inventario == n].sort_values(col).reset_index(drop=True)
        ax.scatter(fila[col], [n] * len(fila), s=110, color=AZUL, edgecolor="white", linewidth=1.5, zorder=3)
        if n > 1:   # en las filas de 2 y 3 cafés se nombra cada barrio
            for i, r in fila.iterrows():
                ax.annotate(r["barrio"], (r[col], n), xytext=(0, 11 if i % 2 == 0 else -13),
                            textcoords="offset points", ha="center", va="center", fontsize=7.5, color=TINTA2)
    ax.set_title(tit, fontsize=11)
    ax.set_xlabel(xl)
    ax.set_yticks([1, 2, 3], ["1 café", "2 cafés", "3 cafés"])
    ax.set_ylim(0.5, 3.6)
    ax.grid(axis="x", color=REJILLA, linewidth=0.6)
    for n in (1, 2, 3):
        ax.axhline(n, color=REJILLA, lw=0.6, zorder=0)
axs[0].set_ylabel("Cafés del inventario en el barrio")
fig.suptitle("¿Hay más cafés del inventario donde el suelo vale más o se valorizó más?", x=0.01, ha="left",
             fontsize=13, fontweight="semibold", y=1.02)
nota(fig, "Un punto por barrio (16 barrios con al menos un café del inventario). En las dos comparaciones, barrios "
          "con 1, 2 o 3 cafés aparecen a lo largo de todo el rango: no se distingue un patrón claro.\n"
          "El conteo describe el inventario de 25 cafés seleccionados (1 a 3 por barrio y ningún barrio con cero "
          "cafés verificados), no la concentración de cafeterías en Bogotá. " + FUENTE_SUELO)
guardar(fig, "fig02_cafes_por_barrio_vs_suelo_y_valorizacion.png")


def barras_pregunta(ax, serie, orden, colores, titulo):
    vc = serie.value_counts().reindex(orden).fillna(0).astype(int)
    yy = np.arange(len(orden))[::-1]
    ax.barh(yy, 100 * vc.values / len(serie), color=colores, height=0.62)
    for yi, n in zip(yy, vc.values):
        ax.text(100 * n / len(serie) + 1, yi, f"{n} de {len(serie)} ({100 * n / len(serie):.0f} %)",
                va="center", fontsize=9, color=TINTA)
    ax.set_yticks(yy, orden)
    ax.tick_params(axis="y", length=0)
    ax.set_xlim(0, 90)
    ax.set_xticks([])
    ax.spines["bottom"].set_visible(False)
    ax.set_title(titulo, fontsize=11)


# ---- Figura 3: indicadores separados de precios, arriendos y salida de vecinos
fig, axs = plt.subplots(3, 1, figsize=(8.5, 7.4), gridspec_kw={"hspace": 0.6})
for ax, q, tit in zip(axs, ["p5", "p6", "p7"],
                      ["P5. Precios de comida, tiendas y servicios frente a hace 3 años",
                       "P6. Arriendos en los últimos 3 años", "P7. ¿Ha visto irse a vecinos de siempre?"]):
    barras_pregunta(ax, enc[q + "_c"], list(P[q].values()), [AZUL, AZUL, AZUL, GRIS_NOSE], tit)
fig.suptitle("Tres indicadores separados: precios, arriendos y salida de vecinos", x=0.01, ha="left",
             fontsize=13, fontweight="semibold", y=1.0)
a6, b6, _ = conteo(enc, *IND["P6 · Arriendos subieron mucho"])
a7, b7, _ = conteo(enc, *IND["P7 · Vio irse vecinos (muchos o algunos)"])
nota(fig, "Base: las 57 respuestas válidas; en esta figura el gris (no sabe / no puede comparar) se muestra y no se excluye. "
          f"Entre quienes sí pudieron responder: arriendos subieron mucho {a6} de {b6}; vio irse vecinos {a7} de {b7}.\n"
          "Ninguna persona eligió «más baratos» ni «bajaron». " + PERCEPCION + "\n" + FUENTE_ENC)
guardar(fig, "fig04_precios_arriendos_vecinos_p5_p7.png")

# ---- Figura 4: contraste por perfil (P9), con conteos y con quienes no pudieron opinar
orden9 = ["Empeoró mucho", "Empeoró un poco", "Sigue igual", "Mejoró un poco", "Mejoró mucho", "No puede opinar"]
col9 = [DIV_ROJO, DIV_ROJO_C, NEUTRO, DIV_AZUL_C, DIV_AZUL, GRIS_NOSE]
grupos = ["Vive, trabaja o estudia aquí", "Visitante (frecuente u ocasional)"]
fig, ax = plt.subplots(figsize=(9.5, 3.4))
for i, g in enumerate(grupos[::-1]):
    s = enc.loc[enc.perfil == g, "p9_c"]
    vc = s.value_counts().reindex(orden9).fillna(0).astype(int)
    left = 0
    for cat, c in zip(orden9, col9):
        n = vc[cat]
        ax.barh(i, n, left=left, color=c, height=0.6, edgecolor="white", linewidth=2)
        if n >= 2:
            ax.text(left + n / 2, i, str(n), ha="center", va="center", fontsize=9,
                    color="white" if c in (DIV_ROJO, DIV_AZUL) else TINTA)
        left += n
    mej = vc["Mejoró mucho"] + vc["Mejoró un poco"]
    base = len(s) - vc["No puede opinar"]
    ax.text(left + 0.6, i, f"Mejoró: {mej} de {base} que pudieron valorar\nNo pudo opinar: {vc['No puede opinar']} de {len(s)}",
            va="center", fontsize=8.5, color=TINTA2)
ax.set_yticks([0, 1], [f"{g}\n(n = {(enc.perfil == g).sum()})" for g in grupos[::-1]])
ax.tick_params(axis="y", length=0)
ax.set_xlim(0, 48)
ax.set_xticks([])
ax.spines["bottom"].set_visible(False)
ax.legend(handles=[plt.Rectangle((0, 0), 1, 1, color=c) for c in col9], labels=orden9, ncol=6, frameon=False,
          fontsize=8, loc="upper center", bbox_to_anchor=(0.42, -0.02), handlelength=1.2)
ax.set_title("P9. ¿Cómo ha cambiado el sector en los últimos 3 años?, según la relación con el sector")
nota(fig, "Las barras cuentan personas (números dentro de cada segmento). En gris, quienes no conocían el sector antes "
          "y no pudieron opinar.\n" + PERCEPCION + " " + FUENTE_ENC, y=-0.14)
guardar(fig, "fig05_cambio_percibido_p9_por_perfil.png")

# ---- Figura 5: percepciones por nivel del suelo del local, con denominador en cada barra
TIT_CORTO = dict(zip(IND, ["P5 · Precios mucho más caros", "P6 · Arriendos subieron mucho",
                            "P7 · Vio irse vecinos", "P9 · El sector mejoró"]))
fig, axs = plt.subplots(1, 4, figsize=(12.5, 3.6), sharey=True, gridspec_kw={"wspace": 0.08})
for ax, (k, (c, si, fuera)), col in zip(axs, IND.items(), [AZUL, NARANJA, AQUA, VIOLETA]):
    yy = np.arange(3)[::-1]
    for y, nivel in zip(yy, NIVELES):
        a, bse, _ = conteo(enc[enc.nivel_suelo == nivel], c, si, fuera)
        pct = 100 * a / bse if bse else 0
        ax.barh(y, 100, color="#f1f0ec", height=0.6)
        ax.barh(y, pct, color=col, height=0.6)
        ax.text(102, y, f"{a} de {bse}", va="center", fontsize=8.5, color=TINTA)
    ax.set_xlim(0, 135)
    ax.set_xticks([])
    ax.spines["bottom"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.set_title(TIT_CORTO[k], fontsize=9.5)
axs[0].set_yticks(np.arange(3)[::-1], [f"Suelo: {n}\n({(enc.nivel_suelo == n).sum()} respuestas)" for n in NIVELES])
fig.suptitle("Percepciones según el valor del suelo del barrio del local", x=0.01, ha="left",
             fontsize=13, fontweight="semibold", y=1.07)
nota(fig, "Grupos por valor del suelo 2025 del barrio del local (millones COP/m²). Cada barra: personas que responden "
          "«sí» sobre quienes pudieron responder esa pregunta "
          "(se excluye «no sabe / no puede comparar / no puede opinar»\nsolo en esta figura). P7 incluye «muchos» y «algunos»; "
          "P9 incluye «mucho» y «un poco». Bases de 5 a 19 personas: diferencias exploratorias. " + PERCEPCION, y=-0.04)
guardar(fig, "fig06_percepciones_por_nivel_de_suelo.png")

# ---- Figura 6: perfil (P1) y negocios nuevos (P2)
fig, axs = plt.subplots(2, 1, figsize=(8.5, 5.8), gridspec_kw={"hspace": 0.55})
barras_pregunta(axs[0], enc["p1_c"], list(P["p1"].values()), [AZUL] * 4, "P1. ¿Cuál es su relación con este sector?")
barras_pregunta(axs[1], enc["p2_c"], list(P["p2"].values()), [AZUL] * 3 + [GRIS_NOSE],
                "P2. ¿Qué negocios nuevos ha visto abrir en los últimos 3 años?")
fig.suptitle("Quién respondió y qué cambio comercial percibe", x=0.01, ha="left", fontsize=13, fontweight="semibold", y=1.0)
nota(fig, "Base: las 57 respuestas válidas; el gris (no conoce el sector) se muestra y no se excluye. Muestra no "
          "probabilística, tomada en o cerca de cafeterías.\n" + FUENTE_ENC)
guardar(fig, "fig03_perfil_p1_y_negocios_p2.png")

# ---- Figura 7 (secundaria): nivel del suelo frente a su variación, por barrio
fig, ax = plt.subplots(figsize=(9, 5.8))
ax.scatter(b["v"], b["var_valor_suelo_16_25_pct"], s=40 + 50 * b["cafes_inventario"], color=AZUL,
           edgecolor="white", linewidth=1.5, zorder=3)
off = {"El Nogal": (7, 0), "Chapinero Alto": (6, 4), "Sucre / Marly": (-7, 0), "Emaús / Zona G": (6, -10),
       "Granada / Pardo Rubio": (-6, -12)}
for _, r in b.iterrows():
    dx, dy = off.get(r["barrio"], (7, 0))
    ax.annotate(r["barrio"], (r["v"], r["var_valor_suelo_16_25_pct"]), xytext=(dx, dy), textcoords="offset points",
                fontsize=8, color=TINTA2, ha="left" if dx > 0 else "right", va="center")
ax.grid(color=REJILLA, linewidth=0.6)
ax.set_xlabel("Valor del suelo 2025 (millones COP/m²)")
ax.set_ylabel("Variación del valor del suelo 2016-2025 (%)")
ax.set_title("Resultado secundario: los barrios de suelo más caro se valorizaron menos en porcentaje")
nota(fig, "Un punto por barrio (16); el tamaño indica cuántos cafés del inventario hay en él. Puede reflejar un efecto de "
          "base (un valor alto crece menos en porcentaje).\nEs una comparación en un solo momento: no muestra causas. "
          + FUENTE_SUELO)
guardar(fig, "fig07_secundaria_suelo_vs_valorizacion.png")

# ---- Figura 8 (secundaria): precio del filtrado frente al suelo, por local
fig, ax = plt.subplots(figsize=(8.5, 5.5))
ax.scatter(inv["valor_suelo_m2_2025_cop"] / 1e6, inv["precio_filtrado_ref_cop"] / 1e3, s=70, color=AZUL,
           edgecolor="white", linewidth=1.5, zorder=3)
for _, r in inv[inv.id_cafe.isin(["CAF_015", "CAF_017", "CAF_004"])].iterrows():
    abajo = r["id_cafe"] == "CAF_004"
    ax.annotate(r["nombre_corto"], (r["valor_suelo_m2_2025_cop"] / 1e6, r["precio_filtrado_ref_cop"] / 1e3),
                xytext=(0, -13) if abajo else (7, 0), textcoords="offset points", fontsize=8, color=TINTA2,
                va="center", ha="center" if abajo else "left")
ax.grid(color=REJILLA, linewidth=0.6)
ax.set_xlabel("Valor del suelo 2025 del barrio (millones COP/m²)")
ax.set_ylabel("Precio de referencia del café filtrado (miles COP)")
ax.set_title("Resultado secundario: el filtrado suele costar más en barrios de suelo más caro")
nota(fig, "Un punto por café (25). Los cafés de un mismo barrio comparten el valor del suelo. Comparación en un solo "
          "momento: no muestra qué ocurrió primero.\n" + FUENTE_SUELO + " Precio: referencia de carta (Entrega 3a).")
guardar(fig, "fig08_secundaria_precio_filtrado_vs_suelo.png")

print("Listo. Salidas en", SALIDA.relative_to(RAIZ))
