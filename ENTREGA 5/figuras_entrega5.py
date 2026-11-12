"""
Entrega 5 - Figuras y tablas del proyecto
"Cafeterías de especialidad, valor del suelo y valorización en barrios de Bogotá"
Haciendo Economía 2026-2, Universidad del Rosario.

Qué hace este script (todo desde los archivos del repositorio, sin pasos manuales):
  1. Lee la base RAW de la Entrega 4 (68 registros) y aplica la regla de depuración
     documentada en el informe: se conservan los registros "En muestra (14 locales)"
     y "Adicional del marco (25 locales)" que tienen coordenadas -> 57 válidos.
  2. Lee el directorio secundario de 25 cafeterías (Entrega 3a).
  3. Construye tablas: frecuencias P1-P9 sobre n = 57, cobertura por local y un
     resumen por barrio (unidad de análisis para el valor del suelo).
  4. Exporta los 25 locales del marco como GeoJSON (capa geográfica que "pega"
     con las bases mediante id_cafe).
  5. Genera 9 figuras en PNG.

Uso (desde la raíz del repositorio):
    python "ENTREGA 5/scripts/figuras_entrega5.py"

Requisitos: Python 3.9+, pandas, openpyxl, matplotlib, scipy (ver requirements.txt).
"""
from pathlib import Path
import json

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize
from matplotlib.lines import Line2D
from scipy.stats import spearmanr

# --------------------------------------------------------------------------- rutas
RAIZ = Path(__file__).resolve().parents[2]
RAW_XLSX = RAIZ / "ENTREGA 4" / "Base_Datos_RAW_Entrega4.xlsx"
HOJA_RAW = "1_Datos_RAW_Completos_68"
MARCO_CSV = RAIZ / "ENTREGA 3" / "base_secundaria_cafeterias_3a-v3.csv"

SALIDA = RAIZ / "ENTREGA 5"
FIG = SALIDA / "figuras"
TAB = SALIDA / "tablas"
DAT = SALIDA / "datos"
for p in (FIG, TAB, DAT):
    p.mkdir(parents=True, exist_ok=True)

ESTADOS_VALIDOS = ["En muestra (14 locales)", "Adicional del marco (25 locales)"]

# --------------------------------------------------------------------------- estilo
AZUL, NARANJA, AQUA = "#2a78d6", "#eb6834", "#1baf7a"
GRIS_NOSE = "#b9b8b2"          # "No sé / no puede opinar": siempre gris neutro
TINTA, TINTA2, TINTA3 = "#0b0b0b", "#52514e", "#8a8984"
REJILLA = "#e6e5e0"
RAMPA_AZUL = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]
CMAP_SUELO = LinearSegmentedColormap.from_list("suelo", RAMPA_AZUL)
DIV_ROJO, DIV_ROJO_C = "#c0392b", "#e88a80"     # empeoró mucho / un poco
DIV_AZUL, DIV_AZUL_C = "#1c5cab", "#86b6ef"     # mejoró mucho / un poco
NEUTRO = "#dedcd6"                              # sigue igual

plt.rcParams.update({
    "font.family": ["Inter", "DejaVu Sans"],
    "font.size": 10,
    "axes.edgecolor": REJILLA,
    "axes.labelcolor": TINTA2,
    "axes.titlesize": 13,
    "axes.titleweight": "semibold",
    "axes.titlecolor": TINTA,
    "axes.titlelocation": "left",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "xtick.color": TINTA2,
    "ytick.color": TINTA2,
    "axes.grid": False,
    "figure.facecolor": "white",
    "savefig.dpi": 200,
    "savefig.bbox": "tight",
})

FUENTE_ENC = "Fuente: encuesta de percepción, trabajo de campo 8-9 nov. 2026 (n = 57 válidas)."
FUENTE_SUELO = ("Valor del suelo: valores de referencia por barrio del directorio de la Entrega 3a "
                "(UAECD, serie 2016-2025).")


def nota(fig, texto, y=-0.02):
    fig.text(0.01, y, texto, ha="left", va="top", fontsize=8, color=TINTA3, wrap=True)


def guardar(fig, nombre):
    fig.savefig(FIG / nombre)
    plt.close(fig)
    print(f"  figura -> {nombre}")


def corto(nombre):
    """Nombre de local corto para etiquetas."""
    return (nombre.replace(" (Coffee Master 2026)", "").replace(" Coffee Roasters", "")
            .replace("Librería Café", "Librería").replace(" & Panadería Artesanal", "")
            .replace("Café Amor Perfecto", "Amor Perfecto").replace(" - Sede ", " - "))


# =========================================================================== 1. datos
print("1) Leyendo y depurando la base RAW")
raw = pd.read_excel(RAW_XLSX, sheet_name=HOJA_RAW)
raw.columns = [c.strip() for c in raw.columns]
ren = {c: f"p{c[0]}" for c in raw.columns if c[:2] in {f"{i}." for i in range(1, 10)}}
ren["¿Acepta participar voluntariamente?"] = "consentimiento"
raw = raw.rename(columns=ren)
assert len(raw) == 68, "La base RAW debería tener 68 registros"


def categoria_depuracion(e):
    if e == "En muestra (14 locales)":
        return "Válido: local de la muestra principal (14)"
    if e == "Adicional del marco (25 locales)":
        return "Válido: local adicional del marco de 25"
    if str(e).startswith("N/A"):
        return "Excluido: error de digitación (no identifica un local)"
    return "Excluido: fuera de las 7 zonas o prueba inicial"


raw["depuracion"] = raw["estado_registro_raw"].map(categoria_depuracion)
enc = raw[raw["estado_registro_raw"].isin(ESTADOS_VALIDOS) & raw["latitud"].notna()].copy()
assert len(enc) == 57, f"Se esperaban 57 válidos y hay {len(enc)}"
assert enc["consentimiento"].str.startswith("Sí").all()

# Base limpia publicada (sin nombres de encuestadores; los nombres de LOCALES sí se conservan)
cols_pub = ["id_encuesta", "Marca temporal", "consentimiento", "estado_registro_raw", "id_cafe",
            "nombre_establecimiento_verificado", "zona_estudio", "barrio_verificado_catastro",
            "localidad", "cod_barrio", "latitud", "longitud", "valor_suelo_m2_2025_cop",
            "var_valor_suelo_16_25_pct", "precio_filtrado_ref_cop", "rating_gmaps"] + \
           [f"p{i}" for i in range(1, 10)]
enc[cols_pub].to_csv(DAT / "encuesta_limpia_57.csv", index=False, encoding="utf-8-sig")

marco = pd.read_csv(MARCO_CSV, encoding="utf-8-sig")
assert len(marco) == 25
n_local = enc.groupby("id_cafe").size().rename("n_encuestas")
marco = marco.merge(n_local, on="id_cafe", how="left").fillna({"n_encuestas": 0})
marco["n_encuestas"] = marco["n_encuestas"].astype(int)
marco["grupo_campo"] = np.select(
    [marco["id_cafe"].isin(enc.loc[enc.estado_registro_raw == ESTADOS_VALIDOS[0], "id_cafe"]),
     marco["n_encuestas"] > 0],
    ["Muestra principal", "Adicional con encuestas"], "Sin encuestas")
marco["nombre_corto"] = marco["nombre_establecimiento"].map(corto)

# ----------------------------------------------------------------- etiquetas cortas
P = {
    "p1": {"Vivo aquí.": "Vive aquí",
           "Trabajo o estudio aquí.": "Trabaja o estudia aquí",
           "No vivo ni trabajo aquí, pero vengo con frecuencia.": "Visita con frecuencia",
           "Vengo de vez en cuando o es la primera vez que vengo.": "Visita ocasional / primera vez"},
    "p2": {"Sobre todo cafés, restaurantes modernos o tiendas de diseño.": "Cafés, restaurantes modernos o diseño",
           "Sobre todo negocios de siempre (tiendas de barrio, panaderías, talleres).": "Negocios de siempre",
           "Casi no he visto negocios nuevos.": "Casi ningún negocio nuevo",
           "No conozco el sector lo suficiente para saberlo.": "No conoce el sector"},
    "p5": {"Están mucho más caros.": "Mucho más caros",
           "Siguen más o menos igual.": "Más o menos igual",
           "Están más baratos.": "Más baratos",
           "No conocía el sector hace 3 años, no puedo comparar.": "No puede comparar"},
    "p6": {"Han subido mucho.": "Subieron mucho",
           "Siguen más o menos igual.": "Más o menos igual",
           "Han bajado.": "Bajaron",
           "No sé, no conozco lo suficiente el sector.": "No sabe"},
    "p7": {"Sí, muchos.": "Sí, muchos",
           "Sí, algunos.": "Sí, algunos",
           "No, casi todos siguen aquí.": "No, casi todos siguen",
           "No sé, no conozco lo suficiente el sector.": "No sabe"},
    "p9": {"Ha mejorado mucho.": "Mejoró mucho",
           "Ha mejorado un poco.": "Mejoró un poco",
           "Sigue igual, no veo diferencia.": "Sigue igual",
           "Ha empeorado un poco.": "Empeoró un poco",
           "Ha empeorado mucho.": "Empeoró mucho",
           "No lo conocía antes, no puedo opinar.": "No puede opinar"},
}
for q, m in P.items():
    enc[q + "_c"] = enc[q].map(m)
    assert enc[q + "_c"].notna().all(), f"Categoría no mapeada en {q}"

# =========================================================================== 2. tablas
print("2) Tablas")
filas = []
for i in range(1, 10):
    q = f"p{i}"
    vc = enc[q].value_counts()
    for cat, n in vc.items():
        filas.append({"pregunta": q.upper(), "categoria": cat, "n": n,
                      "pct_sobre_57": round(100 * n / len(enc), 1)})
pd.DataFrame(filas).to_csv(TAB / "frecuencias_p1_p9_n57.csv", index=False, encoding="utf-8-sig")

marco.sort_values(["zona_estudio", "id_cafe"])[
    ["id_cafe", "nombre_establecimiento", "zona_estudio", "barrio", "cod_barrio",
     "muestra_campo", "grupo_campo", "n_encuestas"]
].to_csv(TAB / "cobertura_por_local.csv", index=False, encoding="utf-8-sig")


def pct_si(s, si, nose):
    """% de 'si' entre quienes NO respondieron 'nose'; devuelve (pct, n_base)."""
    base = s[~s.isin(nose)]
    return (100 * base.isin(si).mean() if len(base) else np.nan, len(base))


# Nota: en el directorio, BAR_0205 aparece con dos nombres ("El Nogal (UPZ El Refugio)" y
# "El Nogal / Zona T") y dos valores del suelo distintos (8,90 y 8,75 M). Se agrega por código
# de barrio, se usa el nombre "El Nogal" y el promedio de ambos valores (ver README).
marco["barrio_std"] = np.where(marco["cod_barrio"] == "BAR_0205", "El Nogal", marco["barrio"])
barrio_marco = (marco.groupby("cod_barrio", as_index=False)
                .agg(barrio=("barrio_std", "first"), zona_estudio=("zona_estudio", "first"),valor_suelo_m2_2025_cop=("valor_suelo_m2_2025_cop", "mean"),
                     var_valor_suelo_16_25_pct=("var_valor_suelo_16_25_pct", "mean"),
                     cafes_en_marco=("id_cafe", "size"),
                     precio_filtrado_medio=("precio_filtrado_ref_cop", "mean")))
res = []
for cod, g in enc.groupby("cod_barrio"):
    p5, _ = pct_si(g["p5_c"], ["Mucho más caros"], ["No puede comparar"])
    p6, n6 = pct_si(g["p6_c"], ["Subieron mucho"], ["No sabe"])
    p7, n7 = pct_si(g["p7_c"], ["Sí, muchos", "Sí, algunos"], ["No sabe"])
    p9, n9 = pct_si(g["p9_c"], ["Mejoró mucho", "Mejoró un poco"], ["No puede opinar"])
    res.append({"cod_barrio": cod, "n_encuestas": len(g), "locales_encuestados": g["id_cafe"].nunique(),
                "pct_precios_mucho_mas_caros": p5, "n_base_p5": (~g["p5_c"].eq("No puede comparar")).sum(),
                "pct_arriendos_subieron": p6, "n_base_p6": n6,
                "pct_vio_irse_vecinos": p7, "n_base_p7": n7,
                "pct_dice_mejoro": p9, "n_base_p9": n9})
barrio = barrio_marco.merge(pd.DataFrame(res), on="cod_barrio", how="left")
barrio["n_encuestas"] = barrio["n_encuestas"].fillna(0).astype(int)
barrio.round(1).to_csv(TAB / "resumen_por_barrio.csv", index=False, encoding="utf-8-sig")

# =========================================================================== 3. GeoJSON
print("3) Capa geográfica")
geo = {"type": "FeatureCollection",
       "name": "cafeterias_marco_25",
       "crs": {"type": "name", "properties": {"name": "urn:ogc:def:crs:OGC:1.3:CRS84"}},
       "features": []}
for _, r in marco.iterrows():
    props = {k: (v.item() if hasattr(v, "item") else v) for k, v in r.items()
             if k not in ("latitud", "longitud")}
    geo["features"].append({"type": "Feature", "properties": props,
                            "geometry": {"type": "Point",
                                         "coordinates": [float(r["longitud"]), float(r["latitud"])]}})
(DAT / "cafeterias_marco_25.geojson").write_text(json.dumps(geo, ensure_ascii=False, indent=1),
                                                 encoding="utf-8")

# =========================================================================== 4. figuras
print("4) Figuras")

# ---- Figura 1: de 68 registros a 57 válidos
orden = ["Válido: local de la muestra principal (14)",
         "Válido: local adicional del marco de 25",
         "Excluido: fuera de las 7 zonas o prueba inicial",
         "Excluido: error de digitación (no identifica un local)"]
cnt = raw["depuracion"].value_counts().reindex(orden).fillna(0).astype(int)
fig, ax = plt.subplots(figsize=(8.5, 3.2))
colores = [AZUL, AQUA, GRIS_NOSE, GRIS_NOSE]
y = np.arange(len(orden))[::-1]
ax.barh(y, cnt.values, color=colores, height=0.6)
for yi, v in zip(y, cnt.values):
    ax.text(v + 0.8, yi, f"{v}", va="center", color=TINTA, fontsize=10)
ax.set_yticks(y, orden)
ax.set_xlim(0, 58)
ax.set_xlabel("Número de registros")
ax.tick_params(axis="y", length=0)
ax.set_title("De 68 registros recibidos a 57 encuestas válidas")
nota(fig, "Válidos = 50 + 7 = 57 (con coordenadas y dentro del marco de 25 locales). Excluidos = 9 + 2 = 11. "
          "67 registros son del 8 nov. y 1 del 9 nov.\n" + "Fuente: Base_Datos_RAW_Entrega4.xlsx, hoja "
          "1_Datos_RAW_Completos_68.", y=-0.1)
guardar(fig, "fig01_depuracion_68_a_57.png")

# ---- Figura 2: mapa del marco de 25 locales por valor del suelo
fig, ax = plt.subplots(figsize=(9.5, 7.6))
norm = Normalize(marco["valor_suelo_m2_2025_cop"].min() / 1e6, marco["valor_suelo_m2_2025_cop"].max() / 1e6)
con = marco[marco.n_encuestas > 0]
sin = marco[marco.n_encuestas == 0]
sc = ax.scatter(con["longitud"], con["latitud"], s=60 + 45 * con["n_encuestas"],
                c=con["valor_suelo_m2_2025_cop"] / 1e6, cmap=CMAP_SUELO, norm=norm,
                edgecolor="white", linewidth=1.5, zorder=3)
ax.scatter(sin["longitud"], sin["latitud"], s=60, facecolor="white",
           edgecolor=sin["valor_suelo_m2_2025_cop"].map(lambda v: CMAP_SUELO(norm(v / 1e6))),
           linewidth=2, zorder=3)
zc = marco.groupby("zona_estudio")[["longitud", "latitud"]].mean()
pos_txt = {"Usaquén": (-74.018, 4.703), "El Refugio y borde oriental": (-74.036, 4.682),
           "Chapinero central y Zona Rosa": (-74.036, 4.667), "Chapinero Alto y Pardo Rubio": (-74.036, 4.632),
           "Quinta Camacho y Teusaquillo": (-74.112, 4.672), "Salitre - Nicolás de Federmán": (-74.118, 4.628),
           "La Macarena": (-74.040, 4.606)}
for z, (tx, ty) in pos_txt.items():
    mz = marco[marco.zona_estudio == z]          # la flecha apunta al local de la zona más cercano a su centro
    k = ((mz.longitud - zc.loc[z, "longitud"]) ** 2 + (mz.latitud - zc.loc[z, "latitud"]) ** 2).idxmin()
    ax.annotate(z, xy=(mz.loc[k, "longitud"], mz.loc[k, "latitud"]), xytext=(tx, ty), fontsize=8.5,
                color=TINTA2, ha="left", va="center",
                arrowprops=dict(arrowstyle="-", color=TINTA3, lw=0.7, shrinkA=2, shrinkB=4))
ax.set_aspect(1 / np.cos(np.deg2rad(4.65)))
ax.set_xlim(-74.135, -73.995)
ax.set_ylim(4.598, 4.712)
ax.set_xlabel("Longitud")
ax.set_ylabel("Latitud")
ax.grid(color=REJILLA, linewidth=0.6)
cb = fig.colorbar(sc, ax=ax, shrink=0.55, pad=0.02)
cb.set_label("Valor del suelo 2025 (millones COP/m²)", color=TINTA2)
cb.outline.set_visible(False)
leg = [Line2D([], [], marker="o", ls="", ms=np.sqrt(60 + 45 * k), mfc="#9ec5f4", mec="white",
              label=f"{k} encuesta{'s' if k > 1 else ''}") for k in (1, 5, 9)]
leg.append(Line2D([], [], marker="o", ls="", ms=np.sqrt(60), mfc="white", mec="#6da7ec", mew=2,
                  label="Sin encuestas"))
ax.legend(handles=leg, loc="upper left", frameon=False, fontsize=8.5, labelspacing=1.2,
          title="Tamaño = encuestas válidas", title_fontsize=8.5, borderpad=1.0)
ax.set_title("Marco de 25 cafeterías: ubicación y valor del suelo del barrio")
nota(fig, "Marco = 25 cafeterías de especialidad seleccionadas en 7 zonas (Entrega 3a); no es el total de "
          "cafés existentes en cada barrio.\n17 locales tienen encuestas válidas; 8 no. " + FUENTE_SUELO,
     y=0.04)
guardar(fig, "fig02_mapa_marco_valor_suelo.png")

# ---- Figura 3: encuestas por local (cobertura)
m3 = marco.sort_values(["grupo_campo", "n_encuestas"], ascending=[True, True])
orden_g = {"Muestra principal": 0, "Adicional con encuestas": 1, "Sin encuestas": 2}
m3 = marco.assign(o=marco.grupo_campo.map(orden_g)).sort_values(["o", "n_encuestas"], ascending=[False, True])
fig, ax = plt.subplots(figsize=(8.5, 8))
col_g = {"Muestra principal": AZUL, "Adicional con encuestas": AQUA, "Sin encuestas": GRIS_NOSE}
yy = np.arange(len(m3))
ax.barh(yy, m3["n_encuestas"], color=m3["grupo_campo"].map(col_g), height=0.65)
for yi, (v, g) in enumerate(zip(m3["n_encuestas"], m3["grupo_campo"])):
    ax.text(v + 0.15, yi, "0" if v == 0 else str(v), va="center", fontsize=9, color=TINTA2 if v == 0 else TINTA)
ax.set_yticks(yy, [f"{n}  ({z})" for n, z in zip(m3["nombre_corto"], m3["barrio"])], fontsize=8.5)
ax.tick_params(axis="y", length=0)
ax.set_xlabel("Encuestas válidas")
ax.set_xlim(0, 10.5)
ax.legend(handles=[plt.Rectangle((0, 0), 1, 1, color=c) for c in col_g.values()],
          labels=[f"{k} ({(marco.grupo_campo == k).sum()} locales, "
                  f"{marco.loc[marco.grupo_campo == k, 'n_encuestas'].sum()} encuestas)" for k in col_g],
          loc="lower right", frameon=False, fontsize=8.5)
ax.set_title("Encuestas válidas por local del marco")
nota(fig, "Muestra principal = 14 locales del plan de análisis (Entrega 3b). Adicionales = locales del marco "
          "de 25 que no estaban en la muestra principal.\n" + FUENTE_ENC)
guardar(fig, "fig03_encuestas_por_local.png")

# ---- Figura 4: valor del suelo vs valorización, por barrio del marco
fig, ax = plt.subplots(figsize=(9, 6))
b = barrio_marco.copy()
b["v"] = b["valor_suelo_m2_2025_cop"] / 1e6
con_enc = b["cod_barrio"].isin(enc["cod_barrio"])
ax.scatter(b.loc[con_enc, "v"], b.loc[con_enc, "var_valor_suelo_16_25_pct"], s=40 + 50 * b.loc[con_enc, "cafes_en_marco"],
           color=AZUL, edgecolor="white", linewidth=1.5, zorder=3, label="Barrio con encuestas")
ax.scatter(b.loc[~con_enc, "v"], b.loc[~con_enc, "var_valor_suelo_16_25_pct"], s=40 + 50 * b.loc[~con_enc, "cafes_en_marco"],
           facecolor="white", edgecolor=AZUL, linewidth=1.8, zorder=3, label="Barrio sin encuestas")
off = {"El Nogal": (7, 0), "Chapinero Alto": (6, 4),
       "Sucre / Marly": (-7, 0), "Emaús / Zona G": (6, -10), "Granada / Pardo Rubio": (-6, -12)}
for _, r in b.iterrows():
    dx, dy = off.get(r["barrio"], (7, 0))
    ax.annotate(f"{r['barrio']} ({int(r['cafes_en_marco'])})", (r["v"], r["var_valor_suelo_16_25_pct"]),
                xytext=(dx, dy), textcoords="offset points", fontsize=8, color=TINTA2,
                ha="left" if dx > 0 else "right", va="center")
rho, pv = spearmanr(b["v"], b["var_valor_suelo_16_25_pct"])
ax.text(0.99, 0.97, f"Spearman ρ = {rho:.2f} (n = {len(b)} barrios)", transform=ax.transAxes,
        ha="right", va="top", fontsize=9, color=TINTA2)
ax.grid(color=REJILLA, linewidth=0.6)
ax.set_xlabel("Valor del suelo 2025 (millones COP/m²)")
ax.set_ylabel("Variación del valor del suelo 2016-2025 (%)")
ax.legend(frameon=False, loc="lower left", fontsize=8.5)
ax.set_title("Los barrios más caros del marco son los que menos se valorizaron")
nota(fig, "Unidad: barrio (16 barrios con al menos una cafetería del marco). Entre paréntesis y tamaño del punto: "
          "número de cafeterías del marco en el barrio.\nAsociación descriptiva, no causal. " + FUENTE_SUELO)
guardar(fig, "fig04_suelo_vs_valorizacion_barrio.png")

# ---- Figura 5: precio del café filtrado vs valor del suelo, por local
fig, ax = plt.subplots(figsize=(8.5, 5.5))
ax.scatter(marco["valor_suelo_m2_2025_cop"] / 1e6, marco["precio_filtrado_ref_cop"] / 1e3, s=70,
           color=AZUL, edgecolor="white", linewidth=1.5, zorder=3)
for _, r in marco.iterrows():
    if r["id_cafe"] in ("CAF_015", "CAF_017", "CAF_004"):     # extremos: más caro, segundo, más barato
        ax.annotate(r["nombre_corto"], (r["valor_suelo_m2_2025_cop"] / 1e6, r["precio_filtrado_ref_cop"] / 1e3),
                    xytext=(7, 0) if r["id_cafe"] != "CAF_004" else (0, -13), textcoords="offset points",
                    fontsize=8, color=TINTA2, va="center", ha="left" if r["id_cafe"] != "CAF_004" else "center")
rho, _ = spearmanr(marco["valor_suelo_m2_2025_cop"], marco["precio_filtrado_ref_cop"])
ax.text(0.01, 0.97, f"Spearman ρ = {rho:.2f} (n = 25 locales)", transform=ax.transAxes,
        ha="left", va="top", fontsize=9, color=TINTA2)
ax.grid(color=REJILLA, linewidth=0.6)
ax.set_xlabel("Valor del suelo 2025 del barrio (millones COP/m²)")
ax.set_ylabel("Precio de referencia del café filtrado (miles COP)")
ax.set_title("El café filtrado es más caro donde el suelo vale más")
nota(fig, "Unidad: local (25 cafeterías del marco). Varios locales comparten el valor del suelo de su barrio, "
          "por lo que no son observaciones independientes.\n" + FUENTE_SUELO + " Precio: referencia de carta (Entrega 3a).")
guardar(fig, "fig05_precio_filtrado_vs_suelo.png")


# ---- utilidades de barras horizontales de una pregunta
def barras_pregunta(ax, serie, orden, colores, titulo):
    vc = serie.value_counts().reindex(orden).fillna(0).astype(int)
    yy = np.arange(len(orden))[::-1]
    ax.barh(yy, 100 * vc.values / len(serie), color=colores, height=0.62)
    for yi, n in zip(yy, vc.values):
        ax.text(100 * n / len(serie) + 1, yi, f"{100 * n / len(serie):.0f} % ({n})", va="center", fontsize=9, color=TINTA)
    ax.set_yticks(yy, orden)
    ax.tick_params(axis="y", length=0)
    ax.set_xlim(0, 85)
    ax.set_xticks([])
    ax.spines["bottom"].set_visible(False)
    ax.set_title(titulo, fontsize=11)


# ---- Figura 6: perfil de los encuestados (P1) y negocios nuevos (P2)
fig, axs = plt.subplots(2, 1, figsize=(8.5, 5.8), gridspec_kw={"hspace": 0.55})
barras_pregunta(axs[0], enc["p1_c"], list(P["p1"].values()), [AZUL] * 4,
                "P1. ¿Cuál es su relación con este sector?")
barras_pregunta(axs[1], enc["p2_c"], list(P["p2"].values()), [AZUL] * 3 + [GRIS_NOSE],
                "P2. ¿Qué negocios nuevos ha visto abrir en los últimos 3 años?")
fig.suptitle("Quién respondió y qué cambio comercial percibe", x=0.01, ha="left",
             fontsize=13, fontweight="semibold", y=1.0)
nota(fig, "Porcentajes sobre n = 57 (número de respuestas entre paréntesis). Gris = no conoce el sector. "
          "Muestra no probabilística, tomada en o cerca de cafeterías.\n" + FUENTE_ENC)
guardar(fig, "fig06_perfil_p1_y_negocios_p2.png")

# ---- Figura 7: precios, arriendos y vecinos (P5-P7), cada uno por separado
fig, axs = plt.subplots(3, 1, figsize=(8.5, 7.4), gridspec_kw={"hspace": 0.6})
barras_pregunta(axs[0], enc["p5_c"], list(P["p5"].values()), [AZUL, AZUL, AZUL, GRIS_NOSE],
                "P5. Precios de comida, tiendas y servicios frente a hace 3 años")
barras_pregunta(axs[1], enc["p6_c"], list(P["p6"].values()), [AZUL, AZUL, AZUL, GRIS_NOSE],
                "P6. Arriendos en los últimos 3 años")
barras_pregunta(axs[2], enc["p7_c"], list(P["p7"].values()), [AZUL, AZUL, AZUL, GRIS_NOSE],
                "P7. ¿Ha visto irse a vecinos de siempre?")
p6c = enc.loc[enc.p6_c != "No sabe", "p6_c"]
p7c = enc.loc[enc.p7_c != "No sabe", "p7_c"]
fig.suptitle("Valorización percibida: precios, arriendos y salida de vecinos", x=0.01, ha="left",
             fontsize=13, fontweight="semibold", y=1.0)
nota(fig, f"Porcentajes sobre n = 57. Entre quienes sí conocen el sector: {100 * (p6c == 'Subieron mucho').mean():.0f} % "
          f"dice que los arriendos subieron mucho (n = {len(p6c)}) y "
          f"{100 * p7c.isin(['Sí, muchos', 'Sí, algunos']).mean():.0f} % vio irse vecinos (n = {len(p7c)}).\n"
          "Ninguna persona eligió «más baratos» ni «bajaron». Son percepciones, no precios ni arriendos observados. "
          + FUENTE_ENC)
guardar(fig, "fig07_precios_arriendos_vecinos_p5_p7.png")

# ---- Figura 8: percepciones por barrio, ordenadas por valorización del suelo
bb = barrio[barrio.n_encuestas > 0].sort_values("var_valor_suelo_16_25_pct").reset_index(drop=True)
fig, ax = plt.subplots(figsize=(9, 6.2))
series = [("pct_precios_mucho_mas_caros", "n_base_p5", "Precios mucho más caros (P5)", AZUL, "o"),
          ("pct_arriendos_subieron", "n_base_p6", "Arriendos subieron mucho (P6)", NARANJA, "s"),
          ("pct_vio_irse_vecinos", "n_base_p7", "Vio irse vecinos (P7)", AQUA, "D")]
yy = np.arange(len(bb))
for k, (col, ncol, lab, c, mk) in enumerate(series):
    off_y = (k - 1) * 0.22
    ax.scatter(bb[col], yy + off_y, color=c, marker=mk, s=55, edgecolor="white", linewidth=1.2, zorder=3, label=lab)
ax.set_yticks(yy, [f"{r.barrio}  · {r.var_valor_suelo_16_25_pct:.0f} %  (n = {r.n_encuestas})"
                   for r in bb.itertuples()], fontsize=8.5)
ax.tick_params(axis="y", length=0)
ax.set_xlim(-3, 103)
ax.set_xlabel("% de quienes respondieron (excluye «no sabe / no puede comparar»)")
ax.set_xticks([0, 25, 50, 75, 100])
for x in (0, 25, 50, 75, 100):
    ax.axvline(x, color=REJILLA, lw=0.6, zorder=0)
ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.45, -0.1), ncol=3, fontsize=8.5)
ax.set_title("Percepciones por barrio, ordenadas por valorización del suelo 2016-2025")
nota(fig, "Unidad: barrio (11 barrios con encuestas). Junto a cada barrio: variación del suelo 2016-2025 y número de "
          "encuestas. Con n tan pequeños, las diferencias entre barrios son solo exploratorias.\n"
          "Cada indicador se reporta por separado; no se combinan en un índice. " + FUENTE_ENC, y=-0.07)
guardar(fig, "fig08_percepciones_por_barrio.png")

# ---- Figura 9: cambio general percibido (P9) según relación con el sector (P1)
grupos = {"Vive, trabaja o estudia aquí": ["Vive aquí", "Trabaja o estudia aquí"],
          "Visitante (frecuente u ocasional)": ["Visita con frecuencia", "Visita ocasional / primera vez"]}
orden9 = ["Empeoró mucho", "Empeoró un poco", "Sigue igual", "Mejoró un poco", "Mejoró mucho", "No puede opinar"]
col9 = [DIV_ROJO, DIV_ROJO_C, NEUTRO, DIV_AZUL_C, DIV_AZUL, GRIS_NOSE]
fig, ax = plt.subplots(figsize=(9, 3.6))
filas9 = list(grupos.items()) + [("Total", sum(grupos.values(), []))]
for i, (g, cats) in enumerate(filas9[::-1]):
    s = enc.loc[enc.p1_c.isin(cats), "p9_c"]
    pct = s.value_counts().reindex(orden9).fillna(0) / len(s) * 100
    left = 0
    for cat, c in zip(orden9, col9):
        w = pct[cat]
        ax.barh(i, w, left=left, color=c, height=0.6, edgecolor="white", linewidth=2)
        if w >= 7:
            ax.text(left + w / 2, i, f"{w:.0f}%", ha="center", va="center", fontsize=8.5,
                    color="white" if c in (DIV_ROJO, DIV_AZUL) else TINTA)
        left += w
    ax.text(101, i, f"n = {len(s)}", va="center", fontsize=8.5, color=TINTA2)
ax.set_yticks(range(len(filas9)), [g for g, _ in filas9[::-1]])
ax.tick_params(axis="y", length=0)
ax.set_xlim(0, 110)
ax.set_xticks([])
ax.spines["bottom"].set_visible(False)
ax.legend(handles=[plt.Rectangle((0, 0), 1, 1, color=c) for c in col9], labels=orden9, ncol=6,
          frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.47, -0.02), handlelength=1.2)
ax.set_title("P9. ¿Cómo ha cambiado el sector en los últimos 3 años?")
nota(fig, "Porcentajes dentro de cada grupo. Gris = no conocía el sector antes. " + FUENTE_ENC, y=-0.12)
guardar(fig, "fig09_cambio_percibido_p9_por_perfil.png")

print("Listo. Figuras en", FIG.relative_to(RAIZ))
