# ----------------------------------------
# Configuración
# ----------------------------------------
CONFIG = {
    "columnas_energeticas": {
        "entradas": ["Potencia estándar", "Trabajo útil", "Calor disipado", "Calentamiento", "Reaction Duty positiva"],
        "salidas": ["Enfriamiento", "Pérdida de potencia estándar", "Pérdidas al ambiente", "Reaction Duty negativa"],
        "extra": ["Entalpía (kW-h)"]
    },
    "columnas": {
        "flujo_masico": "Flujo masico (kg/u.o)",
        "entalpia": "Entalpía (kW-h)",
        "operacion": "Operación",
        "origen": "Origen",
        "destino": "Destino",
        "componente": "Componente",
        "stream": "Stream"
    },
    "colores": {
        "entrada_energia": "rgba(255,100,100,0.8)",
        "salida_energia": "rgba(80,160,255,0.8)",
        "flujo_masico": "rgba(160,160,160,0.8)",
        "nodo_por_defecto": "lightblue",
        "nodo_entrada": "red",
        "nodo_salida": "blue"
    }
}

# ----------------------------------------
# Imports
# ----------------------------------------
import pandas as pd
import os, re, webbrowser
from tkinter import filedialog, Tk
from dash import Dash, dcc, html, Input, Output
import plotly.graph_objects as go

# ----------------------------------------
# Helpers
# ----------------------------------------
_excel_cache = {"df": None, "mtime": 0}

def seleccionar_archivo(default_path=None):
    root = Tk()
    root.withdraw()
    archivo = filedialog.askopenfilename(
        title="Selecciona tu archivo Excel",
        initialdir=os.path.dirname(default_path) if default_path else None,
        initialfile=os.path.basename(default_path) if default_path else None,
        filetypes=[("Archivos de Excel", "*.xlsx *.xls *.xlsm *.xlsb"), ("Todos los archivos", "*.*")]
    )
    if not archivo:
        print("No se seleccionó ningún archivo. Saliendo...")
        exit()
    return archivo

def leer_excel_cache(ruta_archivo):
    global _excel_cache
    mtime_actual = os.path.getmtime(ruta_archivo)
    if _excel_cache["df"] is None or mtime_actual != _excel_cache["mtime"]:
        df = pd.read_excel(ruta_archivo, sheet_name="Tabla Sankey")
        df.columns = df.columns.str.strip()
        _excel_cache.update({"df": df, "mtime": mtime_actual})
    return _excel_cache["df"]

def set_alpha_rgba(rgba_str, alpha):
    return re.sub(
        r"rgba\(\s*(\d+\s*,\s*\d+\s*,\s*\d+)\s*,\s*[\d.]+\s*\)",
        lambda m: f"rgba({m.group(1)},{alpha})",
        rgba_str
    )


COLORES_COMPONENTES_50 = [
    "#FFCC99", "#99CC33", "#66CC66", "#33CC99", "#33CCCC",  # verdes/azules
    
    "#FFCC33", "#FF9933", "#FF6666", "#CC6666", "#66CCFF",  # cálidos/azules
    "#FF3300", "#FF6600", "#FF9900", "#FFCC00", "#FFFF00",  # naranjas/amarillos
    "#CCFF33", "#99FF33", "#66FF33", "#33FF33", "#33FF66",  # verdes brillantes
    "#33FF99", "#33CCFF", "#3399FF", "#3366FF", "#3333FF",  # cian/azul intenso
    "#808080", "#A9A9A9", "#C0C0C0", "#FFFFFF", "#B2FF59",  # grises y blanco/neón
    "#FF33FF", "#FF33CC", "#FF3399", "#FF3366", "#FF3333",  # rosa/rojos
    "#FFB300", "#FF8033", "#FF4D4D", "#FF6666", "#FF99CC",  # cálidos
    "#3399FF", "#3366FF", "#6699FF", "#00FFFF", "#33FFCC",  # azules/cian
    "#CC0000", "#990000", "#660000", "#330000", "#000000"   # rojos oscuros y negro
]



def generar_colores_componentes(componentes):
    n = len(componentes)
    if n > len(COLORES_COMPONENTES_50):
        raise ValueError(f"Solo hay {len(COLORES_COMPONENTES_50)} colores definidos, tienes {n} componentes")
    return dict(zip(componentes, COLORES_COMPONENTES_50[:n]))

# ----------------------------------------
# Función para preparar datos Sankey
# ----------------------------------------
def cargar_datos(df, tipo="masa",
                 componentes_filtrados=None, nodos_filtrados=None,
                 columnas_energia_seleccionadas=None,
                 slider_entrada_agrupamiento=1.0,
                 slider_salida_agrupamiento=1.0):

    c = CONFIG["columnas"]
    colores = CONFIG["colores"]
    energia = CONFIG["columnas_energeticas"]

    df = df.copy()
    df = df[df[c["origen"]].notna() & df[c["destino"]].notna()]

    # --- Masa ---
    if tipo == "masa":
        if componentes_filtrados:
            df = df[df[c["componente"]].isin(componentes_filtrados)]
        if nodos_filtrados:
            nodos_set = set(nodos_filtrados)
            df = df[df[c["origen"]].isin(nodos_set) | df[c["destino"]].isin(nodos_set)]

        nodos = sorted(set(df[c["origen"]]).union(df[c["destino"]]))
        label_idx = {n: i for i, n in enumerate(nodos)}

        color_dict = generar_colores_componentes(df[c["componente"]].unique())


        enlaces = df[[c["origen"], c["destino"], c["flujo_masico"], c["stream"], c["componente"]]].copy()
        enlaces[c["flujo_masico"]] = pd.to_numeric(enlaces[c["flujo_masico"]], errors='coerce')
        enlaces = enlaces[enlaces[c["flujo_masico"]] > 0]

        enlaces["source"] = enlaces[c["origen"]].map(label_idx)
        enlaces["target"] = enlaces[c["destino"]].map(label_idx)
        enlaces["value"] = enlaces[c["flujo_masico"]]
        enlaces["label"] = enlaces[c["stream"]] + " (" + enlaces[c["componente"]] + ")"
        enlaces["color"] = enlaces[c["componente"]].map(color_dict).fillna(colores["flujo_masico"])

        return {
            "type": "sankey",
            "valueformat": ".2f",
            "valuesuffix": " kg/u.o",
            "node": {"pad": 90, "thickness": 17, "line": {"color": "black", "width": 0.5},
                     "label": nodos, "color": [colores["nodo_por_defecto"]] * len(nodos)},
            "link": {
                "source": enlaces["source"].tolist(),
                "target": enlaces["target"].tolist(),
                "value": enlaces["value"].tolist(),
                "label": enlaces["label"].tolist(),
                "color": enlaces["color"].tolist()
            },
            "max_salida": enlaces["value"].max(),
            "max_entrada": enlaces["value"].max(),
        }

    # --- Energía ---
    entradas_default = energia["entradas"]
    salidas_default = energia["salidas"]
    if columnas_energia_seleccionadas and "Todos" not in columnas_energia_seleccionadas:
        entradas = [col for col in columnas_energia_seleccionadas if col in entradas_default]
        salidas = [col for col in columnas_energia_seleccionadas if col in salidas_default]
    else:
        entradas = entradas_default
        salidas = salidas_default

    columnas_existentes = [col for col in entradas + salidas + energia["extra"] if col in df.columns]
    df_ent = df[df[columnas_existentes].apply(pd.to_numeric, errors='coerce').notna().any(axis=1)]

    nodos = set(df_ent[c["origen"]]).union(df_ent[c["destino"]])
    if c["operacion"] in df_ent.columns:
        nodos.update(df_ent[c["operacion"]].dropna().apply(lambda x: x.split(":")[0].strip()))
    nodos = sorted(n for n in nodos if isinstance(n, str) and n.strip())
    label_idx = {n: i for i, n in enumerate(nodos)}
    get_or_add = lambda n: label_idx.setdefault(n, len(label_idx))

    enlaces_list = []

    # Entalpía general (gris, no afectado por opacidad)
    if c["entalpia"] in df_ent.columns:
        df_ent[c["entalpia"]] = pd.to_numeric(df_ent[c["entalpia"]], errors='coerce')
        ent_df = df_ent[df_ent[c["entalpia"]].notna()]
        for _, row in ent_df.iterrows():
            enlaces_list.append({
                "source": get_or_add(row[c["origen"]]),
                "target": get_or_add(row[c["destino"]]),
                "value": abs(row[c["entalpia"]]),
                "label": row[c["stream"]],
                "color": colores["flujo_masico"]
            })

    # Entradas y salidas energéticas con agrupamiento (afectadas por slider)
    for _, row in df_ent.iterrows():
        op = row.get(c["operacion"])
        if pd.isna(op): continue
        op = op.split(":")[0].strip()
        op_idx = get_or_add(op)

        for col in entradas:
            val = pd.to_numeric(row.get(col), errors='coerce')
            if pd.notna(val):
                valor_max = df_ent[col].max() + 1
                nodo_origen = col if val < slider_entrada_agrupamiento * valor_max else f"{col} ({row[c['stream']]})"
                enlaces_list.append({
                    "source": get_or_add(nodo_origen),
                    "target": op_idx,
                    "value": val,
                    "label": f"{nodo_origen} a {op}",
                    "color": colores["entrada_energia"]
                })

        for col in salidas:
            val = pd.to_numeric(row.get(col), errors='coerce')
            if pd.notna(val):
                valor_max = df_ent[col].max() + 1
                nodo_destino = col if val < slider_salida_agrupamiento * valor_max else f"{col} ({row[c['stream']]})"
                enlaces_list.append({
                    "source": op_idx,
                    "target": get_or_add(nodo_destino),
                    "value": val,
                    "label": nodo_destino,
                    "color": colores["salida_energia"]
                })

    nodos_final = [n for n, _ in sorted(label_idx.items(), key=lambda x: x[1])]
    nodo_colores = [
        colores["nodo_entrada"] if any(e in n for e in entradas) else
        colores["nodo_salida"] if any(s in n for s in salidas) else
        colores["nodo_por_defecto"]
        for n in nodos_final
    ]

    df_enlaces = pd.DataFrame(enlaces_list)
    max_salida = df_enlaces["value"].max() if not df_enlaces.empty else 0
    max_entrada = df_enlaces["value"].max() if not df_enlaces.empty else 0

    return {
        "type": "sankey",
        "valueformat": ".2f",
        "valuesuffix": " kW",
        "node": {"pad": 44, "thickness": 15, "line": {"color": "black", "width": 0.5},
                 "label": nodos_final, "color": nodo_colores},
        "link": {
            "source": df_enlaces["source"].tolist(),
            "target": df_enlaces["target"].tolist(),
            "value": df_enlaces["value"].tolist(),
            "label": df_enlaces["label"].tolist(),
            "color": df_enlaces["color"].tolist()
        },
        "max_salida": max_salida,
        "max_entrada": max_entrada,
    }

# ----------------------------------------
# Inicialización Dash
# ----------------------------------------
app = Dash(__name__)
ruta_archivo_global = seleccionar_archivo()
df_global = leer_excel_cache(ruta_archivo_global)

componentes_unicos = sorted(df_global[CONFIG["columnas"]["componente"]].dropna().unique())
nodos_unicos = sorted(set(df_global[CONFIG["columnas"]["origen"]].dropna().unique())
                      .union(df_global[CONFIG["columnas"]["destino"]].dropna().unique()))

# Agregar operaciones si existen
if CONFIG["columnas"]["operacion"] in df_global.columns:
    nodos_unicos += list(df_global[CONFIG["columnas"]["operacion"]].dropna().apply(lambda x: x.split(":")[0].strip()))
    nodos_unicos = sorted(set(nodos_unicos))

# ----------------------------------------
# Layout
# ----------------------------------------
app.layout = html.Div([

    html.H3("Diagrama Sankey"),

    dcc.Graph(id="graph", style={"height": "100vh", "width": "100vw"}),

    html.Div([

        html.Div([
            html.Label("Opacidad de enlaces"),
            dcc.Slider(id='slider_opacidad', min=0.1, max=1, value=0.8, step=0.05)
        ], style={"margin-bottom": "10px"}),

        html.Div([
            dcc.Dropdown(id='tipo_selector', options=[
                {"label": "Flujo masico (kg/u.o.)", "value": "masa"},
                {"label": "Energía (kW)", "value": "energia"}
            ], value="masa", clearable=False)
        ], style={"margin-bottom": "10px"}),

        html.Div([
            dcc.Dropdown(
                id="columnas_energia_selector",
                placeholder="Selecciona columnas energéticas a mostrar",
                multi=True,
                options=[{"label": "Ninguno", "value": "Ninguno"}] +
                        [{"label": col, "value": col} for col in CONFIG["columnas_energeticas"]["entradas"] +
                         CONFIG["columnas_energeticas"]["salidas"]],
                value=["Ninguno"]
            )
        ], id="columnas_wrapper", style={"margin-bottom": "10px"}),

        html.Div([
            html.Label("Agrupamiento nodos entradas"),
            dcc.Slider(id='slider_entrada_agrupamiento', min=0, max=1, value=1, step=0.01)
        ], id="slider_entrada_wrapper", style={"margin-bottom": "10px"}),

        html.Div([
            html.Label("Agrupamiento nodos salidas"),
            dcc.Slider(id='slider_salida_agrupamiento', min=0, max=1, value=1, step=0.01)
        ], id="slider_salida_wrapper", style={"margin-bottom": "10px"}),

        html.Div([
            dcc.Dropdown(id='componente_selector', placeholder="Filtra por componente(s)", multi=True,
                         options=[{"label": c, "value": c} for c in componentes_unicos])
        ], id="componentes_wrapper", style={"margin-bottom": "10px"}),

        html.Div([
            dcc.Dropdown(id='nodo_selector', placeholder="Filtra por nodo(s)", multi=True,
                         options=[{"label": n, "value": n} for n in nodos_unicos])
        ], id="nodos_wrapper", style={"margin-bottom": "10px"}),

    ], style={"width": "10 vh", "margin": "auto"})
])

# ----------------------------------------
# Callbacks
# ----------------------------------------
@app.callback(
    Output("graph", "figure"),
    Output("componente_selector", "options"),
    Output("nodo_selector", "options"),
    Input("slider_opacidad", "value"),
    Input("tipo_selector", "value"),
    Input("columnas_energia_selector", "value"),
    Input("slider_entrada_agrupamiento", "value"),
    Input("slider_salida_agrupamiento", "value"),
    Input("componente_selector", "value"),
    Input("nodo_selector", "value")
)
def actualizar_grafico(opacidad, tipo, columnas_energia, slider_entrada, slider_salida, componentes_seleccionados, nodos_seleccionados):
    df_actualizado = leer_excel_cache(ruta_archivo_global)

    data = cargar_datos(
        df_actualizado,
        tipo=tipo,
        columnas_energia_seleccionadas=columnas_energia,
        slider_entrada_agrupamiento=slider_entrada,
        slider_salida_agrupamiento=slider_salida,
        componentes_filtrados=componentes_seleccionados,
        nodos_filtrados=nodos_seleccionados
    )

    # Ajuste de opacidad solo para enlaces de energía
    if tipo == "energia":
        link_colors = pd.Series(data["link"]["color"])
        color_entrada_base = CONFIG["colores"]["entrada_energia"][:-(len(",0.8)"))]
        color_salida_base = CONFIG["colores"]["salida_energia"][:-(len(",0.8)"))]

        def ajustar_opacidad(c):
            if c.startswith(color_entrada_base) or c.startswith(color_salida_base):
                return set_alpha_rgba(c, opacidad)
            return c

        link_colors = link_colors.apply(ajustar_opacidad)
        data["link"]["color"] = link_colors.tolist()

    componentes_unicos = sorted(df_actualizado[CONFIG["columnas"]["componente"]].dropna().unique())
    nodos_unicos = sorted(set(df_actualizado[CONFIG["columnas"]["origen"]].dropna().unique())
                          .union(df_actualizado[CONFIG["columnas"]["destino"]].dropna().unique()))
    if CONFIG["columnas"]["operacion"] in df_actualizado.columns:
        nodos_unicos += list(df_actualizado[CONFIG["columnas"]["operacion"]].dropna().apply(lambda x: x.split(":")[0].strip()))
        nodos_unicos = sorted(set(nodos_unicos))

    return (
        go.Figure(go.Sankey(link=data["link"], node=data["node"], arrangement="snap")).update_layout(font_size=10),
        [{"label": c, "value": c} for c in componentes_unicos],
        [{"label": n, "value": n} for n in nodos_unicos]
    )

@app.callback(
    Output("columnas_wrapper", "style"),
    Output("slider_entrada_wrapper", "style"),
    Output("slider_salida_wrapper", "style"),
    Output("componentes_wrapper", "style"),
    Output("nodos_wrapper", "style"),
    Input("tipo_selector", "value")
)
def mostrar_ocultar_controles(tipo):
    if tipo == "masa":
        return {"display": "none"}, {"display": "none"}, {"display": "none"}, {"display": "block"}, {"display": "block"}
    else:
        return {"display": "block"}, {"display": "block"}, {"display": "block"}, {"display": "none"}, {"display": "none"}

# ----------------------------------------
# Lanzamiento
# ----------------------------------------
if __name__ == "__main__":
    try:
        webbrowser.open_new("http://127.0.0.1:8050/")
        app.run(debug=False, port=8050, use_reloader=False)
    except Exception as e:
        print("Error al ejecutar la app:", e)
        exit()


