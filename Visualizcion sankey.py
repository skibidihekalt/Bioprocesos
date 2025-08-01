import threading
import webbrowser
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from dash import Dash, dcc, html, Input, Output
import tkinter as tk
from tkinter import filedialog

# --------- Selección de archivo Excel ---------
def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()
    archivo = filedialog.askopenfilename(
        title="Selecciona tu archivo Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not archivo:
        print("No se seleccionó ningún archivo.")
        input("Presiona ENTER para salir...")
        exit()
    return archivo

ruta_archivo_global = seleccionar_archivo()
df_global = pd.read_excel(ruta_archivo_global, sheet_name="Tabla Sankey")
df_global.columns = df_global.columns.str.strip()

# --------- Helpers ---------
def to_float(val):
    try:
        return float(str(val).replace(",", ".")) if pd.notna(val) else None
    except ValueError:
        return None

def escalar_valores(valores, objetivo=100.0):
    return objetivo / max(valores) if valores else 1.0

def asignar_colores(lista):
    colores = px.colors.qualitative.Plotly * ((len(lista) // len(px.colors.qualitative.Plotly)) + 1)
    return dict(zip(lista, colores))

componentes_unicos = sorted(df_global["Componente"].dropna().unique())
nodos_unicos = sorted(set(df_global["Origen"].dropna().unique())
                      .union(df_global["Destino"].dropna().unique()))
if "Operación" in df_global.columns:
    nodos_unicos += list(df_global["Operación"].dropna().apply(lambda x: x.split(":")[0].strip()))
    nodos_unicos = sorted(set(nodos_unicos))

# --------- Carga de datos ---------
def cargar_datos(df, tipo="masa", umbral_salida=50, umbral_entrada=50,
                 componentes_filtrados=None, nodos_filtrados=None,
                 columnas_energia_seleccionadas=None):

    df = df.copy()
    df = df[df["Origen"].notna() & df["Destino"].notna()]

    if nodos_filtrados:
        nodos_set = set(nodos_filtrados)
        df = df[df["Origen"].isin(nodos_set) | df["Destino"].isin(nodos_set)]

    if tipo == "masa":
        if componentes_filtrados:
            df = df[df["Componente"].isin(componentes_filtrados)]

        nodos = sorted(set(df["Origen"]).union(df["Destino"]))
        label_idx = {n: i for i, n in enumerate(nodos)}
        color_dict = asignar_colores(df["Componente"].unique())

        flujo_vals = df["Flujo masico (kg/u.o)"].apply(to_float).dropna().tolist()
        escalar = escalar_valores(flujo_vals)

        source, target, value, label, color = [], [], [], [], []
        for _, row in df.iterrows():
            flujo = to_float(row["Flujo masico (kg/u.o)"])
            if flujo and flujo > 0:
                source.append(label_idx[row["Origen"]])
                target.append(label_idx[row["Destino"]])
                value.append(flujo * escalar)
                label.append(f"{row['Stream']} ({row['Componente']})")
                color.append(color_dict.get(row["Componente"], "rgba(160,160,160,0.8)"))

        return {
            "type": "sankey",
            "valueformat": ".2f",
            "valuesuffix": " kg/u.o (normalizado)",
            "node": {"pad": 100, "thickness": 15, "line": {"color": "black", "width": 0.5},
                     "label": nodos, "color": ["lightblue"] * len(nodos)},
            "link": {"source": source, "target": target, "value": value, "label": label, "color": color},
            "max_salida": max(value, default=0) + 1,
            "max_entrada": max(value, default=0) + 1,
        }

    # Energía
    entradas_default = ["Potencia estándar", "Trabajo útil", "Calor disipado", "Calentamiento", "Reaction Duty positiva"]
    salidas_default = ["Enfriamiento", "Pérdida de potencia estándar", "Pérdidas al ambiente", "Reaction Duty negativa"]

    if columnas_energia_seleccionadas:
        entradas = [col for col in columnas_energia_seleccionadas if col in entradas_default]
        salidas = [col for col in columnas_energia_seleccionadas if col in salidas_default]
    else:
        entradas = entradas_default
        salidas = salidas_default

    columnas_existentes = [col for col in entradas + salidas + ["Entalpía (kW-h)"] if col in df.columns]
    df_ent = df[df[columnas_existentes].apply(lambda col: col.map(to_float)).notna().any(axis=1)]

    nodos = set(df_ent["Origen"]).union(df_ent["Destino"])
    if "Operación" in df_ent.columns:
        nodos.update(df_ent["Operación"].dropna().apply(lambda x: x.split(":")[0].strip()))
    nodos = sorted(n for n in nodos if isinstance(n, str) and n.strip())
    label_idx = {n: i for i, n in enumerate(nodos)}
    get_or_add = lambda n: label_idx.setdefault(n, len(label_idx))

    source, target, value, label, color = [], [], [], [], []
    max_entrada, max_salida = 0, 0
    temp_values, temp_source, temp_target, temp_label, temp_color = [], [], [], [], []

    for _, row in df_ent.iterrows():
        ent = to_float(row.get("Entalpía (kW-h)"))
        if ent:
            temp_source.append(get_or_add(row["Origen"]))
            temp_target.append(get_or_add(row["Destino"]))
            temp_values.append(abs(ent))
            temp_label.append(row["Stream"])
            temp_color.append("rgba(160,160,160,0.8)")

    escalar = escalar_valores(temp_values)
    value.extend([v * escalar for v in temp_values])
    source.extend(temp_source)
    target.extend(temp_target)
    label.extend(temp_label)
    color.extend(temp_color)
    max_entrada = max(temp_values, default=0)
    max_salida = max(temp_values, default=0)

    for _, row in df_ent.iterrows():
        op = row.get("Operación")
        if pd.isna(op): continue
        op = op.split(":")[0].strip()
        op_idx = get_or_add(op)

        for col in entradas:
            val = to_float(row.get(col))
            if val:
                max_entrada = max(max_entrada, val)
                origen = f"{col} en {op}" if (val * escalar) >= umbral_entrada else col
                source.append(get_or_add(origen))
                target.append(op_idx)
                value.append(val * escalar)
                label.append(f"{origen} a {op}")
                color.append("rgba(255,100,100,0.8)")

        for col in salidas:
            val = to_float(row.get(col))
            if val:
                max_salida = max(max_salida, val)
                destino = f"{col} en {op}" if (val * escalar) >= umbral_salida else col
                source.append(op_idx)
                target.append(get_or_add(destino))
                value.append(val * escalar)
                label.append(destino)
                color.append("rgba(80,160,255,0.8)")

    nodos_final = [n for n, _ in sorted(label_idx.items(), key=lambda x: x[1])]
    nodo_colores = [
        "red" if any(e in n for e in entradas) else
        "blue" if any(s in n for s in salidas) else
        "lightblue"
        for n in nodos_final
    ]

    return {
        "type": "sankey",
        "valueformat": ".2f",
        "valuesuffix": " kW",
        "node": {"pad":25, "thickness": 10, "line": {"color": "black", "width": 0.5},
                 "label": nodos_final, "color": nodo_colores},
        "link": {"source": source, "target": target, "value": value, "label": label, "color": color},
        "max_salida": (max_salida * escalar) + 1,
        "max_entrada": (max_entrada * escalar) + 1,
    }

# --------- Layout de Dash ---------
app = Dash(__name__)
app.layout = html.Div([
    html.H3("Diagrama Sankey"),
    dcc.Graph(id="graph", style={"height": "90vh", "width": "95vw"}),
    dcc.Slider(id='slider', min=0.1, max=1, value=0.8, step=0.1),

    dcc.Dropdown(id='tipo_selector', options=[
        {"label": "Flujo masico (kg/u.o.)", "value": "masa"},
        {"label": "Energía (kW)", "value": "energia"}
    ], value="energia", clearable=False),

    dcc.Dropdown(id="componente_selector", placeholder="Filtra por componente(s)", multi=True,
                 options=[{"label": c, "value": c} for c in componentes_unicos]),

    dcc.Dropdown(id="nodo_selector", placeholder="Filtra por nodo(s)", multi=True,
                 options=[{"label": n, "value": n} for n in nodos_unicos]),

    dcc.Dropdown(id="columnas_energia_selector", placeholder="Selecciona columnas energéticas a mostrar", multi=True,
                 options=[{"label": col, "value": col} for col in [
                     "Trabajo útil", "Calor disipado", "Enfriamiento", "Pérdidas al ambiente",
                     "Calentamiento", "Ninguno", "Pérdida de potencia estándar",
                     "Reaction Duty positiva", "Reaction Duty negativa", "Potencia estándar",
                 ]],
                 value=[
                     "Trabajo útil", "Calor disipado", "Enfriamiento", "Calentamiento", "Reaction Duty positiva", "Reaction Duty negativa",
]
    ),

    html.Div(id='umbral_entrada_wrapper', children=[
        dcc.Slider(id='umbral_entrada_slider', min=0, max=200, value=50, step=1)
    ]),
    html.Div(id='umbral_salida_wrapper', children=[
        dcc.Slider(id='umbral_salida_slider', min=0, max=200, value=50, step=1)
    ]),
])

# --------- Callbacks ---------
@app.callback(
    Output("graph", "figure"),
    Input("slider", "value"),
    Input("tipo_selector", "value"),
    Input("umbral_salida_slider", "value"),
    Input("umbral_entrada_slider", "value"),
    Input("componente_selector", "value"),
    Input("nodo_selector", "value"),
    Input("columnas_energia_selector", "value")
)
def actualizar_grafico(opacidad, tipo, umbral_salida, umbral_entrada,
                       componentes, nodos, columnas_energia):
    try:
        df_actualizado = pd.read_excel(ruta_archivo_global, sheet_name="Tabla Sankey")
        df_actualizado.columns = df_actualizado.columns.str.strip()
        data = cargar_datos(
            df_actualizado, tipo, umbral_salida, umbral_entrada,
            componentes, nodos, columnas_energia_seleccionadas=columnas_energia
        )
        link = data["link"].copy()
        link["color"] = [c.replace("0.8", str(opacidad)) for c in link["color"]]
        return go.Figure(go.Sankey(link=link, node=data["node"], arrangement="snap")).update_layout(
            font_size=10
        )
    except Exception as e:
        print("Error:", e)
        return go.Figure()

@app.callback(Output("componente_selector", "style"),
              Input("tipo_selector", "value"))
def mostrar_selector_componentes(tipo):
    return {"display": "block"} if tipo == "masa" else {"display": "none"}

@app.callback(Output("nodo_selector", "style"),
              Input("tipo_selector", "value"))
def mostrar_selector_nodos(tipo):
    return {"display": "block"} if tipo == "masa" else {"display": "none"}

@app.callback(Output("columnas_energia_selector", "style"),
              Input("tipo_selector", "value"))
def mostrar_columnas_selector(tipo):
    return {"display": "block"} if tipo == "energia" else {"display": "none"}

@app.callback(
    Output("umbral_entrada_wrapper", "style"),
    Output("umbral_salida_wrapper", "style"),
    Input("tipo_selector", "value")
)
def mostrar_ocultar_umbral(tipo):
    if tipo == "masa":
        return {"display": "none"}, {"display": "none"}
    return {"display": "block"}, {"display": "block"}

@app.callback(
    Output("umbral_salida_slider", "max"),
    Output("umbral_salida_slider", "value"),
    Output("umbral_entrada_slider", "max"),
    Output("umbral_entrada_slider", "value"),
    Input("tipo_selector", "value")
)
def actualizar_sliders(tipo):
    try:
        df_actualizado = pd.read_excel(ruta_archivo_global, sheet_name="Tabla Sankey")
        df_actualizado.columns = df_actualizado.columns.str.strip()
        data = cargar_datos(df_actualizado, tipo)
        return (
            max(10, int(data.get("max_salida", 100))),
            int(data.get("max_salida", 100)),
            max(10, int(data.get("max_entrada", 100))),
            int(data.get("max_entrada", 100))
        )
    except:
        return 100, 20, 100, 20

# --------- Lanzamiento de la app ---------
if __name__ == "__main__":
    try:
        threading.Timer(1, lambda: webbrowser.open_new("http://127.0.0.1:8050/")).start()
        app.run(debug=False)
    except Exception as e:
        print("Error al ejecutar la app:", e)
        input("Presiona ENTER para salir...")

