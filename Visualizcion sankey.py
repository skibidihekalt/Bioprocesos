import threading
import webbrowser

from dash import Dash, dcc, html, Input, Output
import plotly.graph_objects as go
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# === FORZAR SELECCIÓN DE ARCHIVO AL INICIO ===
root = tk.Tk()
root.withdraw()
ruta_archivo_global = filedialog.askopenfilename(title="Selecciona tu archivo Excel", filetypes=[("Excel files", "*.xlsx")])
if not ruta_archivo_global:
    print("No se seleccionó ningún archivo.")
    input("Presiona ENTER para salir...")
    exit()

# === FUNCIÓN PARA CARGAR Y PROCESAR DATOS DE EXCEL ===
def cargar_datos(path_excel, tipo="masa"):
    df = pd.read_excel(path_excel, sheet_name="Tabla Sankey")
    df.columns = [col.strip() for col in df.columns]
    source, target, value, label, color = [], [], [], [], []

    if tipo == "masa":
        flujos_df = df[df["Origen"].notna() & df["Destino"].notna()]
        nodos = sorted(n for n in set(flujos_df["Origen"]).union(set(flujos_df["Destino"])) if isinstance(n, str) and n.strip() != "")
        label_idx = {name: i for i, name in enumerate(nodos)}
        for _, row in flujos_df.iterrows():
            try:
                flujo = float(str(row["Flujo masico (kg/u.o.)"]).replace(",", "."))
                if flujo == 0: continue
                source.append(label_idx[row["Origen"]])
                target.append(label_idx[row["Destino"]])
                value.append(flujo)
                label.append(row["Stream"])
                color.append("rgba(160,160,160,0.8)")
            except: continue
        nodo_colores = ["lightblue"] * len(nodos)
        unidad = "kg/h"

    else:
        flujos_df = df[df["Origen"].notna() & df["Destino"].notna()]
        nodos = set(flujos_df["Origen"]).union(set(flujos_df["Destino"]))
        if "Operación" in df.columns:
            nodos.update(df["Operación"].dropna().apply(lambda x: x.split(":")[0]))
        nodos = sorted(n for n in nodos if isinstance(n, str) and n.strip() != "")
        if "Origen externa de energía" not in nodos:
            nodos.insert(0, "Origen externa de energía")
        label_idx = {name: i for i, name in enumerate(nodos)}

        for _, row in flujos_df.iterrows():
            try:
                ent = float(str(row["Entalpía (kW)"]).replace(",", "."))
                if ent == 0: continue
                source.append(label_idx[row["Origen"]])
                target.append(label_idx[row["Destino"]])
                value.append(abs(ent))
                label.append(row["Stream"])
                color.append("rgba(160,160,160,0.8)")
            except: continue

        for _, row in df.iterrows():
            try:
                val = float(str(row["W + Q"]).replace(",", "."))
                operacion = row["Operación"]
                if pd.isna(val) or pd.isna(operacion): continue
                eq = operacion.split(":")[0].strip()
            except: continue

            if eq not in label_idx:
                nodos.append(eq)
                label_idx[eq] = len(label_idx)

            if val > 0:
                source.append(label_idx["Origen externa de energía"])
                target.append(label_idx[eq])
                value.append(val)
                label.append(f"Energía a {operacion}")
                color.append("rgba(255,100,100,0.8)")
            elif val < 0:
                destino = f"Pérdida de energía en {operacion}"
                if destino not in label_idx:
                    nodos.append(destino)
                    label_idx[destino] = len(label_idx)
                source.append(label_idx[eq])
                target.append(label_idx[destino])
                value.append(abs(val))
                label.append(destino)
                color.append("rgba(80,160,255,0.8)")
        nodo_colores = []
        for l in nodos:
            if l == "Origen externa de energía":
                nodo_colores.append("red")
            elif l.startswith("Pérdida"):
                nodo_colores.append("blue")
            else:
                nodo_colores.append("lightblue")
        unidad = "kW"

    return {
        "type": "sankey",
        "valueformat": ".2f",
        "valuesuffix": f" {unidad}",
        "node": {"pad": 15, "thickness": 15, "line": {"color": "black", "width": 0.5}, "label": nodos, "color": nodo_colores},
        "link": {"source": source, "target": target, "value": value, "label": label, "color": color}
    }

# === INTERFAZ WEB CON DASH ===
app = Dash(__name__)
app.layout = html.Div([
    html.H3("Diagrama Sankey"),
    dcc.Graph(id="graph", style={"height": "1200px", "width": "2400px"}),
    html.Div([
        html.Label("Opacidad global"),
        dcc.Slider(id='slider', min=0.1, max=1, value=0.8, step=0.1)
    ], style={'margin-bottom': '20px'}),
    html.Div([
        html.Label("Selecciona tipo de diagrama:"),
        dcc.Dropdown(id='tipo_selector',
            options=[
                {"label": "Flujo masico (kg/u.o.)", "value": "masa"},
                {"label": "Energía (kW)", "value": "energia"},
            ],
            value="masa",
            clearable=False
        )
    ]),
    html.Div([
        html.Label("Selecciona una corriente para resaltar:"),
        dcc.Dropdown(id="stream_selector", options=[], value=None, placeholder="Selecciona un flujo")
    ])
])

@app.callback(
    Output("graph", "figure"),
    Input("slider", "value"),
    Input("stream_selector", "value"),
    Input("tipo_selector", "value")
)
def actualizar_grafico(opacidad, flujo_resaltado, tipo):
    try:
        data = cargar_datos(ruta_archivo_global, tipo)
    except Exception as e:
        print("Error al generar Sankey:", e)
        return go.Figure()
    node = data["node"].copy()
    link = data["link"].copy()
    link["color"] = [
        "rgba(0,200,0,1)" if lbl == flujo_resaltado else c.replace("0.8", str(opacidad))
        for lbl, c in zip(link["label"], link["color"])
    ]
    fig = go.Figure(go.Sankey(link=link, node=node, arrangement="snap"))
    fig.update_layout(font_size=10, height=1200, width=2400)
    return fig

@app.callback(
    Output("stream_selector", "options"),
    Input("tipo_selector", "value")
)
def actualizar_dropdown(tipo):
    try:
        data = cargar_datos(ruta_archivo_global, tipo)
        return [{"label": s, "value": s} for s in sorted(set(data["link"]["label"]))]    
    except:
        return []

if __name__ == "__main__":
    def abrir_navegador():
        webbrowser.open_new("http://127.0.0.1:8050/")
    threading.Timer(1, abrir_navegador).start()
    app.run(debug=False)
