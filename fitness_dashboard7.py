"""
FITNESS OFFICE CHALLENGE - Dashboard Dash
Alimentado por: fitness_registro.xlsx
Ejecutar: python fitness_dashboard.py
Abrir: http://localhost:8050
"""

import dash
from dash import dcc, html, dash_table, Input, Output
import pandas as pd
from io import StringIO
import plotly.graph_objects as go
from datetime import date, timedelta
import os

# ── CONFIG ────────────────────────────────────────────────────────────────────
EXCEL_PATH       = "fitness_registro.xlsx"
MULTA_POR_DIA    = 100
DIAS_META_SEMANA = 4
META_DIAS_MES    = 16
ANIO_RETO        = 2026

# ── PALETA ────────────────────────────────────────────────────────────────────
C_AZUL     = "#2D2D6B"
C_AZUL2    = "#5C5CA8"
C_VERDE    = "#276221"
C_VERDE_BG = "#C6EFCE"
C_ROJO     = "#9C0006"
C_ROJO_BG  = "#FFC7CE"
C_GRIS     = "#F5F5F5"
C_BLANCO   = "#FFFFFF"

REGLAS = [
    ("🗓️", "Duración",     "1 mes — prueba piloto"),
    ("💪", "Meta semanal", "4 días de actividad física por semana (lun–dom)"),
    ("📋", "Registro",     "Anotar SI o NO en el Excel cada día hábil"),
    ("💸", "Multa",        "$100 por cada día faltante para llegar a 4 en la semana"),
    ("📌", "Ejemplo",      "Si hiciste 3 días esa semana → 1 día de multa = $100"),
    ("✅", "Cumplimiento", "Semana perfecta = 4 días realizados, sin multa"),
]

DIAS_ES = {0:"Lun", 1:"Mar", 2:"Mier", 3:"Jue", 4:"Vie", 5:"Sab", 6:"Dom"}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def col_a_fecha(col):
    try:
        p = col.strip().split()[-1].split("/")
        dia, mes = int(p[0]), int(p[1])
        # Ajustar año: meses 6-12 → 2026, meses 1-5 → 2027 (si el reto cruza año)
        anio = ANIO_RETO if mes >= 6 else ANIO_RETO + 1
        return date(anio, mes, dia)
    except Exception:
        return None

def es_dia_valido(col):
    """Todos los días cuentan (lun-dom), excepto columnas sin fecha reconocible."""
    return col_a_fecha(col) is not None

def num_semana(col, inicio):
    f = col_a_fecha(col)
    if f is None:
        return None
    lun_ini = inicio - timedelta(days=inicio.weekday())
    lun_f   = f - timedelta(days=f.weekday())
    n = (lun_f - lun_ini).days // 7 + 1
    return n if n >= 1 else None

def domingo_semana(col):
    f = col_a_fecha(col)
    if f is None:
        return None
    return f + timedelta(days=(6 - f.weekday()))  # domingo = weekday 6

# ── LEER DATOS ────────────────────────────────────────────────────────────────
def leer_datos(path):
    if not os.path.exists(path):
        return pd.DataFrame()

    df_raw = pd.read_excel(path, sheet_name="Registro", header=1)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    col_nombre = next((c for c in df_raw.columns if "articipante" in c), df_raw.columns[1])

    excluir = {"#", col_nombre, "Dias OK", "Días OK"}
    cols_raw = [c for c in df_raw.columns
                if c not in excluir and not c.startswith("Unnamed")]

    # Si pandas convirtió alguna cabecera a datetime, formatearla de vuelta
    rename_map = {}
    for c in cols_raw:
        try:
            d = pd.to_datetime(c)
            rename_map[c] = f"{DIAS_ES[d.weekday()]} {d.day:02d}/{d.month:02d}"
        except Exception:
            pass
    if rename_map:
        df_raw = df_raw.rename(columns=rename_map)
        cols_raw = [rename_map.get(c, c) for c in cols_raw]

    # Todos los días con fecha válida (lun-dom)
    cols_habiles = [c for c in cols_raw if es_dia_valido(c)]

    fechas_validas = [col_a_fecha(c) for c in cols_habiles if col_a_fecha(c)]
    if not fechas_validas:
        return pd.DataFrame()
    inicio = min(fechas_validas)

    rows = []
    for _, row in df_raw.iterrows():
        nombre = row[col_nombre]
        if pd.isna(nombre) or str(nombre).strip() == "":
            continue
        nombre = str(nombre).strip()
        for col in cols_habiles:
            v = row[col]
            if hasattr(v, "iloc"):
                v = v.iloc[0]
            val = str(v).strip().upper() if pd.notna(v) else ""
            val = val if val in ("SI", "NO") else None
            dom = domingo_semana(col)
            rows.append({
                "participante": nombre,
                "dia_col":      col,
                "estado":       val,
                "semana":       num_semana(col, inicio),
                "domingo_sem":  dom.isoformat() if dom else None,
            })

    return pd.DataFrame(rows)

# ── KPIs ─────────────────────────────────────────────────────────────────────
def calcular_kpis(df):
    if df.empty:
        return pd.DataFrame()

    hoy = date.today()
    resumen = []

    for nombre, grp in df.groupby("participante"):
        reg      = grp[grp["estado"].notna()]
        si_total = int((reg["estado"] == "SI").sum())
        no_total = int((reg["estado"] == "NO").sum())

        multa_total = 0
        for semana, sg in grp.groupby("semana"):
            if pd.isna(semana):
                continue
            # Semana cerrada = el domingo ya pasó
            dom_str = sg["domingo_sem"].dropna().iloc[0] if sg["domingo_sem"].notna().any() else None
            if dom_str is None:
                continue
            domingo = date.fromisoformat(dom_str)
            if domingo >= hoy:
                continue  # semana todavía en curso

            si_sem    = int((sg["estado"] == "SI").sum())
            # Faltantes = cuántos días le faltaron para llegar a 4 (máx 4)
            faltantes = max(0, DIAS_META_SEMANA - si_sem)
            multa_total += faltantes * MULTA_POR_DIA

        pct = round(si_total / META_DIAS_MES * 100, 1)
        resumen.append({
            "Participante":   nombre,
            "Días ✓":         si_total,
            "Días ✗":         no_total,
            "Días reg.":      si_total + no_total,
            "% Cumplimiento": pct,
            "Multa ($)":      multa_total,
        })

    return pd.DataFrame(resumen).sort_values("% Cumplimiento", ascending=False)

# ── GRÁFICAS ──────────────────────────────────────────────────────────────────
def layout_base(titulo, height=350):
    return dict(
        title=dict(text=titulo, font=dict(color=C_AZUL, size=14, family="Calibri"), x=0.02),
        plot_bgcolor=C_BLANCO, paper_bgcolor=C_BLANCO,
        font=dict(family="Calibri", color="#333"),
        margin=dict(l=40, r=30, t=50, b=40),
        height=height,
    )

def grafica_cumplimiento(kpis):
    colores = [C_VERDE if p >= 100 else C_AZUL2 if p >= 75 else C_ROJO
               for p in kpis["% Cumplimiento"]]
    fig = go.Figure(go.Bar(
        x=kpis["Participante"], y=kpis["% Cumplimiento"],
        marker_color=colores,
        text=[f"{v}%" for v in kpis["% Cumplimiento"]],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Cumplimiento: %{y}%<extra></extra>",
    ))
    fig.add_hline(y=100, line_dash="dot", line_color=C_VERDE, line_width=2,
                  annotation_text="Meta 100%", annotation_font_color=C_VERDE)
    fig.update_layout(**layout_base("% Cumplimiento mensual"))
    fig.update_yaxes(range=[0, 130], ticksuffix="%")
    return fig

def grafica_multas(kpis):
    km = kpis[kpis["Multa ($)"] > 0].sort_values("Multa ($)", ascending=True)
    if km.empty:
        fig = go.Figure()
        fig.add_annotation(text="Sin multas acumuladas 🎉",
                           xref="paper", yref="paper", x=0.5, y=0.5,
                           showarrow=False, font=dict(size=18, color=C_VERDE))
        fig.update_layout(**layout_base("Multas acumuladas ($)"))
        return fig
    fig = go.Figure(go.Bar(
        y=km["Participante"], x=km["Multa ($)"], orientation="h",
        marker_color=C_ROJO,
        text=[f"${v:,}" for v in km["Multa ($)"]],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Multa: $%{x:,}<extra></extra>",
    ))
    fig.update_layout(**layout_base("Multas acumuladas ($)"))
    fig.update_xaxes(tickprefix="$")
    return fig

def grafica_heatmap(df):
    if df.empty:
        return go.Figure()
    df = df.copy()
    df["dia_col"] = df["dia_col"].astype(str)
    pivot = df.pivot_table(index="participante", columns="dia_col",
                           values="estado",
                           aggfunc=lambda x: x.iloc[0] if len(x) else None)

    def orden(col):
        try:
            p = col.split()[-1].split("/")
            return (int(p[1]), int(p[0]))
        except Exception:
            return (99, 99)
    pivot = pivot.reindex(sorted(pivot.columns, key=orden), axis=1)

    z, custom = [], []
    for _, row in pivot.iterrows():
        fz, fc = [], []
        for v in row:
            if v == "SI":   fz.append(1);  fc.append("SI ✓")
            elif v == "NO": fz.append(-1); fc.append("NO ✗")
            else:           fz.append(0);  fc.append("Sin registro")
        z.append(fz); custom.append(fc)

    fig = go.Figure(go.Heatmap(
        z=z, x=list(pivot.columns), y=list(pivot.index),
        colorscale=[[0, C_ROJO_BG], [0.5, "#F0F0F0"], [1, C_VERDE_BG]],
        showscale=False,
        hovertemplate="<b>%{y}</b><br>%{x}<br>%{customdata}<extra></extra>",
        customdata=custom,
    ))
    fig.update_layout(**layout_base("Actividad diaria — mapa de calor", height=320))
    fig.update_xaxes(tickangle=45, tickfont_size=9)
    return fig

def panel_reglas():
    items = []
    for icono, titulo, desc in REGLAS:
        items.append(html.Div(
            style={"display":"flex","gap":"12px","alignItems":"flex-start",
                   "padding":"10px 0","borderBottom":"1px solid #EEEEEE"},
            children=[
                html.Span(icono, style={"fontSize":"1.3rem","minWidth":"28px"}),
                html.Div([
                    html.Span(titulo+": ", style={"fontWeight":"bold","color":C_AZUL,"fontSize":"0.85rem"}),
                    html.Span(desc,        style={"color":"#444","fontSize":"0.85rem"}),
                ]),
            ]))
    return html.Div(
        style={"backgroundColor":C_BLANCO,"borderRadius":"10px",
               "boxShadow":"0 1px 6px rgba(0,0,0,0.08)","padding":"16px 20px","height":"100%"},
        children=[
            html.H3("📌 Reglas del reto",
                    style={"color":C_AZUL,"marginTop":"0","fontSize":"1rem","marginBottom":"4px"}),
            html.Div(items),
        ])

# ── LAYOUT ────────────────────────────────────────────────────────────────────
app = dash.Dash(__name__, title="Fitness Challenge Dashboard")
server = app.server

app.layout = html.Div(
    style={"fontFamily":"Calibri, sans-serif","backgroundColor":C_GRIS,"minHeight":"100vh"},
    children=[
        html.Div(style={"backgroundColor":C_AZUL,"padding":"20px 30px",
                        "display":"flex","alignItems":"center","gap":"16px"}, children=[
            html.Span("🏋️", style={"fontSize":"2rem"}),
            html.Div([
                html.H1("FITNESS OFFICE CHALLENGE",
                        style={"color":C_BLANCO,"margin":"0","fontSize":"1.6rem","letterSpacing":"2px"}),
                html.P("Tablero de seguimiento · Prueba piloto",
                       style={"color":"#BBBBEE","margin":"4px 0 0","fontSize":"0.85rem"}),
            ]),
            html.Div(style={"marginLeft":"auto","display":"flex","gap":"12px","alignItems":"center"}, children=[
                html.Button("🔄 Actualizar", id="btn-refresh",
                            style={"backgroundColor":C_AZUL2,"color":C_BLANCO,"border":"none",
                                   "padding":"8px 18px","borderRadius":"6px","cursor":"pointer","fontSize":"0.9rem"}),
                dcc.Interval(id="intervalo", interval=30_000, n_intervals=0),
                html.Span(id="lbl-actualizacion", style={"color":"#BBBBEE","fontSize":"0.8rem"}),
            ]),
        ]),

        html.Div(id="kpi-cards",
                 style={"display":"flex","gap":"16px","padding":"20px 30px 0","flexWrap":"wrap"}),

        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1fr",
                        "gap":"16px","padding":"16px 30px 0"}, children=[
            html.Div(dcc.Graph(id="graf-cumplimiento"),
                     style={"backgroundColor":C_BLANCO,"borderRadius":"10px",
                            "boxShadow":"0 1px 6px rgba(0,0,0,0.08)","padding":"8px"}),
            html.Div(dcc.Graph(id="graf-multas"),
                     style={"backgroundColor":C_BLANCO,"borderRadius":"10px",
                            "boxShadow":"0 1px 6px rgba(0,0,0,0.08)","padding":"8px"}),
        ]),

        html.Div(style={"display":"grid","gridTemplateColumns":"1fr 1.8fr",
                        "gap":"16px","padding":"16px 30px 0"}, children=[
            html.Div(id="panel-reglas"),
            html.Div(dcc.Graph(id="graf-heatmap"),
                     style={"backgroundColor":C_BLANCO,"borderRadius":"10px",
                            "boxShadow":"0 1px 6px rgba(0,0,0,0.08)","padding":"8px"}),
        ]),

        html.Div(style={"padding":"16px 30px 30px"}, children=[
            html.Div(style={"backgroundColor":C_BLANCO,"borderRadius":"10px",
                            "boxShadow":"0 1px 6px rgba(0,0,0,0.08)","padding":"16px"}, children=[
                html.H3("Detalle por participante",
                        style={"color":C_AZUL,"marginTop":"0","fontSize":"1rem"}),
                html.Div(id="tabla-detalle"),
            ]),
        ]),

        dcc.Store(id="store-datos"),
    ])

# ── CALLBACKS ─────────────────────────────────────────────────────────────────
@app.callback(
    Output("store-datos", "data"),
    Output("lbl-actualizacion", "children"),
    Input("intervalo", "n_intervals"),
    Input("btn-refresh", "n_clicks"),
)
def actualizar_datos(n_intervals, n_clicks):
    from datetime import datetime
    df = leer_datos(EXCEL_PATH)
    ahora = datetime.now().strftime("%H:%M:%S")
    if df.empty:
        return None, f"Sin datos · {ahora}"
    return df.to_json(orient="records"), f"Actualizado: {ahora}"


@app.callback(
    Output("kpi-cards", "children"),
    Output("graf-cumplimiento", "figure"),
    Output("graf-multas", "figure"),
    Output("graf-heatmap", "figure"),
    Output("tabla-detalle", "children"),
    Output("panel-reglas", "children"),
    Input("store-datos", "data"),
)
def actualizar_ui(data_json):
    empty_fig = go.Figure()
    empty_fig.update_layout(paper_bgcolor=C_BLANCO, plot_bgcolor=C_BLANCO, height=350)

    if not data_json:
        msg = html.Div(f"📂 No se encontró '{EXCEL_PATH}'.",
                       style={"color":C_ROJO,"padding":"20px"})
        return [], empty_fig, empty_fig, empty_fig, msg, panel_reglas()

    df   = pd.read_json(StringIO(data_json), orient="records")
    kpis = calcular_kpis(df)

    avg_cumpl    = round(kpis["% Cumplimiento"].mean(), 1) if not kpis.empty else 0
    sin_multa    = int((kpis["Multa ($)"] == 0).sum()) if not kpis.empty else 0
    total_multas = int(kpis["Multa ($)"].sum()) if not kpis.empty else 0

    def card(icono, valor, etiqueta, color=C_AZUL):
        return html.Div(
            style={"backgroundColor":color,"borderRadius":"10px","padding":"16px 20px",
                   "minWidth":"160px","flex":"1","boxShadow":"0 2px 8px rgba(0,0,0,0.1)"},
            children=[
                html.Div(icono, style={"fontSize":"1.6rem","marginBottom":"4px"}),
                html.Div(str(valor), style={"fontSize":"1.8rem","fontWeight":"bold",
                                             "color":C_BLANCO,"lineHeight":"1.1"}),
                html.Div(etiqueta, style={"fontSize":"0.75rem","color":C_BLANCO,
                                           "opacity":"0.85","marginTop":"2px"}),
            ])

    cards = [
        card("📊", f"{avg_cumpl}%",      "Cumplimiento promedio"),
        card("✅", f"{sin_multa}/{len(kpis)}", "Participantes sin multa", C_VERDE),
        card("💸", f"${total_multas:,}", "Total multas acumuladas",  C_ROJO),
    ]

    cols_tabla = [c for c in kpis.columns if not c.startswith("_")]
    tabla = dash_table.DataTable(
        data=kpis[cols_tabla].to_dict("records"),
        columns=[{"name": c, "id": c} for c in cols_tabla],
        style_header={"backgroundColor":C_AZUL,"color":C_BLANCO,
                      "fontWeight":"bold","textAlign":"center","fontSize":"12px"},
        style_data={"textAlign":"center","fontSize":"12px"},
        style_data_conditional=[
            {"if":{"filter_query":"{Multa ($)} > 0","column_id":"Multa ($)"},
             "color":C_ROJO,"fontWeight":"bold"},
            {"if":{"filter_query":"{% Cumplimiento} >= 100","column_id":"% Cumplimiento"},
             "color":C_VERDE,"fontWeight":"bold"},
            {"if":{"row_index":"odd"},"backgroundColor":C_GRIS},
        ],
        style_table={"overflowX":"auto"},
        sort_action="native", page_action="none",
    )

    return (cards,
            grafica_cumplimiento(kpis),
            grafica_multas(kpis),
            grafica_heatmap(df),
            tabla,
            panel_reglas())

# ── RUN ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "="*55)
    print("  FITNESS OFFICE CHALLENGE - Dashboard")
    print("="*55)
    print(f"  Excel: {os.path.abspath(EXCEL_PATH)}")
    print("  URL:   http://localhost:8050")
    print("  Auto-refresh cada 30 segundos")
    print("="*55 + "\n")
    app.run(debug=False, host="0.0.0.0", port=8050)
