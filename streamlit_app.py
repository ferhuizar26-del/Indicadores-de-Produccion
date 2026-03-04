import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Indicadores de Producción", layout="wide")

st.title("Dashboard de Control de Producción")

uploaded = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded:

    df = pd.read_excel(uploaded, sheet_name="INYECCION_Ordenada", header=3)
    df.columns = df.columns.astype(str).str.replace("\n", " ", regex=False).str.strip()
    

    # Columnas importantes
    COL_FECHA = "Fecha DD/MM/AA"
    COL_SEMANA = "Semana"
    COL_MAQUINA = "Nombre de la maquina "
    COL_TURNO = "Turno"
    COL_OPERADOR = "Operador "

    COL_DISP = "DISPONIBILIDAD"
    COL_EFIC = "EFICIENCIA"
    COL_CAL = "CALIDAD"
    COL_OEE = "OEE"

    # Convertir fecha
    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors='coerce')

    # Convertir indicadores
    for c in [COL_DISP, COL_EFIC, COL_CAL, COL_OEE]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Si están en %
    for c in [COL_DISP, COL_EFIC, COL_CAL, COL_OEE]:
        if df[c].max() > 1.5:
            df[c] = df[c] / 100

    # -------------------
    # FILTROS
    # -------------------

    st.sidebar.header("Filtros")

    maquinas = st.sidebar.multiselect(
        "Máquina",
        options=df[COL_MAQUINA].dropna().unique(),
        default=df[COL_MAQUINA].dropna().unique()
    )

    turnos = st.sidebar.multiselect(
        "Turno",
        options=df[COL_TURNO].dropna().unique(),
        default=df[COL_TURNO].dropna().unique()
    )

    operadores = st.sidebar.multiselect(
        "Operador",
        options=df[COL_OPERADOR].dropna().unique(),
        default=df[COL_OPERADOR].dropna().unique()
    )

    df = df[
        (df[COL_MAQUINA].isin(maquinas)) &
        (df[COL_TURNO].isin(turnos)) &
        (df[COL_OPERADOR].isin(operadores))
    ]

    # -------------------
    # KPIs
    # -------------------

    st.subheader("Indicadores actuales")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Disponibilidad", f"{df[COL_DISP].mean()*100:.1f}%")
    col2.metric("Eficiencia", f"{df[COL_EFIC].mean()*100:.1f}%")
    col3.metric("Calidad", f"{df[COL_CAL].mean()*100:.1f}%")
    col4.metric("OEE", f"{df[COL_OEE].mean()*100:.1f}%")

    # -------------------
    # AGRUPACIÓN POR FECHA
    # -------------------

    df_group = df.groupby(COL_FECHA)[[COL_DISP, COL_EFIC, COL_CAL, COL_OEE]].mean().reset_index()

    st.subheader("Tendencia de indicadores")

    c1, c2 = st.columns(2)

    with c1:
        fig = px.line(df_group, x=COL_FECHA, y=COL_DISP, markers=True, title="Disponibilidad")
        st.plotly_chart(fig, use_container_width=True)

        fig = px.line(df_group, x=COL_FECHA, y=COL_EFIC, markers=True, title="Eficiencia")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig = px.line(df_group, x=COL_FECHA, y=COL_CAL, markers=True, title="Calidad")
        st.plotly_chart(fig, use_container_width=True)

        fig = px.line(df_group, x=COL_FECHA, y=COL_OEE, markers=True, title="OEE")
        st.plotly_chart(fig, use_container_width=True)

else:

    st.info("Sube el Excel para generar el dashboard.")
