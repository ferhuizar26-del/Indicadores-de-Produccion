import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Indicadores de Producción", layout="wide")
st.title("Dashboard de Control de Producción")

uploaded = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded:
    # Lee hoja y respeta que header está en fila 4 (index 3)
    df = pd.read_excel(uploaded, sheet_name="INYECCION_Ordenada", header=3)
    df.columns = df.columns.astype(str).str.replace("\n", " ", regex=False).str.strip()

    def find_col(expected: str):
        exp = expected.strip().lower()
        for c in df.columns:
            if c.strip().lower() == exp:
                return c
        for c in df.columns:
            if exp in c.strip().lower():
                return c
        return None

    # Columnas (detectadas)
    COL_FECHA = find_col("Fecha DD/MM/AA") or "Fecha DD/MM/AA"
    COL_SEMANA = find_col("Semana") or "Semana"
    COL_MAQUINA = find_col("Nombre de la maquina") or "Nombre de la maquina"
    COL_TURNO = find_col("Turno") or "Turno"
    # COL_OPERADOR = find_col("Operador")  # opcional (lo quitamos del filtro por ahora)

    COL_DISP = find_col("DISPONIBILIDAD") or "DISPONIBILIDAD"
    COL_EFIC = find_col("EFICIENCIA") or "EFICIENCIA"
    COL_CAL = find_col("CALIDAD") or "CALIDAD"
    COL_OEE = find_col("OEE") or "OEE"      # OJO: si tu columna se llama "OEE." cámbialo a "OEE."
    if COL_OEE not in df.columns and "OEE." in df.columns:
        COL_OEE = "OEE."

    # Validación mínima
    needed = [COL_FECHA, COL_SEMANA, COL_MAQUINA, COL_TURNO, COL_DISP, COL_EFIC, COL_CAL, COL_OEE]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        st.error(f"Faltan columnas en el Excel: {missing}\n\nColumnas detectadas: {list(df.columns)}")
        st.stop()

    # Tipos
    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce")

    for c in [COL_DISP, COL_EFIC, COL_CAL, COL_OEE]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Normaliza a porcentaje (0-100)
    # - Si viene 0-1 -> multiplica por 100
    # - Si viene 0-100 -> deja igual
    for c in [COL_DISP, COL_EFIC, COL_CAL, COL_OEE]:
        mx = df[c].max(skipna=True)
        if pd.notna(mx) and mx <= 1.5:
            df[c] = df[c] * 100

    # -------------------
    # FILTROS (planta)
    # -------------------
    st.sidebar.header("Filtros")

    maquinas = st.sidebar.multiselect(
        "Máquina",
        options=sorted(df[COL_MAQUINA].dropna().unique().tolist()),
        default=sorted(df[COL_MAQUINA].dropna().unique().tolist())
    )

    turnos = st.sidebar.multiselect(
        "Turno",
        options=sorted(df[COL_TURNO].dropna().unique().tolist()),
        default=sorted(df[COL_TURNO].dropna().unique().tolist())
    )

    # (Quitamos operador para uso “planta”. Si algún día lo necesitan, lo reactivamos.)
    df = df[
        (df[COL_MAQUINA].isin(maquinas)) &
        (df[COL_TURNO].isin(turnos))
    ].copy()

    # -------------------
    # KPIs (promedio del filtro actual)
    # -------------------
    st.subheader("Indicadores General")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Disponibilidad", f"{df[COL_DISP].mean():.1f}%")
    col2.metric("Eficiencia", f"{df[COL_EFIC].mean():.1f}%")
    col3.metric("Calidad", f"{df[COL_CAL].mean():.1f}%")
    col4.metric("OEE", f"{df[COL_OEE].mean():.1f}%")

    # -------------------
    # TENDENCIA POR SEMANA
    # -------------------
    st.subheader("Tendencia semanal de indicadores")

    df_group = (
        df.groupby(COL_SEMANA, as_index=False)[[COL_DISP, COL_EFIC, COL_CAL, COL_OEE]]
        .mean(numeric_only=True)
    )

    # Ordena semana (si es numérica)
    df_group[COL_SEMANA] = pd.to_numeric(df_group[COL_SEMANA], errors="ignore")
    df_group = df_group.sort_values(COL_SEMANA)

    def plot_indicator(y_col, title):
        fig = px.line(
            df_group,
            x=COL_SEMANA,
            y=y_col,
            markers=True,
            title=title
        )
        fig.update_layout(
            xaxis_title="Semana",
            yaxis_title="%",
            yaxis=dict(range=[0, 100]),
            margin=dict(l=10, r=10, t=50, b=10)
        )
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        plot_indicator(COL_DISP, "Disponibilidad (%) por semana")
        plot_indicator(COL_EFIC, "Eficiencia (%) por semana")
    with c2:
        plot_indicator(COL_CAL, "Calidad (%) por semana")
        plot_indicator(COL_OEE, "OEE (%) por semana")

else:
    st.info("Sube el Excel para generar el dashboard.")
