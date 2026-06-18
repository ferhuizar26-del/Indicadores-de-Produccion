import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Indicadores de Producción", layout="wide")
st.markdown("""
<style>
div[data-testid="stMetric"] {
    padding: 10px 10px;
}
</style>
""", unsafe_allow_html=True)

st.title("Control de Producción Industrial MEFI")

uploaded = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded:
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

    # -------------------
    # COLUMNAS
    # -------------------
    COL_FECHA    = find_col("Fecha DD/MM/AA") or "Fecha DD/MM/AA"
    COL_SEMANA   = find_col("Semana") or "Semana"
    COL_MAQUINA  = find_col("Nombre de la maquina") or "Nombre de la maquina"
    COL_TURNO    = find_col("Turno") or "Turno"
    COL_PRODUCTO = find_col("Descripción") or "Descripción"

    COL_TOTAL    = find_col("Total de Producción (pza)") or "Total de Producción (pza)"
    COL_MALA     = find_col("Producción Mala (pza)") or "Producción Mala (pza)"
    COL_T_MUERTO = find_col("Tiempo Muerto (min)") or "Tiempo Muerto (min)"
    COL_T_DISP   = find_col("Tiempo Disponible Teorico (min)") or "Tiempo Disponible Teorico (min)"
    COL_SCRAP_GR = find_col("Peso SCRAP (gr)") or "Peso SCRAP (gr)"

    COL_DISP = find_col("DISPONIBILIDAD") or "DISPONIBILIDAD"
    COL_EFIC = find_col("EFICIENCIA") or "EFICIENCIA"
    COL_CAL  = find_col("CALIDAD") or "CALIDAD"
    COL_OEE  = find_col("OEE") or "OEE"
    if COL_OEE not in df.columns and "OEE." in df.columns:
        COL_OEE = "OEE."

    needed = [
        COL_FECHA, COL_SEMANA, COL_MAQUINA, COL_TURNO, COL_PRODUCTO,
        COL_TOTAL, COL_MALA, COL_T_MUERTO, COL_T_DISP,
        COL_DISP, COL_EFIC, COL_CAL, COL_OEE
    ]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        st.error(f"Faltan columnas en el Excel: {missing}\n\nColumnas detectadas: {list(df.columns)}")
        st.stop()

    # -------------------
    # LIMPIEZA Y TIPOS
    # -------------------
    df[COL_FECHA]   = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce")
    df[COL_SEMANA]  = pd.to_numeric(df[COL_SEMANA], errors="coerce")

    df[COL_MAQUINA] = (
        df[COL_MAQUINA].astype(str).str.strip()
        .str.replace(r"\s+", " ", regex=True).str.title()
    )
    df[COL_TURNO]   = df[COL_TURNO].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df = df.dropna(subset=[COL_TURNO])
    df[COL_TURNO]   = pd.to_numeric(df[COL_TURNO], errors="coerce").astype("Int64")
    df[COL_PRODUCTO]= (
        df[COL_PRODUCTO].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    )

    # ---------------------------------------------------------------
    # FUNCIÓN PARA LIMPIAR VALORES QUE VIENEN COMO "97%", "22,5%", etc.
    # ---------------------------------------------------------------
    def clean_numeric(series: pd.Series) -> pd.Series:
        """
        Convierte columnas que pueden llegar como:
          - número puro: 0.97 o 97
          - string con %: "97%" o "97,5%"
          - string con coma decimal: "132,86"
        → siempre devuelve float.
        """
        s = series.astype(str).str.strip()
        has_pct = s.str.contains("%", na=False)
        s = s.str.replace("%", "", regex=False).str.replace(",", ".", regex=False).str.strip()
        result = pd.to_numeric(s, errors="coerce")
        # Si venían con % y el valor resultante es <= 1.5 → ya estaba en fracción, multiplicar x100
        # Si venían con % y el valor > 1.5 → ya estaba en escala 0-100, dejar igual
        # (No ajustamos aquí, se unifica abajo)
        return result, has_pct

    numeric_simple = [COL_TOTAL, COL_MALA, COL_T_MUERTO, COL_T_DISP, COL_SCRAP_GR]
    for c in numeric_simple:
        if c in df.columns:
            df[c], _ = clean_numeric(df[c])

    # Indicadores: limpiar y normalizar a escala 0-100
    for c in [COL_DISP, COL_EFIC, COL_CAL, COL_OEE]:
        df[c], had_pct = clean_numeric(df[c])
        # Si NO venían con % y el máximo es <= 1.5 → estaban en fracción 0-1
        if not had_pct.any():
            mx = df[c].max(skipna=True)
            if pd.notna(mx) and mx <= 1.5:
                df[c] = df[c] * 100
        # Si venían con % el valor ya está en 0-100 (ej. 97, 22, 100)
        # pero si alguien puso "0.97%" lo normalizamos:
        else:
            mx = df[c].max(skipna=True)
            if pd.notna(mx) and mx <= 1.5:
                df[c] = df[c] * 100

    # Métricas derivadas
    df["SCRAP_%"]          = (df[COL_MALA] / df[COL_TOTAL]) * 100
    df["SCRAP_%"]          = df["SCRAP_%"].replace([float("inf"), -float("inf")], pd.NA)
    df["TIEMPO_MUERTO_HRS"]= df[COL_T_MUERTO] / 60

    # -------------------
    # FILTROS
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
    productos = st.sidebar.multiselect(
        "Producto",
        options=sorted(df[COL_PRODUCTO].dropna().unique().tolist()),
        default=sorted(df[COL_PRODUCTO].dropna().unique().tolist())
    )

    meta_oee   = st.sidebar.number_input("Meta OEE (%)", value=80.0)
    meta_disp  = st.sidebar.number_input("Meta Disponibilidad (%)", value=80.0)
    meta_efic  = st.sidebar.number_input("Meta Eficiencia (%)", value=80.0)
    meta_cal   = st.sidebar.number_input("Meta Calidad (%)", value=97.0)
    meta_scrap = st.sidebar.number_input("Meta Scrap (%)", value=2.0)
    meta_tm    = st.sidebar.number_input("Meta Tiempo muerto semanal (hrs)", value=12.0)

    df = df[
        (df[COL_MAQUINA].isin(maquinas)) &
        (df[COL_TURNO].isin(turnos)) &
        (df[COL_PRODUCTO].isin(productos))
    ].copy()

    if df.empty:
        st.warning("No hay datos con esos filtros.")
        st.stop()

    # -------------------
    # DEBUG RÁPIDO (descomenta si necesitas verificar)
    # -------------------
    # with st.expander("🔍 Diagnóstico de columnas (debug)"):
    #     st.write(df[[COL_DISP, COL_EFIC, COL_CAL, COL_OEE]].describe())
    #     st.dataframe(df[[COL_DISP, COL_EFIC, COL_CAL, COL_OEE]].head(10))

    # -------------------
    # KPIs GENERALES
    # -------------------
    st.subheader("Indicadores generales")

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Disponibilidad", f"{df[COL_DISP].mean():.1f}%")
    k2.metric("Eficiencia",     f"{df[COL_EFIC].mean():.1f}%")
    k3.metric("Calidad",        f"{df[COL_CAL].mean():.1f}%")
    k4.metric("OEE",            f"{df[COL_OEE].mean():.1f}%")
    k5.metric("Scrap",          f"{df['SCRAP_%'].mean():.2f}%")
    k6.metric("Tiempo muerto",  f"{df['TIEMPO_MUERTO_HRS'].sum():.1f} hrs")

    # -------------------
    # AGRUPACIÓN SEMANAL
    # -------------------
    weekly = (
        df.groupby(COL_SEMANA, as_index=False)
        .agg({
            COL_DISP:           "mean",
            COL_EFIC:           "mean",
            COL_CAL:            "mean",
            COL_OEE:            "mean",
            "SCRAP_%":          "mean",
            "TIEMPO_MUERTO_HRS":"sum"
        })
        .sort_values(COL_SEMANA)
    )

    def plot_percent_line(data, y_col, title, meta=None):
        fig = px.line(data, x=COL_SEMANA, y=y_col, markers=True, title=title)
        if meta is not None:
            fig.add_hline(y=meta, line_dash="dash", annotation_text=f"Meta {meta:.1f}%")
        fig.update_layout(
            xaxis_title="Semana", yaxis_title="%",
            yaxis=dict(range=[0, 110]),          # ← 110 para que el 100% no quede pegado al borde
            margin=dict(l=10, r=10, t=50, b=10)
        )
        st.plotly_chart(fig, use_container_width=True)

    def plot_hours_line(data, y_col, title, meta=None):
        fig = px.line(data, x=COL_SEMANA, y=y_col, markers=True, title=title)
        if meta is not None:
            fig.add_hline(y=meta, line_dash="dash", annotation_text=f"Meta {meta:.1f} hrs")
        fig.update_layout(
            xaxis_title="Semana", yaxis_title="Horas",
            margin=dict(l=10, r=10, t=50, b=10)
        )
        st.plotly_chart(fig, use_container_width=True)

    # -------------------
    # TABS
    # -------------------
    tab1, tab2, tab3 = st.tabs(["Vista general", "Scrap y tiempo muerto", "Análisis por producto"])

    with tab1:
        st.subheader("Tendencia semanal de indicadores")
        c1, c2 = st.columns(2)
        with c1:
            plot_percent_line(weekly, COL_DISP, "Disponibilidad (%) por semana", meta_disp)
            plot_percent_line(weekly, COL_EFIC, "Eficiencia (%) por semana",     meta_efic)
        with c2:
            plot_percent_line(weekly, COL_CAL,  "Calidad (%) por semana",        meta_cal)
            plot_percent_line(weekly, COL_OEE,  "OEE (%) por semana",            meta_oee)

    with tab2:
        st.subheader("Scrap y tiempo muerto")
        c1, c2 = st.columns(2)
        with c1:
            plot_percent_line(weekly, "SCRAP_%",           "Histórico semanal de scrap (%)",            meta_scrap)
        with c2:
            plot_hours_line(weekly,  "TIEMPO_MUERTO_HRS", "Histórico semanal de tiempo muerto (hrs)",  meta_tm)

        st.markdown("### Detalle por semana")
        semanas_disponibles = sorted(df[COL_SEMANA].dropna().unique().tolist())
        semana_sel = st.selectbox("Selecciona una semana", semanas_disponibles, index=len(semanas_disponibles)-1)
        df_semana  = df[df[COL_SEMANA] == semana_sel].copy()

        d1, d2 = st.columns(2)
        with d1:
            scrap_maquina = (
                df_semana.groupby(COL_MAQUINA, as_index=False)["SCRAP_%"]
                .mean().sort_values("SCRAP_%", ascending=False)
            )
            fig = px.bar(scrap_maquina, x=COL_MAQUINA, y="SCRAP_%",
                         title=f"% Scrap por máquina - Semana {int(semana_sel)}", text_auto=".2f")
            fig.add_hline(y=meta_scrap, line_dash="dash", annotation_text=f"Meta {meta_scrap:.1f}%")
            fig.update_layout(yaxis_title="% Scrap", xaxis_title="Máquina")
            st.plotly_chart(fig, use_container_width=True)

        with d2:
            tm_turno = (
                df_semana.groupby(COL_TURNO, as_index=False)["TIEMPO_MUERTO_HRS"]
                .sum().sort_values("TIEMPO_MUERTO_HRS", ascending=False)
            )
            fig = px.bar(tm_turno, x=COL_TURNO, y="TIEMPO_MUERTO_HRS",
                         title=f"Tiempo muerto por turno - Semana {int(semana_sel)}", text_auto=".2f")
            fig.update_layout(yaxis_title="Horas", xaxis_title="Turno")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Histórico por turno")
        hist_turno = (
            df.groupby([COL_SEMANA, COL_TURNO], as_index=False)
            .agg({"SCRAP_%": "mean", "TIEMPO_MUERTO_HRS": "sum"})
            .sort_values([COL_SEMANA, COL_TURNO])
        )
        # Turno como string para que Plotly lo trate como categoría y use colores distintos
        hist_turno[COL_TURNO] = hist_turno[COL_TURNO].astype(str)

        h1, h2 = st.columns(2)
        with h1:
            fig = px.line(hist_turno, x=COL_SEMANA, y="SCRAP_%", color=COL_TURNO,
                          markers=True, title="Histórico de scrap por turno")
            fig.add_hline(y=meta_scrap, line_dash="dash", annotation_text=f"Meta {meta_scrap:.1f}%")
            fig.update_layout(yaxis_title="% Scrap", xaxis_title="Semana")
            st.plotly_chart(fig, use_container_width=True)

        with h2:
            fig = px.line(hist_turno, x=COL_SEMANA, y="TIEMPO_MUERTO_HRS", color=COL_TURNO,
                          markers=True, title="Histórico de tiempo muerto por turno")
            fig.update_layout(yaxis_title="Horas", xaxis_title="Semana")
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        st.subheader("Indicadores por producto")

        producto_kpis = (
            df.groupby(COL_PRODUCTO, as_index=False)
            .agg({
                COL_OEE:            "mean",
                COL_DISP:           "mean",
                COL_EFIC:           "mean",
                COL_CAL:            "mean",
                "SCRAP_%":          "mean",
                "TIEMPO_MUERTO_HRS":"sum"
            })
            .sort_values(COL_OEE, ascending=False)
        )

        p1, p2 = st.columns(2)
        with p1:
            fig = px.bar(producto_kpis, x=COL_PRODUCTO, y=COL_OEE,
                         title="OEE promedio por producto", text_auto=".1f")
            fig.add_hline(y=meta_oee, line_dash="dash", annotation_text=f"Meta {meta_oee:.1f}%")
            fig.update_layout(yaxis_title="OEE (%)", xaxis_title="Producto")
            st.plotly_chart(fig, use_container_width=True)

        with p2:
            fig = px.bar(producto_kpis, x=COL_PRODUCTO, y="SCRAP_%",
                         title="Scrap promedio por producto", text_auto=".2f")
            fig.add_hline(y=meta_scrap, line_dash="dash", annotation_text=f"Meta {meta_scrap:.1f}%")
            fig.update_layout(yaxis_title="% Scrap", xaxis_title="Producto")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Tendencia semanal por producto")
        productos_top = producto_kpis[COL_PRODUCTO].dropna().head(5).tolist()
        productos_sel = st.multiselect(
            "Selecciona productos para comparar",
            options=sorted(df[COL_PRODUCTO].dropna().unique().tolist()),
            default=productos_top
        )

        if productos_sel:
            hist_prod = (
                df[df[COL_PRODUCTO].isin(productos_sel)]
                .groupby([COL_SEMANA, COL_PRODUCTO], as_index=False)
                .agg({COL_OEE: "mean", "SCRAP_%": "mean"})
                .sort_values([COL_SEMANA, COL_PRODUCTO])
            )
            q1, q2 = st.columns(2)
            with q1:
                fig = px.line(hist_prod, x=COL_SEMANA, y=COL_OEE, color=COL_PRODUCTO,
                              markers=True, title="OEE semanal por producto")
                fig.add_hline(y=meta_oee, line_dash="dash", annotation_text=f"Meta {meta_oee:.1f}%")
                fig.update_layout(yaxis_title="OEE (%)", xaxis_title="Semana")
                st.plotly_chart(fig, use_container_width=True)
            with q2:
                fig = px.line(hist_prod, x=COL_SEMANA, y="SCRAP_%", color=COL_PRODUCTO,
                              markers=True, title="Scrap semanal por producto")
                fig.add_hline(y=meta_scrap, line_dash="dash", annotation_text=f"Meta {meta_scrap:.1f}%")
                fig.update_layout(yaxis_title="% Scrap", xaxis_title="Semana")
                st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Sube el Excel para generar el dashboard.")
