# app.py
import io, zipfile
from datetime import timedelta
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode  # streamlit-aggrid
from streamlit_extras.stylable_container import stylable_container  # streamlit-extras

# ---------------------------- Config & Estilo ----------------------------
st.set_page_config(page_title="Control de Horas", layout="wide")
# Paleta: blanco, negro, beige + acento
PRIMARY = "#0D6EFD"   # acento (azul vivo)
BEIGE   = "#F5F0E6"
BLACK   = "#111111"
WHITE   = "#FFFFFF"

CUSTOM_CSS = f"""
<style>
.stApp {{ background: {WHITE}; }}
.block-container {{ padding-top: 1.2rem; }}
h1, h2, h3, h4 {{ color: {BLACK}; }}
.stButton>button {{
  background:{PRIMARY}; color:{WHITE}; border-radius: 12px; border: none;
}}
hr {{ border: 0; height: 1px; background: #e8e6e0; }}
.ag-theme-streamlit {{ --ag-background-color: {WHITE}; --ag-odd-row-background-color: #faf9f7; }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

st.title("⏱️ Control de Horas (por día, semana, mes)")

st.caption(
    "Sube el Excel original. La app usa la **fila 7** como encabezado. "
    "Fichajes sin salida se marcan como **Sin registro** y **no suman** en totales."
)

uploaded = st.file_uploader("Sube el fichero Excel", type=["xlsx", "xls"])

# ---------------------------- Utilidades ----------------------------
def parse_dt(x):
    return pd.to_datetime(x, errors="coerce")

def td_to_hhmmss(td):
    if pd.isna(td):
        return "Sin registro"
    total_seconds = int(td.total_seconds())
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

def to_excel_bytes(dfs: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, index=False, sheet_name=sheet[:31])
    buf.seek(0)
    return buf.read()

def build_outputs(df_raw: pd.DataFrame):
    # Map robusto de columnas
    cols_map = {c.strip().lower(): c for c in df_raw.columns}
    need = ["usuario", "nombre", "apellidos", "inicio", "fin"]
    for key in need:
        if key not in cols_map:
            raise KeyError(
                f"Falta columna '{key}'. Encontradas: {list(df_raw.columns)}"
            )
    cu, cn, ca, ci, cf = (
        cols_map["usuario"], cols_map["nombre"], cols_map["apellidos"],
        cols_map["inicio"], cols_map["fin"]
    )

    df = df_raw.copy()
    df["Inicio_dt"] = parse_dt(df[ci])
    df["Fin_dt"] = parse_dt(df[cf])

    # Invalidar 01/01/0001
    df.loc[df["Fin_dt"].notna() & (df["Fin_dt"].dt.year == 1), "Fin_dt"] = pd.NaT

    # Campos formateados
    df["Fecha"] = df["Inicio_dt"].dt.strftime("%d/%m/%Y")
    df["Hora inicio"] = df["Inicio_dt"].dt.strftime("%H:%M:%S")
    df["Hora fin"] = df["Fin_dt"].apply(lambda x: "Sin registro" if pd.isna(x) else x.strftime("%H:%M:%S"))

    # Duración
    df["Dur_td"] = df["Fin_dt"] - df["Inicio_dt"]
    df["Total horas"] = df["Dur_td"].apply(td_to_hhmmss)

    # Semana ISO y Mes/Año
    iso = df["Inicio_dt"].dt.isocalendar()
    df["Semana"] = (iso.year.astype(str) + "-W" + iso.week.astype(str).str.zfill(2))
    df["Año"] = df["Inicio_dt"].dt.year
    df["Mes"] = df["Inicio_dt"].dt.to_period("M").astype(str)

    # Orden
    tabla = df[["Semana","Año","Mes","Fecha",cu,cn,ca,"Hora inicio","Hora fin","Total horas","Dur_td","Inicio_dt"]].copy()
    tabla = tabla.sort_values(by=[cu,"Semana","Fecha","Hora inicio"]).reset_index(drop=True)

    # Subtotales (Usuario, Semana)
    bloques = []
    for (usuario, semana), g in tabla.groupby([cu,"Semana"], sort=False):
        g2 = g.drop(columns=["Dur_td","Inicio_dt"]).copy()
        bloques.append(g2)
        total_seconds = int(g["Dur_td"].dropna().dt.total_seconds().sum())
        subtotal = td_to_hhmmss(pd.Timedelta(seconds=total_seconds))
        bloques.append(pd.DataFrame({
            "Semana":[semana], "Año":[""], "Mes":[""], "Fecha":[""],
            cu:[f"Subtotal {usuario}"], cn:[""], ca:[""],
            "Hora inicio":[""], "Hora fin":[""], "Total horas":[subtotal]
        }))
    tabla_final = pd.concat(bloques, ignore_index=True)

    # Totales semana (1 fila por usuario-semana)
    tmp = tabla.dropna(subset=["Dur_td"]).copy()
    totales_semana = (tmp.groupby([cu,cn,ca,"Semana"], as_index=False)["Dur_td"].sum())
    totales_semana["Total horas semana"] = totales_semana["Dur_td"].apply(td_to_hhmmss)
    totales_semana = totales_semana.drop(columns=["Dur_td"])

    # Totales mes
    totales_mes = (tmp.groupby([cu,cn,ca,"Año","Mes"], as_index=False)["Dur_td"].sum())
    totales_mes["Total horas mes"] = totales_mes["Dur_td"].apply(td_to_hhmmss)
    totales_mes = totales_mes.drop(columns=["Dur_td"])

    # Renombres homogéneos
    tabla_final = tabla_final.rename(columns={cu:"Usuario", cn:"Nombre", ca:"Apellidos"})
    tabla_final['Nombre'] = tabla_final['Nombre'].str.upper()
    tabla_final['Apellidos'] = tabla_final['Apellidos'].str.upper()
    totales_semana = totales_semana.rename(columns={cu:"Usuario", cn:"Nombre", ca:"Apellidos"})
    totales_semana['Nombre'] = totales_semana['Nombre'].str.upper()
    totales_semana['Apellidos'] = totales_semana['Apellidos'].str.upper()
    totales_mes = totales_mes.rename(columns={cu:"Usuario", cn:"Nombre", ca:"Apellidos"})
    totales_mes['Nombre'] = totales_mes['Nombre'].str.upper()
    totales_mes['Apellidos'] = totales_mes['Apellidos'].str.upper()

    return tabla, tabla_final, totales_semana, totales_mes

def aggrid(df, height=420, bold_subtotals=False):
    # Configurar la tabla
    gob = GridOptionsBuilder.from_dataframe(df)
    gob.configure_default_column(resizable=True, filter=True, sortable=True)
    
    # Si necesitamos resaltar subtotales, aplicamos CSS personalizado
    if bold_subtotals and 'Usuario' in df.columns:
        # Agregar CSS personalizado para filas de subtotal
        st.markdown("""
        <style>
        .ag-theme-streamlit .ag-row[row-id*="Subtotal"] {
            font-weight: bold !important;
        }
        .ag-theme-streamlit .ag-cell[col-id="Usuario"][title*="Subtotal"] {
            font-weight: bold !important;
        }
        .ag-theme-streamlit .ag-cell[title*="Subtotal"] {
            font-weight: bold !important;
        }
        </style>
        """, unsafe_allow_html=True)
    
    gob.configure_grid_options(domLayout='normal')
    grid_options = gob.build()
    
    # Mostrar la tabla
    AgGrid(df, gridOptions=grid_options, height=height, theme="streamlit")

# ---------------------------- Flujo principal ----------------------------
if uploaded is None:
    st.info("Sube el Excel para comenzar.")
    st.stop()

try:
    df_raw = pd.read_excel(uploaded, header=6)  # fila 7 como encabezado
    tabla, tabla_final, tot_sem, tot_mes = build_outputs(df_raw)
    st.success("Excel cargado y procesado.")
except Exception as e:
    st.error(f"Error procesando el archivo: {e}")
    st.stop()

# Filtros (multiselección)
# Crear nombres completos únicos
tabla["Nombre_Completo"] = tabla["Nombre"] + " " + tabla["Apellidos"]
nombres_completos = sorted(tabla["Nombre_Completo"].dropna().unique().tolist())
semanas = sorted(tabla["Semana"].dropna().unique().tolist())

st.markdown("### 🔍 Filtros")
col_f1, col_f2 = st.columns([2,1])
with col_f1:
    sel_nombres_completos = st.multiselect("Filtrar por nombre y apellidos", nombres_completos, default=nombres_completos)
with col_f2:
    sel_weeks = st.multiselect("Filtrar semanas (ISO)", semanas, default=semanas)

mask = (tabla["Nombre_Completo"].isin(sel_nombres_completos) & 
        tabla["Semana"].isin(sel_weeks))
tabla_f = tabla.loc[mask].copy()

# Resumen filtrado con subtotales
tabla_final_f = []
for (usuario, semana), g in tabla_f.groupby(["Usuario","Semana"]):
    g2 = g.copy()
    g2["Total horas"] = g2["Dur_td"].apply(td_to_hhmmss)
    g2 = g2[["Semana","Año","Mes","Fecha","Usuario","Nombre","Apellidos","Inicio_dt","Hora inicio","Hora fin","Total horas","Dur_td"]]
    g2 = g2.sort_values(["Fecha","Hora inicio"])
    tabla_final_f.append(g2.drop(columns=["Inicio_dt","Dur_td"]))
    total_seconds = int(g2["Dur_td"].dropna().dt.total_seconds().sum())
    subtotal = td_to_hhmmss(pd.Timedelta(seconds=total_seconds))
    tabla_final_f.append(pd.DataFrame({
        "Semana":[semana], "Año":[""], "Mes":[""], "Fecha":[""],
        "Usuario":[f"Subtotal {usuario}"], "Nombre":[""], "Apellidos":[""],
        "Hora inicio":[""], "Hora fin":[""], "Total horas":[subtotal]
    }))
tabla_final_f = pd.concat(tabla_final_f, ignore_index=True) if tabla_final_f else pd.DataFrame()

# Solo aplicar transformaciones si el DataFrame no está vacío y tiene las columnas necesarias
if not tabla_final_f.empty and 'Nombre' in tabla_final_f.columns:
    tabla_final_f['Nombre'] = tabla_final_f['Nombre'].str.upper()
if not tabla_final_f.empty and 'Apellidos' in tabla_final_f.columns:
    tabla_final_f['Apellidos'] = tabla_final_f['Apellidos'].str.upper()

# Totales filtrados
tmp_f = tabla_f.dropna(subset=["Dur_td"]).copy()
tot_sem_f = (tmp_f.groupby(["Usuario","Nombre","Apellidos","Semana"], as_index=False)["Dur_td"].sum())
tot_sem_f["Total horas semana"] = tot_sem_f["Dur_td"].apply(td_to_hhmmss)
tot_sem_f = tot_sem_f.drop(columns=["Dur_td"])
tot_sem_f['Nombre'] = tot_sem_f['Nombre'].str.upper()
tot_sem_f['Apellidos'] = tot_sem_f['Apellidos'].str.upper()

tot_mes_f = (tmp_f.groupby(["Usuario","Nombre","Apellidos","Año","Mes"], as_index=False)["Dur_td"].sum())
tot_mes_f["Total horas mes"] = tot_mes_f["Dur_td"].apply(td_to_hhmmss)
tot_mes_f = tot_mes_f.drop(columns=["Dur_td"])
tot_mes_f['Nombre'] = tot_mes_f['Nombre'].str.upper()
tot_mes_f['Apellidos'] = tot_mes_f['Apellidos'].str.upper()

# ---------------------------- UI: Tablas ----------------------------
st.subheader("📋 Resumen por día (con subtotales semanales)")
with stylable_container(
    key="card1",
    css_styles="""
        {
          background: #F5F0E6;
          border: 1px solid #e5e2db;
          border-radius: 16px;
          padding: 1rem 1.2rem;
          box-shadow: 0 2px 14px rgba(0,0,0,0.05);
        }
    """
):
    if not tabla_final_f.empty:
        # oculto Año y Mes para simplificar
        df_show = tabla_final_f.drop(columns=["Año","Mes"], errors="ignore")
        aggrid(df_show, 420, bold_subtotals=True)
    else:
        st.info("Sin filas para los filtros seleccionados.")

st.subheader("🗓️ Totales por semana")
with stylable_container(
    key="card2",
    css_styles="""
        {
          background: #F5F0E6;
          border: 1px solid #e5e2db;
          border-radius: 16px;
          padding: 1rem 1.2rem;
          box-shadow: 0 2px 14px rgba(0,0,0,0.05);
        }
    """
):
    aggrid(tot_sem_f, 350)

st.subheader("🗓️ Totales por mes")
with stylable_container(
    key="card3",
    css_styles="""
        {
          background: #F5F0E6;
          border: 1px solid #e5e2db;
          border-radius: 16px;
          padding: 1rem 1.2rem;
          box-shadow: 0 2px 14px rgba(0,0,0,0.05);
        }
    """
):
    aggrid(tot_mes_f, 350)

# ---------------------------- UI: Gráficos ----------------------------
st.subheader("📈 Gráficos")
gcol1, gcol2 = st.columns(2)

# Horas por día
with gcol1, stylable_container(
    key="card4",
    css_styles="""
        {
          background: #F5F0E6;
          border: 1px solid #e5e2db;
          border-radius: 16px;
          padding: 1rem 1.2rem;
          box-shadow: 0 2px 14px rgba(0,0,0,0.05);
        }
    """
):
    daily = tmp_f.copy()
    if not daily.empty:
        daily["Fecha_dt"] = daily["Inicio_dt"].dt.date
        daily["Horas"] = (daily["Dur_td"].dt.total_seconds()/3600).round(0).astype(int)
        daily_total = daily.groupby('Fecha_dt', as_index=False)['Horas'].sum()
        fig = px.bar(daily_total, x="Fecha_dt", y="Horas",
                     title="Horas totales por día (filtrado)",
                     labels={"Fecha_dt":"Fecha","Horas":"Horas totales"})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sin datos para gráfico diario.")

# Top personas (total horas en rango de filtros)
with gcol2, stylable_container(
    key="card5",
    css_styles="""
        {
          background: #F5F0E6;
          border: 1px solid #e5e2db;
          border-radius: 16px;
          padding: 1rem 1.2rem;
          box-shadow: 0 2px 14px rgba(0,0,0,0.05);
        }
    """
):
    top = (tmp_f.groupby(["Usuario", "Nombre", "Apellidos"], as_index=False)["Dur_td"].sum())
    if not top.empty:
        top["Horas"] = (top["Dur_td"].dt.total_seconds()/3600).round(0).astype(int)
        # Crear nombre completo para mostrar en el gráfico
        top["Nombre_Apellido"] = top["Nombre"] + " " + top["Apellidos"].str.split().str[0]
        top = top.sort_values("Horas", ascending=False).head(15)
        fig2 = px.bar(top, x="Nombre_Apellido", y="Horas", title="Top personas por horas (rango filtrado)",
                     hover_data={"Usuario": True, "Nombre_Apellido": False})
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Sin datos para Top personas.")

# Horas por semana
with stylable_container(
    key="card6",
    css_styles="""
        {
          background: #F5F0E6;
          border: 1px solid #e5e2db;
          border-radius: 16px;
          padding: 1rem 1.2rem;
          box-shadow: 0 2px 14px rgba(0,0,0,0.05);
        }
    """
):
    by_week = (tmp_f.groupby(["Semana","Usuario"], as_index=False)["Dur_td"].sum())
    if not by_week.empty:
        by_week["Horas"] = (by_week["Dur_td"].dt.total_seconds()/3600).round(0).astype(int)
        by_week_total = by_week.groupby('Semana', as_index=False)['Horas'].sum()
        fig3 = px.bar(by_week_total, x="Semana", y="Horas",
                      title="Horas totales por semana", barmode="group")
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Sin datos para gráfico semanal.")

# Horas por mes
with stylable_container(
    key="card7",
    css_styles="""
        {
          background: #F5F0E6;
          border: 1px solid #e5e2db;
          border-radius: 16px;
          padding: 1rem 1.2rem;
          box-shadow: 0 2px 14px rgba(0,0,0,0.05);
        }
    """
):
    by_month = (tmp_f.groupby(["Año","Mes","Usuario"], as_index=False)["Dur_td"].sum())
    if not by_month.empty:
        by_month["Horas"] = (by_month["Dur_td"].dt.total_seconds()/3600).round(0).astype(int)
        by_month["Periodo"] = by_month["Mes"]
        by_month_total = by_month.groupby('Periodo', as_index=False)['Horas'].sum()
        fig4 = px.bar(by_month_total, x="Periodo", y="Horas",
                      title="Horas totales por mes", barmode="group")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Sin datos para gráfico mensual.")

# ---------------------------- Descargas ----------------------------
st.subheader("📥 Descargas")
col_d1, col_d2 = st.columns([1,1])

with col_d1:
    # Excel global (Resumen + Totales semana + Totales mes)
    dfs_global = {
        "Resumen": tabla_final_f if not tabla_final_f.empty else pd.DataFrame(),
        "Totales semana": tot_sem_f,
        "Totales mes": tot_mes_f,
    }
    st.download_button(
        "Descargar Excel (global filtrado)",
        data=to_excel_bytes(dfs_global),
        file_name="horas_resumen_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col_d2:
    # ZIP con un Excel por usuario filtrado
    usuarios_filtrados = tabla_f["Usuario"].dropna().unique().tolist()
    if usuarios_filtrados:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for u in usuarios_filtrados:
                g = tabla_f[tabla_f["Usuario"] == u].copy()
                if g.empty:
                    continue
                g["Total horas"] = g["Dur_td"].apply(td_to_hhmmss)
                g_out = g[["Semana","Año","Mes","Fecha","Usuario","Nombre","Apellidos","Hora inicio","Hora fin","Total horas"]]
                # subtotales por semana
                subt_list = []
                for semana, sg in g.groupby("Semana"):
                    tot_sec = int(sg["Dur_td"].dropna().dt.total_seconds().sum())
                    subt_list.append({"Usuario":u,"Semana":semana,"Subtotal":td_to_hhmmss(pd.Timedelta(seconds=tot_sec))})
                subt_df = pd.DataFrame(subt_list)
                xbytes = to_excel_bytes({"Resumen": g_out, "Subtotales semana": subt_df})
                zf.writestr(f"{u.replace('@','_at_')}.xlsx", xbytes)
        zip_buf.seek(0)
        st.download_button(
            "Descargar ZIP (un Excel por trabajador)",
            data=zip_buf.getvalue(),
            file_name="horas_por_trabajador.zip",
            mime="application/zip",
        )
    else:
        st.info("No hay usuarios en los datos filtrados para generar el ZIP por trabajador.")
