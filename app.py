# app.py
import io, zipfile
from datetime import timedelta
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode  # streamlit-aggrid
from streamlit_extras.stylable_container import stylable_container  # streamlit-extras
from geopy.distance import geodesic  # Para calcular distancias geográficas

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
        workbook = writer.book
        
        for sheet, df in dfs.items():
            # Crear una copia del DataFrame para Excel
            df_excel = df.copy()
            
            # Guardar las URLs originales antes de modificar
            urls_inicio = {}
            urls_fin = {}
            
            if 'Mapa Inicio' in df_excel.columns:
                for idx, url in enumerate(df_excel['Mapa Inicio']):
                    if pd.notna(url) and url != '':
                        urls_inicio[idx] = url
                df_excel['Mapa Inicio'] = df_excel['Mapa Inicio'].apply(
                    lambda x: '🗺️ Ver ubicación' if pd.notna(x) and x != '' else x
                )
            
            if 'Mapa Fin' in df_excel.columns:
                for idx, url in enumerate(df_excel['Mapa Fin']):
                    if pd.notna(url) and url != '':
                        urls_fin[idx] = url
                df_excel['Mapa Fin'] = df_excel['Mapa Fin'].apply(
                    lambda x: '🗺️ Ver ubicación' if pd.notna(x) and x != '' else x
                )
            
            # Escribir el DataFrame
            df_excel.to_excel(writer, index=False, sheet_name=sheet[:31])
            worksheet = writer.sheets[sheet[:31]]
            
            # Crear hipervínculos reales para las columnas de mapas
            if urls_inicio or urls_fin:
                # Encontrar las columnas de mapas
                col_inicio = None
                col_fin = None
                
                for col_idx, col_name in enumerate(df_excel.columns):
                    if col_name == 'Mapa Inicio':
                        col_inicio = col_idx
                    elif col_name == 'Mapa Fin':
                        col_fin = col_idx
                
                # Agregar hipervínculos para Mapa Inicio
                if col_inicio is not None and urls_inicio:
                    for row_idx, url in urls_inicio.items():
                        cell_row = row_idx + 1  # +1 porque la fila 0 es el encabezado
                        worksheet.write_url(cell_row, col_inicio, url, string='🗺️ Ver ubicación')
                
                # Agregar hipervínculos para Mapa Fin
                if col_fin is not None and urls_fin:
                    for row_idx, url in urls_fin.items():
                        cell_row = row_idx + 1  # +1 porque la fila 0 es el encabezado
                        worksheet.write_url(cell_row, col_fin, url, string='🗺️ Ver ubicación')
    
    buf.seek(0)
    return buf.read()

def calcular_distancia_geografica(lat1, lon1, lat2, lon2):
    """Calcula la distancia en metros entre dos puntos geográficos."""
    try:
        if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
            return None
        if lat1 == 0 and lon1 == 0 and lat2 == 0 and lon2 == 0:
            return None
        punto1 = (lat1, lon1)
        punto2 = (lat2, lon2)
        distancia = geodesic(punto1, punto2).meters
        return round(distancia, 2)
    except:
        return None

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
    
    # Columnas de geolocalización (opcionales)
    geo_cols = {}
    geo_keys = ["latitud", "longitud", "latitud fin", "longitud fin"]
    for key in geo_keys:
        if key in cols_map:
            geo_cols[key] = cols_map[key]

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
    
    # Función para generar enlaces de Google Maps
    def generar_enlace_maps(lat, lon):
        if pd.isna(lat) or pd.isna(lon) or lat == 0 or lon == 0:
            return None
        return f"https://www.google.com/maps?q={lat},{lon}"
    
    # Enlaces de Google Maps para ubicaciones de inicio y fin
    if len(geo_cols) >= 2:
        df["Mapa Inicio"] = df.apply(lambda row: generar_enlace_maps(
            row[geo_cols["latitud"]], row[geo_cols["longitud"]]
        ), axis=1)
    else:
        df["Mapa Inicio"] = None
        
    if len(geo_cols) == 4:
        df["Mapa Fin"] = df.apply(lambda row: generar_enlace_maps(
            row[geo_cols["latitud fin"]], row[geo_cols["longitud fin"]]
        ), axis=1)
    else:
        df["Mapa Fin"] = None
    
    # Distancia geográfica (si están disponibles las columnas y no hay 'Sin registro')
    if len(geo_cols) == 4:
        def calcular_distancia_con_validacion(row):
            # No calcular distancia si hay 'Sin registro'
            if pd.isna(row["Fin_dt"]) or row["Hora fin"] == "Sin registro":
                return None
            return calcular_distancia_geografica(
                row[geo_cols["latitud"]], row[geo_cols["longitud"]],
                row[geo_cols["latitud fin"]], row[geo_cols["longitud fin"]]
            )
        df["Distancia (m)"] = df.apply(calcular_distancia_con_validacion, axis=1)
    else:
        df["Distancia (m)"] = None

    # Semana ISO y Mes/Año
    iso = df["Inicio_dt"].dt.isocalendar()
    df["Semana"] = (iso.year.astype(str) + "-W" + iso.week.astype(str).str.zfill(2))
    df["Año"] = df["Inicio_dt"].dt.year
    df["Mes"] = df["Inicio_dt"].dt.to_period("M").astype(str)

    # Orden
    columnas_tabla = ["Semana","Año","Mes","Fecha",cu,cn,ca,"Hora inicio","Hora fin","Total horas","Distancia (m)","Mapa Inicio","Mapa Fin","Dur_td","Inicio_dt"]
    tabla = df[columnas_tabla].copy()
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
            "Hora inicio":[""], "Hora fin":[""], "Total horas":[subtotal],
            "Distancia (m)":[None], "Mapa Inicio":[None], "Mapa Fin":[None]
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
    # Crear una copia del DataFrame para mostrar enlaces como HTML
    df_display = df.copy()
    
    # Mantener las URLs originales para el cellRenderer
    # No modificamos las columnas aquí, el cellRenderer manejará la visualización
    
    # Configurar la tabla
    gob = GridOptionsBuilder.from_dataframe(df_display)
    gob.configure_default_column(resizable=True, filter=True, sortable=True)
    
    # Configurar columnas de mapas con cellRenderer personalizado
    cellRenderer_maps = JsCode("""
    class LinkRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            if (params.value && params.value !== '') {
                this.eGui.innerHTML = '<a href="' + params.value + '" target="_blank" style="color: #0D6EFD; text-decoration: underline; cursor: pointer;">📍 Ver mapa</a>';
            } else {
                this.eGui.innerHTML = '';
            }
        }
        getGui() {
            return this.eGui;
        }
    }
    """)
    
    if 'Mapa Inicio' in df_display.columns:
        gob.configure_column('Mapa Inicio', width=150, cellRenderer=cellRenderer_maps)
    if 'Mapa Fin' in df_display.columns:
        gob.configure_column('Mapa Fin', width=150, cellRenderer=cellRenderer_maps)
    
    # Si necesitamos resaltar subtotales, aplicamos CSS personalizado
    if bold_subtotals and 'Usuario' in df_display.columns:
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
    
    # Mostrar la tabla con HTML habilitado
    AgGrid(df_display, gridOptions=grid_options, height=height, theme="streamlit", allow_unsafe_jscode=True, enable_enterprise_modules=False)

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
    g2 = g2[["Semana","Año","Mes","Fecha","Usuario","Nombre","Apellidos","Inicio_dt","Hora inicio","Hora fin","Total horas","Distancia (m)","Mapa Inicio","Mapa Fin","Dur_td"]]
    g2 = g2.sort_values(["Fecha","Hora inicio"])
    tabla_final_f.append(g2.drop(columns=["Inicio_dt","Dur_td"]))
    total_seconds = int(g2["Dur_td"].dropna().dt.total_seconds().sum())
    subtotal = td_to_hhmmss(pd.Timedelta(seconds=total_seconds))
    tabla_final_f.append(pd.DataFrame({
        "Semana":[semana], "Año":[""], "Mes":[""], "Fecha":[""],
        "Usuario":[f"Subtotal {usuario}"], "Nombre":[""], "Apellidos":[""],
        "Hora inicio":[""], "Hora fin":[""], "Total horas":[subtotal],
        "Distancia (m)":[None], "Mapa Inicio":[None], "Mapa Fin":[None]
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
                g_out = g[["Semana","Año","Mes","Fecha","Usuario","Nombre","Apellidos","Hora inicio","Hora fin","Total horas","Distancia (m)","Mapa Inicio","Mapa Fin"]]
                
                # Guardar las URLs originales antes de modificar para Excel individual
                urls_inicio_individual = {}
                urls_fin_individual = {}
                
                if 'Mapa Inicio' in g_out.columns:
                    for idx, url in enumerate(g_out['Mapa Inicio']):
                        if pd.notna(url) and url != '':
                            urls_inicio_individual[idx] = url
                    g_out['Mapa Inicio'] = g_out['Mapa Inicio'].apply(
                        lambda x: '🗺️ Ver ubicación' if pd.notna(x) and x != '' else x
                    )
                
                if 'Mapa Fin' in g_out.columns:
                    for idx, url in enumerate(g_out['Mapa Fin']):
                        if pd.notna(url) and url != '':
                            urls_fin_individual[idx] = url
                    g_out['Mapa Fin'] = g_out['Mapa Fin'].apply(
                        lambda x: '🗺️ Ver ubicación' if pd.notna(x) and x != '' else x
                    )
                
                # subtotales por semana
                subt_list = []
                for semana, sg in g.groupby("Semana"):
                    tot_sec = int(sg["Dur_td"].dropna().dt.total_seconds().sum())
                    subt_list.append({"Usuario":u,"Semana":semana,"Subtotal":td_to_hhmmss(pd.Timedelta(seconds=tot_sec))})
                subt_df = pd.DataFrame(subt_list)
                
                # Crear Excel individual con hipervínculos reales
                buf_individual = io.BytesIO()
                with pd.ExcelWriter(buf_individual, engine="xlsxwriter") as writer:
                    workbook_individual = writer.book
                    
                    # Escribir hojas
                    g_out.to_excel(writer, index=False, sheet_name="Resumen")
                    subt_df.to_excel(writer, index=False, sheet_name="Subtotales semana")
                    
                    # Agregar hipervínculos en la hoja Resumen
                    worksheet_resumen = writer.sheets["Resumen"]
                    
                    if urls_inicio_individual or urls_fin_individual:
                        # Encontrar las columnas de mapas
                        col_inicio_ind = None
                        col_fin_ind = None
                        
                        for col_idx, col_name in enumerate(g_out.columns):
                            if col_name == 'Mapa Inicio':
                                col_inicio_ind = col_idx
                            elif col_name == 'Mapa Fin':
                                col_fin_ind = col_idx
                        
                        # Agregar hipervínculos para Mapa Inicio
                        if col_inicio_ind is not None and urls_inicio_individual:
                            for row_idx, url in urls_inicio_individual.items():
                                cell_row = row_idx + 1  # +1 porque la fila 0 es el encabezado
                                worksheet_resumen.write_url(cell_row, col_inicio_ind, url, string='🗺️ Ver ubicación')
                        
                        # Agregar hipervínculos para Mapa Fin
                        if col_fin_ind is not None and urls_fin_individual:
                            for row_idx, url in urls_fin_individual.items():
                                cell_row = row_idx + 1  # +1 porque la fila 0 es el encabezado
                                worksheet_resumen.write_url(cell_row, col_fin_ind, url, string='🗺️ Ver ubicación')
                
                buf_individual.seek(0)
                xbytes = buf_individual.read()
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
