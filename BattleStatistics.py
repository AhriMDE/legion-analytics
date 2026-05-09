import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="BOD Activity Tracker v11", layout="wide")

st.sidebar.title("🛡️ Panel de Control BOD")
st.title("📊 Reporte de Actividad Battle of Dawn")
st.markdown("Motor de búsqueda universal optimizado para múltiples columnas y legiones.")

# ==========================================
# 1. FUNCIÓN DE PROCESAMIENTO (ESCÁNER UNIVERSAL)
# ==========================================
def process_bod_file(uploaded_file):
    # Leemos todo el archivo como texto para evitar errores de formato
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None, dtype=str)
    
    all_data = []
    legion_pattern = r"^LEGI[OÓ]N\s*[1-5]$"
    # Palabras clave que NO son alianzas (para evitar falsos positivos)
    invalid_alliances = {"SCORE", "RANK", "DATE", "TOTAL", "TBD", "UTC", "PLAYER", "NAN"}
    
    # Recorremos cada celda del Excel buscando legiones
    for r in range(raw_df.shape[0]):
        for c in range(raw_df.shape[1]):
            cell_val = str(raw_df.iloc[r, c]).strip().upper()
            
            if re.match(legion_pattern, cell_val):
                # ¡Encontramos una tabla de Legión!
                legion_name = str(raw_df.iloc[r, c]).strip()
                
                # Datos de cabecera
                date_val = str(raw_df.iloc[r, c+1]).strip() if c+1 < raw_df.shape[1] else "TBD"
                hour_val = str(raw_df.iloc[r, c+2]).strip() if c+2 < raw_df.shape[1] else "TBD"
                
                # Limpieza de Fecha y Hora
                date_clean = date_val.split(" ")[0].replace("-", "/") if date_val.lower() != "nan" else "TBD"
                if hour_val.lower() == "nan": 
                    hour_clean = "TBD"
                else:
                    h_num = hour_val.split(".")[0]
                    hour_clean = f"{h_num} UTC" if "UTC" not in h_num.upper() else h_num

                # Búsqueda de Alianza: Escaneamos hacia ARRIBA revisando TODA LA FILA
                alliance_name = "Default"
                for search_r in range(r - 1, -1, -1):
                    found = False
                    for search_c in range(raw_df.shape[1]):
                        potential = str(raw_df.iloc[search_r, search_c]).strip().upper()
                        # Si tiene entre 2 y 5 letras, es solo texto y no es una palabra inválida, es la Alianza
                        if 2 <= len(potential) <= 5 and potential.isalpha() and potential not in invalid_alliances:
                            alliance_name = str(raw_df.iloc[search_r, search_c]).strip().upper()
                            found = True
                            break
                    if found:
                        break
                
                # Extracción de Jugadores: Escaneamos hacia ABAJO desde r+2
                for k in range(r + 2, raw_df.shape[0]):
                    rank_cell = str(raw_df.iloc[k, c]).strip().split('.')[0]
                    player_cell = str(raw_df.iloc[k, c+1]).strip() if c+1 < raw_df.shape[1] else "nan"
                    score_cell = str(raw_df.iloc[k, c+2]).strip() if c+2 < raw_df.shape[1] else "0"
                    
                    if rank_cell.lower() == "nan" or rank_cell == "": break
                    if not rank_cell.isdigit(): continue
                    if player_cell.lower() in ["nan", "", "player", "joueur"]: continue
                    
                    try:
                        score = float(score_cell.replace(" ", "").replace(",", ""))
                    except:
                        score = 0.0
                        
                    all_data.append({
                        'Alliance': alliance_name,
                        'Date': date_clean,
                        'Schedule': hour_clean,
                        'Legion': legion_name,
                        'Player': player_cell,
                        'Score': score,
                        'Status': 'Active' if score > 0 else 'Inactive'
                    })
                    
    return pd.DataFrame(all_data)

# ==========================================
# 2. CARGA Y FILTROS
# ==========================================
uploaded_file = st.sidebar.file_uploader("📂 Subir Excel (Pestaña BOD)", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df = process_bod_file(uploaded_file)
        
        if df.empty:
            st.error("⚠️ No se detectaron datos. Revisa el formato.")
            st.stop()

        total_events = df['Date'].nunique()

        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alianzas:", alliances, default=alliances)
        
        dates = sorted(df['Date'].unique(), reverse=True)
        sel_dates = st.sidebar.multiselect("📅 Seleccionar Fechas:", dates, default=dates[:2])

        if not sel_dates:
            st.warning("Selecciona al menos una fecha para ver el resumen.")
            st.stop()

        df_filtered = df[(df['Alliance'].isin(sel_alliances)) & (df['Date'].isin(sel_dates))]

        # ==========================================
        # SECCIÓN 1: RESUMEN DE ACTIVIDAD
        # ==========================================
        st.header(f"1. Resumen Semanal")
        
        if not df_filtered.empty:
            summary = df_filtered.groupby(['Alliance', 'Legion', 'Schedule']).agg(
                Total_Players=('Player', 'count'),
                Total_Score=('Score', 'sum')
            ).reset_index()

            status_counts = df_filtered.groupby(['Alliance', 'Legion', 'Schedule', 'Status']).size().unstack(fill_value=0).reset_index()
            if 'Active' not in status_counts.columns: status_counts['Active'] = 0
            if 'Inactive' not in status_counts.columns: status_counts['Inactive'] = 0

            summary = summary.merge(status_counts[['Alliance', 'Legion', 'Schedule', 'Active', 'Inactive']], on=['Alliance', 'Legion', 'Schedule'])
            summary['% Participation'] = (summary['Active'] / summary['Total_Players']) * 100

            summary = summary[['Alliance', 'Legion', 'Schedule', 'Total_Players', 'Active', 'Inactive', 'Total_Score', '% Participation']]

            st.dataframe(
                summary.style.format({'Total_Score': '{:,.0f}', '% Participation': '{:.1f}%'}),
                use_container_width=True, hide_index=True
            )

            # --- GRÁFICA CON SELECTOR ---
            st.subheader("📊 Comparativa de Actividad")
            if len(sel_alliances) > 0:
                chart_alliance = st.selectbox("Selecciona Alianza para la gráfica:", sel_alliances)
                
                chart_data = df_filtered[df_filtered['Alliance'] == chart_alliance].groupby(['Legion', 'Status']).size().reset_index(name='Count')
                
                if not chart_data.empty:
                    fig = px.bar(
                        chart_data, x="Legion", y="Count", color="Status",
                        title=f"Jugadores Activos vs Inactivos - {chart_alliance}",
                        barmode="group",
                        color_discrete_map={'Active': '#2ECC71', 'Inactive': '#E74C3C'},
                        text_auto=True
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info(f"No hay datos para {chart_alliance}.")

        # ==========================================
        # SECCIÓN 2: RANKING DE TEMPORADA
        # ==========================================
        st.divider()
        st.header("2. Perfil de Jugadores (Acumulado)")
        
        player_base = df[df['Alliance'].isin(sel_alliances)]
        
        if not player_base.empty:
            p_ranking = player_base.groupby(['Player', 'Alliance']).agg(
                Total_Score=('Score', 'sum'),
                Attendances=('Status', lambda x: (x == 'Active').sum()),
                Favorite_Hour=('Schedule', lambda x: x.mode().iloc[0] if not x.mode().empty else "N/A")
            ).reset_index()

            p_ranking['Participation %'] = (p_ranking['Attendances'] / total_events) * 100
            p_ranking = p_ranking.sort_values('Total_Score', ascending=False)

            st.dataframe(
                p_ranking.style.format({'Total_Score': '{:,.0f}', 'Participation %': '{:.1f}%'}),
                use_container_width=True, hide_index=True
            )

        # ==========================================
        # EXPORTACIÓN
        # ==========================================
        st.sidebar.divider()
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if not df_filtered.empty: summary.to_excel(writer, sheet_name='Summary', index=False)
            if not player_base.empty: p_ranking.to_excel(writer, sheet_name='Ranking', index=False)
            
        st.sidebar.download_button("📥 Descargar Excel", buffer, "BOD_Report.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Sube tu archivo Excel para comenzar.")
