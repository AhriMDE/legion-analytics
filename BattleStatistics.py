import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="BOD Activity Tracker v5", layout="wide")

st.sidebar.title("🛡️ BOD Control Panel")
st.title("📊 Battle of Dawn (BOD) Activity Report")
st.markdown("Seguimiento de actividad, participación y legiones.")

# ==========================================
# 1. PROCESSING FUNCTION (BULLETPROOF TEXT LOGIC)
# ==========================================
def process_bod_file(uploaded_file):
    # dtype=str obliga al programa a leer todo exactamente como está en Excel (evita el error 1.0)
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None, dtype=str)
    
    all_data = []
    current_alliance = None
    current_date_range = None
    current_schedule = "TBD"
    legion_counter = 0
    expecting_alliance = False
    
    # Busca "2026/05/09 - 2026/05/10" tolerando espacios
    date_range_pattern = r"(\d{4}/\d{1,2}/\d{1,2}\s*-\s*\d{4}/\d{1,2}/\d{1,2})"

    for i, row in raw_df.iterrows():
        # Limpiar celdas ignorando los valores "nan"
        cells = [str(val).strip() for val in row if pd.notna(val) and str(val).strip().lower() != "nan"]
        row_str = " ".join(cells)
        
        if not row_str:
            continue

        # 1. Detectar Fecha
        dr_match = re.search(date_range_pattern, row_str)
        if dr_match:
            current_date_range = dr_match.group(1).replace(" ", "")
            expecting_alliance = True
            legion_counter = 0 # Reinicia las legiones por cada bloque de fecha
            continue

        # 2. Detectar Alianza
        if expecting_alliance and cells:
            current_alliance = cells[0]
            expecting_alliance = False
            continue

        # 3. Detectar Horario (Busca UTC en cualquier parte de la fila)
        if "UTC" in row_str.upper():
            legion_counter += 1
            # Extrae la celda exacta que tiene la palabra UTC
            utc_cells = [c for c in cells if "UTC" in c.upper()]
            current_schedule = utc_cells[0] if utc_cells else row_str
            continue
        
        # 4. Extraer Filas de Jugadores (0:Rank, 1:Player, 2:Score)
        if len(row) >= 3:
            # Forzamos a string y quitamos espacios
            rank_cell = str(row[0]).strip()
            player_cell = str(row[1]).strip()
            score_cell = str(row[2]).strip()

            # Evita encabezados o celdas vacías
            invalid_words = ["nan", "", "player", "joueur", "rank", "date"]
            
            # Limpiamos el rank por si Excel le puso un decimal oculto (ej. "1.0" -> "1")
            clean_rank = rank_cell.split('.')[0]

            if clean_rank.isdigit() and player_cell.lower() not in invalid_words:
                # Es un jugador válido. Procesar Score.
                try:
                    clean_score = score_cell.replace(" ", "").replace(",", "")
                    score = float(clean_score)
                except:
                    score = 0.0

                all_data.append({
                    'Alliance': current_alliance,
                    'Date_Range': current_date_range,
                    'Schedule': current_schedule,
                    'Legion': f"Legion {legion_counter}",
                    'Player': player_cell,
                    'Score': score,
                    'Status': 'Active' if score > 0 else 'Inactive'
                })

    return pd.DataFrame(all_data)

# ==========================================
# 2. DATA LOADING & FILTERS
# ==========================================
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel (BOD Sheet)", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df = process_bod_file(uploaded_file)
        
        if df.empty:
            st.error("⚠️ No se pudo extraer información. Verifica el formato de tu Excel.")
            st.stop()

        total_weeks = df['Date_Range'].nunique()

        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alianzas globales:", alliances, default=alliances)
        
        date_ranges = sorted(df['Date_Range'].unique(), reverse=True)
        sel_range = st.sidebar.selectbox("📅 Ver Semana Específica:", date_ranges)

        df_week = df[(df['Alliance'].isin(sel_alliances)) & (df['Date_Range'] == sel_range)]

        # ==========================================
        # SECTION 1: WEEKLY SUMMARY
        # ==========================================
        st.header(f"1. Resumen Semanal: {sel_range}")
        
        if not df_week.empty:
            # Agrupación base para la tabla
            summary = df_week.groupby(['Alliance', 'Legion', 'Schedule']).agg(
                Total_Players=('Player', 'count'),
                Total_Score=('Score', 'sum')
            ).reset_index()

            # Conteo de Activos e Inactivos
            status_summary = df_week.groupby(['Alliance', 'Legion', 'Schedule', 'Status']).size().unstack(fill_value=0).reset_index()
            
            if 'Active' not in status_summary.columns: status_summary['Active'] = 0
            if 'Inactive' not in status_summary.columns: status_summary['Inactive'] = 0

            summary = summary.merge(status_summary[['Alliance', 'Legion', 'Schedule', 'Active', 'Inactive']], on=['Alliance', 'Legion', 'Schedule'])
            summary['% Participation'] = (summary['Active'] / summary['Total_Players']) * 100

            # Orden de las columnas
            summary = summary[['Alliance', 'Legion', 'Schedule', 'Total_Players', 'Active', 'Inactive', 'Total_Score', '% Participation']]

            st.dataframe(
                summary.style.format({'Total_Score': '{:,.0f}', '% Participation': '{:.1f}%'}),
                use_container_width=True, hide_index=True
            )

            # --- GRÁFICA CON SELECTOR DE ALIANZA ---
            st.subheader("📊 Comparativa de Actividad por Legión")
            
            # Si hay más de una alianza en el filtro, mostramos el selector
            if len(sel_alliances) > 0:
                chart_alliance = st.selectbox("🎯 Selecciona la Alianza para ver en la gráfica:", sel_alliances)
                
                chart_data = df_week[df_week['Alliance'] == chart_alliance].groupby(['Legion', 'Status']).size().reset_index(name='Count')
                
                if not chart_data.empty:
                    fig = px.bar(
                        chart_data, 
                        x="Legion", 
                        y="Count", 
                        color="Status",
                        title=f"Jugadores Activos vs Inactivos - {chart_alliance}",
                        labels={"Count": "Número de Jugadores", "Legion": "Legión"},
                        barmode="group",
                        color_discrete_map={'Active': '#2ECC71', 'Inactive': '#E74C3C'}
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info(f"No hay datos de la alianza {chart_alliance} en esta semana.")

        # ==========================================
        # SECTION 2: PLAYER RANKING (SEASON)
        # ==========================================
        st.divider()
        st.header("2. Ranking y Perfil de Jugadores (Temporada)")
        
        player_base = df[df['Alliance'].isin(sel_alliances)]
        
        p_ranking = player_base.groupby(['Player', 'Alliance']).agg(
            Total_Score=('Score', 'sum'),
            Participations=('Status', lambda x: (x == 'Active').sum()),
            Favorite_Hour=('Schedule', lambda x: x.mode().iloc[0] if not x.mode().empty else "N/A")
        ).reset_index()

        p_ranking['Participation %'] = (p_ranking['Participations'] / total_weeks) * 100
        p_ranking = p_ranking.sort_values('Total_Score', ascending=False)

        st.dataframe(
            p_ranking.style.format({'Total_Score': '{:,.0f}', 'Participation %': '{:.1f}%'}),
            use_container_width=True, hide_index=True
        )

        # ==========================================
        # EXPORTS
        # ==========================================
        st.sidebar.divider()
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if not df_week.empty:
                summary.to_excel(writer, sheet_name='Weekly_Summary', index=False)
            p_ranking.to_excel(writer, sheet_name='Season_Ranking', index=False)
            
        st.sidebar.download_button("📥 Descargar Reporte", buffer, "BOD_Activity_Report.xlsx")

    except Exception as e:
        st.error(f"Error en el análisis: {e}")
else:
    st.info("Sube el archivo Excel con la pestaña 'BOD' para comenzar.")


