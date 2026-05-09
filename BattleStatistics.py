import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="BOD Activity Tracker v2", layout="wide")

st.sidebar.title("🛡️ BOD Control Panel")
st.title("📊 Battle of Dawn (BOD) Activity Report")
st.markdown("Seguimiento detallado de actividad por alianza, legión y jugador.")

# ==========================================
# 1. PROCESSING FUNCTION (ROBUST LOGIC)
# ==========================================
def process_bod_file(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None)
    
    all_data = []
    current_alliance = None
    current_date_range = None
    current_schedule = "TBD" # Default if not found
    expecting_alliance = False
    
    # Matches YYYY/MM/DD-YYYY/MM/DD
    date_range_pattern = r"(\d{4}/\d{1,2}/\d{1,2}-\d{4}/\d{1,2}/\d{1,2})"

    for i, row in raw_df.iterrows():
        # Clean values
        cells = [str(val).strip() for val in row if pd.notna(val)]
        row_str = " ".join(cells)
        
        if not row_str:
            continue

        # 1. Detect Date Range
        dr_match = re.search(date_range_pattern, row_str)
        if dr_match:
            current_date_range = dr_match.group(1)
            expecting_alliance = True
            continue

        # 2. Detect Alliance (The row immediately after the date)
        if expecting_alliance and cells:
            current_alliance = cells[0]
            expecting_alliance = False
            continue

        # 3. Detect Schedule (Looking for UTC)
        if "UTC" in row_str.upper():
            for c in cells:
                if "UTC" in c.upper():
                    current_schedule = c
                    break
            continue
        
        # 4. Extract Data Rows (Mapping: 0:Rank, 1:Player, 2:Score, 3:Legion)
        # We need at least Player and Legion
        if len(row) >= 4:
            rank_cell = str(row[0]).strip()
            player_cell = str(row[1]).strip()
            score_cell = str(row[2]).strip()
            legion_cell = str(row[3]).strip()

            # Verify if Rank is a number to confirm it's a data row
            if rank_cell.isdigit() and player_cell != "nan" and player_cell != "":
                
                # Score Cleanup (removes spaces in numbers like "644 645")
                try:
                    clean_score = score_cell.replace(" ", "").replace(",", "")
                    score = float(clean_score)
                except:
                    score = 0.0

                all_data.append({
                    'Alliance': current_alliance,
                    'Date_Range': current_date_range,
                    'Schedule': current_schedule,
                    'Player': player_cell,
                    'Score': score,
                    'Legion': legion_cell,
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
            st.error("⚠️ No se pudo extraer información. Verifica el formato del archivo.")
            st.stop()

        # Get total weeks for participation percentage
        total_weeks = df['Date_Range'].nunique()

        # Sidebar Filters
        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alianzas:", alliances, default=alliances)
        
        date_ranges = sorted(df['Date_Range'].unique(), reverse=True)
        sel_range = st.sidebar.selectbox("📅 Ver Semana Específica:", date_ranges)

        # Filters for the Weekly View
        df_week = df[(df['Alliance'].isin(sel_alliances)) & (df['Date_Range'] == sel_range)]

        # ==========================================
        # SECTION 1: WEEKLY SUMMARY (BY ALLIANCE/LEGION)
        # ==========================================
        st.header(f"1. Resumen Semanal: {sel_range}")
        
        if not df_week.empty:
            # Grouping stats
            summary = df_week.groupby(['Alliance', 'Legion', 'Schedule']).agg(
                Players=('Player', 'nunique'),
                Total_Score=('Score', 'sum')
            ).reset_index()

            # Active count
            active_counts = df_week[df_week['Status'] == 'Active'].groupby(['Alliance', 'Legion', 'Schedule']).size().reset_index(name='Active_Players')
            summary = summary.merge(active_counts, on=['Alliance', 'Legion', 'Schedule'], how='left').fillna(0)
            summary['% Participation'] = (summary['Active_Players'] / summary['Players']) * 100

            st.dataframe(
                summary.style.format({'Total_Score': '{:,.0f}', '% Participation': '{:.1f}%'}),
                use_container_width=True, hide_index=True
            )

            # Visuals
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("👥 Jugadores Activos por Legión")
                fig_p = px.bar(summary, x="Legion", y="Active_Players", color="Alliance", barmode="group", text_auto=True)
                st.plotly_chart(fig_p, use_container_width=True)
            with col2:
                st.subheader("🔥 Contribución de Puntos")
                fig_s = px.pie(summary, values='Total_Score', names='Alliance', hole=.4)
                st.plotly_chart(fig_s, use_container_width=True)

        # ==========================================
        # SECTION 2: CUMULATIVE PLAYER STATS (SEASON)
        # ==========================================
        st.divider()
        st.header("2. Ranking y Perfil de Jugadores (Toda la Temporada)")
        st.markdown(f"Estadísticas calculadas sobre un total de **{total_weeks}** eventos registrados.")
        
        # Individual Stats
        player_base = df[df['Alliance'].isin(sel_alliances)]
        
        # Function to get favorite hour
        def get_fav_hour(x):
            return x.mode().iloc[0] if not x.mode().empty else "N/A"

        p_ranking = player_base.groupby(['Player', 'Alliance']).agg(
            Total_Score=('Score', 'sum'),
            Participations=('Status', lambda x: (x == 'Active').sum()),
            Favorite_Hour=('Schedule', get_fav_hour)
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
            if not summary.empty:
                summary.to_excel(writer, sheet_name='Summary', index=False)
            p_ranking.to_excel(writer, sheet_name='Season_Ranking', index=False)
            
        st.sidebar.download_button("📥 Descargar Reporte Completo", buffer, "BOD_Activity_Season.xlsx")

    except Exception as e:
        st.error(f"Error en el análisis: {e}")
else:
    st.info("Sube el archivo Excel con la pestaña 'BOD' para comenzar.")
