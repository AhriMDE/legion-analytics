import streamlit as st
import pandas as pd
import plotly.express as px
import io

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="BOD Activity Tracker v9", layout="wide")

st.sidebar.title("🛡️ BOD Control Panel")
st.title("📊 Battle of Dawn (BOD) Activity Report")
st.markdown("Activity tracking optimized for exact column mapping and multiple dates.")

# ==========================================
# 1. PROCESSING FUNCTION (BULLETPROOF LOGIC)
# ==========================================
def process_bod_file(uploaded_file):
    # dtype=str forces Pandas to read everything as text
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None, dtype=str)
    
    all_data = []
    current_alliance = "Default"
    current_date = "TBD"
    current_schedule = "TBD"
    current_legion = "Legion 1"

    for i, row in raw_df.iterrows():
        # Leer estrictamente las columnas A, B y C
        col_a = str(row[0]).strip()
        col_b = str(row[1]).strip()
        col_c = str(row[2]).strip()
        
        # Ignorar filas completamente vacías
        if col_a.lower() == "nan" and col_b.lower() == "nan":
            continue

        # 1. Detectar Alianza (Col A tiene el nombre corto, Col B está vacía)
        if 2 <= len(col_a) <= 5 and col_a.isalpha() and col_a.isupper() and col_b.lower() == "nan":
            current_alliance = col_a
            continue

        # 2. Detectar Encabezado de Legión (Inmune a acentos)
        col_a_norm = col_a.upper().replace("Ó", "O").replace("É", "E")
        
        if col_a_norm.startswith("LEGION"):
            current_legion = col_a # Guarda el nombre original
            
            # Limpiar Fecha (Toma solo la parte de la fecha)
            raw_date = col_b if col_b.lower() != "nan" else "TBD"
            if len(raw_date) >= 10:
                current_date = raw_date[:10].replace("-", "/")
            else:
                current_date = raw_date
            
            # Limpiar Horario (Convierte "1.0" o "1" en "1 UTC")
            sched = col_c if col_c.lower() != "nan" else "TBD"
            if sched.endswith(".0"):
                sched = sched[:-2] 
                
            if sched.isdigit():
                sched = f"{sched} UTC"
            elif "UTC" not in sched.upper() and sched != "TBD":
                sched = f"{sched} UTC"
                
            current_schedule = sched
            continue

        # 3. Detectar Filas de Jugadores
        clean_rank = col_a.split('.')[0]
        
        invalid_headers = ["player", "joueur", "jugador", "rank", "date", "rango", "score", "puntuación", "puntuacion"]
        
        if clean_rank.isdigit() and col_b.lower() != "nan" and col_b.lower() not in invalid_headers:
            player_name = col_b
            score_str = col_c
            
            if score_str.lower() == "nan" or score_str == "":
                score = 0.0
            else:
                try:
                    clean_score = score_str.replace(" ", "").replace(",", "")
                    score = float(clean_score)
                except:
                    score = 0.0

            all_data.append({
                'Alliance': current_alliance,
                'Date': current_date,
                'Schedule': current_schedule,
                'Legion': current_legion,
                'Player': player_name,
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
            st.error("⚠️ No data could be extracted. Please check the Excel format.")
            st.stop()

        total_events = df['Date'].nunique()

        # Filters
        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alliances:", alliances, default=alliances)
        
        dates = sorted(df['Date'].unique(), reverse=True)
        # Seleccionamos por defecto las dos fechas más recientes (ideal para eventos de fin de semana)
        default_dates = dates[:2] if len(dates) >= 2 else dates
        
        # ¡CAMBIO CLAVE AQUÍ! Ahora es multiselect
        sel_dates = st.sidebar.multiselect("📅 Select Dates (Choose multiple for weekends):", dates, default=default_dates)

        # Filtramos por las fechas seleccionadas
        df_filtered = df[(df['Alliance'].isin(sel_alliances)) & (df['Date'].isin(sel_dates))]

        # ==========================================
        # SECTION 1: WEEKLY SUMMARY
        # ==========================================
        # Unimos las fechas para el título
        dates_title = ", ".join(sel_dates)
        st.header(f"1. Event Summary: {dates_title}")
        
        if not df_filtered.empty:
            # Base table
            summary = df_filtered.groupby(['Alliance', 'Legion', 'Schedule']).agg(
                Total_Players=('Player', 'count'),
                Total_Score=('Score', 'sum')
            ).reset_index()

            # Active/Inactive split
            status_summary = df_filtered.groupby(['Alliance', 'Legion', 'Schedule', 'Status']).size().unstack(fill_value=0).reset_index()
            
            if 'Active' not in status_summary.columns: status_summary['Active'] = 0
            if 'Inactive' not in status_summary.columns: status_summary['Inactive'] = 0

            summary = summary.merge(status_summary[['Alliance', 'Legion', 'Schedule', 'Active', 'Inactive']], on=['Alliance', 'Legion', 'Schedule'])
            summary['% Participation'] = (summary['Active'] / summary['Total_Players']) * 100

            # Column ordering
            summary = summary[['Alliance', 'Legion', 'Schedule', 'Total_Players', 'Active', 'Inactive', 'Total_Score', '% Participation']]

            st.dataframe(
                summary.style.format({'Total_Score': '{:,.0f}', '% Participation': '{:.1f}%'}),
                use_container_width=True, hide_index=True
            )

            # --- DYNAMIC CHART ---
            st.subheader("📊 Activity Comparison per Legion")
            
            if len(sel_alliances) > 0:
                chart_alliance = st.selectbox("🎯 Select Alliance for the chart:", sel_alliances)
                
                chart_data = df_filtered[df_filtered['Alliance'] == chart_alliance].groupby(['Legion', 'Status']).size().reset_index(name='Count')
                
                if not chart_data.empty:
                    fig = px.bar(
                        chart_data, 
                        x="Legion", 
                        y="Count", 
                        color="Status",
                        title=f"Active vs Inactive Players - {chart_alliance}",
                        labels={"Count": "Number of Players", "Legion": "Legion"},
                        barmode="group",
                        color_discrete_map={'Active': '#2ECC71', 'Inactive': '#E74C3C'},
                        text_auto=True
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info(f"No data available for {chart_alliance} on this date.")

        # ==========================================
        # SECTION 2: PLAYER RANKING (SEASON)
        # ==========================================
        st.divider()
        st.header("2. Season Player Ranking")
        
        player_base = df[df['Alliance'].isin(sel_alliances)]
        
        p_ranking = player_base.groupby(['Player', 'Alliance']).agg(
            Total_Score=('Score', 'sum'),
            Participations=('Status', lambda x: (x == 'Active').sum()),
            Favorite_Hour=('Schedule', lambda x: x.mode().iloc[0] if not x.mode().empty else "N/A")
        ).reset_index()

        p_ranking['Participation %'] = (p_ranking['Participations'] / total_events) * 100
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
            if not df_filtered.empty:
                summary.to_excel(writer, sheet_name='Event_Summary', index=False)
            p_ranking.to_excel(writer, sheet_name='Season_Ranking', index=False)
            df.to_excel(writer, sheet_name='Raw_Data', index=False)
            
        st.sidebar.download_button("📥 Download Full Report", buffer, "BOD_Activity_Report.xlsx")

    except Exception as e:
        st.error(f"Analysis error: {e}")
else:
    st.info("Upload the Excel file with the 'BOD' sheet to begin.")
