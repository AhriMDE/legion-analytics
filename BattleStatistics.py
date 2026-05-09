import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="BOD Activity Tracker", layout="wide")

st.sidebar.title("🛡️ BOD Control Panel")
st.title("📊 Battle of Dawn (BOD) Activity Report")
st.markdown("Advanced parsing logic with strict Legion detection (Anti-Legionalpha protection).")

# ==========================================
# 1. PROCESSING FUNCTION (STRICT LEGION SCANNER)
# ==========================================
def process_bod_file(uploaded_file):
    # dtype=str forces Pandas to read everything as text
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None, dtype=str)
    
    all_data = []
    current_alliance = "Default"
    current_date = "TBD"
    current_schedule = "TBD"
    current_legion = "Legion 1"
    
    # Patrones estrictos
    date_pattern = r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})"
    # ESTO ES LA MAGIA: Busca exactamente "LEGION 1" al "5", acepta la "Ó" y espacios.
    legion_pattern = r"^LEGI[OÓ]N\s*[1-5]$" 

    for i, row in raw_df.iterrows():
        # Get all non-empty cells in the row
        cells = [str(val).strip() for val in row if pd.notna(val) and str(val).strip().lower() not in ["nan", ""]]
        
        if not cells:
            continue

        row_str_upper = " ".join(cells).upper()

        # 1. Detect Alliance
        if len(cells) == 1 and 2 <= len(cells[0]) <= 5 and cells[0].isalpha() and cells[0].isupper():
            current_alliance = cells[0].upper()
            continue

        # 2. Detect Legion Header Row
        # Ahora solo se activa si alguna celda coincide EXACTAMENTE con "Legion 1", "Legion 2", etc.
        has_legion_header = any(re.match(legion_pattern, c.upper()) for c in cells)

        if has_legion_header:
            for c in cells:
                c_up = c.upper()
                clean_c = c.split(".")[0] if c.endswith(".0") else c

                if re.match(legion_pattern, c_up):
                    current_legion = c # Guarda el nombre (ej. Legion 1)
                elif re.search(date_pattern, c):
                    current_date = re.search(date_pattern, c).group(1).replace("-", "/")
                elif "UTC" in c_up:
                    current_schedule = c
                elif clean_c.isdigit(): 
                    current_schedule = f"{clean_c} UTC"
            continue

        # 3. Detect Player Data
        rank_idx = -1
        for idx, c in enumerate(cells):
            clean_c = c.split(".")[0]
            if clean_c.isdigit():
                rank_idx = idx
                break
                
        if rank_idx != -1 and len(cells) > rank_idx + 1:
            clean_rank = cells[rank_idx].split(".")[0]
            player_name = cells[rank_idx + 1]
            
            invalid_words = ["player", "joueur", "jugador", "rank", "date", "score", "utc"]
            
            # Verificamos que el jugador no sea una palabra inválida Y que NO sea exactamente un nombre de legión
            if player_name.lower() not in invalid_words and not re.match(legion_pattern, player_name.upper()):
                score = 0.0
                if len(cells) > rank_idx + 2:
                    score_str = cells[rank_idx + 2]
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
        default_dates = dates[:2] if len(dates) >= 2 else dates
        sel_dates = st.sidebar.multiselect("📅 Select Dates (Choose multiple for weekends):", dates, default=default_dates)

        if not sel_dates:
            st.warning("⚠️ Please select at least one date in the sidebar.")
            st.stop()

        df_filtered = df[(df['Alliance'].isin(sel_alliances)) & (df['Date'].isin(sel_dates))]

        # ==========================================
        # SECTION 1: WEEKLY SUMMARY
        # ==========================================
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
        
        if not player_base.empty:
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
        # DEBUG SECTION & EXPORTS
        # ==========================================
        st.divider()
        with st.expander("🛠️ Raw Data View (Debug)"):
            st.dataframe(df_filtered)
            
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
