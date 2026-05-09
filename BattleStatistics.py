import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="BOD Activity Tracker v6", layout="wide")

st.sidebar.title("🛡️ BOD Control Panel")
st.title("📊 Battle of Dawn (BOD) Activity Report")
st.markdown("Activity tracking optimized for the new simplified Excel format.")

# ==========================================
# 1. PROCESSING FUNCTION (NEW FORMAT LOGIC)
# ==========================================
def process_bod_file(uploaded_file):
    # dtype=str prevents Excel from turning Rank 1 into 1.0, keeping everything as plain text
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None, dtype=str)
    
    all_data = []
    current_alliance = "Default"
    current_date = "TBD"
    current_schedule = "TBD"
    current_legion = "Legion 1"
    
    date_pattern = r"(\d{4}/\d{1,2}/\d{1,2})"

    for i, row in raw_df.iterrows():
        # Clean and extract non-empty cells
        cells = [str(val).strip() for val in row if pd.notna(val) and str(val).strip().lower() != "nan"]
        
        if not cells:
            continue

        row_str = " ".join(cells).upper()

        # 1. Detect Alliance (If a cell is just "DL" or "KUT" alone in a row)
        if len(cells) == 1 and 2 <= len(cells[0]) <= 5 and cells[0].isalpha():
            current_alliance = cells[0].upper()
            continue

        # 2. Detect Legion, Date, and Schedule in the new header format
        if "LEGION" in row_str:
            for c in cells:
                c_up = c.upper()
                if "LEGION" in c_up:
                    current_legion = c
                elif re.search(date_pattern, c):
                    current_date = re.search(date_pattern, c).group(1)
                elif "UTC" in c_up:
                    current_schedule = c
            continue # Move to the next row (the table headers)

        # 3. Extract Player Data Rows
        if len(cells) >= 3:
            rank_cell = cells[0].split('.')[0] # Clean hidden decimals like "1.0"
            player_cell = cells[1]
            score_cell = cells[2]

            invalid_words = ["player", "joueur", "rank", "date"]
            
            # If Col A is a number and Col B is not a header word
            if rank_cell.isdigit() and player_cell.lower() not in invalid_words:
                try:
                    clean_score = score_cell.replace(" ", "").replace(",", "")
                    score = float(clean_score)
                except:
                    score = 0.0

                all_data.append({
                    'Alliance': current_alliance,
                    'Date': current_date,
                    'Schedule': current_schedule,
                    'Legion': current_legion,
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
            st.error("⚠️ No data could be extracted. Please check the Excel format.")
            st.stop()

        total_events = df['Date'].nunique()

        # Filters
        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alliances:", alliances, default=alliances)
        
        dates = sorted(df['Date'].unique(), reverse=True)
        sel_date = st.sidebar.selectbox("📅 Select Date:", dates)

        df_filtered = df[(df['Alliance'].isin(sel_alliances)) & (df['Date'] == sel_date)]

        # ==========================================
        # SECTION 1: WEEKLY SUMMARY
        # ==========================================
        st.header(f"1. Event Summary: {sel_date}")
        
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

