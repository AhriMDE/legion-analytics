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
st.markdown("Simplified tracking focused on player scores and alliance activity.")

# ==========================================
# 1. PROCESSING FUNCTION (FIXED COLUMN MAPPING)
# ==========================================
def process_bod_file(uploaded_file):
    # Read the entire "BOD" sheet
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None)
    
    all_data = []
    current_alliance = None
    current_date_range = None
    current_schedule = None
    expecting_alliance = False
    
    # Matches YYYY/MM/DD-YYYY/MM/DD
    date_range_pattern = r"(\d{4}/\d{1,2}/\d{1,2}-\d{4}/\d{1,2}/\d{1,2})"

    for i, row in raw_df.iterrows():
        # Clean values to handle NaNs and spaces
        cells = [str(val).strip() for val in row]
        row_str = " ".join([c for c in cells if c != "nan"]).strip()
        
        if not row_str:
            continue

        # 1. Detect Date Range
        dr_match = re.search(date_range_pattern, row_str)
        if dr_match:
            current_date_range = dr_match.group(1)
            expecting_alliance = True
            current_schedule = None
            continue

        # 2. Detect Alliance (row below date)
        if expecting_alliance:
            for c in cells:
                if c != "nan" and c != "":
                    current_alliance = c
                    break
            expecting_alliance = False
            continue

        # 3. Detect Schedule (cell containing UTC)
        if "UTC" in row_str.upper():
            for c in cells:
                if "UTC" in c.upper():
                    current_schedule = c
                    break
            continue
        
        # 4. Skip Header Row (Rank, Player, Score, Legion)
        if cells[0].lower() == "rank" or cells[1].lower() == "player":
            continue
        
        # 5. Extract Data Rows
        # Mapping: 0:Rank, 1:Player, 2:Score, 3:Legion
        if current_alliance and len(cells) >= 4:
            rank_val = cells[0]
            player_name = cells[1]
            score_raw = cells[2]
            legion_val = cells[3]

            # Verify if it's a numeric data row (Rank is a number)
            if rank_val.isdigit() and player_name != "nan" and player_name != "":
                
                # Clean Score (handle spaces like "644 645")
                try:
                    clean_score = score_raw.replace(" ", "").replace(",", "")
                    score = pd.to_numeric(clean_score, errors='coerce')
                except:
                    score = 0

                data_row = {
                    'Alliance': current_alliance,
                    'Date_Range': current_date_range,
                    'Schedule': current_schedule if current_schedule else "TBD",
                    'Rank': int(rank_val),
                    'Player': player_name,
                    'Score': score if not pd.isna(score) else 0,
                    'Legion': legion_val
                }
                all_data.append(data_row)

    cols = ['Alliance', 'Date_Range', 'Schedule', 'Rank', 'Player', 'Score', 'Legion']
    return pd.DataFrame(all_data, columns=cols)

# ==========================================
# 2. DATA LOADING & FILTERS
# ==========================================
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel (BOD Sheet)", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df = process_bod_file(uploaded_file)
        
        if df.empty:
            st.error("⚠️ No data could be extracted. Please check the Rank and Player columns.")
            st.stop()

        # Status definition based purely on activity (Score > 0)
        df['Status'] = df['Score'].apply(lambda x: 'Active' if x > 0 else 'Inactive')

        # Sidebar Filters
        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alliances:", alliances, default=alliances)
        
        date_ranges = sorted(df['Date_Range'].unique(), reverse=True)
        sel_range = st.sidebar.selectbox("📅 Date Range:", date_ranges)

        # Apply Filters
        df_filtered = df[(df['Alliance'].isin(sel_alliances)) & (df['Date_Range'] == sel_range)]

        # ==========================================
        # SECTION 1: ACTIVITY SUMMARY
        # ==========================================
        st.header(f"1. Alliance Activity: {sel_range}")
        
        if not df_filtered.empty:
            # Stats by Alliance and Legion
            summary = df_filtered.groupby(['Alliance', 'Legion']).agg(
                Total_Players=('Player', 'nunique'),
                Total_Score=('Score', 'sum'),
                Preferred_Time=('Schedule', lambda x: x.mode().iloc[0] if not x.mode().empty else "-")
            ).reset_index()

            # Calculate Active count
            active_counts = df_filtered[df_filtered['Status'] == 'Active'].groupby(['Alliance', 'Legion']).size().reset_index(name='Active_Players')
            summary = summary.merge(active_counts, on=['Alliance', 'Legion'], how='left').fillna(0)
            summary['Participation %'] = (summary['Active_Players'] / summary['Total_Players']) * 100

            st.dataframe(
                summary.style.format({'Total_Score': '{:,.0f}', 'Participation %': '{:.1f}%'}),
                use_container_width=True, hide_index=True
            )

            # Charts
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("👥 Active Players by Legion")
                fig_p = px.bar(summary, x="Legion", y="Active_Players", color="Alliance", barmode="group", text_auto=True)
                st.plotly_chart(fig_p, use_container_width=True)
            with col2:
                st.subheader("🔥 Score Contribution")
                fig_s = px.pie(summary, values='Total_Score', names='Alliance', hole=.4)
                st.plotly_chart(fig_s, use_container_width=True)
        else:
            st.warning("No data found for the current selection.")

        # ==========================================
        # SECTION 2: PLAYER RANKING
        # ==========================================
        st.divider()
        st.header("2. Top Player Ranking (Season)")
        st.caption("Showing performance based on personal scores across all events.")
        
        # Aggregate player scores for the season
        p_ranking = df[df['Alliance'].isin(sel_alliances)].groupby(['Player', 'Alliance']).agg(
            Cumulative_Score=('Score', 'sum'),
            Average_Score=('Score', 'mean'),
            Appearances=('Status', 'count'),
            Times_Active=('Status', lambda x: (x == 'Active').sum())
        ).reset_index().sort_values('Cumulative_Score', ascending=False)

        # Show top 25 slots for rewards
        st.dataframe(
            p_ranking.head(25).style.format({'Cumulative_Score': '{:,.0f}', 'Average_Score': '{:,.1f}'}),
            use_container_width=True, hide_index=True
        )

        # ==========================================
        # SIDEBAR: EXPORTS
        # ==========================================
        st.sidebar.divider()
        st.sidebar.header("📥 Exports")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if not summary.empty:
                summary.to_excel(writer, sheet_name='Legion_Summary', index=False)
            p_ranking.to_excel(writer, sheet_name='Season_Ranking', index=False)
            df.to_excel(writer, sheet_name='Full_Activity_Log', index=False)
            
        st.sidebar.download_button("💾 Download Activity Report", buffer, "BOD_Activity_Log.xlsx")

    except Exception as e:
        st.error(f"Analysis error: {e}")
else:
    st.info("Please upload the BOD Excel file to start.")
