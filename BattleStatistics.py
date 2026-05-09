import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="BOD Activity Tracker v14", layout="wide")

st.sidebar.title("🛡️ BOD Control Panel")
st.title("📊 Battle of Dawn Activity Report")
st.markdown("Universal search engine optimized for multiple columns and legions.")

# ==========================================
# 1. PROCESSING FUNCTION (UNIVERSAL SCANNER)
# ==========================================
def process_bod_file(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None, dtype=str)
    
    all_data = []
    legion_pattern = r"^LEGI[OÓ]N\s*[1-5]$"
    invalid_alliances = {"SCORE", "RANK", "DATE", "TOTAL", "TBD", "UTC", "PLAYER", "NAN"}
    
    for r in range(raw_df.shape[0]):
        for c in range(raw_df.shape[1]):
            cell_val = str(raw_df.iloc[r, c]).strip().upper()
            
            if re.match(legion_pattern, cell_val):
                legion_name = str(raw_df.iloc[r, c]).strip()
                
                date_val = str(raw_df.iloc[r, c+1]).strip() if c+1 < raw_df.shape[1] else "TBD"
                hour_val = str(raw_df.iloc[r, c+2]).strip() if c+2 < raw_df.shape[1] else "TBD"
                
                date_clean = date_val.split(" ")[0].replace("-", "/") if date_val.lower() != "nan" else "TBD"
                if hour_val.lower() == "nan": 
                    hour_clean = "TBD"
                else:
                    h_num = hour_val.split(".")[0]
                    hour_clean = f"{h_num} UTC" if "UTC" not in h_num.upper() else h_num

                alliance_name = "Default"
                for search_r in range(r - 1, -1, -1):
                    found = False
                    for search_c in range(raw_df.shape[1]):
                        potential = str(raw_df.iloc[search_r, search_c]).strip().upper()
                        if 2 <= len(potential) <= 5 and potential.isalpha() and potential not in invalid_alliances:
                            alliance_name = str(raw_df.iloc[search_r, search_c]).strip().upper()
                            found = True
                            break
                    if found:
                        break
                
                for k in range(r + 2, raw_df.shape[0]):
                    rank_cell = str(raw_df.iloc[k, c]).strip().split('.')[0]
                    player_cell = str(raw_df.iloc[k, c+1]).strip() if c+1 < raw_df.shape[1] else "nan"
                    score_cell = str(raw_df.iloc[k, c+2]).strip() if c+2 < raw_df.shape[1] else "0"
                    
                    if rank_cell.lower() == "nan" or rank_cell == "": break
                    if not rank_cell.isdigit(): continue
                    if player_cell.lower() in ["nan", "", "player", "joueur", "jugador"]: continue
                    
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
# 2. DATA LOADING & FILTERS
# ==========================================
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel (BOD Sheet)", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df = process_bod_file(uploaded_file)
        
        if df.empty:
            st.error("⚠️ No data detected. Please check the format.")
            st.stop()

        total_events = df['Date'].nunique()
        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alliances:", alliances, default=alliances)
        
        dates = sorted(df['Date'].unique(), reverse=True)
        sel_dates = st.sidebar.multiselect("📅 Select Dates:", dates, default=dates[:1])

        if not sel_dates:
            st.warning("Select at least one date to view the summary.")
            st.stop()

        df_filtered = df[(df['Alliance'].isin(sel_alliances)) & (df['Date'].isin(sel_dates))]

        # ==========================================
        # SECTION 1: ACTIVITY SUMMARY & CHART
        # ==========================================
        st.header("1. Weekly Summary")
        
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

            st.subheader("📊 Activity Comparison")
            if len(sel_alliances) > 0:
                chart_alliance = st.selectbox("Select Alliance for the chart:", sel_alliances)
                
                chart_data = df_filtered[df_filtered['Alliance'] == chart_alliance].groupby(['Legion', 'Status']).size().reset_index(name='Count')
                
                if not chart_data.empty:
                    fig = px.bar(
                        chart_data, x="Legion", y="Count", color="Status",
                        title=f"Active vs Inactive Players - {chart_alliance}",
                        barmode="group",
                        color_discrete_map={'Active': '#2ECC71', 'Inactive': '#E74C3C'},
                        text_auto=True
                    )
                    st.plotly_chart(fig, use_container_width=True)

            # ==========================================
            # SECTION 1.5: DETAILED PLAYER TABLES (WITH TABS)
            # ==========================================
            st.divider()
            st.header("📋 Detailed Player Lists by Legion")
            
            if len(sel_alliances) > 0:
                # Create a tab for each selected alliance
                tabs = st.tabs(sel_alliances)
                
                for idx, alliance in enumerate(sel_alliances):
                    with tabs[idx]:
                        df_detailed = df_filtered[df_filtered['Alliance'] == alliance]
                        legions_list = sorted(df_detailed['Legion'].unique())
                        
                        if legions_list:
                            active_dict = {}
                            inactive_dict = {}
                            
                            for legion in legions_list:
                                legion_subset = df_detailed[df_detailed['Legion'] == legion]
                                schedule = legion_subset['Schedule'].iloc[0] if not legion_subset.empty else ""
                                
                                schedule_short = schedule.replace(" UTC", "").replace(":00:00", ":00")
                                col_name = f"{legion} ({schedule_short})"
                                
                                active_players = legion_subset[legion_subset['Status'] == 'Active']['Player'].tolist()
                                inactive_players = legion_subset[legion_subset['Status'] == 'Inactive']['Player'].tolist()
                                
                                active_dict[col_name] = active_players
                                inactive_dict[col_name] = inactive_players
                            
                            # Pad the lists so they are all the same length
                            max_active = max([len(v) for v in active_dict.values()]) if active_dict else 0
                            for k in active_dict:
                                active_dict[k].extend([""] * (max_active - len(active_dict[k])))
                                
                            max_inactive = max([len(v) for v in inactive_dict.values()]) if inactive_dict else 0
                            for k in inactive_dict:
                                inactive_dict[k].extend([""] * (max_inactive - len(inactive_dict[k])))
                            
                            df_active = pd.DataFrame(active_dict)
                            df_inactive = pd.DataFrame(inactive_dict)
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown("✅ **Active Players**")
                                st.dataframe(df_active, use_container_width=True, hide_index=True)
                            with col2:
                                st.markdown("💤 **Inactive Players**")
                                st.dataframe(df_inactive, use_container_width=True, hide_index=True)
                        else:
                            st.info(f"No detailed data available for {alliance} on the selected dates.")
            else:
                st.info("Please select at least one alliance in the sidebar.")

        # ==========================================
        # SECTION 2: SEASON RANKING
        # ==========================================
        st.divider()
        st.header("2. Season Player Ranking")
        
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
        # EXPORTS
        # ==========================================
        st.sidebar.divider()
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if not df_filtered.empty: summary.to_excel(writer, sheet_name='Summary', index=False)
            if not player_base.empty: p_ranking.to_excel(writer, sheet_name='Ranking', index=False)
            
        st.sidebar.download_button("📥 Download Excel", buffer, "BOD_Report.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload your Excel file to begin.")
