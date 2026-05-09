import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="BOD Analytics - Multi Alliance", layout="wide")

st.sidebar.title("🛡️ BOD Control Panel")
st.title("📊 Battle of Dawn (BOD) Report")
st.markdown("Advanced analytics for multiple alliances based on the simplified log format.")

# ==========================================
# 1. PROCESSING FUNCTION (SIMPLIFIED LOGIC)
# ==========================================
def process_bod_file(uploaded_file):
    # Read the entire "BOD" sheet without headers
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None)
    
    all_data = []
    current_alliance = None
    current_date_range = None
    current_schedule = None
    expecting_alliance = False
    
    # Matches exactly "YYYY/MM/DD-YYYY/MM/DD"
    date_range_pattern = r"(\d{4}/\d{1,2}/\d{1,2}-\d{4}/\d{1,2}/\d{1,2})"

    for i, row in raw_df.iterrows():
        # Convert row to string to search
        row_str = " ".join([str(val) for val in row if pd.notna(val)]).strip()
        
        # Skip completely empty rows
        if not row_str:
            continue

        # 1. If we found the date range in the previous step, this row is the Alliance
        if expecting_alliance:
            first_val = [str(val) for val in row if pd.notna(val)][0]
            current_alliance = str(first_val).strip()
            expecting_alliance = False
            continue

        # 2. Check for Date Range
        dr_match = re.search(date_range_pattern, row_str)
        if dr_match:
            current_date_range = dr_match.group(1)
            expecting_alliance = True # The next row will contain the alliance name
            current_schedule = None   # Reset schedule for the new block
            continue
            
        # 3. Check for Schedule (e.g., "10 May 1 UTC")
        if "UTC" in row_str.upper():
            # Find the exact cell with UTC to avoid extra blank spaces
            for val in row:
                if pd.notna(val) and "UTC" in str(val).upper():
                    current_schedule = str(val).strip()
                    break
            continue
        
        # 4. Extract Data Rows
        # We assume data columns: 0:Date, 1:Hour, 2:Player, 3:Score, 4:Result, 5:Legion
        if current_alliance and len(row) > 5:
            col_date = str(row[0]).strip()
            col_player = str(row[2]).strip()
            col_legion = str(row[5]).strip()
            
            # Check if it's a real data row (Date column starts with a number, Player is not empty)
            if col_date[0:1].isdigit() and pd.notna(row[2]) and col_player != "" and col_player.lower() != "joueur":
                
                # Use the UTC schedule if found, otherwise fallback to the row's hour column
                final_hour = current_schedule if current_schedule else str(row[1])

                data_row = {
                    'Alliance': current_alliance,
                    'Date_Range': current_date_range,
                    'Date': col_date[:10],
                    'Hour': final_hour,     
                    'Player': col_player,   
                    'Score': row[3],
                    'Result': str(row[4]),
                    'Legion': col_legion
                }
                all_data.append(data_row)

    cols = ['Alliance', 'Date_Range', 'Date', 'Hour', 'Player', 'Score', 'Result', 'Legion']
    return pd.DataFrame(all_data, columns=cols)

# ==========================================
# 2. DATA LOADING & FILTERS
# ==========================================
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel (BOD Sheet)", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        df = process_bod_file(uploaded_file)
        
        if df.empty:
            st.error("⚠️ No data could be extracted. Check your format: Date range in one cell, Alliance directly below it.")
            st.stop()

        # --- ADDITIONAL CLEANING ---
        def clean_hour(val):
            val_str = str(val).strip()
            if 'UTC' in val_str.upper():
                return val_str # Keeps "10 May 1 UTC" intact
            if 'H' in val_str.upper():
                num = ''.join(filter(str.isdigit, val_str))
                if num:
                    return f"{int(num):02d}:00"
            return val_str
            
        df['Hour'] = df['Hour'].apply(clean_hour)
        
        # Clean Score
        if df['Score'].dtype == 'object':
            df['Score'] = df['Score'].astype(str).str.replace(',', '').str.replace('.', '')
        df['Score'] = pd.to_numeric(df['Score'], errors='coerce').fillna(0)
        
        # Define Status and Is_Win helper
        df['Status'] = df['Score'].apply(lambda x: 'Active' if x > 0 else 'Inactive')
        df['Is_Win'] = df['Result'].astype(str).apply(lambda x: 1 if 'Victory' in x.strip() else 0)

        # Sidebar Filters
        alliances = sorted(df['Alliance'].unique())
        sel_alliances = st.sidebar.multiselect("🛡️ Alliances:", alliances, default=alliances)
        
        date_ranges = sorted(df['Date_Range'].unique(), reverse=True)
        sel_range = st.sidebar.selectbox("📅 Date Range:", date_ranges)

        # Apply Filters
        df_filtered = df[(df['Alliance'].isin(sel_alliances)) & (df['Date_Range'] == sel_range)]

        # --- SAFETY NETS ---
        summary = pd.DataFrame()
        p_stats = pd.DataFrame()

        # ==========================================
        # SECTION 1: EVENT SUMMARY
        # ==========================================
        st.header(f"1. Battle Summary: {sel_range}")
        
        if not df_filtered.empty:
            summary = df_filtered.groupby(['Alliance', 'Legion']).agg(
                Players=('Player', 'nunique'),
                Total_Score=('Score', 'sum'),
                Result=('Result', lambda x: x.mode().iloc[0] if not x.mode().empty else "-"),
                Schedule=('Hour', lambda x: x.mode().iloc[0] if not x.mode().empty else "-")
            ).reset_index()

            active_counts = df_filtered[df_filtered['Status'] == 'Active'].groupby(['Alliance', 'Legion']).size().reset_index(name='Active_Players')
            summary = summary.merge(active_counts, on=['Alliance', 'Legion'], how='left').fillna(0)
            summary['% Part.'] = (summary['Active_Players'] / summary['Players']) * 100

            st.dataframe(
                summary.style.format({'Total_Score': '{:,.0f}', '% Part.': '{:.1f}%'})
                .background_gradient(subset=['% Part.'], cmap='RdYlGn'),
                use_container_width=True, hide_index=True
            )

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("👥 Participation by Alliance")
                fig_p = px.bar(summary, x="Legion", y="Active_Players", color="Alliance", barmode="group", text_auto=True)
                st.plotly_chart(fig_p, use_container_width=True)
            with col2:
                st.subheader("🔥 Top Scores by Alliance")
                fig_s = px.pie(summary, values='Total_Score', names='Alliance', hole=.4)
                st.plotly_chart(fig_s, use_container_width=True)
        else:
            st.warning("No data found for the selected alliance and date range.")

        # ==========================================
        # SECTION 2: SCHEDULE PERFORMANCE
        # ==========================================
        st.divider()
        st.header("2. Effective Schedule Analysis")
        
        match_data = df[['Alliance', 'Hour', 'Is_Win', 'Date_Range']].drop_duplicates()
        
        if not match_data.empty:
            hourly_win = match_data.groupby('Hour').agg(
                Matches=('Is_Win', 'count'),
                Wins=('Is_Win', 'sum')
            ).reset_index()
            hourly_win['Winrate'] = (hourly_win['Wins'] / hourly_win['Matches']) * 100

            fig_h = px.bar(hourly_win, x='Hour', y='Winrate', color='Winrate', 
                           title="Win Probability by Time of Day",
                           color_continuous_scale='RdYlGn', text_auto='.1f')
            st.plotly_chart(fig_h, use_container_width=True)

        # ==========================================
        # SECTION 3: PLAYER DETAILS
        # ==========================================
        st.divider()
        st.header("3. Player Ranking (Season)")
        
        if not df[df['Alliance'].isin(sel_alliances)].empty:
            p_stats = df[df['Alliance'].isin(sel_alliances)].groupby(['Player', 'Alliance']).agg(
                Total_Points=('Score', 'sum'),
                Average_Score=('Score', 'mean'),
                Attendances=('Status', lambda x: (x == 'Active').sum())
            ).reset_index().sort_values('Total_Points', ascending=False)

            st.dataframe(p_stats.head(20), use_container_width=True, hide_index=True)

        # ==========================================
        # SIDEBAR: EXPORTS
        # ==========================================
        st.sidebar.divider()
        st.sidebar.header("📥 Download Reports")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if not summary.empty:
                summary.to_excel(writer, sheet_name='Summary', index=False)
            if not p_stats.empty:
                p_stats.to_excel(writer, sheet_name='Player_Ranking', index=False)
            df.to_excel(writer, sheet_name='Raw_Data', index=False)
            
        st.sidebar.download_button("💾 Download BOD Report", buffer, "BOD_Report.xlsx")

    except Exception as e:
        st.error(f"Critical error: {e}. Check formatting rules.")
else:
    st.info("Upload the Excel file with the 'BOD' sheet to begin the analysis.")
