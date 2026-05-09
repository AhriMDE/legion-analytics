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
st.markdown("Advanced analytics for multiple alliances based on the new log format.")

# ==========================================
# 1. PROCESSING FUNCTION (BULLETPROOF LOGIC)
# ==========================================
def process_bod_file(uploaded_file):
    # Read the entire "BOD" sheet without headers to process it manually
    raw_df = pd.read_excel(uploaded_file, sheet_name="BOD", header=None)
    
    all_data = []
    current_alliance = None
    current_date_range = None
    
    # Nuevo buscador: Tolera mayúsculas (To/TO/to), espacios variables y nombres de alianza largos
    header_pattern = r"(\d+/\d+)\s*to\s*(\d+/\d+)\s+(.+)"

    for i, row in raw_df.iterrows():
        # Convert the entire row to a single string to find the header anywhere
        row_str = " ".join([str(val) for val in row if pd.notna(val)])
        
        # 1. Detect Alliance and Date Range header (re.IGNORECASE hace que ignore mayúsculas)
        match = re.search(header_pattern, row_str, re.IGNORECASE)
        if match:
            current_date_range = f"{match.group(1)} - {match.group(2)}"
            current_alliance = match.group(3).strip()
            continue
        
        # 2. Detect start of table (headers row)
        if str(row[0]).strip().lower() == "date":
            continue 
            
        # 3. Extract data if we are inside an alliance block
        # Check if there is a Player (row[2]) and a Legion (row[5]) to confirm it's a data row
        if current_alliance and pd.notna(row[2]) and str(row[2]).strip() != "" and pd.notna(row[5]):
            data_row = {
                'Alliance': current_alliance,
                'Date_Range': current_date_range,
                'Date': str(row[0])[:10], # Keeps only YYYY-MM-DD
                'Hour': str(row[1]),     
                'Player': str(row[2]),   
                'Score': row[3],
                'Result': str(row[4]),
                'Legion': row[5]
            }
            all_data.append(data_row)

    # Red de seguridad: Siempre devolver la tabla con las columnas correctas, aunque esté vacía
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
            st.error("⚠️ No data could be extracted. Check that the sheet is named 'BOD' and headers look like '5/9 to 5/10 DL'.")
            st.stop()

        # --- ADDITIONAL CLEANING ---
        # Clean Hour formatting
        df['Hour'] = df['Hour'].apply(lambda x: f"{int(x.replace('H','').strip()):02d}:00" if 'H' in str(x).upper() else x)
        
        # Clean Score
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

        # --- SAFETY NETS (Initialize empty dataframes) ---
        summary = pd.DataFrame()
        p_stats = pd.DataFrame()

        # ==========================================
        # SECTION 1: EVENT SUMMARY
        # ==========================================
        st.header(f"1. Battle Summary: {sel_range}")
        
        if not df_filtered.empty:
            # Technical grouping
            summary = df_filtered.groupby(['Alliance', 'Legion']).agg(
                Players=('Player', 'nunique'),
                Total_Score=('Score', 'sum'),
                Result=('Result', lambda x: x.mode().iloc[0] if not x.mode().empty else "-"),
                Schedule=('Hour', lambda x: x.mode().iloc[0] if not x.mode().empty else "-")
            ).reset_index()

            # Add active players count
            active_counts = df_filtered[df_filtered['Status'] == 'Active'].groupby(['Alliance', 'Legion']).size().reset_index(name='Active_Players')
            summary = summary.merge(active_counts, on=['Alliance', 'Legion'], how='left').fillna(0)
            summary['% Part.'] = (summary['Active_Players'] / summary['Players']) * 100

            st.dataframe(
                summary.style.format({'Total_Score': '{:,.0f}', '% Part.': '{:.1f}%'})
                .background_gradient(subset=['% Part.'], cmap='RdYlGn'),
                use_container_width=True, hide_index=True
            )

            # Activity Charts
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
        
        # Winrate by hour
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
            # Only save the tables if they were actually created
            if not summary.empty:
                summary.to_excel(writer, sheet_name='Summary', index=False)
            if not p_stats.empty:
                p_stats.to_excel(writer, sheet_name='Player_Ranking', index=False)
            df.to_excel(writer, sheet_name='Raw_Data', index=False)
            
        st.sidebar.download_button("💾 Download BOD Report", buffer, "BOD_Report.xlsx")

    except Exception as e:
        st.error(f"Critical error: {e}. Ensure the sheet is named 'BOD'.")
else:
    st.info("Upload the Excel file with the 'BOD' sheet to begin the analysis.")
