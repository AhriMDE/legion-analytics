import streamlit as st
import pandas as pd
import plotly.express as px
import io

# ==========================================
# STREAMLIT PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="Legion Analytics", layout="wide")

# --- SIDEBAR CONFIGURATION ---
st.sidebar.title("🎛️ Control Panel")
st.sidebar.markdown("Upload file and select week to filter.")

# Main Title
st.title("📊 Legion & Player Activity Report")
st.markdown("Select a week in the sidebar to filter the **Legion Summary**. The **Player Activity** section shows global history.")

# ==========================================
# 1. FILE UPLOADER (SIDEBAR)
# ==========================================
uploaded_file = st.sidebar.file_uploader("📂 Upload Excel File", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        # Load the specific sheet
        df = pd.read_excel(uploaded_file, sheet_name='Données Brutes')
        
        # ==========================================
        # --- DATA CLEANING & PREPROCESSING ---
        # ==========================================
        
        # 1. Clean Time (Heure)
        def clean_hour(val):
            val_str = str(val).upper().replace('H', '').strip()
            if val_str.isdigit():
                return f"{int(val_str):02d}:00"
            return val
        
        df['Heure'] = df['Heure'].apply(clean_hour)

        # 2. Clean Date & Create Week Label
        temp_datetime = pd.to_datetime(df['Date'], dayfirst=True)
        df['Date'] = temp_datetime.dt.date
        df['Year'] = temp_datetime.dt.isocalendar().year
        df['Week'] = temp_datetime.dt.isocalendar().week

        # Create "Week Label"
        week_date_map = {}
        for (year, week), group in df.groupby(['Year', 'Week']):
            unique_dates = sorted(group['Date'].unique())
            date_strs = [d.strftime('%d/%m') for d in unique_dates]
            dates_joined = ", ".join(date_strs)
            label = f"{year} W{week:02d} ({dates_joined})"
            week_date_map[(year, week)] = label

        df['Week_Label'] = df.apply(lambda row: week_date_map.get((row['Year'], row['Week'])), axis=1)

        # 3. Clean Score
        if df['Score'].dtype == 'object':
            df['Score'] = df['Score'].astype(str).str.replace(',', '').str.replace('.', '')
        df['Score'] = pd.to_numeric(df['Score'], errors='coerce').fillna(0)

        # 4. Define Status
        df['Status'] = df['Score'].apply(lambda x: 'Active' if x > 0 else 'Inactive')

        # ==========================================
        # --- WEEK SELECTOR (SIDEBAR) ---
        # ==========================================
        unique_weeks = sorted(df['Week_Label'].unique(), reverse=True)
        selected_week = st.sidebar.selectbox("📅 Select Week to Analyze:", unique_weeks)

        # Create a filtered dataframe ONLY for the Legion section
        df_week_filtered = df[df['Week_Label'] == selected_week]

        # ==========================================
        # 2. LEGION SUMMARY SECTION (WEEKLY)
        # ==========================================
        st.header(f"1. Legion Summary: {selected_week}")
        st.caption("These statistics apply only to the selected week.")
        
        # --- A. STATS & CHART ---
        
        # AJUSTE 1: Añadimos 'Result' a la agregación para obtener Victoria/Derrota
        base_stats = df_week_filtered.groupby('Legion').agg(
            Total_Players=('Joueur', 'nunique'),
            Total_Score=('Score', 'sum'),
            Result=('Result', lambda x: x.mode().iloc[0] if not x.mode().empty else "N/A")
        )

        status_counts = df_week_filtered.groupby(['Legion', 'Status']).size().unstack(fill_value=0)
        
        # Ensure cols exist
        if 'Active' not in status_counts.columns: status_counts['Active'] = 0
        if 'Inactive' not in status_counts.columns: status_counts['Inactive'] = 0
            
        final_summary_table = base_stats.join(status_counts[['Active', 'Inactive']]).reset_index()

        # AJUSTE 2: Calcular Porcentaje de Participación
        final_summary_table['Participation_Rate'] = (final_summary_table['Active'] / final_summary_table['Total_Players']) * 100

        # Add Total Row
        # Nota: El Total de participación se recalcula basado en los totales, no en el promedio
        total_active = final_summary_table['Active'].sum()
        total_players_sum = final_summary_table['Total_Players'].sum()
        total_participation_rate = (total_active / total_players_sum * 100) if total_players_sum > 0 else 0

        total_row = pd.DataFrame({
            'Legion': ['TOTAL'],
            'Total_Players': [total_players_sum],
            'Total_Score': [final_summary_table['Total_Score'].sum()],
            'Result': ['-'], # No aplica resultado para el total
            'Active': [total_active],
            'Inactive': [final_summary_table['Inactive'].sum()],
            'Participation_Rate': [total_participation_rate]
        })
        final_summary_table = pd.concat([final_summary_table, total_row], ignore_index=True)

        col1, col2 = st.columns([1.5, 2]) # Ajusté un poco el ancho para que quepan las columnas nuevas
        
        with col1:
            st.subheader("Stats by Legion")
            st.dataframe(
                final_summary_table.style.format({
                    'Total_Score': lambda x: f"{x:,.0f}".replace(",", " "),
                    'Total_Players': '{:.0f}', 
                    'Active': '{:.0f}', 
                    'Inactive': '{:.0f}',
                    'Participation_Rate': '{:.1f}%' # Formato porcentaje
                }), 
                use_container_width=True,
                column_order=['Legion', 'Result', 'Total_Players', 'Total_Score', 'Active', 'Inactive', 'Participation_Rate'] # Reordenar columnas
            )

        with col2:
            st.subheader("Weekly Activity Chart")
            weekly_activity = df_week_filtered.groupby(['Legion', 'Status']).size().reset_index(name='Count')
            fig = px.bar(
                weekly_activity, x="Legion", y="Count", color="Status", 
                title=f"Legion Health - {selected_week}",
                labels={"Count": "Number of Players"},
                color_discrete_map={'Active': '#0052cc', 'Inactive': '#EF553B'},
                barmode='stack'
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- WEEKLY HOURLY CHART ---
        st.write("---")
        st.subheader(f"⏰ Active Participations by Schedule ({selected_week})")
        weekly_hourly_stats = df_week_filtered[df_week_filtered['Score'] > 0].groupby('Heure').size().reset_index(name='Active_Players')
        
        fig_weekly_hourly = px.bar(
            weekly_hourly_stats, x='Heure', y='Active_Players',
            title=f"Sum of Active Players by Hour - {selected_week}",
            labels={'Active_Players': 'Total Active Players'},
            color_discrete_sequence=['#00CC96'], text='Active_Players'
        )
        fig_weekly_hourly.update_traces(textposition='outside')
        st.plotly_chart(fig_weekly_hourly, use_container_width=True)

        # --- B. DETAILED LISTS ---
        st.divider()
        st.subheader("📋 Detailed Player Lists by Legion (with Schedule)")

        legion_time_map = df_week_filtered.groupby('Legion')['Heure'].agg(
            lambda x: x.mode().iloc[0] if not x.mode().empty else "N/A"
        )

        def create_legion_table(status_filter):
            subset = df_week_filtered[df_week_filtered['Status'] == status_filter]
            all_legions = sorted(df_week_filtered['Legion'].unique())
            data_dict = {}
            max_length = 0
            for legion in all_legions:
                players = subset[subset['Legion'] == legion]['Joueur'].tolist()
                time_val = legion_time_map.get(legion, "N/A")
                col_name = f"Legion {legion} ({time_val})"
                data_dict[col_name] = players
                if len(players) > max_length: max_length = len(players)
            for key in data_dict: data_dict[key] += [""] * (max_length - len(data_dict[key]))
            return pd.DataFrame(data_dict)

        df_active = create_legion_table('Active')
        df_inactive = create_legion_table('Inactive')

        col_act, col_inact = st.columns(2)
        with col_act:
            st.markdown("✅ **Active Players**")
            st.dataframe(df_active, use_container_width=True, hide_index=True)
        with col_inact:
            st.markdown("💤 **Inactive Players**")
            st.dataframe(df_inactive, use_container_width=True, hide_index=True)

        # ==========================================
        # 3. INDIVIDUAL PLAYER SUMMARY (GLOBAL)
        # ==========================================
        st.divider()
        st.header("2. Individual Player Activity (Global History)")
        st.caption("These statistics are cumulative (All Time).")

        def calculate_win_rate(series):
            wins = series.apply(lambda x: 1 if 'Victory' in str(x) else 0).sum()
            total = len(series)
            return round((wins / total) * 100, 2) if total > 0 else 0
        def get_preferred_schedule(series):
            return series.mode()[0] if not series.mode().empty else "N/A"

        player_stats = df.groupby('Joueur').agg(
            Total_Score=('Score', 'sum'),
            Average_Score=('Score', 'mean'),
            Win_Rate=('Result', calculate_win_rate),
            Participations=('Score', lambda x: (x > 0).sum()),
            Absences=('Score', lambda x: (x == 0).sum()),
            Preferred_Schedule=('Heure', get_preferred_schedule)
        ).reset_index()
        player_stats['Average_Score'] = player_stats['Average_Score'].round(2)
        
        st.dataframe(
            player_stats.style.format({
                'Total_Score': lambda x: f"{x:,.0f}".replace(",", " "),
                'Average_Score': lambda x: f"{x:,.2f}".replace(",", " "),
                'Win_Rate': '{:.2f}', 'Participations': '{:.0f}', 'Absences': '{:.0f}'
            }), use_container_width=True
        )

        # --- GLOBAL HOURLY CHART ---
        st.subheader("🌍 Total Participation by Schedule (Global)")
        global_hourly_stats = df[df['Score'] > 0].groupby('Heure').size().reset_index(name='Active_Players')
        fig_global_hourly = px.bar(
            global_hourly_stats, x='Heure', y='Active_Players',
            title="Sum of Active Participations by Hour (All Time)",
            labels={'Active_Players': 'Total Participations'},
            color_discrete_sequence=['#AB63FA'], text='Active_Players'
        )
        fig_global_hourly.update_traces(textposition='outside')
        st.plotly_chart(fig_global_hourly, use_container_width=True)

        # --- SCHEDULE MATRIX ---
        st.subheader("🕰️ Active Participation by Schedule Matrix")
        active_global_df = df[df['Score'] > 0]
        schedule_pivot = active_global_df.pivot_table(
            index='Joueur', columns='Heure', values='Score', aggfunc='count', fill_value=0
        )
        schedule_pivot['TOTAL'] = schedule_pivot.sum(axis=1)
        schedule_pivot = schedule_pivot.sort_values('TOTAL', ascending=False)
        st.dataframe(schedule_pivot, use_container_width=True)

        # ==========================================
        # 4. RAW DATA SECTION (RESTORED)
        # ==========================================
        st.divider()
        st.header("3. Raw Data")
        with st.expander("📂 Click to see raw data from Excel"):
            st.dataframe(df)

        # ==========================================
        # --- SIDEBAR: EXPORTS ---
        # ==========================================
        st.sidebar.divider()
        st.sidebar.header("📥 Download Reports")
        
        # --- 1. WEEKLY REPORT (BUFFER) ---
        buffer_week = io.BytesIO()
        with pd.ExcelWriter(buffer_week, engine='xlsxwriter') as writer:
            final_summary_table.to_excel(writer, sheet_name='Summary', index=False)
            df_active.to_excel(writer, sheet_name='Active', index=False)
            df_inactive.to_excel(writer, sheet_name='Inactive', index=False)
            weekly_hourly_stats.to_excel(writer, sheet_name='Hourly_Stats', index=False)
            df_week_filtered.to_excel(writer, sheet_name='Raw_Data', index=False)
            
        st.sidebar.download_button(
            label="📄 Download Weekly Report",
            data=buffer_week,
            file_name=f"Weekly_Report_{selected_week.replace(' ', '_').replace('(', '').replace(')', '')}.xlsx",
            mime="application/vnd.ms-excel"
        )

        # --- 2. GLOBAL REPORT (BUFFER) ---
        buffer_global = io.BytesIO()
        with pd.ExcelWriter(buffer_global, engine='xlsxwriter') as writer:
            # Stats Globales
            player_stats.to_excel(writer, sheet_name='Global_Stats', index=False)
            schedule_pivot.to_excel(writer, sheet_name='Schedule_Matrix')
            global_hourly_stats.to_excel(writer, sheet_name='Global_Hourly', index=False)
            # Todo el historial
            df.to_excel(writer, sheet_name='Full_History_Raw', index=False)

        st.sidebar.download_button(
            label="💾 Download Full History",
            data=buffer_global,
            file_name="Legion_Global_History.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Error processing the file: {e}")

else:
    st.info("Please upload an Excel file in the sidebar to begin.")
    
