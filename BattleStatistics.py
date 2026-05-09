import streamlit as st
import pandas as pd
import plotly.express as px
import io

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Alliance & Legion Analytics", layout="wide")

# --- BARRA LATERAL ---
st.sidebar.title("🛡️ Panel de Control")
st.sidebar.markdown("Sube el archivo y filtra por Alianza y Semana.")

# Título Principal
st.title("📊 Reporte de Actividad: Multi-Alianza")
st.markdown("Este reporte permite analizar el rendimiento de múltiples alianzas y sus legiones.")

# ==========================================
# 1. CARGA Y PROCESAMIENTO DE DATOS
# ==========================================
uploaded_file = st.sidebar.file_uploader("📂 Cargar archivo Excel", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    try:
        # Leemos la primera hoja (sheet_name=0) para evitar errores de nombre
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # --- LIMPIEZA DE DATOS ---
        
        # 1. Limpiar Hora
        def clean_hour(val):
            val_str = str(val).upper().replace('H', '').strip()
            if val_str.isdigit():
                return f"{int(val_str):02d}:00"
            return val
        df['Heure'] = df['Heure'].apply(clean_hour)

        # 2. Fechas y Etiquetas de Semana
        temp_datetime = pd.to_datetime(df['Date'], dayfirst=True)
        df['Date'] = temp_datetime.dt.date
        df['Year'] = temp_datetime.dt.isocalendar().year
        df['Week'] = temp_datetime.dt.isocalendar().week

        week_date_map = {}
        for (year, week), group in df.groupby(['Year', 'Week']):
            unique_dates = sorted(group['Date'].unique())
            date_strs = [d.strftime('%d/%m') for d in unique_dates]
            label = f"{year} W{week:02d} ({', '.join(date_strs)})"
            week_date_map[(year, week)] = label

        df['Week_Label'] = df.apply(lambda row: week_date_map.get((row['Year'], row['Week'])), axis=1)

        # 3. Limpiar Score y Status
        if df['Score'].dtype == 'object':
            df['Score'] = df['Score'].astype(str).str.replace(',', '').str.replace('.', '')
        df['Score'] = pd.to_numeric(df['Score'], errors='coerce').fillna(0)
        df['Status'] = df['Score'].apply(lambda x: 'Activo' if x > 0 else 'Inactivo')
        
        # Auxiliar para victorias (1=Victoria, 0=Otros)
        df['Is_Win'] = df['Result'].astype(str).apply(lambda x: 1 if 'Victory' in x.strip() else 0)

        # ==========================================
        # FILTROS GLOBALES (SIDEBAR)
        # ==========================================
        
        # Filtro de Alianza
        unique_alliances = sorted(df['Alliance'].unique())
        selected_alliances = st.sidebar.multiselect(
            "🛡️ Seleccionar Alianzas:", 
            unique_alliances, 
            default=unique_alliances
        )

        # Filtro de Semana
        unique_weeks = sorted(df['Week_Label'].unique(), reverse=True)
        selected_week = st.sidebar.selectbox("📅 Seleccionar Semana:", unique_weeks)

        # APLICAR FILTRO DE ALIANZA AL DF GLOBAL
        df_filtered = df[df['Alliance'].isin(selected_alliances)]
        
        # DF filtrado por semana para la sección 1
        df_week_filtered = df_filtered[df_filtered['Week_Label'] == selected_week]

        # ==========================================
        # SECCIÓN 1: RESUMEN SEMANAL POR ALIANZA
        # ==========================================
        st.header(f"1. Resumen Semanal: {selected_week}")
        
        if not df_week_filtered.empty:
            # Agrupación por Alianza y Legión
            base_stats = df_week_filtered.groupby(['Alliance', 'Legion']).agg(
                Jugadores=('Joueur', 'nunique'),
                Score_Total=('Score', 'sum'),
                Resultado=('Result', lambda x: x.mode().iloc[0] if not x.mode().empty else "-"),
                Horario=('Heure', lambda x: x.mode().iloc[0] if not x.mode().empty else "-")
            )

            status_counts = df_week_filtered.groupby(['Alliance', 'Legion', 'Status']).size().unstack(fill_value=0)
            if 'Activo' not in status_counts.columns: status_counts['Activo'] = 0
            if 'Inactivo' not in status_counts.columns: status_counts['Inactivo'] = 0
                
            summary_table = base_stats.join(status_counts[['Activo', 'Inactivo']]).reset_index()
            summary_table['%_Part'] = (summary_table['Activo'] / summary_table['Jugadores']) * 100

            st.subheader("📊 Estadísticas por Alianza y Legión")
            st.dataframe(
                summary_table.style.format({
                    'Score_Total': lambda x: f"{x:,.0f}".replace(",", " "),
                    '%_Part': '{:.1f}%'
                }).background_gradient(subset=['%_Part'], cmap='RdYlGn', vmin=0, vmax=100),
                use_container_width=True,
                hide_index=True
            )

            # Gráfico Comparativo
            st.write("---")
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                st.subheader("👥 Actividad por Alianza")
                fig_act = px.bar(
                    df_week_filtered.groupby(['Alliance', 'Status']).size().reset_index(name='Cant'),
                    x="Alliance", y="Cant", color="Status", barmode="group",
                    color_discrete_map={'Activo': '#0052cc', 'Inactivo': '#EF553B'}
                )
                st.plotly_chart(fig_act, use_container_width=True)
                
            with col_chart2:
                st.subheader("⏰ Jugadores Activos por Hora")
                hourly_act = df_week_filtered[df_week_filtered['Score'] > 0].groupby(['Heure', 'Alliance']).size().reset_index(name='Activos')
                fig_hour = px.line(hourly_act, x='Heure', y='Activos', color='Alliance', markers=True)
                st.plotly_chart(fig_hour, use_container_width=True)

        else:
            st.warning("No hay datos para las alianzas seleccionadas en esta semana.")

        # ==========================================
        # SECCIÓN 2: RENDIMIENTO HISTÓRICO (TEMPORADA)
        # ==========================================
        st.divider()
        st.header("2. Rendimiento de Temporada (Global)")
        
        # Datos a nivel de partida (Unicos por Fecha, Alianza, Legión y Hora)
        match_level_df = df_filtered[['Date', 'Alliance', 'Legion', 'Heure', 'Is_Win']].drop_duplicates()

        col_win1, col_win2 = st.columns([1, 1])

        with col_win1:
            st.subheader("🏆 Winrate por Alianza")
            alliance_winrate = match_level_df.groupby('Alliance').agg(
                Partidas=('Is_Win', 'count'),
                Victorias=('Is_Win', 'sum')
            ).reset_index()
            alliance_winrate['Winrate'] = (alliance_winrate['Victorias'] / alliance_winrate['Partidas']) * 100
            
            st.dataframe(
                alliance_winrate.style.format({'Winrate': '{:.1f}%'})
                .background_gradient(subset=['Winrate'], cmap='RdYlGn', vmin=0, vmax=100),
                use_container_width=True, hide_index=True
            )

        with col_win2:
            st.subheader("🕰️ Winrate por Horario")
            time_winrate = match_level_df.groupby('Heure').agg(
                Partidas=('Is_Win', 'count'),
                Victorias=('Is_Win', 'sum')
            ).reset_index()
            time_winrate['Winrate'] = (time_winrate['Victorias'] / time_winrate['Partidas']) * 100
            
            fig_time_win = px.bar(time_winrate, x='Heure', y='Winrate', color='Winrate', color_continuous_scale='RdYlGn')
            st.plotly_chart(fig_time_win, use_container_width=True)

        # ==========================================
        # SECCIÓN 3: ESTADÍSTICAS INDIVIDUALES
        # ==========================================
        st.divider()
        st.header("3. Actividad Individual de Jugadores")
        
        player_stats = df_filtered.groupby(['Joueur', 'Alliance']).agg(
            Total_Score=('Score', 'sum'),
            Media_Score=('Score', 'mean'),
            Participaciones=('Score', lambda x: (x > 0).sum()),
            Horario_Pref=('Heure', lambda x: x.mode().iloc[0] if not x.mode().empty else "-")
        ).reset_index().sort_values('Total_Score', ascending=False)

        st.dataframe(
            player_stats.style.format({'Total_Score': '{:,.0f}', 'Media_Score': '{:,.1f}'}),
            use_container_width=True, hide_index=True
        )

        # ==========================================
        # EXPORTACIÓN
        # ==========================================
        st.sidebar.divider()
        st.sidebar.header("📥 Descargar Reportes")
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if not df_week_filtered.empty: summary_table.to_excel(writer, sheet_name='Resumen_Semanal', index=False)
            alliance_winrate.to_excel(writer, sheet_name='Winrate_Alianzas', index=False)
            player_stats.to_excel(writer, sheet_name='Stats_Jugadores', index=False)
            
        st.sidebar.download_button(
            label="💾 Descargar Excel Completo",
            data=buffer,
            file_name=f"Reporte_Alianzas_{selected_week}.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
else:
    st.info("Por favor, sube el archivo Excel en la barra lateral para comenzar.")

