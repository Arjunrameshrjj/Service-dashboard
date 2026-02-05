import streamlit as st
import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import json
import numpy as np
from io import BytesIO
import xlsxwriter

# Set page config
st.set_page_config(
    page_title="Service Analytics Dashboard",
    page_icon="📚",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .header-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 0.5rem;
        color: white;
        margin-bottom: 2rem;
    }
    .kpi-container {
        display: flex;
        justify-content: center;
        gap: 20px;
        margin: 20px 0;
        flex-wrap: wrap;
    }
    .kpi-box {
        min-width: 220px;
        max-width: 220px;
        height: 130px;
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        color: white;
        display: flex;
        flex-direction: column;
        justify-content: center;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }
    .kpi-title {
        font-size: 14px;
        opacity: 0.9;
        margin-bottom: 8px;
        font-weight: 500;
    }
    .kpi-value {
        font-size: 34px;
        font-weight: 700;
        margin: 5px 0;
    }
    .kpi-sub {
        font-size: 13px;
        opacity: 0.85;
    }
    .kpi-box-blue { background: linear-gradient(135deg, #4A6FA5, #166088); }
    .kpi-box-green { background: linear-gradient(135deg, #2E8B57, #3CB371); }
    .kpi-box-orange { background: linear-gradient(135deg, #FF7A59, #FFA500); }
    .kpi-box-purple { background: linear-gradient(135deg, #8A2BE2, #9370DB); }
    .kpi-box-teal { background: linear-gradient(135deg, #20B2AA, #48D1CC); }
    .section-header {
        background: linear-gradient(90deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1.5rem 0 1rem 0;
        border-left: 4px solid #4dabf7;
    }
</style>
""", unsafe_allow_html=True)

def render_kpi(title, value, subtitle="", color_class=""):
    """Render KPI card."""
    return f'<div class="kpi-box {color_class}"><div class="kpi-title">{title}</div><div class="kpi-value">{value}</div><div class="kpi-sub">{subtitle}</div></div>'

def render_leader_card(title, name, value, subtitle, icon="🏆"):
    """Render a fancy leader card."""
    return f"""
    <div style="
        background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%);
        border-radius: 15px;
        padding: 20px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 10px 20px rgba(0,0,0,0.1);
        display: flex;
        align-items: center;
        justify-content: space-between;
    ">
        <div>
            <div style="font-size: 14px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9;">{title}</div>
            <div style="font-size: 32px; font-weight: 800; margin: 5px 0;">{name}</div>
            <div style="font-size: 18px; opacity: 0.9;">{value}</div>
            <div style="font-size: 14px; opacity: 0.8;">{subtitle}</div>
        </div>
        <div style="font-size: 50px; background: rgba(255,255,255,0.2); border-radius: 50%; width: 80px; height: 80px; display: flex; align-items: center; justify-content: center;">
            {icon}
        </div>
    </div>
    """

def render_kpi_row(kpis):
    """Render row of KPI cards."""
    return f'<div class="kpi-container">{"".join(kpis)}</div>'

def fetch_data_from_sheets(api_url):
    """Fetch data from Google Apps Script API."""
    try:
        response = requests.get(f"{api_url}?action=getData", timeout=30)
        response.raise_for_status()
        data = response.json()
        
        if isinstance(data, dict) and 'data' in data:
            return pd.DataFrame(data['data']), None
        else:
            return None, data.get('error', 'Unknown error')
    
    except Exception as e:
        return None, str(e)

def normalize_dataframe(df):
    """Normalize column names and data."""
    if df.empty:
        return df
    
    # Standardize column names (Map from Apps Script keys to Internal names)
    column_mapping = {
        'sl_no': 'SL_NO',
        'student_name': 'Student_Name',
        'course': 'Course',
        'contact_number': 'Contact',
        'package_hours': 'Package_Hours',
        'mail_id': 'Email',
        'joining_date': 'Joining_Date',
        'status': 'Status',
        'started_date': 'Started_Date',
        'completed_date': 'Completed_Date',
        'tutor_name': 'Tutor',
        'team_name': 'Team',
        'sheet_name': 'Month',
        'new_old': 'New_Old'
    }
    
    df = df.rename(columns=column_mapping)
    
    # Normalize Status
    if 'Status' in df.columns:
        def clean_status(x):
            s = str(x).lower().strip()
            if 'started' in s and 'not' not in s:
                return 'Started'
            elif 'yes' in s:
                return 'Started'
            # Check for dates in Started_Date if status is ambiguous? 
            # For now, strict 'not started' check
            return 'Not Started'

        df['Status_Clean'] = df['Status'].apply(clean_status)
    
    # Cross-verify with Started_Date if available
    if 'Started_Date' in df.columns and 'Status_Clean' in df.columns:
        # If started date is present, force status to Started
        df.loc[df['Started_Date'].astype(str).str.strip().str.len() > 4, 'Status_Clean'] = 'Started'
    
    # Add completion status
    if 'Completed_Date' in df.columns:
        # More robust completion check
        df['Is_Completed'] = df['Completed_Date'].apply(
            lambda x: bool(x) and str(x).strip() not in ['', 'nan', 'None', 'NaT']
        )
        df['Completion_Status'] = df['Is_Completed'].apply(
            lambda x: 'Completed' if x else 'In Progress'
        )
    else:
        # If no Completed_Date column, mark all as In Progress
        df['Is_Completed'] = False
        df['Completion_Status'] = 'In Progress'
    
    return df

def calculate_kpis(df):
    """Calculate KPIs."""
    if df.empty:
        return {}
    
    total_students = len(df)
    
    # Status breakdown
    started = len(df[df['Status_Clean'] == 'Started']) if 'Status_Clean' in df.columns else 0
    not_started = total_students - started
    
    # Completion
    completed = df['Is_Completed'].sum() if 'Is_Completed' in df.columns else 0
    in_progress = started - completed
    
    # Rates
    start_rate = round((started / total_students * 100), 1) if total_students > 0 else 0
    completion_rate = round((completed / total_students * 100), 1) if total_students > 0 else 0
    
    # Top performers
    top_tutor = df['Tutor'].value_counts().index[0] if 'Tutor' in df.columns and not df['Tutor'].empty else 'N/A'
    top_team = df['Team'].value_counts().index[0] if 'Team' in df.columns and not df['Team'].empty else 'N/A'
    top_course = df['Course'].value_counts().index[0] if 'Course' in df.columns and not df['Course'].empty else 'N/A'
    
    return {
        'total_students': total_students,
        'started': started,
        'not_started': not_started,
        'completed': completed,
        'in_progress': in_progress,
        'start_rate': start_rate,
        'completion_rate': completion_rate,
        'top_tutor': str(top_tutor)[:20],
        'top_team': str(top_team)[:20],
        'top_course': str(top_course)[:20]
    }

def create_tutor_performance(df):
    """Create tutor performance metrics."""
    if df.empty or 'Tutor' not in df.columns:
        return pd.DataFrame()
    
    tutor_stats = df.groupby('Tutor').agg(
        Total_Students=('Student_Name', 'count'),
        Started=('Status_Clean', lambda x: (x == 'Started').sum()),
        Completed=('Is_Completed', 'sum'),
        In_Progress=('Completion_Status', lambda x: (x == 'In Progress').sum())
    ).reset_index()
    
    tutor_stats['Start_Rate_%'] = np.where(
        tutor_stats['Total_Students'] > 0,
        (tutor_stats['Started'] / tutor_stats['Total_Students'] * 100).round(1),
        0
    )
    
    tutor_stats['Completion_Rate_%'] = np.where(
        tutor_stats['Started'] > 0,
        (tutor_stats['Completed'] / tutor_stats['Started'] * 100).round(1),
        0
    )
    
    return tutor_stats.sort_values('Total_Students', ascending=False)

def create_team_performance(df):
    """Create team performance metrics."""
    if df.empty or 'Team' not in df.columns:
        return pd.DataFrame()
    
    team_stats = df.groupby('Team').agg(
        Total_Students=('Student_Name', 'count'),
        Started=('Status_Clean', lambda x: (x == 'Started').sum()),
        Completed=('Is_Completed', 'sum')
    ).reset_index()
    
    team_stats['Completion_Rate_%'] = np.where(
        team_stats['Started'] > 0,
        (team_stats['Completed'] / team_stats['Started'] * 100).round(1),
        0
    )
    
    return team_stats.sort_values('Total_Students', ascending=False)

def create_course_analysis(df):
    """Create course-wise analysis."""
    if df.empty or 'Course' not in df.columns:
        return pd.DataFrame()
    
    course_stats = df.groupby('Course').agg(
        Total_Students=('Student_Name', 'count'),
        Started=('Status_Clean', lambda x: (x == 'Started').sum()),
        Completed=('Is_Completed', lambda x: x.astype(int).sum())
    ).reset_index()
    
    course_stats['Completion_Rate_%'] = np.where(
        course_stats['Started'] > 0,
        (course_stats['Completed'] / course_stats['Started'] * 100).round(1),
        0
    )
    
    return course_stats.sort_values('Total_Students', ascending=False)

def create_excel_report(df, kpis, tutor_perf, team_perf, course_analysis):
    """Create Excel report."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'align': 'center',
            'border': 1
        })
        
        # Executive Summary
        summary_data = {
            'Metric': [
                'Total Students',
                'Started',
                'Not Started',
                'Completed',
                'In Progress',
                'Start Rate %',
                'Completion Rate %'
            ],
            'Value': [
                kpis['total_students'],
                kpis['started'],
                kpis['not_started'],
                kpis['completed'],
                kpis['in_progress'],
                kpis['start_rate'],
                kpis['completion_rate']
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
        
        worksheet = writer.sheets['Executive Summary']
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Raw Data
        if not df.empty:
            df.to_excel(writer, sheet_name='Raw Data', index=False)
        
        # Tutor Performance
        if not tutor_perf.empty:
            tutor_perf.to_excel(writer, sheet_name='Tutor Performance', index=False)
        
        # Team Performance
        if not team_perf.empty:
            team_perf.to_excel(writer, sheet_name='Team Performance', index=False)
        
        # Course Analysis
        if not course_analysis.empty:
            course_analysis.to_excel(writer, sheet_name='Course Analysis', index=False)
    
    return output.getvalue()

def main():
    # Header
    st.markdown("""
        <div class="header-container">
            <h1 style="margin: 0; font-size: 2.5rem;">📚 Service Analytics Dashboard</h1>
            <p style="margin: 0.5rem 0 0 0; font-size: 1.2rem;">Student Progress & Performance Tracking</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'data_df' not in st.session_state:
        st.session_state.data_df = None
    if 'api_url' not in st.session_state:
        st.session_state.api_url = ""
    
    # Sidebar
    with st.sidebar:
        st.markdown("## 🔧 Configuration")
        
        st.markdown("### 🔗 Google Apps Script URL")
        api_url = st.text_input(
            "Enter your Web App URL:",
            value=st.session_state.api_url if st.session_state.api_url else "https://script.google.com/macros/s/AKfycbzTQAKD41vvCK-An5VpqacvdoR-6MWFhmtaB45KWyVhuiUZ-dgUWDFaogsNB5sdJ_Ad/exec",
            placeholder="https://script.google.com/macros/s/..../exec",
            help="Deploy the Apps Script and paste the Web App URL here"
        )
        
        if api_url:
            st.session_state.api_url = api_url
        
        st.divider()
        
        st.markdown("## ⚡ Quick Actions")
        
        if st.button("🚀 Fetch Data", type="primary", use_container_width=True, disabled=not api_url):
            with st.spinner("Fetching data from Google Sheets..."):
                df, error = fetch_data_from_sheets(api_url)
                
                if df is not None:
                    df = normalize_dataframe(df)
                    st.session_state.data_df = df
                    st.success(f"✅ Loaded {len(df)} records!")
                    st.rerun()
                else:
                    st.error(f"❌ Error: {error}")
        
        if st.button("🔄 Refresh", use_container_width=True, disabled=st.session_state.data_df is None):
            st.rerun()
        
        if st.button("🗑️ Clear", use_container_width=True):
            st.session_state.clear()
            st.rerun()
        
        st.divider()
        
        st.markdown("### 📖 Setup Instructions")
        st.info("""
        1. Open your Google Sheet
        2. Go to Extensions > Apps Script
        3. Paste the provided script
        4. Update SHEET_NAME
        5. Deploy as Web App
        6. Copy URL and paste above
        """)
    
    # Main content
    if st.session_state.data_df is not None and not st.session_state.data_df.empty:
        df = st.session_state.data_df
        
        # Filter Logic
        if 'Month' in df.columns:
            months = sorted(df['Month'].unique().astype(str).tolist())
            
            with st.sidebar:
                st.divider()
                st.markdown("## 📅 Filter Data")
                selected_month = st.selectbox(
                    "Select Month/Sheet",
                    options=["All Time"] + months,
                    index=0
                )
            
            if selected_month != "All Time":
                df = df[df['Month'] == selected_month]
                st.info(f"📅 Showing data for: **{selected_month}** ({len(df)} records)")

        kpis = calculate_kpis(df)
        
        # KPI Dashboard
        st.markdown('<div class="section-header"><h2>🏆 Executive Dashboard</h2></div>', unsafe_allow_html=True)
        
        st.markdown(
            render_kpi_row([
                render_kpi("Total Students", f"{kpis['total_students']:,}", "Enrolled", "kpi-box-blue"),
                render_kpi("Started", f"{kpis['started']:,}", f"{kpis['start_rate']}% start rate", "kpi-box-green"),
                render_kpi("Completed", f"{kpis['completed']:,}", f"{kpis['completion_rate']}% completion", "kpi-box-purple"),
                render_kpi("In Progress", f"{kpis['in_progress']:,}", "Active students", "kpi-box-orange"),
            ]),
            unsafe_allow_html=True
        )
        
        st.divider()
        
        # Download Section
        st.markdown("### 📥 Download Options")
        
        col1, col2 = st.columns(2)
        
        with col1:
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📄 Download CSV",
                csv,
                f"service_data_{datetime.now().strftime('%Y%m%d')}.csv",
                "text/csv",
                use_container_width=True
            )
        
        with col2:
            if st.button("💎 Generate Excel Report", use_container_width=True, type="primary"):
                with st.spinner("Creating Excel report..."):
                    tutor_perf = create_tutor_performance(df)
                    team_perf = create_team_performance(df)
                    course_analysis = create_course_analysis(df)
                    
                    excel_data = create_excel_report(df, kpis, tutor_perf, team_perf, course_analysis)
                    
                    st.download_button(
                        "⬇️ Download Excel",
                        excel_data,
                        f"Service_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        st.divider()
        
        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 Overview",
            "👨‍🏫 Tutor Performance",
            "👥 Team Performance",
            "📚 Course Analysis"
        ])
        
        with tab1:
            st.markdown('<div class="section-header"><h3>📊 Student Overview</h3></div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if 'Status_Clean' in df.columns:
                    status_counts = df['Status_Clean'].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Count']
                    
                    fig = px.pie(status_counts, values='Count', names='Status', 
                                title='Student Status Distribution', hole=0.3)
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                if 'Completion_Status' in df.columns:
                    comp_counts = df['Completion_Status'].value_counts().reset_index()
                    comp_counts.columns = ['Status', 'Count']
                    
                    fig = px.bar(comp_counts, x='Status', y='Count',
                               title='Completion Status', color='Count')
                    st.plotly_chart(fig, use_container_width=True)
            
            st.dataframe(df, use_container_width=True, height=300)
        
        with tab2:
            st.markdown('<div class="section-header"><h3>👨‍🏫 Tutor Performance</h3></div>', unsafe_allow_html=True)
            
            tutor_perf = create_tutor_performance(df)
            
            if not tutor_perf.empty:
                st.dataframe(tutor_perf, use_container_width=True, height=400)
                
                fig = px.bar(tutor_perf.head(10), x='Tutor', y='Total_Students',
                           title='Top 10 Tutors by Student Count',
                           color='Completion_Rate_%', color_continuous_scale='Greens')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            st.markdown('<div class="section-header"><h3>👥 Team Performance Targets</h3></div>', unsafe_allow_html=True)
            
            team_perf = create_team_performance(df)
            
            if not team_perf.empty:
                # Champion Card
                top_team = team_perf.iloc[0]
                st.markdown(render_leader_card(
                    "Top Performing Team", 
                    top_team['Team'], 
                    f"{top_team['Total_Students']} Students",
                    f"Completion Rate: {top_team['Completion_Rate_%']}%"
                ), unsafe_allow_html=True)

                col1, col2 = st.columns([2, 1])
                
                with col1:
                    # Fancy Horizontal Bar Chart
                    fig = px.bar(
                        team_perf, 
                        y='Team', 
                        x='Total_Students',
                        orientation='h',
                        title='Team Student Volume',
                        text='Total_Students',
                        color='Completion_Rate_%', 
                        color_continuous_scale='Viridis',
                        labels={'Total_Students': 'Number of Students', 'Completion_Rate_%': 'Completion %'}
                    )
                    fig.update_traces(textposition='outside')
                    fig.update_layout(yaxis={'categoryorder':'total ascending'}, height=500)
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    st.markdown("### 📋 Detailed Metrics")
                    st.dataframe(
                        team_perf,
                        column_config={
                            "Completion_Rate_%": st.column_config.ProgressColumn(
                                "Completion Rate",
                                format="%.1f%%",
                                min_value=0,
                                max_value=100,
                            ),
                        },
                        use_container_width=True, 
                        height=500
                    )
        
        with tab4:
            st.markdown('<div class="section-header"><h3>📚 Course Popularity & Analysis</h3></div>', unsafe_allow_html=True)
            
            course_analysis = create_course_analysis(df)
            
            if not course_analysis.empty:
                # Champion Card
                top_course = course_analysis.iloc[0]
                st.markdown(render_leader_card(
                    "Most Popular Course", 
                    top_course['Course'], 
                    f"{top_course['Total_Students']} Students",
                    f"{top_course['Completion_Rate_%']}% Completion",
                    icon="📚"
                ), unsafe_allow_html=True)

                # Simple horizontal bar chart - Course Popularity
                fig = px.bar(
                    course_analysis,
                    y='Course',
                    x='Total_Students',
                    orientation='h',
                    title='Course Popularity (Bar Length = Students, Color = Completion Rate)',
                    text='Total_Students',
                    color='Completion_Rate_%',
                    color_continuous_scale='RdYlGn',
                    labels={'Total_Students': 'Number of Students', 'Completion_Rate_%': 'Completion %'}
                )
                fig.update_traces(textposition='outside')
                fig.update_layout(
                    yaxis={'categoryorder': 'total ascending'},
                    height=max(500, len(course_analysis) * 35),
                    plot_bgcolor='white',
                    margin=dict(l=150, r=50, t=80, b=50)
                )
                st.plotly_chart(fig, use_container_width=True)
                
                with st.expander("🔎 View Course Details Table"):
                    st.dataframe(course_analysis, use_container_width=True)
    
    else:
        st.info("👈 Enter your Google Apps Script URL in the sidebar and click 'Fetch Data'")

if __name__ == "__main__":
    main()
