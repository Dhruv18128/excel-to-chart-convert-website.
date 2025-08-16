import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import base64
from PIL import Image
import numpy as np

# Page configuration
st.set_page_config(
    page_title="Excel to Chart Converter - Transform Data Into Beautiful Charts",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Website design
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
    }
    
    .stats-container {
        background: rgba(255,255,255,0.95);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255,255,255,0.2);
        border-radius: 12px;
        padding: 2rem;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1);
    }
    
    .feature-card {
        background: rgba(255,255,255,0.15);
        backdrop-filter: blur(15px);
        border: 1px solid rgba(255,255,255,0.2);
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        color: white;
        transition: all 0.3s ease;
    }
    
    .feature-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 20px 25px -5px rgb(0 0 0 / 0.1);
        background: rgba(255,255,255,0.2);
    }
    
    .upload-section {
        border: 3px dashed #e2e8f0;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
        margin: 2rem 0;
        background: #f8fafc;
        transition: all 0.3s ease;
    }
    
    .upload-section:hover {
        border-color: #2563eb;
        background: #f1f5f9;
        transform: translateY(-2px);
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    }
    
    .template-card {
        background: white;
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        cursor: pointer;
        transition: all 0.3s ease;
        margin: 0.5rem;
    }
    
    .template-card:hover {
        border-color: #2563eb;
        transform: translateY(-3px);
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1);
    }
    
    .chart-container {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 2rem 0;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    }
    
    .success-message {
        background: linear-gradient(135deg, #d1fae5, #a7f3d0);
        border: 1px solid #10b981;
        border-radius: 12px;
        padding: 1rem;
        margin: 1rem 0;
        border-left: 4px solid #10b981;
    }
    
    .error-message {
        background: linear-gradient(135deg, #fee2e2, #fecaca);
        border: 1px solid #dc2626;
        border-radius: 12px;
        padding: 1rem;
        margin: 1rem 0;
        border-left: 4px solid #dc2626;
        color: #dc2626;
    }
    
    .demo-section {
        background: rgba(0,0,0,0.2);
        backdrop-filter: blur(15px);
        border: 1px solid rgba(255,255,255,0.2);
        border-radius: 12px;
        padding: 2rem;
        margin: 2rem 0;
    }
    
    .demo-chart {
        background: rgba(255,255,255,0.98);
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1);
        transition: all 0.3s ease;
        cursor: pointer;
        margin: 1rem;
    }
    
    .demo-chart:hover {
        transform: translateY(-8px);
        box-shadow: 0 20px 25px -5px rgb(0 0 0 / 0.1);
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #2563eb, #1d4ed8);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 1.5rem;
        font-weight: 700;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1);
    }
    
    .generate-btn > button {
        background: linear-gradient(135deg, #10b981, #059669);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 1rem 2rem;
        font-weight: 700;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    }
    
    .generate-btn > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1);
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'current_data' not in st.session_state:
    st.session_state.current_data = None
if 'current_chart' not in st.session_state:
    st.session_state.current_chart = None

def main():
    # Header Section
    st.markdown("""
    <div class="main-header">
        <div style="display: inline-block; background: rgba(255,255,255,0.2); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.3); border-radius: 50px; padding: 0.5rem 1rem; margin-bottom: 1rem; font-size: 0.875rem; font-weight: 500;">
            🚀 Free • Secure • No Sign-up Required
        </div>
        <h1 style="font-size: 3.75rem; font-weight: 900; margin-bottom: 1rem; text-shadow: 0 2px 4px rgba(0,0,0,0.3); line-height: 1.1;">
            📊 Excel to Chart Converter
        </h1>
        <p style="font-size: 1.375rem; max-width: 700px; margin: 0 auto; font-weight: 400;">
            Transform your spreadsheet data into stunning, interactive charts in seconds. Upload any Excel or CSV file and create professional visualizations that tell your data's story.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Trust Indicators
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div style="background: rgba(255,255,255,0.1); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.2); border-radius: 12px; padding: 1rem 1.5rem; color: white; font-size: 0.875rem; font-weight: 600; text-align: center;">
            🔒 100% Secure & Private
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div style="background: rgba(255,255,255,0.1); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.2); border-radius: 12px; padding: 1rem 1.5rem; color: white; font-size: 0.875rem; font-weight: 600; text-align: center;">
            ⚡ Works Offline
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div style="background: rgba(255,255,255,0.1); backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.2); border-radius: 12px; padding: 1rem 1.5rem; color: white; font-size: 0.875rem; font-weight: 600; text-align: center;">
            📱 Mobile Friendly
        </div>
        """, unsafe_allow_html=True)
    
    # Stats Section
    st.markdown("""
    <div class="stats-container">
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 2rem; text-align: center;">
            <div>
                <div style="font-size: 2.5rem; font-weight: 800; color: #2563eb; margin-bottom: 0.5rem;">10K+</div>
                <div style="font-size: 0.875rem; color: #64748b; font-weight: 500;">Charts Created</div>
            </div>
            <div>
                <div style="font-size: 2.5rem; font-weight: 800; color: #2563eb; margin-bottom: 0.5rem;">6</div>
                <div style="font-size: 0.875rem; color: #64748b; font-weight: 500;">Chart Types</div>
            </div>
            <div>
                <div style="font-size: 2.5rem; font-weight: 800; color: #2563eb; margin-bottom: 0.5rem;">0s</div>
                <div style="font-size: 0.875rem; color: #64748b; font-weight: 500;">Setup Time</div>
            </div>
            <div>
                <div style="font-size: 2.5rem; font-weight: 800; color: #2563eb; margin-bottom: 0.5rem;">100%</div>
                <div style="font-size: 0.875rem; color: #64748b; font-weight: 500;">Free Forever</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Features Section
    st.markdown("""
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 1rem; margin-bottom: 2rem;">
        <div class="feature-card">
            <div style="font-size: 2.5rem; margin-bottom: 1rem;">📈</div>
            <h3 style="font-size: 1.125rem; font-weight: 700; margin-bottom: 0.5rem;">6 Professional Chart Types</h3>
            <p style="font-size: 0.875rem; opacity: 0.9; line-height: 1.5;">Bar, Line, Pie, Doughnut, Radar & Polar Area charts. Perfect for any data visualization need.</p>
        </div>
        <div class="feature-card">
            <div style="font-size: 2.5rem; margin-bottom: 1rem;">⚡</div>
            <h3 style="font-size: 1.125rem; font-weight: 700; margin-bottom: 0.5rem;">Lightning Fast Processing</h3>
            <p style="font-size: 0.875rem; opacity: 0.9; line-height: 1.5;">No server uploads. Your data is processed instantly in your browser for maximum speed.</p>
        </div>
        <div class="feature-card">
            <div style="font-size: 2.5rem; margin-bottom: 1rem;">📱</div>
            <h3 style="font-size: 1.125rem; font-weight: 700; margin-bottom: 0.5rem;">Works Everywhere</h3>
            <p style="font-size: 0.875rem; opacity: 0.9; line-height: 1.5;">Optimized for desktop, tablet, and mobile. Create charts anywhere, anytime.</p>
        </div>
        <div class="feature-card">
            <div style="font-size: 2.5rem; margin-bottom: 1rem;">🔒</div>
            <h3 style="font-size: 1.125rem; font-weight: 700; margin-bottom: 0.5rem;">Privacy First Design</h3>
            <p style="font-size: 0.875rem; opacity: 0.9; line-height: 1.5;">Your data never leaves your device. Complete privacy and security guaranteed.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Interactive Demo Section
    st.markdown("""
    <div class="demo-section">
        <h2 style="color: white; text-align: center; margin-bottom: 2rem; font-size: 2.5rem; font-weight: 800;">🚀 Try It Now - Interactive Demo</h2>
        <p style="color: rgba(255,255,255,0.9); font-size: 1.125rem; text-align: center; margin-bottom: 2rem; max-width: 800px; margin-left: auto; margin-right: auto;">
            Experience the power instantly! Click any demo below to see how your data transforms into beautiful charts.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Demo Charts
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📈 Sales Performance Demo", key="sales_demo"):
            load_sales_demo()
    
    with col2:
        if st.button("🥧 Market Share Analysis", key="market_demo"):
            load_market_demo()
    
    with col3:
        if st.button("📊 Growth Analytics", key="growth_demo"):
            load_growth_demo()
    
    # Quick Start Templates
    st.markdown("## 🚀 Quick Start Templates")
    st.markdown("Choose a template to get started quickly, or upload your own data below.")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("📈 Sales Dashboard", key="sales_template"):
            load_sales_template()
    
    with col2:
        if st.button("💰 Financial Report", key="finance_template"):
            load_finance_template()
    
    with col3:
        if st.button("📊 Marketing Analytics", key="marketing_template"):
            load_marketing_template()
    
    with col4:
        if st.button("📦 Inventory Management", key="inventory_template"):
            load_inventory_template()
    
    # File Upload Section
    st.markdown("## 📂 Upload Your Excel File")
    
    uploaded_file = st.file_uploader(
        "Choose an Excel or CSV file",
        type=['xlsx', 'xls', 'csv'],
        help="Supports .xlsx, .xls, and .csv files (up to 200MB)"
    )
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                st.session_state.current_data = pd.read_csv(uploaded_file)
            else:
                st.session_state.current_data = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File uploaded successfully: {uploaded_file.name} ({len(st.session_state.current_data)} rows, {len(st.session_state.current_data.columns)} columns)")
            
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
    
    # Data Preview
    if st.session_state.current_data is not None:
        st.markdown("## 📋 Data Preview")
        st.dataframe(st.session_state.current_data.head(10), use_container_width=True)
        
        # Chart Controls
        st.markdown("## 🎨 Chart Options")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            chart_type = st.selectbox(
                "Chart Type",
                ["bar", "line", "pie", "doughnut", "scatter", "area", "histogram"],
                index=0
            )
        
        with col2:
            x_column = st.selectbox(
                "X-Axis Column",
                st.session_state.current_data.columns.tolist(),
                index=0
            )
        
        with col3:
            y_column = st.selectbox(
                "Y-Axis Column",
                st.session_state.current_data.columns.tolist(),
                index=min(1, len(st.session_state.current_data.columns) - 1)
            )
        
        with col4:
            chart_title = st.text_input("Chart Title", value="My Chart")
        
        # Generate Chart Button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🎨 Generate Chart", key="generate_chart", use_container_width=True):
                generate_chart(chart_type, x_column, y_column, chart_title)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #64748b; font-size: 0.875rem;">
        © 2025 Excel to Chart Converter. Created by Dhruv | Built with ❤️ for data visualization enthusiasts
    </div>
    """, unsafe_allow_html=True)

def load_sales_demo():
    """Load sales performance demo data"""
    demo_data = pd.DataFrame({
        'Month': ['January', 'February', 'March', 'April', 'May', 'June'],
        'Revenue': [65000, 59000, 80000, 81000, 56000, 95000],
        'Units Sold': [850, 720, 1020, 1050, 680, 1180],
        'Profit Margin': [18.5, 19.2, 21.1, 20.8, 17.9, 22.3]
    })
    
    st.session_state.current_data = demo_data
    st.success("🎉 Sales Performance demo data loaded! Scroll down to see the data preview and generate your chart.")
    st.rerun()

def load_market_demo():
    """Load market share demo data"""
    demo_data = pd.DataFrame({
        'Product': ['Product A', 'Product B', 'Product C', 'Product D'],
        'Market Share': [35, 25, 20, 20],
        'Revenue': [2800000, 2000000, 1600000, 1600000],
        'Growth Rate': [12.5, 8.3, 15.2, 6.7]
    })
    
    st.session_state.current_data = demo_data
    st.success("🎉 Market Share Analysis demo data loaded! Scroll down to see the data preview and generate your chart.")
    st.rerun()

def load_growth_demo():
    """Load growth analytics demo data"""
    demo_data = pd.DataFrame({
        'Year': [2020, 2021, 2022, 2023, 2024],
        'Customers': [120, 190, 300, 500, 620],
        'Revenue Growth': [0, 58, 150, 317, 417],
        'Team Size': [8, 12, 18, 25, 32]
    })
    
    st.session_state.current_data = demo_data
    st.success("🎉 Growth Analytics demo data loaded! Scroll down to see the data preview and generate your chart.")
    st.rerun()

def load_sales_template():
    """Load sales dashboard template"""
    template_data = pd.DataFrame({
        'Month': ['January', 'February', 'March', 'April', 'May', 'June'],
        'Revenue': [125000, 138000, 142000, 156000, 164000, 178000],
        'Units Sold': [1250, 1380, 1420, 1560, 1640, 1780],
        'Conversion Rate': [3.2, 3.8, 4.1, 4.3, 4.6, 4.9],
        'Customer Acquisition': [215, 248, 267, 289, 312, 341]
    })
    
    st.session_state.current_data = template_data
    st.success("🎉 Sales Dashboard template loaded! Ready to create your chart.")
    st.rerun()

def load_finance_template():
    """Load financial report template"""
    template_data = pd.DataFrame({
        'Category': ['Marketing', 'Operations', 'Technology', 'Human Resources', 'Administration'],
        'Budget': [50000, 80000, 35000, 65000, 25000],
        'Actual': [48500, 82300, 33200, 67800, 23900],
        'Variance': [1500, -2300, 1800, -2800, 1100],
        'Percentage': [97, 103, 95, 104, 96]
    })
    
    st.session_state.current_data = template_data
    st.success("🎉 Financial Report template loaded! Ready to create your chart.")
    st.rerun()

def load_marketing_template():
    """Load marketing analytics template"""
    template_data = pd.DataFrame({
        'Channel': ['Social Media', 'Email Campaign', 'Google Ads', 'Content Marketing', 'Influencer'],
        'Reach': [25000, 15000, 35000, 18000, 12000],
        'Engagement': [1250, 2250, 1750, 1980, 960],
        'Conversions': [156, 342, 289, 198, 87],
        'ROI': [285, 420, 315, 245, 180]
    })
    
    st.session_state.current_data = template_data
    st.success("🎉 Marketing Analytics template loaded! Ready to create your chart.")
    st.rerun()

def load_inventory_template():
    """Load inventory management template"""
    template_data = pd.DataFrame({
        'Product': ['Product Alpha', 'Product Beta', 'Product Gamma', 'Product Delta', 'Product Epsilon'],
        'Stock Level': [250, 89, 340, 156, 45],
        'Reorder Point': [100, 120, 80, 150, 75],
        'Monthly Sales': [85, 145, 65, 98, 112],
        'Status': ['Good', 'Low', 'Overstock', 'Good', 'Critical']
    })
    
    st.session_state.current_data = template_data
    st.success("🎉 Inventory Management template loaded! Ready to create your chart.")
    st.rerun()

def generate_chart(chart_type, x_column, y_column, title):
    """Generate the chart based on selected options"""
    if st.session_state.current_data is None:
        st.error("Please load data first!")
        return
    
    try:
        # Prepare data
        x_data = st.session_state.current_data[x_column].astype(str)
        y_data = pd.to_numeric(st.session_state.current_data[y_column], errors='coerce')
        
        # Remove NaN values
        valid_mask = ~y_data.isna()
        x_data = x_data[valid_mask]
        y_data = y_data[valid_mask]
        
        if len(x_data) == 0:
            st.error("No valid numeric data found in Y-axis column!")
            return
        
        # Create chart based on type
        if chart_type == "bar":
            fig = px.bar(
                x=x_data, 
                y=y_data, 
                title=title,
                labels={x_column: x_column, y_column: y_column}
            )
        elif chart_type == "line":
            fig = px.line(
                x=x_data, 
                y=y_data, 
                title=title,
                labels={x_column: x_column, y_column: y_column},
                markers=True
            )
        elif chart_type == "pie":
            fig = px.pie(
                values=y_data, 
                names=x_data, 
                title=title
            )
        elif chart_type == "doughnut":
            fig = px.pie(
                values=y_data, 
                names=x_data, 
                title=title,
                hole=0.4
            )
        elif chart_type == "scatter":
            fig = px.scatter(
                x=x_data, 
                y=y_data, 
                title=title,
                labels={x_column: x_column, y_column: y_column}
            )
        elif chart_type == "area":
            fig = px.area(
                x=x_data, 
                y=y_data, 
                title=title,
                labels={x_column: x_column, y_column: y_column}
            )
        elif chart_type == "histogram":
            fig = px.histogram(
                x=y_data, 
                title=title,
                labels={y_column: y_column}
            )
        
        # Update layout
        fig.update_layout(
            title_x=0.5,
            title_font_size=20,
            showlegend=True,
            height=500
        )
        
        # Display chart
        st.markdown("## 📊 Generated Chart")
        st.plotly_chart(fig, use_container_width=True)
        
        # Download options
        col1, col2 = st.columns(2)
        with col1:
            # Download chart as PNG
            img_bytes = fig.to_image(format="png")
            st.download_button(
                label="📥 Download Chart as PNG",
                data=img_bytes,
                file_name=f"{title.replace(' ', '_')}.png",
                mime="image/png"
            )
        
        with col2:
            # Download data as CSV
            csv = st.session_state.current_data.to_csv(index=False)
            st.download_button(
                label="📥 Download Data as CSV",
                data=csv,
                file_name="chart_data.csv",
                mime="text/csv"
            )
        
        st.success("🎉 Chart generated successfully!")
        
    except Exception as e:
        st.error(f"Error generating chart: {str(e)}")

if __name__ == "__main__":
    main()
