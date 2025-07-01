import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import geopandas as gpd
import math
from io import BytesIO
import re

# Custom CSS for the dashboard
st.markdown("""
<style>
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        max-width: none;
    }
    
    h1 {
        color: #2c3e50;
        text-align: center;
        font-weight: 700;
        margin-bottom: 2rem;
    }
    
    h2, h3 {
        color: #34495e;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
    }
    
    .stButton > button {
        background: linear-gradient(45deg, #3498db, #2980b9);
        color: white;
        border: none;
        border-radius: 25px;
        padding: 0.5rem 2rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

def parse_gps_coordinates(gps_str):
    """Enhanced GPS coordinate parsing that handles multiple formats"""
    if pd.isna(gps_str):
        return None, None
    
    gps_str = str(gps_str).strip()
    lat, lon = None, None
    
    # Format 1: "8.6103181,-12.2029534"
    if ',' in gps_str and not ' ' in gps_str:
        try:
            parts = gps_str.split(',')
            if len(parts) == 2:
                lat = float(parts[0].strip())
                lon = float(parts[1].strip())
        except ValueError:
            pass
    
    # Format 2: "8.6103181 -12.2029534" (space separated)
    elif ' ' in gps_str and ',' not in gps_str:
        try:
            parts = gps_str.split()
            if len(parts) == 2:
                lat = float(parts[0].strip())
                lon = float(parts[1].strip())
        except ValueError:
            pass
    
    # Format 3: Other formats with parentheses, etc.
    else:
        # Extract numbers using regex
        numbers = re.findall(r'-?\d+\.?\d*', gps_str)
        if len(numbers) >= 2:
            try:
                lat = float(numbers[0])
                lon = float(numbers[1])
            except ValueError:
                pass
    
    return lat, lon

def extract_gps_data_from_excel(df):
    """Extract GPS data from the Excel file"""
    # Create empty lists to store extracted data
    districts, chiefdoms, gps_locations = [], [], []
    
    # Process each row in the "Scan QR code" column
    for idx, qr_text in enumerate(df["Scan QR code"]):
        if pd.isna(qr_text):
            districts.append(None)
            chiefdoms.append(None)
            gps_locations.append(None)
            continue
            
        # Extract values using regex patterns
        district_match = re.search(r"District:\s*([^\n]+)", str(qr_text))
        district = district_match.group(1).strip() if district_match else None
        districts.append(district)
        
        chiefdom_match = re.search(r"Chiefdom:\s*([^\n]+)", str(qr_text))
        chiefdom = chiefdom_match.group(1).strip() if chiefdom_match else None
        chiefdoms.append(chiefdom)
        
        # Get GPS Location from the corresponding row
        if "GPS Location" in df.columns:
            gps_locations.append(df["GPS Location"].iloc[idx])
        else:
            gps_locations.append(None)
    
    # Create a new DataFrame with extracted values
    extracted_df = pd.DataFrame({
        "District": districts,
        "Chiefdom": chiefdoms,
        "GPS_Location": gps_locations
    })
    
    return extracted_df

def create_chiefdom_subplot_dashboard(gdf, extracted_df, district_name, cols=4):
    """Create subplot dashboard for all chiefdoms in a district"""
    
    # Filter shapefile for the district
    district_gdf = gdf[gdf['FIRST_DNAM'] == district_name].copy()
    
    if len(district_gdf) == 0:
        st.error(f"No chiefdoms found for {district_name} district in shapefile")
        return None
    
    # Get unique chiefdoms from shapefile
    chiefdoms = sorted(district_gdf['FIRST_CHIE'].dropna().unique())
    
    st.write(f"**Found {len(chiefdoms)} chiefdoms in {district_name} District:**")
    for i, chiefdom in enumerate(chiefdoms):
        st.write(f"{i+1}. {chiefdom}")
    
    # Calculate rows needed
    rows = math.ceil(len(chiefdoms) / cols)
    
    # Create subplot figure
    fig, axes = plt.subplots(rows, cols, figsize=(cols*5, rows*4))
    fig.suptitle(f'{district_name} District - All Chiefdoms with GPS Locations', 
                 fontsize=20, fontweight='bold', y=0.98)
    
    # Ensure axes is always 2D array
    if rows == 1:
        axes = axes.reshape(1, -1)
    elif cols == 1:
        axes = axes.reshape(-1, 1)
    
    # Plot each chiefdom
    for idx, chiefdom in enumerate(chiefdoms):
        row = idx // cols
        col = idx % cols
        ax = axes[row, col]
        
        # Filter shapefile for this specific chiefdom
        chiefdom_gdf = district_gdf[district_gdf['FIRST_CHIE'] == chiefdom].copy()
        
        # Plot chiefdom boundary
        chiefdom_gdf.plot(ax=ax, color='lightblue', edgecolor='navy', alpha=0.7, linewidth=2)
        
        # Filter GPS data for this district and chiefdom
        district_data = extracted_df[extracted_df["District"].str.upper() == district_name.upper()].copy()
        chiefdom_data = district_data[district_data["Chiefdom"].str.contains(chiefdom, case=False, na=False)].copy()
        
        # Extract and plot GPS coordinates for this chiefdom
        coords_extracted = []
        if len(chiefdom_data) > 0 and not chiefdom_data["GPS_Location"].isna().all():
            gps_data = chiefdom_data["GPS_Location"].dropna()
            
            for gps_val in gps_data:
                if pd.notna(gps_val):
                    lat, lon = parse_gps_coordinates(gps_val)
                    
                    # Validate coordinates for Sierra Leone
                    if lat is not None and lon is not None:
                        if 6.0 <= lat <= 11.0 and -14.0 <= lon <= -10.0:
                            coords_extracted.append([lat, lon])
        
        # Plot GPS points if available
        if coords_extracted:
            lats, lons = zip(*coords_extracted)
            ax.scatter(lons, lats, c='red', s=100, alpha=1.0, 
                      edgecolors='white', linewidth=2, zorder=100, marker='o')
            
            # Add labels for each GPS point
            for i, (lat, lon) in enumerate(coords_extracted):
                ax.annotate(f'S{i+1}', (lon, lat), xytext=(3, 3), 
                           textcoords='offset points', fontsize=8, 
                           fontweight='bold', color='red',
                           bbox=dict(boxstyle='round,pad=0.2', facecolor='white', alpha=0.8))
        
        # Set title and labels
        ax.set_title(f'{chiefdom}\n({len(coords_extracted)} schools)', 
                    fontsize=12, fontweight='bold', pad=10)
        ax.set_xlabel('Longitude', fontsize=10)
        ax.set_ylabel('Latitude', fontsize=10)
        
        # Add grid
        ax.grid(True, alpha=0.3, linestyle='--')
        
        # Set equal aspect ratio and tight layout
        ax.set_aspect('equal')
        
        # Set bounds to chiefdom extent with some padding
        bounds = chiefdom_gdf.total_bounds
        padding = 0.01
        ax.set_xlim(bounds[0] - padding, bounds[2] + padding)
        ax.set_ylim(bounds[1] - padding, bounds[3] + padding)
    
    # Hide empty subplots
    total_plots = rows * cols
    for idx in range(len(chiefdoms), total_plots):
        row = idx // cols
        col = idx % cols
        axes[row, col].set_visible(False)
    
    plt.tight_layout()
    plt.subplots_adjust(top=0.95)
    
    return fig

# Streamlit App
st.title("🗺️ Chiefdom GPS Dashboard - BO and BOMBALI Districts")
st.markdown("**Comprehensive view of all chiefdoms with GPS school locations**")

# File upload section
st.sidebar.header("📁 File Upload")
uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=['xlsx'])

if uploaded_file is None:
    # Default file path
    uploaded_file = "sbd first_submission_clean.xlsx"

# Load the data
try:
    if isinstance(uploaded_file, str):
        df_original = pd.read_excel(uploaded_file)
        st.sidebar.success("✅ Default file loaded successfully!")
    else:
        df_original = pd.read_excel(uploaded_file)
        st.sidebar.success("✅ Uploaded file loaded successfully!")
    
    # Extract GPS data
    extracted_df = extract_gps_data_from_excel(df_original)
    st.sidebar.info(f"📊 Extracted {len(extracted_df)} records")
    
except Exception as e:
    st.error(f"❌ Error loading Excel file: {e}")
    st.stop()

# Load shapefile
try:
    gdf = gpd.read_file("Chiefdom2021.shp")
    st.sidebar.success("✅ Shapefile loaded successfully!")
    
    # Display shapefile info
    st.sidebar.info(f"🗺️ Shapefile contains {len(gdf)} chiefdoms")
    
except Exception as e:
    st.sidebar.error(f"❌ Could not load shapefile: {e}")
    st.stop()

# Dashboard Controls
st.sidebar.header("⚙️ Dashboard Settings")
columns = st.sidebar.slider("Number of columns", min_value=2, max_value=6, value=4)
show_data_info = st.sidebar.checkbox("Show data information", value=True)

if show_data_info:
    # Display data information
    st.subheader("📊 Data Overview")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_records = len(extracted_df)
        st.metric("Total Records", f"{total_records:,}")
    
    with col2:
        bo_records = len(extracted_df[extracted_df["District"].str.upper() == "BO"])
        st.metric("BO District", f"{bo_records:,}")
    
    with col3:
        bombali_records = len(extracted_df[extracted_df["District"].str.upper() == "BOMBALI"])
        st.metric("BOMBALI District", f"{bombali_records:,}")
    
    with col4:
        total_gps = len(extracted_df[extracted_df["GPS_Location"].notna()])
        st.metric("GPS Records", f"{total_gps:,}")

# Create dashboards
st.header("🗺️ District Dashboards")

# BO District Dashboard
st.subheader("1a. BO District - All Chiefdoms")
st.write("**Individual maps for each chiefdom in BO District with GPS school locations**")

with st.spinner("Generating BO District dashboard..."):
    try:
        fig_bo = create_chiefdom_subplot_dashboard(gdf, extracted_df, "BO", columns)
        if fig_bo:
            st.pyplot(fig_bo)
            
            # Save figure option
            buffer_bo = BytesIO()
            fig_bo.savefig(buffer_bo, format='png', dpi=300, bbox_inches='tight')
            buffer_bo.seek(0)
            
            st.download_button(
                label="📥 Download BO District Dashboard",
                data=buffer_bo,
                file_name="BO_District_Chiefdoms_Dashboard.png",
                mime="image/png"
            )
        else:
            st.warning("Could not generate BO District dashboard")
    except Exception as e:
        st.error(f"Error generating BO District dashboard: {e}")

st.divider()

# BOMBALI District Dashboard
st.subheader("1b. BOMBALI District - All Chiefdoms")
st.write("**Individual maps for each chiefdom in BOMBALI District with GPS school locations**")

with st.spinner("Generating BOMBALI District dashboard..."):
    try:
        fig_bombali = create_chiefdom_subplot_dashboard(gdf, extracted_df, "BOMBALI", columns)
        if fig_bombali:
            st.pyplot(fig_bombali)
            
            # Save figure option
            buffer_bombali = BytesIO()
            fig_bombali.savefig(buffer_bombali, format='png', dpi=300, bbox_inches='tight')
            buffer_bombali.seek(0)
            
            st.download_button(
                label="📥 Download BOMBALI District Dashboard",
                data=buffer_bombali,
                file_name="BOMBALI_District_Chiefdoms_Dashboard.png",
                mime="image/png"
            )
        else:
            st.warning("Could not generate BOMBALI District dashboard")
    except Exception as e:
        st.error(f"Error generating BOMBALI District dashboard: {e}")

# Summary Statistics
st.header("📈 Summary Statistics")

# Create summary for both districts
summary_data = []

for district in ["BO", "BOMBALI"]:
    district_gdf = gdf[gdf['FIRST_DNAM'] == district]
    district_data = extracted_df[extracted_df["District"].str.upper() == district.upper()]
    
    # Count GPS locations
    gps_count = len(district_data[district_data["GPS_Location"].notna()])
    
    summary_data.append({
        "District": district,
        "Chiefdoms": len(district_gdf) if len(district_gdf) > 0 else 0,
        "Total Records": len(district_data),
        "GPS Records": gps_count,
        "GPS Coverage": f"{(gps_count/len(district_data)*100) if len(district_data) > 0 else 0:.1f}%"
    })

summary_df = pd.DataFrame(summary_data)
st.dataframe(summary_df, use_container_width=True)

# Raw data preview
if st.checkbox("Show raw data preview"):
    st.subheader("📄 Raw Data Preview")
    st.dataframe(extracted_df.head(20))
    
    # Download raw data
    csv_data = extracted_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Extracted Data as CSV",
        data=csv_data,
        file_name="extracted_gps_data.csv",
        mime="text/csv"
    )

# Additional information
st.sidebar.markdown("---")
st.sidebar.markdown("""
### 📋 Dashboard Features:
- **Subplot Grid**: 4 columns (adjustable) × n rows
- **GPS Plotting**: Red markers for school locations
- **Chiefdom Boundaries**: Light blue with navy borders
- **School Labels**: S1, S2, etc. for each GPS point
- **Download Options**: High-quality PNG exports
""")

st.sidebar.markdown("""
### 🎯 Usage:
1. Adjust number of columns using slider
2. View comprehensive chiefdom grids
3. Download dashboard images
4. Check summary statistics
""")

# Footer
st.markdown("---")
st.markdown("**📊 Chiefdom GPS Dashboard | School-Based Distribution Analysis**")
