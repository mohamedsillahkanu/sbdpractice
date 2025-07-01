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

def create_chiefdom_mapping():
    """Create mapping between GPS data chiefdom names and shapefile FIRST_CHIE names"""
    chiefdom_mapping = {
        # BO District mappings
        "Bo City": "BO TOWN",
        "Badjia": "BADJIA",
        "Bargbo": "BAGBO",
        "Bagbwe": "BAGBWE(BAGBE)",
        "Baoma": "BOAMA",
        "Bongor": "BONGOR",
        "Bumpe Ngao": "BUMPE NGAO",
        "Gbo": "GBO",
        "Jaiama": "JAIAMA",
        "Kakua": "KAKUA",
        "Komboya": "KOMBOYA",
        "Lugbu": "LUGBU",
        "Niawa Lenga": "NIAWA LENGA",
        "Selenga": "SELENGA",
        "Tikonko": "TIKONKO",
        "Valunia": "VALUNIA",
        "Wonde": "WONDE",
        
        # BOMBALI District mappings
        "Biriwa": "BIRIWA",
        "Bombali Sebora": "BOMBALI SEBORA",
        "Bombali Serry": "BOMBALI SIARI",
        "Gbanti (Bombali)": "GBANTI",
        "Gbanti": "GBANTI",
        "Gbendembu": "GBENDEMBU",
        "Kamaranka": "KAMARANKA",
        "Magbaimba Ndohahun": "MAGBAIMBA NDORWAHUN",
        "Makarie": "MAKARI",
        "Mara": "MARA",
        "Ngowahun": "N'GOWAHUN",
        "Paki Masabong": "PAKI MASABONG",
        "Safroko Limba": "SAFROKO LIMBA",
        "Makeni City": "MAKENI CITY",
    }
    return chiefdom_mapping

def map_chiefdom_name(chiefdom_name, mapping):
    """Map chiefdom name from GPS data to shapefile name"""
    if pd.isna(chiefdom_name):
        return None
    
    chiefdom_name = str(chiefdom_name).strip()
    
    # Direct match
    if chiefdom_name in mapping:
        return mapping[chiefdom_name]
    
    # Case-insensitive match
    for key, value in mapping.items():
        if key.upper() == chiefdom_name.upper():
            return value
    
    # Partial match (contains)
    for key, value in mapping.items():
        if key.upper() in chiefdom_name.upper() or chiefdom_name.upper() in key.upper():
            return value
    
    # Return original if no mapping found
    return chiefdom_name

def extract_gps_data_from_excel(df):
    """Extract GPS data from the Excel file"""
    # Create empty lists to store extracted data
    districts, chiefdoms, gps_locations = [], [], []
    
    # Get chiefdom mapping
    chiefdom_mapping = create_chiefdom_mapping()
    
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
        original_chiefdom = chiefdom_match.group(1).strip() if chiefdom_match else None
        
        # Map chiefdom name to match shapefile
        mapped_chiefdom = map_chiefdom_name(original_chiefdom, chiefdom_mapping)
        chiefdoms.append(mapped_chiefdom)
        
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

def generate_target_school_data(chiefdoms):
    """Generate target school data based on real chiefdom data"""
    
    # Real target data provided
    target_data = {
        # BO District
        "BADJIA": 9,
        "BAGBWE(BAGBE)": 18,
        "BOAMA": 56,
        "BAGBO": 31,  # Bargbo
        "BO TOWN": 86,  # Bo City
        "BONGOR": 18,
        "BUMPE NGAO": 63,  # Bumpeh
        "GBO": 10,
        "JAIAMA": 25,
        "KAKUA": 164,
        "KOMBOYA": 17,
        "LUGBU": 32,
        "NIAWA LENGA": 25,
        "SELENGA": 7,
        "TIKONKO": 89,  # Tinkoko
        "VALUNIA": 38,
        "WONDE": 13,
        
        # BOMBALI District  
        "BIRIWA": 48,
        "BOMBALI SEBORA": 44,
        "BOMBALI SIARI": 7,  # Bombali Serry Chiefdom
        "GBANTI": 40,  # Gbanti (Bombali)
        "GBENDEMBU": 30,
        "KAMARANKA": 13,
        "MAGBAIMBA NDORWAHUN": 17,  # Magbaimba Ndohahun
        "MAKARI": 54,  # Makarie
        "MAKENI CITY": 93,
        "MARA": 15,
        "N'GOWAHUN": 28,  # Ngowahun
        "PAKI MASABONG": 29,
        "SAFROKO LIMBA": 36,
    }
    
    # For any chiefdom not in the list, return a default value
    result = {}
    for chiefdom in chiefdoms:
        result[chiefdom] = target_data.get(chiefdom, 20)  # Default to 20 if not found
    
    return result

def get_coverage_color(coverage_percent):
    """Get color based on coverage percentage"""
    if coverage_percent < 20:
        return '#d32f2f'  # Red
    elif coverage_percent < 40:
        return '#f57c00'  # Orange
    elif coverage_percent < 60:
        return '#fbc02d'  # Yellow
    elif coverage_percent < 80:
        return '#388e3c'  # Light Green
    elif coverage_percent < 100:
        return '#1976d2'  # Blue
    else:
        return '#4a148c'  # Purple (100% coverage)

def create_coverage_dashboard(gdf, extracted_df, district_name, cols=4):
    """Create coverage dashboard for all chiefdoms in a district"""
    
    # Filter shapefile for the district
    district_gdf = gdf[gdf['FIRST_DNAM'] == district_name].copy()
    
    if len(district_gdf) == 0:
        st.error(f"No chiefdoms found for {district_name} district in shapefile")
        return None
    
    # Get unique chiefdoms from shapefile
    chiefdoms = sorted(district_gdf['FIRST_CHIE'].dropna().unique())
    
    # Generate real target data
    target_data = generate_target_school_data(chiefdoms)
    
    # Calculate rows needed
    rows = math.ceil(len(chiefdoms) / cols)
    
    # Create subplot figure with increased vertical space
    fig, axes = plt.subplots(rows, cols, figsize=(cols*5, rows*6))
    fig.suptitle(f'{district_name} District - School Coverage Analysis', 
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
        
        # Filter GPS data for this district and chiefdom with exact matching
        district_data = extracted_df[extracted_df["District"].str.upper() == district_name.upper()].copy()
        chiefdom_data = district_data[district_data["Chiefdom"] == chiefdom].copy()
        
        # Count actual schools (records) for this chiefdom
        actual_schools = len(chiefdom_data)
        
        # Get target schools from dummy data
        target_schools = target_data.get(chiefdom, 50)  # Default to 50 if not found
        
        # Calculate coverage percentage
        coverage_percent = (actual_schools / target_schools * 100) if target_schools > 0 else 0
        coverage_percent = min(coverage_percent, 100)  # Cap at 100%
        
        # Get color based on coverage
        coverage_color = get_coverage_color(coverage_percent)
        
        # Plot chiefdom boundary with coverage color
        chiefdom_gdf.plot(ax=ax, color=coverage_color, edgecolor='black', alpha=0.8, linewidth=2)
        
        # Create coverage text
        coverage_text = f"{actual_schools}/{target_schools} ({coverage_percent:.0f}%)"
        
        # Set title with coverage information
        ax.set_title(f'{chiefdom}\n{coverage_text}', 
                    fontsize=11, fontweight='bold', pad=10)
        
        # Remove axis labels and ticks for cleaner look
        ax.set_xticks([])
        ax.set_yticks([])
        ax.set_xlabel('')
        ax.set_ylabel('')
        
        # Remove the box frame
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_visible(False)
        ax.spines['left'].set_visible(False)
        
        # Add light grid
        ax.grid(True, alpha=0.2, linestyle='--')
        
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
    plt.subplots_adjust(top=0.93, hspace=0.4, wspace=0.3)
    
    return fig

# Streamlit App
st.title("📊 Section 2: School Coverage Analysis")
st.markdown("**Survey completion rates comparing actual vs target schools**")

# Color legend
st.markdown("""
### Coverage Color Legend:
- 🔴 **Red**: < 20%
- 🟠 **Orange**: 20-39%
- 🟡 **Yellow**: 40-59%
- 🟢 **Light Green**: 60-79%
- 🔵 **Blue**: 80-99%
- 🟣 **Purple**: 100%+
""")

# File Information
st.info("""
**📁 Embedded Files:** `sbd first_submission_clean.xlsx` | `Chiefdom2021.shp`  
**📊 Layout:** Fixed 4-column grid optimized for Word export
""")

# Load the embedded data files
try:
    # Load Excel file (embedded)
    df_original = pd.read_excel("sbd first_submission_clean.xlsx")
    
    # Extract GPS data with chiefdom mapping
    extracted_df = extract_gps_data_from_excel(df_original)
    st.success(f"✅ Excel file loaded successfully! Found {len(extracted_df)} records.")
    
except Exception as e:
    st.error(f"❌ Error loading Excel file: {e}")
    st.info("💡 Make sure 'sbd first_submission_clean.xlsx' is in the same directory as this app")
    st.stop()

# Load shapefile (embedded)
try:
    gdf = gpd.read_file("Chiefdom2021.shp")
    st.success(f"✅ Shapefile loaded successfully! Found {len(gdf)} features.")
    
except Exception as e:
    st.error(f"❌ Could not load shapefile: {e}")
    st.info("💡 Make sure 'Chiefdom2021.shp' and supporting files (.dbf, .shx, .prj) are in the same directory as this app")
    st.stop()

# Dashboard Settings
st.sidebar.header("⚙️ Dashboard Settings")
columns = st.sidebar.selectbox("Number of columns", [2, 3, 4, 5], index=2)
show_targets = st.sidebar.checkbox("Show target data details", value=True)

if show_targets:
    # Display target data information
    st.subheader("🎯 Target School Data")
    
    # Show some target data
    target_data_all = generate_target_school_data([])
    st.write("Sample target schools by chiefdom:")
    
    # Create a sample table
    sample_targets = {k: v for k, v in list(target_data_all.items())[:10]}
    target_df = pd.DataFrame(list(sample_targets.items()), columns=['Chiefdom', 'Target Schools'])
    st.dataframe(target_df, use_container_width=True)
    
    st.info("💡 Coverage is calculated as: (Actual Schools / Target Schools) × 100%")

# Create dashboards
st.header("📊 School Coverage Dashboards")

# BO District Coverage Dashboard
st.subheader("BO District - School Coverage")

with st.spinner("Generating BO District coverage dashboard..."):
    try:
        fig_bo_coverage = create_coverage_dashboard(gdf, extracted_df, "BO", columns)
        if fig_bo_coverage:
            st.pyplot(fig_bo_coverage)
            
            # Save figure option
            buffer_bo_coverage = BytesIO()
            fig_bo_coverage.savefig(buffer_bo_coverage, format='png', dpi=300, bbox_inches='tight')
            buffer_bo_coverage.seek(0)
            
            st.download_button(
                label="📥 Download BO District Coverage Dashboard",
                data=buffer_bo_coverage,
                file_name="BO_District_Coverage_Dashboard.png",
                mime="image/png"
            )
        else:
            st.warning("Could not generate BO District coverage dashboard")
    except Exception as e:
        st.error(f"Error generating BO District coverage dashboard: {e}")

st.divider()

# BOMBALI District Coverage Dashboard
st.subheader("BOMBALI District - School Coverage")

with st.spinner("Generating BOMBALI District coverage dashboard..."):
    try:
        fig_bombali_coverage = create_coverage_dashboard(gdf, extracted_df, "BOMBALI", columns)
        if fig_bombali_coverage:
            st.pyplot(fig_bombali_coverage)
            
            # Save figure option
            buffer_bombali_coverage = BytesIO()
            fig_bombali_coverage.savefig(buffer_bombali_coverage, format='png', dpi=300, bbox_inches='tight')
            buffer_bombali_coverage.seek(0)
            
            st.download_button(
                label="📥 Download BOMBALI District Coverage Dashboard",
                data=buffer_bombali_coverage,
                file_name="BOMBALI_District_Coverage_Dashboard.png",
                mime="image/png"
            )
        else:
            st.warning("Could not generate BOMBALI District coverage dashboard")
    except Exception as e:
        st.error(f"Error generating BOMBALI District coverage dashboard: {e}")

# Coverage Analysis
st.header("📈 Coverage Analysis")

# Calculate coverage statistics
bo_data = extracted_df[extracted_df["District"].str.upper() == "BO"]
bombali_data = extracted_df[extracted_df["District"].str.upper() == "BOMBALI"]

# Get target data
target_data_all = generate_target_school_data([])
bo_chiefdoms = ['BADJIA', 'BAGBWE(BAGBE)', 'BOAMA', 'BAGBO', 'BO TOWN', 'BONGOR', 'BUMPE NGAO', 'GBO', 'JAIAMA', 'KAKUA', 'KOMBOYA', 'LUGBU', 'NIAWA LENGA', 'SELENGA', 'TIKONKO', 'VALUNIA', 'WONDE']
bombali_chiefdoms = ['BIRIWA', 'BOMBALI SEBORA', 'BOMBALI SIARI', 'GBANTI', 'GBENDEMBU', 'KAMARANKA', 'MAGBAIMBA NDORWAHUN', 'MAKARI', 'MAKENI CITY', 'MARA', 'N\'GOWAHUN', 'PAKI MASABONG', 'SAFROKO LIMBA']

bo_target_total = sum([v for k, v in target_data_all.items() if k in bo_chiefdoms])
bombali_target_total = sum([v for k, v in target_data_all.items() if k in bombali_chiefdoms])

# Coverage metrics
col1, col2, col3, col4 = st.columns(4)

with col1:
    bo_coverage = (len(bo_data) / bo_target_total * 100) if bo_target_total > 0 else 0
    st.metric("BO District Coverage", f"{bo_coverage:.1f}%", f"{len(bo_data)}/{bo_target_total}")

with col2:
    bombali_coverage = (len(bombali_data) / bombali_target_total * 100) if bombali_target_total > 0 else 0
    st.metric("BOMBALI District Coverage", f"{bombali_coverage:.1f}%", f"{len(bombali_data)}/{bombali_target_total}")

with col3:
    total_actual = len(extracted_df)
    total_target = bo_target_total + bombali_target_total
    overall_coverage = (total_actual / total_target * 100) if total_target > 0 else 0
    st.metric("Overall Coverage", f"{overall_coverage:.1f}%", f"{total_actual}/{total_target}")

with col4:
    # Calculate chiefdoms with good coverage (>= 60%)
    good_coverage_count = 0
    total_chiefdoms = 0
    
    for district in ["BO", "BOMBALI"]:
        district_data = extracted_df[extracted_df["District"].str.upper() == district.upper()]
        chiefdoms = district_data['Chiefdom'].dropna().unique()
        
        for chiefdom in chiefdoms:
            chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
            actual = len(chiefdom_data)
            target = target_data_all.get(chiefdom, 20)
            coverage = (actual / target * 100) if target > 0 else 0
            
            if coverage >= 60:
                good_coverage_count += 1
            total_chiefdoms += 1
    
    good_coverage_percent = (good_coverage_count / total_chiefdoms * 100) if total_chiefdoms > 0 else 0
    st.metric("Chiefdoms with Good Coverage", f"{good_coverage_percent:.0f}%", f"{good_coverage_count}/{total_chiefdoms}")

# Detailed coverage table
st.subheader("📋 Detailed Coverage by Chiefdom")

detailed_coverage = []
for district in ["BO", "BOMBALI"]:
    district_data = extracted_df[extracted_df["District"].str.upper() == district.upper()]
    chiefdoms = sorted(district_data['Chiefdom'].dropna().unique())
    
    for chiefdom in chiefdoms:
        chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
        actual = len(chiefdom_data)
        target = target_data_all.get(chiefdom, 20)
        coverage = (actual / target * 100) if target > 0 else 0
        
        # Determine status
        if coverage >= 80:
            status = "✅ Excellent"
        elif coverage >= 60:
            status = "🟢 Good"
        elif coverage >= 40:
            status = "🟡 Fair"
        elif coverage >= 20:
            status = "🟠 Poor"
        else:
            status = "🔴 Critical"
        
        detailed_coverage.append({
            'District': district,
            'Chiefdom': chiefdom,
            'Actual Schools': actual,
            'Target Schools': target,
            'Coverage %': f"{coverage:.1f}%",
            'Status': status
        })

coverage_df = pd.DataFrame(detailed_coverage)
st.dataframe(coverage_df, use_container_width=True)

# Download detailed coverage data
csv_data = coverage_df.to_csv(index=False)
st.download_button(
    label="📥 Download Coverage Analysis CSV",
    data=csv_data,
    file_name="school_coverage_analysis.csv",
    mime="text/csv"
)

# Memory optimization - close matplotlib figures
plt.close('all')

# Footer
st.markdown("---")
st.markdown("**📊 Section 2: School Coverage Analysis | School-Based Distribution Analysis**")
