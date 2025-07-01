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
    """Generate target school data based on actual provided target data"""
    
    # Actual target data provided - mapped to shapefile chiefdom names
    target_data = {
        # BO District - using actual target numbers
        "BADJIA": 9,                    # Badjia
        "BAGBWE(BAGBE)": 18,           # Bagbwe
        "BOAMA": 56,                   # Baoma
        "BAGBO": 31,                   # Bargbo
        "BO TOWN": 86,                 # Bo City
        "BONGOR": 18,                  # Bongor
        "BUMPE NGAO": 63,              # Bumpeh
        "GBO": 10,                     # Gbo
        "JAIAMA": 25,                  # Jaiama
        "KAKUA": 164,                  # Kakua
        "KOMBOYA": 17,                 # Komboya
        "LUGBU": 32,                   # Lugbu
        "NIAWA LENGA": 25,             # Niawa Lenga
        "SELENGA": 7,                  # Selenga
        "TIKONKO": 89,                 # Tinkoko
        "VALUNIA": 38,                 # Valunia
        "WONDE": 13,                   # Wonde
        
        # BOMBALI District - using actual target numbers
        "BIRIWA": 48,                  # Biriwa
        "BOMBALI SEBORA": 44,          # Bombali Sebora
        "BOMBALI SIARI": 7,            # Bombali Serry Chiefdom
        "GBANTI": 40,                  # Gbanti (Bombali)
        "GBENDEMBU": 30,               # Gbendembu
        "KAMARANKA": 13,               # Kamaranka
        "MAGBAIMBA NDORWAHUN": 17,     # Magbaimba Ndohahun
        "MAKARI": 54,                  # Makarie
        "MAKENI CITY": 93,             # Makeni City
        "MARA": 15,                    # Mara
        "N'GOWAHUN": 28,               # Ngowahun
        "PAKI MASABONG": 29,           # Paki Masabong
        "SAFROKO LIMBA": 36,           # Safroko Limba
    }
    
    # Return target data for requested chiefdoms
    result = {}
    for chiefdom in chiefdoms:
        if chiefdom in target_data:
            result[chiefdom] = target_data[chiefdom]
        else:
            # If chiefdom not found, set to 0 to indicate no target data available
            result[chiefdom] = 0
            print(f"Warning: No target data found for chiefdom: {chiefdom}")
    
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
    """Create coverage dashboard optimized for Word document export"""
    
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
    
    # Optimize figure size for Word document (16:10 aspect ratio works well)
    fig_width = 16  # Width for Word document
    fig_height = rows * 3.5  # Height per row optimized for Word
    
    # Create subplot figure optimized for Word export
    fig, axes = plt.subplots(rows, cols, figsize=(fig_width, fig_height))
    fig.suptitle(f'{district_name} District - School Coverage Analysis', 
                 fontsize=18, fontweight='bold', y=0.98)
    
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
        chiefdom_gdf.plot(ax=ax, color=coverage_color, edgecolor='black', alpha=0.8, linewidth=1.5)
        
        # Create coverage text
        coverage_text = f"{actual_schools}/{target_schools} ({coverage_percent:.0f}%)"
        
        # Set title with coverage information (optimized font size for Word)
        ax.set_title(f'{chiefdom}\n{coverage_text}', 
                    fontsize=10, fontweight='bold', pad=8)
        
        # Remove axis labels and ticks for cleaner look
        ax.set_xticks([])
        ax.set_yticks([])
        ax.set_xlabel('')
        ax.set_ylabel('')
        
        # Remove the box frame
        for spine in ax.spines.values():
            spine.set_visible(False)
        
        # Add very light grid
        ax.grid(True, alpha=0.2, linestyle='--', linewidth=0.5)
        
        # Set equal aspect ratio
        ax.set_aspect('equal')
        
        # Set bounds to chiefdom extent with minimal padding for better fit
        bounds = chiefdom_gdf.total_bounds
        padding = 0.005  # Reduced padding for better fit in Word
        ax.set_xlim(bounds[0] - padding, bounds[2] + padding)
        ax.set_ylim(bounds[1] - padding, bounds[3] + padding)
    
    # Hide empty subplots
    total_plots = rows * cols
    for idx in range(len(chiefdoms), total_plots):
        row = idx // cols
        col = idx % cols
        axes[row, col].set_visible(False)
    
    # Optimize layout for Word document
    plt.tight_layout()
    plt.subplots_adjust(top=0.95, hspace=0.35, wspace=0.25)
    
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
    df_original = pd.read_excel("SBD_Submissions_07_01_2025.xlsx")
    
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

# Dashboard Settings - Fixed configuration
columns = 4  # Fixed to 4 columns for optimal Word export
show_targets = True  # Always show target data details

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
                label="📥 Download BO District Coverage Dashboard (PNG)",
                data=buffer_bo_coverage,
                file_name="BO_District_Coverage_Dashboard.png",
                mime="image/png"
            )
            
            # Word Export for BO District
            try:
                from docx import Document
                from docx.shared import Inches, Pt
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                
                # Create Word document
                doc = Document()
                
                # Add title
                title = doc.add_heading('BO District - School Coverage Analysis', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add generation date
                date_para = doc.add_paragraph()
                date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                date_run = date_para.add_run(f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
                date_run.font.size = Pt(12)
                
                doc.add_paragraph()  # Add space
                
                # Save matplotlib figure as PNG and embed in Word
                timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
                png_filename = f"BO_District_Coverage_Dashboard_{timestamp}.png"
                
                # Save PNG file to current directory
                fig_bo_coverage.savefig(png_filename, format='png', dpi=200, 
                                       bbox_inches='tight', facecolor='white', 
                                       edgecolor='none', pad_inches=0.1)
                
                # Add the saved PNG to Word document
                doc.add_picture(png_filename, width=Inches(9.5))  # Fits well in Word page
                
                # Add coverage legend
                doc.add_heading('Coverage Color Legend', level=2)
                legend_items = [
                    "🔴 Red: < 20% coverage",
                    "🟠 Orange: 20-39% coverage", 
                    "🟡 Yellow: 40-59% coverage",
                    "🟢 Light Green: 60-79% coverage",
                    "🔵 Blue: 80-99% coverage",
                    "🟣 Purple: 100%+ coverage"
                ]
                
                for item in legend_items:
                    p = doc.add_paragraph()
                    p.add_run('• ').bold = True
                    p.add_run(item)
                
                # Add summary information
                doc.add_heading('Dashboard Summary', level=2)
                
                bo_data = extracted_df[extracted_df["District"].str.upper() == "BO"]
                target_data_all = generate_target_school_data([])
                bo_chiefdoms = ['BADJIA', 'BAGBWE(BAGBE)', 'BOAMA', 'BAGBO', 'BO TOWN', 'BONGOR', 'BUMPE NGAO', 'GBO', 'JAIAMA', 'KAKUA', 'KOMBOYA', 'LUGBU', 'NIAWA LENGA', 'SELENGA', 'TIKONKO', 'VALUNIA', 'WONDE']
                bo_target_total = sum([v for k, v in target_data_all.items() if k in bo_chiefdoms])
                bo_coverage = (len(bo_data) / bo_target_total * 100) if bo_target_total > 0 else 0
                
                summary_text = f"""
                District: BO
                Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BO'])}
                Actual Schools: {len(bo_data)}
                Target Schools: {bo_target_total}
                Coverage Rate: {bo_coverage:.1f}%
                PNG File Saved: {png_filename}
                """
                
                for line in summary_text.strip().split('\n'):
                    if line.strip():
                        p = doc.add_paragraph()
                        p.add_run('• ').bold = True
                        p.add_run(line.strip())
                
                # Save to BytesIO
                word_buffer = BytesIO()
                doc.save(word_buffer)
                word_data = word_buffer.getvalue()
                
                # Success message
                st.success(f"✅ PNG saved as: {png_filename}")
                
                st.download_button(
                    label="📄 Download BO District Coverage Report (Word)",
                    data=word_data,
                    file_name=f"BO_District_Coverage_Report_{timestamp}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except ImportError:
                st.warning("⚠️ Word export requires python-docx library. Install with: pip install python-docx")
            except Exception as e:
                st.warning(f"⚠️ Word export failed: {str(e)}")
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
                label="📥 Download BOMBALI District Coverage Dashboard (PNG)",
                data=buffer_bombali_coverage,
                file_name="BOMBALI_District_Coverage_Dashboard.png",
                mime="image/png"
            )
            
            # Word Export for BOMBALI District
            try:
                from docx import Document
                from docx.shared import Inches, Pt
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                
                # Create Word document
                doc = Document()
                
                # Add title
                title = doc.add_heading('BOMBALI District - School Coverage Analysis', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add generation date
                date_para = doc.add_paragraph()
                date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                date_run = date_para.add_run(f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
                date_run.font.size = Pt(12)
                
                doc.add_paragraph()  # Add space
                
                # Save matplotlib figure as PNG and embed in Word
                timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
                png_filename = f"BOMBALI_District_Coverage_Dashboard_{timestamp}.png"
                
                # Save PNG file to current directory
                fig_bombali_coverage.savefig(png_filename, format='png', dpi=200, 
                                            bbox_inches='tight', facecolor='white', 
                                            edgecolor='none', pad_inches=0.1)
                
                # Add the saved PNG to Word document
                doc.add_picture(png_filename, width=Inches(9.5))  # Fits well in Word page
                
                # Add coverage legend
                doc.add_heading('Coverage Color Legend', level=2)
                legend_items = [
                    "🔴 Red: < 20% coverage",
                    "🟠 Orange: 20-39% coverage", 
                    "🟡 Yellow: 40-59% coverage",
                    "🟢 Light Green: 60-79% coverage",
                    "🔵 Blue: 80-99% coverage",
                    "🟣 Purple: 100%+ coverage"
                ]
                
                for item in legend_items:
                    p = doc.add_paragraph()
                    p.add_run('• ').bold = True
                    p.add_run(item)
                
                # Add summary information
                doc.add_heading('Dashboard Summary', level=2)
                
                bombali_data = extracted_df[extracted_df["District"].str.upper() == "BOMBALI"]
                target_data_all = generate_target_school_data([])
                bombali_chiefdoms = ['BIRIWA', 'BOMBALI SEBORA', 'BOMBALI SIARI', 'GBANTI', 'GBENDEMBU', 'KAMARANKA', 'MAGBAIMBA NDORWAHUN', 'MAKARI', 'MAKENI CITY', 'MARA', 'N\'GOWAHUN', 'PAKI MASABONG', 'SAFROKO LIMBA']
                bombali_target_total = sum([v for k, v in target_data_all.items() if k in bombali_chiefdoms])
                bombali_coverage = (len(bombali_data) / bombali_target_total * 100) if bombali_target_total > 0 else 0
                
                summary_text = f"""
                District: BOMBALI
                Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BOMBALI'])}
                Actual Schools: {len(bombali_data)}
                Target Schools: {bombali_target_total}
                Coverage Rate: {bombali_coverage:.1f}%
                PNG File Saved: {png_filename}
                """
                
                for line in summary_text.strip().split('\n'):
                    if line.strip():
                        p = doc.add_paragraph()
                        p.add_run('• ').bold = True
                        p.add_run(line.strip())
                
                # Save to BytesIO
                word_buffer = BytesIO()
                doc.save(word_buffer)
                word_data = word_buffer.getvalue()
                
                # Success message
                st.success(f"✅ PNG saved as: {png_filename}")
                
                st.download_button(
                    label="📄 Download BOMBALI District Coverage Report (Word)",
                    data=word_data,
                    file_name=f"BOMBALI_District_Coverage_Report_{timestamp}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except ImportError:
                st.warning("⚠️ Word export requires python-docx library. Install with: pip install python-docx")
            except Exception as e:
                st.warning(f"⚠️ Word export failed: {str(e)}")
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

# Export All Dashboards as Combined Word Document
st.header("📄 Combined Word Export")

if st.button("📋 Generate Combined Coverage Report", help="Generate a comprehensive Word document with both districts"):
    try:
        from docx import Document
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        # Create Word document
        doc = Document()
        
        # Add main title
        title = doc.add_heading('School-Based Distribution (SBD)', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        subtitle = doc.add_heading('School Coverage Analysis Dashboard', level=1)
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add generation date
        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_run = date_para.add_run(f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
        date_run.font.size = Pt(12)
        date_run.bold = True
        
        doc.add_page_break()
        
        # Coverage Legend
        doc.add_heading('Coverage Color Legend', level=1)
        legend_items = [
            "🔴 Red: < 20% coverage (Critical - requires immediate attention)",
            "🟠 Orange: 20-39% coverage (Poor - needs significant improvement)", 
            "🟡 Yellow: 40-59% coverage (Fair - room for improvement)",
            "🟢 Light Green: 60-79% coverage (Good - meeting most targets)",
            "🔵 Blue: 80-99% coverage (Excellent - exceeding expectations)",
            "🟣 Purple: 100%+ coverage (Outstanding - surpassing all targets)"
        ]
        
        for item in legend_items:
            p = doc.add_paragraph()
            p.add_run('• ').bold = True
            p.add_run(item)
        
        doc.add_page_break()
        
        # Executive Summary
        doc.add_heading('Executive Summary', level=1)
        
        bo_data = extracted_df[extracted_df["District"].str.upper() == "BO"]
        bombali_data = extracted_df[extracted_df["District"].str.upper() == "BOMBALI"]
        target_data_all = generate_target_school_data([])
        
        bo_chiefdoms = ['BADJIA', 'BAGBWE(BAGBE)', 'BOAMA', 'BAGBO', 'BO TOWN', 'BONGOR', 'BUMPE NGAO', 'GBO', 'JAIAMA', 'KAKUA', 'KOMBOYA', 'LUGBU', 'NIAWA LENGA', 'SELENGA', 'TIKONKO', 'VALUNIA', 'WONDE']
        bombali_chiefdoms = ['BIRIWA', 'BOMBALI SEBORA', 'BOMBALI SIARI', 'GBANTI', 'GBENDEMBU', 'KAMARANKA', 'MAGBAIMBA NDORWAHUN', 'MAKARI', 'MAKENI CITY', 'MARA', 'N\'GOWAHUN', 'PAKI MASABONG', 'SAFROKO LIMBA']
        
        bo_target_total = sum([v for k, v in target_data_all.items() if k in bo_chiefdoms])
        bombali_target_total = sum([v for k, v in target_data_all.items() if k in bombali_chiefdoms])
        total_target = bo_target_total + bombali_target_total
        total_actual = len(extracted_df)
        overall_coverage = (total_actual / total_target * 100) if total_target > 0 else 0
        
        summary_text = f"""
        This comprehensive dashboard report presents school coverage analysis comparing actual surveyed schools versus target schools for BO and BOMBALI districts:
        
        • Districts Covered: BO, BOMBALI
        • Total Target Schools: {total_target:,}
        • Total Actual Schools: {total_actual:,}
        • Overall Coverage Rate: {overall_coverage:.1f}%
        • BO District Coverage: {(len(bo_data)/bo_target_total*100) if bo_target_total > 0 else 0:.1f}%
        • BOMBALI District Coverage: {(len(bombali_data)/bombali_target_total*100) if bombali_target_total > 0 else 0:.1f}%
        
        Coverage is calculated as: (Actual Schools / Target Schools) × 100%
        Color coding helps identify areas requiring attention and those performing well.
        """
        
        for line in summary_text.strip().split('\n'):
            if line.strip():
                if line.startswith('•'):
                    p = doc.add_paragraph()
                    p.add_run(line.strip())
                else:
                    doc.add_paragraph(line.strip())
        
        doc.add_page_break()
        
        # BO District section
        if 'fig_bo_coverage' in locals():
            doc.add_heading('BO District - School Coverage Analysis', level=1)
            
            # Save BO figure as PNG and embed in Word
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            bo_png_filename = f"BO_District_Coverage_Combined_{timestamp}.png"
            
            # Save PNG file to current directory
            fig_bo_coverage.savefig(bo_png_filename, format='png', dpi=200, 
                                   bbox_inches='tight', facecolor='white', 
                                   edgecolor='none', pad_inches=0.1)
            
            # Add the saved PNG to Word document
            doc.add_picture(bo_png_filename, width=Inches(9.5))  # Fits well in Word page
            
            # BO summary
            doc.add_heading('BO District Summary', level=2)
            bo_coverage = (len(bo_data) / bo_target_total * 100) if bo_target_total > 0 else 0
            
            bo_summary_items = [
                f"Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BO'])}",
                f"Target Schools: {bo_target_total:,}",
                f"Actual Schools: {len(bo_data):,}",
                f"Coverage Rate: {bo_coverage:.1f}%",
                f"PNG File Saved: {bo_png_filename}"
            ]
            
            for item in bo_summary_items:
                p = doc.add_paragraph()
                p.add_run('• ').bold = True
                p.add_run(item)
            
            doc.add_page_break()
        
        # BOMBALI District section
        if 'fig_bombali_coverage' in locals():
            doc.add_heading('BOMBALI District - School Coverage Analysis', level=1)
            
            # Save BOMBALI figure as PNG and embed in Word
            bombali_png_filename = f"BOMBALI_District_Coverage_Combined_{timestamp}.png"
            
            # Save PNG file to current directory
            fig_bombali_coverage.savefig(bombali_png_filename, format='png', dpi=200, 
                                        bbox_inches='tight', facecolor='white', 
                                        edgecolor='none', pad_inches=0.1)
            
            # Add the saved PNG to Word document
            doc.add_picture(bombali_png_filename, width=Inches(9.5))  # Fits well in Word page
            
            # BOMBALI summary
            doc.add_heading('BOMBALI District Summary', level=2)
            bombali_coverage = (len(bombali_data) / bombali_target_total * 100) if bombali_target_total > 0 else 0
            
            bombali_summary_items = [
                f"Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BOMBALI'])}",
                f"Target Schools: {bombali_target_total:,}",
                f"Actual Schools: {len(bombali_data):,}",
                f"Coverage Rate: {bombali_coverage:.1f}%",
                f"PNG File Saved: {bombali_png_filename}"
            ]
            
            for item in bombali_summary_items:
                p = doc.add_paragraph()
                p.add_run('• ').bold = True
                p.add_run(item)
        
        # Save to BytesIO
        word_buffer = BytesIO()
        doc.save(word_buffer)
        word_data = word_buffer.getvalue()
        
        # Success message showing saved PNG files
        st.success(f"✅ PNG files saved: {bo_png_filename}, {bombali_png_filename}")
        
        st.download_button(
            label="💾 Download Combined Coverage Analysis Report (Word)",
            data=word_data,
            file_name=f"School_Coverage_Analysis_Report_{timestamp}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            help="Download comprehensive Word report with both districts"
        )
        
    except ImportError:
        st.error("❌ Word generation requires python-docx library. Please install it using: pip install python-docx")
    except Exception as e:
        st.error(f"❌ Error generating combined Word document: {str(e)}")

# Download detailed coverage data (CSV removed as requested)
st.subheader("📊 Coverage Analysis Summary")

# Memory optimization - close matplotlib figures
plt.close('all')

# Footer
st.markdown("---")
st.markdown("**📊 Section 2: School Coverage Analysis | School-Based Distribution Analysis**")
