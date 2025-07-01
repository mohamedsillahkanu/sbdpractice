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

def create_chiefdom_subplot_dashboard(gdf, extracted_df, district_name, cols=4):
    """Create subplot dashboard for all chiefdoms in a district"""
    
    # Filter shapefile for the district
    district_gdf = gdf[gdf['FIRST_DNAM'] == district_name].copy()
    
    if len(district_gdf) == 0:
        st.error(f"No chiefdoms found for {district_name} district in shapefile")
        return None
    
    # Get unique chiefdoms from shapefile
    chiefdoms = sorted(district_gdf['FIRST_CHIE'].dropna().unique())
    
    # Calculate rows needed
    rows = math.ceil(len(chiefdoms) / cols)
    
    # Create subplot figure with increased vertical space
    fig, axes = plt.subplots(rows, cols, figsize=(cols*5, rows*6))
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
        
        # Filter GPS data for this district and chiefdom with exact matching
        district_data = extracted_df[extracted_df["District"].str.upper() == district_name.upper()].copy()
        chiefdom_data = district_data[district_data["Chiefdom"] == chiefdom].copy()
        
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
                            coords_extracted.append([lat, lon, str(gps_val)])
        
        # Handle overlapping coordinates by adding small offsets
        def separate_overlapping_points(coords, min_distance=0.001):
            """Separate overlapping GPS points by adding small offsets"""
            if len(coords) <= 1:
                return coords
            
            separated_coords = []
            for i, (lat, lon, gps_str) in enumerate(coords):
                adjusted_lat, adjusted_lon = lat, lon
                
                # Check if this point overlaps with any previous point
                for j, (prev_lat, prev_lon, _) in enumerate(separated_coords):
                    distance = ((adjusted_lat - prev_lat)**2 + (adjusted_lon - prev_lon)**2)**0.5
                    if distance < min_distance:
                        # Add small offset in a circular pattern
                        angle = (i * 2 * 3.14159) / len(coords)
                        offset = min_distance * 1.5
                        adjusted_lat += offset * math.cos(angle)
                        adjusted_lon += offset * math.sin(angle)
                
                separated_coords.append([adjusted_lat, adjusted_lon, gps_str])
            
            return separated_coords
        
        # Separate overlapping points
        if coords_extracted:
            coords_extracted = separate_overlapping_points(coords_extracted)
        
        # Plot GPS points if available
        if coords_extracted:
            lats, lons, gps_strings = zip(*coords_extracted)
            
            # Plot each point individually to ensure they're all visible
            for i, (lat, lon) in enumerate(zip(lats, lons)):
                ax.scatter(lon, lat, c='red', s=100, alpha=1.0, 
                          edgecolors='white', linewidth=2, zorder=100, marker='o')
            
            # Add debug info in title if multiple points exist
            if len(coords_extracted) > 1:
                unique_coords = len(set([(round(lat, 6), round(lon, 6)) for lat, lon, _ in coords_extracted]))
                if unique_coords < len(coords_extracted):
                    debug_info = f" [Debug: {len(coords_extracted)} total, {unique_coords} unique locations]"
                else:
                    debug_info = ""
            else:
                debug_info = ""
        
        # Set title and clean up axes
        ax.set_title(f'{chiefdom}\n({len(coords_extracted)} schools{debug_info if coords_extracted else ""})', 
                    fontsize=12, fontweight='bold', pad=10)
        
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
    plt.subplots_adjust(top=0.93, hspace=0.4, wspace=0.3)
    
    return fig

# Streamlit App
st.title("🗺️ Section 1: GPS School Locations Dashboard")
st.markdown("**Visual mapping of all school GPS coordinates by chiefdom**")

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

# Dashboard Settings - Fixed configuration
columns = 4  # Fixed to 4 columns for optimal Word export
show_data_info = True  # Always show data overview

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
st.header("🗺️ GPS Location Dashboards")

# BO District Dashboard
st.subheader("BO District - All Chiefdoms")

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
                label="📥 Download BO District Dashboard (PNG)",
                data=buffer_bo,
                file_name="BO_District_GPS_Dashboard.png",
                mime="image/png"
            )
            
            # Word Export for BO District
            try:
                from docx import Document
                from docx.shared import Inches, Pt
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                import tempfile
                import os
                
                # Create Word document
                doc = Document()
                
                # Add title
                title = doc.add_heading('BO District - GPS School Locations Dashboard', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add generation date
                date_para = doc.add_paragraph()
                date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                date_run = date_para.add_run(f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
                date_run.font.size = Pt(12)
                
                doc.add_paragraph()  # Add space
                
                # Save matplotlib figure to temporary file for Word export
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                    # Save with optimized settings for Word document
                    fig_bo.savefig(tmp_file.name, format='png', dpi=200, 
                                  bbox_inches='tight', facecolor='white', 
                                  edgecolor='none', pad_inches=0.1)
                    
                    # Add image to Word document with optimal size for page fit
                    doc.add_picture(tmp_file.name, width=Inches(9.5))  # Fits well in Word page
                    
                    # Clean up temp file
                    os.unlink(tmp_file.name)
                
                # Add summary information
                doc.add_heading('Dashboard Summary', level=1)
                
                summary_text = f"""
                District: BO
                Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BO'])}
                Total Records: {len(extracted_df[extracted_df['District'].str.upper() == 'BO'])}
                GPS Records: {len(extracted_df[(extracted_df['District'].str.upper() == 'BO') & extracted_df['GPS_Location'].notna()])}
                GPS Coverage: {(len(extracted_df[(extracted_df['District'].str.upper() == 'BO') & extracted_df['GPS_Location'].notna()]) / len(extracted_df[extracted_df['District'].str.upper() == 'BO']) * 100) if len(extracted_df[extracted_df['District'].str.upper() == 'BO']) > 0 else 0:.1f}%
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
                
                st.download_button(
                    label="📄 Download BO District Dashboard (Word)",
                    data=word_data,
                    file_name="BO_District_GPS_Dashboard.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except ImportError:
                st.warning("⚠️ Word export requires python-docx library. Install with: pip install python-docx")
            except Exception as e:
                st.warning(f"⚠️ Word export failed: {str(e)}")
        else:
            st.warning("Could not generate BO District dashboard")
    except Exception as e:
        st.error(f"Error generating BO District dashboard: {e}")

st.divider()

# BOMBALI District Dashboard
st.subheader("BOMBALI District - All Chiefdoms")

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
                label="📥 Download BOMBALI District Dashboard (PNG)",
                data=buffer_bombali,
                file_name="BOMBALI_District_GPS_Dashboard.png",
                mime="image/png"
            )
            
            # Word Export for BOMBALI District
            try:
                from docx import Document
                from docx.shared import Inches, Pt
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                import tempfile
                import os
                
                # Create Word document
                doc = Document()
                
                # Add title
                title = doc.add_heading('BOMBALI District - GPS School Locations Dashboard', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Add generation date
                date_para = doc.add_paragraph()
                date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                date_run = date_para.add_run(f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
                date_run.font.size = Pt(12)
                
                doc.add_paragraph()  # Add space
                
                # Save matplotlib figure to temporary file for Word export
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                    # Save with optimized settings for Word document
                    fig_bombali.savefig(tmp_file.name, format='png', dpi=200, 
                                       bbox_inches='tight', facecolor='white', 
                                       edgecolor='none', pad_inches=0.1)
                    
                    # Add image to Word document with optimal size for page fit
                    doc.add_picture(tmp_file.name, width=Inches(9.5))  # Fits well in Word page
                    
                    # Clean up temp file
                    os.unlink(tmp_file.name)
                
                # Add summary information
                doc.add_heading('Dashboard Summary', level=1)
                
                summary_text = f"""
                District: BOMBALI
                Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BOMBALI'])}
                Total Records: {len(extracted_df[extracted_df['District'].str.upper() == 'BOMBALI'])}
                GPS Records: {len(extracted_df[(extracted_df['District'].str.upper() == 'BOMBALI') & extracted_df['GPS_Location'].notna()])}
                GPS Coverage: {(len(extracted_df[(extracted_df['District'].str.upper() == 'BOMBALI') & extracted_df['GPS_Location'].notna()]) / len(extracted_df[extracted_df['District'].str.upper() == 'BOMBALI']) * 100) if len(extracted_df[extracted_df['District'].str.upper() == 'BOMBALI']) > 0 else 0:.1f}%
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
                
                st.download_button(
                    label="📄 Download BOMBALI District Dashboard (Word)",
                    data=word_data,
                    file_name="BOMBALI_District_GPS_Dashboard.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except ImportError:
                st.warning("⚠️ Word export requires python-docx library. Install with: pip install python-docx")
            except Exception as e:
                st.warning(f"⚠️ Word export failed: {str(e)}")
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

# Raw data preview (optional)
show_raw_data = st.checkbox("Show raw data preview")
if show_raw_data:
    st.subheader("📄 Raw Data Preview")
    st.dataframe(extracted_df.head(20))

# Export All Dashboards as Combined Word Document
st.header("📄 Combined Word Export")

if st.button("📋 Generate Combined Word Report", help="Generate a comprehensive Word document with both districts"):
    try:
        from docx import Document
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import tempfile
        import os
        
        # Create Word document
        doc = Document()
        
        # Add main title
        title = doc.add_heading('School-Based Distribution (SBD)', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        subtitle = doc.add_heading('GPS School Locations Dashboard', level=1)
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add generation date
        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_run = date_para.add_run(f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
        date_run.font.size = Pt(12)
        date_run.bold = True
        
        doc.add_page_break()
        
        # Executive Summary
        doc.add_heading('Executive Summary', level=1)
        
        bo_records = len(extracted_df[extracted_df['District'].str.upper() == 'BO'])
        bombali_records = len(extracted_df[extracted_df['District'].str.upper() == 'BOMBALI'])
        total_gps = len(extracted_df[extracted_df['GPS_Location'].notna()])
        
        summary_text = f"""
        This comprehensive dashboard report presents GPS school location analysis for BO and BOMBALI districts:
        
        • Districts Covered: BO, BOMBALI
        • Total Records: {len(extracted_df):,}
        • BO District Records: {bo_records:,}
        • BOMBALI District Records: {bombali_records:,}
        • GPS Records: {total_gps:,}
        • Overall GPS Coverage: {(total_gps/len(extracted_df)*100) if len(extracted_df) > 0 else 0:.1f}%
        
        This report contains visual mapping of all school GPS coordinates by chiefdom. Each chiefdom is displayed with its administrative boundaries and red markers indicating school locations.
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
        if 'fig_bo' in locals():
            doc.add_heading('BO District - GPS School Locations', level=1)
            
            # Save BO figure to temporary file and add to document
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                # Save with optimized settings for Word document
                fig_bo.savefig(tmp_file.name, format='png', dpi=200, 
                              bbox_inches='tight', facecolor='white', 
                              edgecolor='none', pad_inches=0.1)
                # Add image with optimal size for Word page
                doc.add_picture(tmp_file.name, width=Inches(9.5))
                os.unlink(tmp_file.name)
            
            # BO summary
            doc.add_heading('BO District Summary', level=2)
            bo_gps_count = len(extracted_df[(extracted_df['District'].str.upper() == 'BO') & extracted_df['GPS_Location'].notna()])
            
            bo_summary_items = [
                f"Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BO'])}",
                f"Total Records: {bo_records:,}",
                f"GPS Records: {bo_gps_count:,}",
                f"GPS Coverage: {(bo_gps_count/bo_records*100) if bo_records > 0 else 0:.1f}%"
            ]
            
            for item in bo_summary_items:
                p = doc.add_paragraph()
                p.add_run('• ').bold = True
                p.add_run(item)
            
            doc.add_page_break()
        
        # BOMBALI District section
        if 'fig_bombali' in locals():
            doc.add_heading('BOMBALI District - GPS School Locations', level=1)
            
            # Save BOMBALI figure to temporary file and add to document
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                fig_bombali.savefig(tmp_file.name, format='png', dpi=300, bbox_inches='tight')
                doc.add_picture(tmp_file.name, width=Inches(10))
                os.unlink(tmp_file.name)
            
            # BOMBALI summary
            doc.add_heading('BOMBALI District Summary', level=2)
            bombali_gps_count = len(extracted_df[(extracted_df['District'].str.upper() == 'BOMBALI') & extracted_df['GPS_Location'].notna()])
            
            bombali_summary_items = [
                f"Total Chiefdoms: {len(gdf[gdf['FIRST_DNAM'] == 'BOMBALI'])}",
                f"Total Records: {bombali_records:,}",
                f"GPS Records: {bombali_gps_count:,}",
                f"GPS Coverage: {(bombali_gps_count/bombali_records*100) if bombali_records > 0 else 0:.1f}%"
            ]
            
            for item in bombali_summary_items:
                p = doc.add_paragraph()
                p.add_run('• ').bold = True
                p.add_run(item)
        
        # Save to BytesIO
        word_buffer = BytesIO()
        doc.save(word_buffer)
        word_data = word_buffer.getvalue()
        
        st.success("✅ Combined Word report generated successfully!")
        st.download_button(
            label="💾 Download Combined GPS Dashboard Report (Word)",
            data=word_data,
            file_name=f"GPS_School_Locations_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            help="Download comprehensive Word report with both districts"
        )
        
    except ImportError:
        st.error("❌ Word generation requires python-docx library. Please install it using: pip install python-docx")
    except Exception as e:
        st.error(f"❌ Error generating combined Word document: {str(e)}")

# Memory optimization - close matplotlib figures
plt.close('all')

# Footer
st.markdown("---")
st.markdown("**📊 Section 1: GPS School Locations | School-Based Distribution Analysis**")
