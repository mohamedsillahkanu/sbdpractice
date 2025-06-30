import streamlit as st
import pandas as pd
import re
import numpy as np
import matplotlib.pyplot as plt
import geopandas as gpd
from io import BytesIO
import base64
import warnings

# Suppress warnings
warnings.filterwarnings('ignore')

# Set page config
st.set_page_config(
    page_title="SBD Analysis Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS with blue and white theme and zoom functionality
st.markdown("""
<style>
    /* Allow zoom functionality */
    .stApp {
        zoom: 1 !important;
        transform: scale(1) !important;
        transform-origin: 0 0 !important;
    }
    
    /* Increase sidebar width */
    section[data-testid="stSidebar"] {
        width: 320px !important;
        background-color: #f8f9fd !important;
    }
    
    /* Main app styling with blue theme */
    .main .block-container {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        margin-left: 0 !important;
        max-width: none !important;
        background-color: white !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important;
        margin-top: 1rem !important;
        margin-bottom: 1rem !important;
    }
    
    /* Title styling */
    h1 {
        color: #2c3e50 !important;
        text-align: center !important;
        font-weight: 700 !important;
        margin-bottom: 2rem !important;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1) !important;
    }
    
    /* Subheader styling */
    h2, h3 {
        color: #34495e !important;
        border-bottom: 2px solid #3498db !important;
        padding-bottom: 0.5rem !important;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(45deg, #3498db, #2980b9) !important;
        color: white !important;
        border: none !important;
        border-radius: 25px !important;
        padding: 0.5rem 2rem !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 15px rgba(52, 152, 219, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        background: linear-gradient(45deg, #2980b9, #1f5f8b) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(52, 152, 219, 0.4) !important;
    }
    
    /* Download button styling */
    .stDownloadButton > button {
        background: linear-gradient(45deg, #27ae60, #229954) !important;
        color: white !important;
        border: none !important;
        border-radius: 25px !important;
        padding: 0.5rem 2rem !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 15px rgba(39, 174, 96, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    
    .stDownloadButton > button:hover {
        background: linear-gradient(45deg, #229954, #1e7e34) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(39, 174, 96, 0.4) !important;
    }
    
    /* Metric styling */
    [data-testid="metric-container"] {
        background: linear-gradient(135deg, #74b9ff, #0984e3) !important;
        border: 1px solid #ddd !important;
        padding: 1rem !important;
        border-radius: 10px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }
    
    [data-testid="metric-container"] > div {
        color: white !important;
    }
    
    /* Dataframe styling */
    .stDataFrame {
        border-radius: 10px !important;
        overflow: hidden !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }
    
    /* Info box styling */
    .stInfo {
        background: linear-gradient(135deg, #74b9ff, #0984e3) !important;
        color: white !important;
        border-radius: 10px !important;
    }
    
    /* Warning box styling */
    .stWarning {
        background: linear-gradient(135deg, #fdcb6e, #e17055) !important;
        color: white !important;
        border-radius: 10px !important;
    }
    
    /* Success box styling */
    .stSuccess {
        background: linear-gradient(135deg, #00b894, #00a085) !important;
        color: white !important;
        border-radius: 10px !important;
    }
</style>
""", unsafe_allow_html=True)

# Function to save maps as PNG and return BytesIO object
def save_map_as_png(fig, filename_prefix):
    """Save matplotlib figure as PNG and return BytesIO object"""
    try:
        buffer = BytesIO()
        fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
        buffer.seek(0)
        
        # Also save to disk for reference
        fig.savefig(f"{filename_prefix}.png", format='png', dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
        
        return buffer
    except Exception as e:
        st.error(f"Error saving map: {e}")
        return None

# Function to generate comprehensive summaries
@st.cache_data
def generate_summaries(df):
    """Generate District, Chiefdom, and Gender summaries"""
    summaries = {}
    
    # Overall Summary
    overall_summary = {
        'total_schools': len(df),
        'total_districts': len(df['District'].dropna().unique()) if 'District' in df.columns else 0,
        'total_chiefdoms': len(df['Chiefdom'].dropna().unique()) if 'Chiefdom' in df.columns else 0,
        'total_boys': 0,
        'total_girls': 0,
        'total_enrollment': 0,
        'total_itn': 0
    }
    
    # Calculate totals using the correct columns
    for class_num in range(1, 6):
        # Total enrollment from "Number of enrollments in class X"
        enrollment_col = f"Number of enrollments in class {class_num}"
        if enrollment_col in df.columns:
            overall_summary['total_enrollment'] += int(df[enrollment_col].fillna(0).sum())
        
        # Boys and girls for gender analysis AND ITN calculation
        boys_col = f"Number of boys in class {class_num}"
        girls_col = f"Number of girls in class {class_num}"
        if boys_col in df.columns:
            overall_summary['total_boys'] += int(df[boys_col].fillna(0).sum())
        if girls_col in df.columns:
            overall_summary['total_girls'] += int(df[girls_col].fillna(0).sum())
    
    # Total ITNs = boys + girls (actual beneficiaries)
    overall_summary['total_itn'] = overall_summary['total_boys'] + overall_summary['total_girls']
    
    # Calculate coverage
    overall_summary['coverage'] = (overall_summary['total_itn'] / overall_summary['total_enrollment'] * 100) if overall_summary['total_enrollment'] > 0 else 0
    overall_summary['itn_remaining'] = overall_summary['total_enrollment'] - overall_summary['total_itn']
    overall_summary['gender_ratio'] = (overall_summary['total_girls'] / overall_summary['total_boys'] * 100) if overall_summary['total_boys'] > 0 else 0
    
    summaries['overall'] = overall_summary
    
    # District Summary
    district_summary = []
    if 'District' in df.columns:
        for district in df['District'].dropna().unique():
            district_data = df[df['District'] == district]
            district_stats = {
                'district': district,
                'schools': len(district_data),
                'chiefdoms': len(district_data['Chiefdom'].dropna().unique()) if 'Chiefdom' in district_data.columns else 0,
                'boys': 0,
                'girls': 0,
                'enrollment': 0,
                'itn': 0
            }
            
            for class_num in range(1, 6):
                # Total enrollment from "Number of enrollments in class X"
                enrollment_col = f"Number of enrollments in class {class_num}"
                if enrollment_col in district_data.columns:
                    district_stats['enrollment'] += int(district_data[enrollment_col].fillna(0).sum())
                
                # Boys and girls for gender analysis AND ITN calculation
                boys_col = f"Number of boys in class {class_num}"
                girls_col = f"Number of girls in class {class_num}"
                if boys_col in district_data.columns:
                    district_stats['boys'] += int(district_data[boys_col].fillna(0).sum())
                if girls_col in district_data.columns:
                    district_stats['girls'] += int(district_data[girls_col].fillna(0).sum())
            
            # Total ITNs = boys + girls (actual beneficiaries)
            district_stats['itn'] = district_stats['boys'] + district_stats['girls']
            
            # Calculate coverage
            district_stats['coverage'] = (district_stats['itn'] / district_stats['enrollment'] * 100) if district_stats['enrollment'] > 0 else 0
            district_stats['itn_remaining'] = district_stats['enrollment'] - district_stats['itn']
            
            district_summary.append(district_stats)
    
    summaries['district'] = district_summary
    
    # Chiefdom Summary
    chiefdom_summary = []
    if 'District' in df.columns and 'Chiefdom' in df.columns:
        for district in df['District'].dropna().unique():
            district_data = df[df['District'] == district]
            for chiefdom in district_data['Chiefdom'].dropna().unique():
                chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
                chiefdom_stats = {
                    'district': district,
                    'chiefdom': chiefdom,
                    'schools': len(chiefdom_data),
                    'boys': 0,
                    'girls': 0,
                    'enrollment': 0,
                    'itn': 0
                }
                
                for class_num in range(1, 6):
                    # Total enrollment from "Number of enrollments in class X"
                    enrollment_col = f"Number of enrollments in class {class_num}"
                    if enrollment_col in chiefdom_data.columns:
                        chiefdom_stats['enrollment'] += int(chiefdom_data[enrollment_col].fillna(0).sum())
                    
                    # Boys and girls for gender analysis AND ITN calculation
                    boys_col = f"Number of boys in class {class_num}"
                    girls_col = f"Number of girls in class {class_num}"
                    if boys_col in chiefdom_data.columns:
                        chiefdom_stats['boys'] += int(chiefdom_data[boys_col].fillna(0).sum())
                    if girls_col in chiefdom_data.columns:
                        chiefdom_stats['girls'] += int(chiefdom_data[girls_col].fillna(0).sum())
                
                # Total ITNs = boys + girls (actual beneficiaries)
                chiefdom_stats['itn'] = chiefdom_stats['boys'] + chiefdom_stats['girls']
                
                # Calculate coverage
                chiefdom_stats['coverage'] = (chiefdom_stats['itn'] / chiefdom_stats['enrollment'] * 100) if chiefdom_stats['enrollment'] > 0 else 0
                chiefdom_stats['itn_remaining'] = chiefdom_stats['enrollment'] - chiefdom_stats['itn']
                
                chiefdom_summary.append(chiefdom_stats)
    
    summaries['chiefdom'] = chiefdom_summary
    
    return summaries

def main():
    # Logo Section - Clean 4 Logo Layout
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        try:
            st.image("NMCP.png", width=230)
            st.markdown('<p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">National Malaria Control Program</p>', unsafe_allow_html=True)
        except:
            st.markdown("""
            <div style="width: 230px; height: 160px; border: 2px dashed #3498db; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #f8f9fd, #e3f2fd); border-radius: 10px; margin: 0 auto;">
                <div style="text-align: center; color: #666; font-size: 11px;">
                    NMCP.png<br>Not Found
                </div>
            </div>
            <p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">National Malaria Control Program</p>
            """, unsafe_allow_html=True)

    with col2:
        try:
            st.image("icf_sl.png", width=230)
            st.markdown('<p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">ICF Sierra Leone</p>', unsafe_allow_html=True)
        except:
            st.markdown("""
            <div style="width: 230px; height: 160px; border: 2px dashed #3498db; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #f8f9fd, #e3f2fd); border-radius: 10px; margin: 0 auto;">
                <div style="text-align: center; color: #666; font-size: 11px;">
                    icf_sl.png<br>Not Found
                </div>
            </div>
            <p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">ICF Sierra Leone</p>
            """, unsafe_allow_html=True)

    with col3:
        try:
            st.image("pmi.png", width=230)
            st.markdown('<p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">PMI Evolve</p>', unsafe_allow_html=True)
        except:
            st.markdown("""
            <div style="width: 230px; height: 160px; border: 2px dashed #3498db; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #f8f9fd, #e3f2fd); border-radius: 10px; margin: 0 auto;">
                <div style="text-align: center; color: #666; font-size: 11px;">
                    pmi.png<br>Not Found
                </div>
            </div>
            <p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">PMI Evolve</p>
            """, unsafe_allow_html=True)

    with col4:
        try:
            st.image("abt.png", width=230)
            st.markdown('<p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">Abt Associates</p>', unsafe_allow_html=True)
        except:
            st.markdown("""
            <div style="width: 230px; height: 160px; border: 2px dashed #3498db; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #f8f9fd, #e3f2fd); border-radius: 10px; margin: 0 auto;">
                <div style="text-align: center; color: #666; font-size: 11px;">
                    abt.png<br>Not Found
                </div>
            </div>
            <p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">Abt Associates</p>
            """, unsafe_allow_html=True)

    st.markdown("---")  # Add a horizontal line separator

    # Streamlit App
    st.title("📊 School Based Distribution of ITNs in SL")

    # Load embedded file directly
    try:
        df_original = pd.read_excel("sbd first_submission_clean.xlsx")
        st.success("✅ Data loaded successfully!")
    except FileNotFoundError:
        st.error("❌ Data file 'sbd first_submission_clean.xlsx' not found. Please ensure the file is in the correct directory.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Error reading data file: {e}")
        st.stop()

    # Validate required columns
    required_columns = ["Scan QR code"]
    missing_columns = [col for col in required_columns if col not in df_original.columns]
    if missing_columns:
        st.error(f"❌ Missing required columns: {missing_columns}")
        st.write("Available columns:", list(df_original.columns))
        st.stop()

    # Load shapefile
    gdf = None
    try:
        gdf = gpd.read_file("Chiefdom2021.shp")
        st.success("✅ Shapefile loaded successfully!")
    except FileNotFoundError:
        st.warning("⚠️ Shapefile 'Chiefdom2021.shp' not found. Maps will not be displayed.")
    except Exception as e:
        st.warning(f"⚠️ Could not load shapefile: {e}")

    # Create empty lists to store extracted data
    districts, chiefdoms, phu_names, community_names, school_names, enrollment_values = [], [], [], [], [], []

    # Process each row in the "Scan QR code" column
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_rows = len(df_original)
    
    for idx, qr_text in enumerate(df_original["Scan QR code"]):
        # Update progress
        progress = (idx + 1) / total_rows
        progress_bar.progress(progress)
        status_text.text(f"Processing QR codes... {idx + 1}/{total_rows}")
        
        if pd.isna(qr_text):
            districts.append(None)
            chiefdoms.append(None)
            phu_names.append(None)
            community_names.append(None)
            school_names.append(None)
            enrollment_values.append(None)
            continue
            
        # Extract values using regex patterns
        qr_str = str(qr_text)
        
        district_match = re.search(r"District:\s*([^\n]+)", qr_str)
        districts.append(district_match.group(1).strip() if district_match else None)
        
        chiefdom_match = re.search(r"Chiefdom:\s*([^\n]+)", qr_str)
        chiefdoms.append(chiefdom_match.group(1).strip() if chiefdom_match else None)
        
        phu_match = re.search(r"PHU name:\s*([^\n]+)", qr_str)
        phu_names.append(phu_match.group(1).strip() if phu_match else None)
        
        community_match = re.search(r"Community name:\s*([^\n]+)", qr_str)
        community_names.append(community_match.group(1).strip() if community_match else None)
        
        school_match = re.search(r"Name of school:\s*([^\n]+)", qr_str)
        school_names.append(school_match.group(1).strip() if school_match else None)

        enrollment_match = re.search(r"Enrollment:\s*([^\n]+)", qr_str)
        enrollment_values.append(enrollment_match.group(1).strip() if enrollment_match else None)

    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()

    # Create a new DataFrame with extracted values
    try:
        extracted_df = pd.DataFrame({
            "District": districts,
            "Chiefdom": chiefdoms,
            "PHU Name": phu_names,
            "Community Name": community_names,
            "School Name": school_names,
            "Enrollment": enrollment_values
        })
        
        # Add all other columns from the original DataFrame
        for column in df_original.columns:
            if column != "Scan QR code":  # Skip the QR code column since we've already processed it
                extracted_df[column] = df_original[column]
                
        st.success(f"✅ Successfully processed {len(extracted_df)} records")
        
    except Exception as e:
        st.error(f"❌ Error creating extracted DataFrame: {e}")
        st.stop()

    # Create sidebar filters
    st.sidebar.header("🔍 Filter Options")

    # Create radio buttons to select which level to group by
    grouping_options = ["District", "Chiefdom", "PHU Name", "Community Name", "School Name"]
    available_options = [opt for opt in grouping_options if opt in extracted_df.columns and not extracted_df[opt].isna().all()]
    
    if available_options:
        grouping_selection = st.sidebar.radio(
            "Select the level for grouping:",
            available_options,
            index=0
        )
    else:
        st.warning("No grouping options available")
        grouping_selection = "District"

    # Dictionary to define the hierarchy for each grouping level
    hierarchy = {
        "District": ["District"],
        "Chiefdom": ["District", "Chiefdom"],
        "PHU Name": ["District", "Chiefdom", "PHU Name"],
        "Community Name": ["District", "Chiefdom", "PHU Name", "Community Name"],
        "School Name": ["District", "Chiefdom", "PHU Name", "Community Name", "School Name"]
    }

    # Initialize filtered dataframe with the full dataset
    filtered_df = extracted_df.copy()

    # Dictionary to store selected values for each level
    selected_values = {}

    # Apply filters based on the hierarchy for the selected grouping level
    if grouping_selection in hierarchy:
        for level in hierarchy[grouping_selection]:
            if level in filtered_df.columns:
                # Filter out None/NaN values and get sorted unique values
                level_values = sorted([x for x in filtered_df[level].dropna().unique() if x is not None])
                
                if level_values:
                    # Create selectbox for this level
                    selected_value = st.sidebar.selectbox(f"Select {level}", level_values)
                    selected_values[level] = selected_value
                    
                    # Apply filter to the dataframe
                    filtered_df = filtered_df[filtered_df[level] == selected_value]

    # Store map images for report
    map_images = {}

    # Display maps section
    if gdf is not None:
        st.subheader("🗺️ Geographic Distribution Maps")
        
        try:
            # OVERALL SIERRA LEONE MAP
            st.write("**Sierra Leone - All Districts Overview**")
            
            # Create overall Sierra Leone map
            fig_overall, ax_overall = plt.subplots(figsize=(16, 10))
            
            # Plot all chiefdoms with gray edges (base layer)
            gdf.plot(ax=ax_overall, color='white', edgecolor='gray', alpha=0.8, linewidth=0.5)
            
            # Plot district boundaries with thick black lines
            if 'FIRST_DNAM' in gdf.columns:
                district_boundaries = gdf.dissolve(by='FIRST_DNAM')
                district_boundaries.plot(ax=ax_overall, facecolor='none', edgecolor='black', linewidth=3, alpha=1.0)
                
                # Add district labels at centroids
                for idx, row in district_boundaries.iterrows():
                    centroid = row.geometry.centroid
                    ax_overall.annotate(
                        idx,  # District name
                        (centroid.x, centroid.y),
                        fontsize=12,
                        fontweight='bold',
                        ha='center',
                        va='center',
                        color='black',
                        bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='black')
                    )
            
            # Extract and plot GPS coordinates
            all_coords_extracted = []
            if "GPS Location" in extracted_df.columns:
                all_gps_data = extracted_df["GPS Location"].dropna()
                
                for idx, gps_val in enumerate(all_gps_data):
                    if pd.notna(gps_val):
                        gps_str = str(gps_val).strip()
                        
                        # Handle the specific format: 8.6103181,-12.2029534
                        if ',' in gps_str:
                            try:
                                parts = gps_str.split(',')
                                if len(parts) == 2:
                                    lat = float(parts[0].strip())
                                    lon = float(parts[1].strip())
                                    
                                    # Check if coordinates are in valid range for Sierra Leone
                                    if 6.0 <= lat <= 11.0 and -14.0 <= lon <= -10.0:
                                        all_coords_extracted.append([lat, lon])
                            except ValueError:
                                continue
            
            # Plot GPS points on the overall map
            if all_coords_extracted:
                lats, lons = zip(*all_coords_extracted)
                
                # Plot GPS points
                ax_overall.scatter(
                    lons, lats,
                    c='#47B5FF',
                    s=100,
                    alpha=0.9,
                    edgecolors='white',
                    linewidth=2,
                    zorder=100,
                    label=f'Schools ({len(all_coords_extracted)})',
                    marker='o'
                )
                
                # Add legend
                ax_overall.legend(fontsize=14, loc='best')
            
            # Customize overall map
            ax_overall.set_title('Sierra Leone - School Distribution by District', fontsize=18, fontweight='bold', pad=20)
            ax_overall.set_xlabel('Longitude', fontsize=14)
            ax_overall.set_ylabel('Latitude', fontsize=14)
            ax_overall.grid(True, alpha=0.3, linestyle='--')
            
            # Set axis limits to show full country
            ax_overall.set_xlim(gdf.total_bounds[0] - 0.1, gdf.total_bounds[2] + 0.1)
            ax_overall.set_ylim(gdf.total_bounds[1] - 0.1, gdf.total_bounds[3] + 0.1)
            
            plt.tight_layout()
            st.pyplot(fig_overall)
            
            # Save overall map
            map_buffer = save_map_as_png(fig_overall, "Sierra_Leone_Overall_Map")
            if map_buffer:
                map_images['sierra_leone_overall'] = map_buffer
            plt.close(fig_overall)
            
        except Exception as e:
            st.error(f"Error creating overall map: {e}")

    # Display Original Data Sample
    st.subheader("📄 Original Data Sample")
    st.dataframe(df_original.head())

    # Display Extracted Data
    st.subheader("📋 Extracted Data")
    st.dataframe(extracted_df)

    # Add download button for CSV
    csv = extracted_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Extracted Data as CSV",
        data=csv,
        file_name="extracted_school_data.csv",
        mime="text/csv"
    )

    # Generate comprehensive summaries
    try:
        summaries = generate_summaries(extracted_df)
        
        # Display Overall Summary
        st.subheader("📊 Overall Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Schools", f"{summaries['overall']['total_schools']:,}")
        with col2:
            st.metric("Total Students", f"{summaries['overall']['total_enrollment']:,}")
        with col3:
            st.metric("Total ITNs", f"{summaries['overall']['total_itn']:,}")
        with col4:
            st.metric("Coverage", f"{summaries['overall']['coverage']:.1f}%")

        col5, col6, col7, col8 = st.columns(4)
        with col5:
            st.metric("Districts", f"{summaries['overall']['total_districts']}")
        with col6:
            st.metric("Chiefdoms", f"{summaries['overall']['total_chiefdoms']}")
        with col7:
            st.metric("Boys", f"{summaries['overall']['total_boys']:,}")
        with col8:
            st.metric("Girls", f"{summaries['overall']['total_girls']:,}")

        # Gender Analysis
        st.subheader("👫 Gender Analysis")
        
        # Check if we have valid gender data
        total_boys = summaries['overall']['total_boys']
        total_girls = summaries['overall']['total_girls']
        
        if total_boys > 0 or total_girls > 0:
            # Overall gender distribution pie chart
            fig_gender, ax_gender = plt.subplots(figsize=(10, 8))
            labels = ['Boys', 'Girls']
            sizes = [total_boys, total_girls]
            colors = ['#4A90E2', '#F39C12']
            
            # Filter out zero values
            non_zero_data = [(label, size, color) for label, size, color in zip(labels, sizes, colors) if size > 0]
            
            if non_zero_data:
                labels_filtered, sizes_filtered, colors_filtered = zip(*non_zero_data)
                
                wedges, texts, autotexts = ax_gender.pie(sizes_filtered, labels=labels_filtered, autopct='%1.1f%%', 
                                                        colors=colors_filtered, startangle=90)
                ax_gender.set_title('Overall Gender Distribution', fontsize=16, fontweight='bold', pad=20)
                plt.setp(autotexts, size=14, weight="bold")
                plt.setp(texts, size=12, weight="bold")
                plt.tight_layout()
                st.pyplot(fig_gender)
                
                # Save gender chart
                gender_buffer = save_map_as_png(fig_gender, "Overall_Gender_Distribution")
                if gender_buffer:
                    map_images['gender_overall'] = gender_buffer
                plt.close(fig_gender)
            else:
                st.warning("No gender data available for pie chart")
        else:
            st.warning("No gender data available for analysis")

        # Gender ratio by district chart
        if summaries['district']:
            districts = [d['district'] for d in summaries['district']]
            boys_counts = [d['boys'] for d in summaries['district']]
            girls_counts = [d['girls'] for d in summaries['district']]
            
            # Check if we have any non-zero values
            if any(boys_counts) or any(girls_counts):
                fig_gender_district, ax_gender_district = plt.subplots(figsize=(14, 8))
                x = np.arange(len(districts))
                width = 0.35
                
                bars1 = ax_gender_district.bar(x - width/2, boys_counts, width, label='Boys', color='#4A90E2', edgecolor='navy', linewidth=1)
                bars2 = ax_gender_district.bar(x + width/2, girls_counts, width, label='Girls', color='#F39C12', edgecolor='darkorange', linewidth=1)
                
                ax_gender_district.set_title('Gender Distribution by District', fontsize=16, fontweight='bold', pad=20)
                ax_gender_district.set_xlabel('Districts', fontsize=12, fontweight='bold')
                ax_gender_district.set_ylabel('Number of Students', fontsize=12, fontweight='bold')
                ax_gender_district.set_xticks(x)
                ax_gender_district.set_xticklabels(districts, rotation=45, ha='right')
                ax_gender_district.legend(fontsize=12)
                ax_gender_district.grid(axis='y', alpha=0.3, linestyle='--')
                
                # Add value labels on bars
                for bar in bars1:
                    height = bar.get_height()
                    if height > 0:
                        ax_gender_district.annotate(f'{int(height):,}',
                                                  xy=(bar.get_x() + bar.get_width() / 2, height),
                                                  xytext=(0, 3),
                                                  textcoords="offset points",
                                                  ha='center', va='bottom', fontsize=10, fontweight='bold')
                
                for bar in bars2:
                    height = bar.get_height()
                    if height > 0:
                        ax_gender_district.annotate(f'{int(height):,}',
                                                  xy=(bar.get_x() + bar.get_width() / 2, height),
                                                  xytext=(0, 3),
                                                  textcoords="offset points",
                                                  ha='center', va='bottom', fontsize=10, fontweight='bold')
                
                plt.tight_layout()
                st.pyplot(fig_gender_district)
                
                # Save gender district chart
                gender_district_buffer = save_map_as_png(fig_gender_district, "Gender_Distribution_by_District")
                if gender_district_buffer:
                    map_images['gender_district'] = gender_district_buffer
                plt.close(fig_gender_district)
            else:
                st.warning("No district gender data available for chart")

        # Enrollment and ITN Distribution Analysis
        st.subheader("📊 Enrollment and ITN Distribution Analysis")
        
        if 'District' in extracted_df.columns:
            # Calculate total enrollment and ITN distribution by district
            district_analysis = []
            
            for district in extracted_df['District'].dropna().unique():
                district_data = extracted_df[extracted_df['District'] == district]
                
                total_enrollment = 0
                total_boys = 0
                total_girls = 0
                
                # Sum enrollments and boys/girls by class using correct columns
                for class_num in range(1, 6):
                    # Use "Number of enrollments in class X" for total students
                    enrollment_col = f"Number of enrollments in class {class_num}"
                    if enrollment_col in district_data.columns:
                        total_enrollment += int(district_data[enrollment_col].fillna(0).sum())
                    
                    # Use boys + girls for total ITNs (actual beneficiaries)
                    boys_col = f"Number of boys in class {class_num}"
                    girls_col = f"Number of girls in class {class_num}"
                    if boys_col in district_data.columns:
                        total_boys += int(district_data[boys_col].fillna(0).sum())
                    if girls_col in district_data.columns:
                        total_girls += int(district_data[girls_col].fillna(0).sum())
                
                total_itn = total_boys + total_girls  # Total ITNs = boys + girls
                itn_remaining = max(0, total_enrollment - total_itn)  # Remaining = enrollment - distributed
                coverage = (total_itn / total_enrollment * 100) if total_enrollment > 0 else 0
                
                district_analysis.append({
                    'District': district,
                    'Total_Enrollment': total_enrollment,
                    'Total_ITN': total_itn,
                    'ITN_Remaining': itn_remaining,
                    'Coverage': coverage
                })
            
            if district_analysis:
                district_df = pd.DataFrame(district_analysis)
                
                # Create enhanced bar chart with enrollment, distributed, and remaining
                fig_enhanced, ax_enhanced = plt.subplots(figsize=(16, 8))
                
                x = np.arange(len(district_df['District']))
                width = 0.25
                
                # Create bars for each category
                bars1 = ax_enhanced.bar(x - width, district_df['Total_Enrollment'], width, 
                                       label='Total Enrollment', color='#47B5FF', edgecolor='navy', linewidth=1)
                bars2 = ax_enhanced.bar(x, district_df['Total_ITN'], width, 
                                       label='ITNs Distributed (Boys + Girls)', color='lightcoral', edgecolor='darkred', linewidth=1)
                bars3 = ax_enhanced.bar(x + width, district_df['ITN_Remaining'], width, 
                                       label='ITNs Remaining', color='hotpink', edgecolor='darkmagenta', linewidth=1)
                
                # Customize the chart
                ax_enhanced.set_title('District Analysis: Enrollment vs ITN Distribution', fontsize=16, fontweight='bold', pad=20)
                ax_enhanced.set_xlabel('Districts', fontsize=12, fontweight='bold')
                ax_enhanced.set_ylabel('Number of Students/ITNs', fontsize=12, fontweight='bold')
                ax_enhanced.set_xticks(x)
                ax_enhanced.set_xticklabels(district_df['District'], rotation=45, ha='right')
                ax_enhanced.legend(fontsize=12)
                ax_enhanced.grid(axis='y', alpha=0.3, linestyle='--')
                
                # Add value labels on bars
                for bar in bars1:
                    height = bar.get_height()
                    if height > 0:
                        ax_enhanced.annotate(f'{int(height):,}',
                                            xy=(bar.get_x() + bar.get_width() / 2, height),
                                            xytext=(0, 3),
                                            textcoords="offset points",
                                            ha='center', va='bottom', fontsize=9, fontweight='bold')
                
                for bar in bars2:
                    height = bar.get_height()
                    if height > 0:
                        ax_enhanced.annotate(f'{int(height):,}',
                                            xy=(bar.get_x() + bar.get_width() / 2, height),
                                            xytext=(0, 3),
                                            textcoords="offset points",
                                            ha='center', va='bottom', fontsize=9, fontweight='bold')
                
                for bar in bars3:
                    height = bar.get_height()
                    if height > 0:
                        ax_enhanced.annotate(f'{int(height):,}',
                                            xy=(bar.get_x() + bar.get_width() / 2, height),
                                            xytext=(0, 3),
                                            textcoords="offset points",
                                            ha='center', va='bottom', fontsize=9, fontweight='bold')
                
                plt.tight_layout()
                st.pyplot(fig_enhanced)
                
                # Save enhanced chart
                enhanced_buffer = save_map_as_png(fig_enhanced, "Enhanced_Enrollment_Analysis")
                if enhanced_buffer:
                    map_images['enhanced_enrollment_analysis'] = enhanced_buffer
                plt.close(fig_enhanced)
                
                # Create overall pie chart for enrollment vs distributed vs remaining
                st.subheader("📊 Overall Distribution Overview (Pie Chart)")
                
                # Calculate overall totals
                overall_enrollment = district_df['Total_Enrollment'].sum()
                overall_distributed = district_df['Total_ITN'].sum()
                overall_remaining = district_df['ITN_Remaining'].sum()
                
                if overall_enrollment > 0 and (overall_distributed > 0 or overall_remaining > 0):
                    fig_overall_pie, ax_overall_pie = plt.subplots(figsize=(10, 8))
                    
                    sizes = []
                    labels = []
                    colors = []
                    
                    if overall_distributed > 0:
                        sizes.append(overall_distributed)
                        labels.append(f'ITNs Distributed\n({overall_distributed:,})')
                        colors.append('lightcoral')
                    
                    if overall_remaining > 0:
                        sizes.append(overall_remaining)
                        labels.append(f'ITNs Remaining\n({overall_remaining:,})')
                        colors.append('hotpink')
                    
                    if sizes:
                        explode = [0.05] + [0] * (len(sizes) - 1)  # Slightly separate the first slice
                        
                        wedges, texts, autotexts = ax_overall_pie.pie(sizes, labels=labels, autopct='%1.1f%%',
                                                                     colors=colors, startangle=90, explode=explode)
                        ax_overall_pie.set_title(f'Overall ITN Distribution Status\nTotal Enrollment: {overall_enrollment:,}', 
                                                fontsize=16, fontweight='bold', pad=20)
                        
                        # Enhance text styling
                        plt.setp(autotexts, size=12, weight="bold", color='white')
                        plt.setp(texts, size=11, weight="bold")
                        
                        plt.tight_layout()
                        st.pyplot(fig_overall_pie)
                        
                        # Save overall pie chart
                        overall_pie_buffer = save_map_as_png(fig_overall_pie, "Overall_Distribution_Pie")
                        if overall_pie_buffer:
                            map_images['overall_distribution_pie'] = overall_pie_buffer
                        plt.close(fig_overall_pie)

        # Display Summary Tables
        st.subheader("📈 District Summary Table")
        if summaries['district']:
            district_summary_df = pd.DataFrame(summaries['district'])
            st.dataframe(district_summary_df)
        else:
            st.info("No district data available")

        st.subheader("📈 Chiefdom Summary Table")
        if summaries['chiefdom']:
            chiefdom_summary_df = pd.DataFrame(summaries['chiefdom'])
            st.dataframe(chiefdom_summary_df)
        else:
            st.info("No chiefdom data available")

        # Summary buttons section
        st.subheader("📊 Summary Reports")
        
        # Create two columns for the summary buttons
        col1, col2 = st.columns(2)
        
        # Button for District Summary
        with col1:
            if st.button("Show District Summary"):
                st.subheader("📈 Summary by District")
                
                if 'District' in extracted_df.columns:
                    # Create aggregation dictionary
                    agg_dict = {}
                    
                    # Add enrollment columns to aggregation
                    for class_num in range(1, 6):
                        total_col = f"Number of enrollments in class {class_num}"
                        boys_col = f"Number of boys in class {class_num}"
                        girls_col = f"Number of girls in class {class_num}"
                        
                        if total_col in extracted_df.columns:
                            agg_dict[total_col] = "sum"
                        if boys_col in extracted_df.columns:
                            agg_dict[boys_col] = "sum"
                        if girls_col in extracted_df.columns:
                            agg_dict[girls_col] = "sum"
                    
                    if agg_dict:
                        # Group by District and aggregate
                        district_summary = extracted_df.groupby("District").agg(agg_dict).reset_index()
                        
                        # Calculate total enrollment
                        district_summary["Total Enrollment"] = 0
                        for class_num in range(1, 6):
                            total_col = f"Number of enrollments in class {class_num}"
                            if total_col in district_summary.columns:
                                district_summary["Total Enrollment"] += district_summary[total_col].fillna(0)
                        
                        # Display summary table
                        st.dataframe(district_summary)
                        
                        # Download button for district summary
                        district_csv = district_summary.to_csv(index=False)
                        st.download_button(
                            label="📥 Download District Summary as CSV",
                            data=district_csv,
                            file_name="district_summary.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No enrollment data available for district summary")
                else:
                    st.info("No district data available")
        
        # Button for Chiefdom Summary
        with col2:
            if st.button("Show Chiefdom Summary"):
                st.subheader("📈 Summary by Chiefdom")
                
                if 'District' in extracted_df.columns and 'Chiefdom' in extracted_df.columns:
                    # Create aggregation dictionary
                    agg_dict = {}
                    
                    # Add enrollment columns to aggregation
                    for class_num in range(1, 6):
                        total_col = f"Number of enrollments in class {class_num}"
                        boys_col = f"Number of boys in class {class_num}"
                        girls_col = f"Number of girls in class {class_num}"
                        
                        if total_col in extracted_df.columns:
                            agg_dict[total_col] = "sum"
                        if boys_col in extracted_df.columns:
                            agg_dict[boys_col] = "sum"
                        if girls_col in extracted_df.columns:
                            agg_dict[girls_col] = "sum"
                    
                    if agg_dict:
                        # Group by District and Chiefdom and aggregate
                        chiefdom_summary = extracted_df.groupby(["District", "Chiefdom"]).agg(agg_dict).reset_index()
                        
                        # Calculate total enrollment
                        chiefdom_summary["Total Enrollment"] = 0
                        for class_num in range(1, 6):
                            total_col = f"Number of enrollments in class {class_num}"
                            if total_col in chiefdom_summary.columns:
                                chiefdom_summary["Total Enrollment"] += chiefdom_summary[total_col].fillna(0)
                        
                        # Display summary table
                        st.dataframe(chiefdom_summary)
                        
                        # Download button for chiefdom summary
                        chiefdom_csv = chiefdom_summary.to_csv(index=False)
                        st.download_button(
                            label="📥 Download Chiefdom Summary as CSV",
                            data=chiefdom_csv,
                            file_name="chiefdom_summary.csv",
                            mime="text/csv"
                        )
                    else:
                        st.info("No enrollment data available for chiefdom summary")
                else:
                    st.info("No chiefdom data available")

        # Final Data Export Section
        st.subheader("📥 Export Complete Dataset")
        st.write("Download the complete extracted dataset in your preferred format:")

        # Create download buttons in columns
        download_col1, download_col2 = st.columns(2)

        with download_col1:
            # CSV Download
            csv_data = extracted_df.to_csv(index=False)
            st.download_button(
                label="📄 Download Complete Data as CSV",
                data=csv_data,
                file_name="complete_extracted_data.csv",
                mime="text/csv",
                help="Download all extracted data in CSV format"
            )

        with download_col2:
            # Excel Download
            try:
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    extracted_df.to_excel(writer, sheet_name='Extracted Data', index=False)
                excel_data = excel_buffer.getvalue()
                
                st.download_button(
                    label="📊 Download Complete Data as Excel",
                    data=excel_data,
                    file_name="complete_extracted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download all extracted data in Excel format"
                )
            except Exception as e:
                st.error(f"Error creating Excel file: {e}")

        # Display final summary
        st.info(f"📋 **Dataset Summary**: {len(extracted_df)} total records processed with comprehensive analysis")
        
        # Display map files saved notification
        if map_images:
            st.success(f"✅ **Visualizations Created**: {len(map_images)} charts and maps have been generated")
            
            # Show list of saved maps
            with st.expander("📁 View Generated Visualizations"):
                for map_name in map_images.keys():
                    st.write(f"• {map_name}")

    except Exception as e:
        st.error(f"❌ Error generating summaries: {e}")
        st.info("Please check your data format and try again.")

if __name__ == "__main__":
    main()
