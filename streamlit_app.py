### PART 1: IMPORTS, CSS, AND CORE FUNCTIONS ###

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

# Updated column mapping based on exact column names provided
COLUMN_MAPPING = {
    'enrollment_cols': {
        1: "How many pupils are enrolled in Class 1?",
        2: "How many pupils are enrolled in Class 2?",
        3: "How many pupils are enrolled in Class 3?",
        4: "How many pupils are enrolled in Class 4?",
        5: "How many pupils are enrolled in Class 5?"
    },
    'boys_cols': {
        1: "How many boys are in Class 1?",
        2: "How many boys are in Class 2?",
        3: "How many boys are in Class 3?",
        4: "How many boys are in Class 4?",
        5: "How many boys are in Class 5?"
    },
    'girls_cols': {
        1: "How many girls are in Class 1?",
        2: "How many girls are in Class 2?",
        3: "How many girls are in Class 3?",
        4: "How many girls are in Class 4?",
        5: "How many girls are in Class 5?"
    },
    'boys_itn_cols': {
        1: "How many boys in Class 1 received ITNs?",
        2: "How many boys in Class 2 received ITNs?",
        3: "How many boys in Class 3 received ITNs?",
        4: "How many boys in Class 4 received ITNs?",
        5: "How many boys in Class 5 received ITNs?"
    },
    'girls_itn_cols': {
        1: "How many girls in Class 1 received ITNs?",
        2: "How many girls in Class 2 received ITNs?",
        3: "How many girls in Class 3 received ITNs?",
        4: "How many girls in Class 4 received ITNs?",
        5: "How many girls in Class 5 received ITNs?"
    },
    'other_cols': {
        'itns_received': "Number of ITNs received in the school",
        'total_itns_distributed': "Total ITNs distributed",
        'itns_remaining': "ITNs remaining",
        'itns_left_absent': "ITNs left at the school for pupils who were absent.",
        'itns_taken_phu': "ITNs taken to the PHU",
        'gps_location': "GPS Location",
        'school_ownership': "School ownership"
    }
}

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

# Function to generate comprehensive summaries with updated column names
@st.cache_data
def generate_summaries(df):
    """Generate District, Chiefdom, and Gender summaries using actual column names"""
    summaries = {}
    
    # Overall Summary
    overall_summary = {
        'total_schools': len(df),
        'total_districts': len(df['District'].dropna().unique()) if 'District' in df.columns else 0,
        'total_chiefdoms': len(df['Chiefdom'].dropna().unique()) if 'Chiefdom' in df.columns else 0,
        'total_boys': 0,
        'total_girls': 0,
        'total_enrollment': 0,
        'total_itn_boys': 0,
        'total_itn_girls': 0,
        'total_itn': 0
    }
    
    # Calculate totals using the updated column names
    for class_num in range(1, 6):
        # Total enrollment
        enrollment_col = COLUMN_MAPPING['enrollment_cols'].get(class_num)
        if enrollment_col and enrollment_col in df.columns:
            overall_summary['total_enrollment'] += int(df[enrollment_col].fillna(0).sum())
        
        # Boys and girls totals
        boys_col = COLUMN_MAPPING['boys_cols'].get(class_num)
        girls_col = COLUMN_MAPPING['girls_cols'].get(class_num)
        if boys_col and boys_col in df.columns:
            overall_summary['total_boys'] += int(df[boys_col].fillna(0).sum())
        if girls_col and girls_col in df.columns:
            overall_summary['total_girls'] += int(df[girls_col].fillna(0).sum())
        
        # ITNs distributed to boys and girls
        boys_itn_col = COLUMN_MAPPING['boys_itn_cols'].get(class_num)
        girls_itn_col = COLUMN_MAPPING['girls_itn_cols'].get(class_num)
        if boys_itn_col and boys_itn_col in df.columns:
            overall_summary['total_itn_boys'] += int(df[boys_itn_col].fillna(0).sum())
        if girls_itn_col and girls_itn_col in df.columns:
            overall_summary['total_itn_girls'] += int(df[girls_itn_col].fillna(0).sum())
    
    # Total ITNs distributed
    overall_summary['total_itn'] = overall_summary['total_itn_boys'] + overall_summary['total_itn_girls']
    
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
                'itn_boys': 0,
                'itn_girls': 0,
                'itn': 0
            }
            
            for class_num in range(1, 6):
                # Total enrollment
                enrollment_col = COLUMN_MAPPING['enrollment_cols'].get(class_num)
                if enrollment_col and enrollment_col in district_data.columns:
                    district_stats['enrollment'] += int(district_data[enrollment_col].fillna(0).sum())
                
                # Boys and girls
                boys_col = COLUMN_MAPPING['boys_cols'].get(class_num)
                girls_col = COLUMN_MAPPING['girls_cols'].get(class_num)
                if boys_col and boys_col in district_data.columns:
                    district_stats['boys'] += int(district_data[boys_col].fillna(0).sum())
                if girls_col and girls_col in district_data.columns:
                    district_stats['girls'] += int(district_data[girls_col].fillna(0).sum())
                
                # ITNs distributed
                boys_itn_col = COLUMN_MAPPING['boys_itn_cols'].get(class_num)
                girls_itn_col = COLUMN_MAPPING['girls_itn_cols'].get(class_num)
                if boys_itn_col and boys_itn_col in district_data.columns:
                    district_stats['itn_boys'] += int(district_data[boys_itn_col].fillna(0).sum())
                if girls_itn_col and girls_itn_col in district_data.columns:
                    district_stats['itn_girls'] += int(district_data[girls_itn_col].fillna(0).sum())
            
            # Total ITNs distributed
            district_stats['itn'] = district_stats['itn_boys'] + district_stats['itn_girls']
            
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
                    'itn_boys': 0,
                    'itn_girls': 0,
                    'itn': 0
                }
                
                for class_num in range(1, 6):
                    # Total enrollment
                    enrollment_col = COLUMN_MAPPING['enrollment_cols'].get(class_num)
                    if enrollment_col and enrollment_col in chiefdom_data.columns:
                        chiefdom_stats['enrollment'] += int(chiefdom_data[enrollment_col].fillna(0).sum())
                    
                    # Boys and girls
                    boys_col = COLUMN_MAPPING['boys_cols'].get(class_num)
                    girls_col = COLUMN_MAPPING['girls_cols'].get(class_num)
                    if boys_col and boys_col in chiefdom_data.columns:
                        chiefdom_stats['boys'] += int(chiefdom_data[boys_col].fillna(0).sum())
                    if girls_col and girls_col in chiefdom_data.columns:
                        chiefdom_stats['girls'] += int(chiefdom_data[girls_col].fillna(0).sum())
                    
                    # ITNs distributed
                    boys_itn_col = COLUMN_MAPPING['boys_itn_cols'].get(class_num)
                    girls_itn_col = COLUMN_MAPPING['girls_itn_cols'].get(class_num)
                    if boys_itn_col and boys_itn_col in chiefdom_data.columns:
                        chiefdom_stats['itn_boys'] += int(chiefdom_data[boys_itn_col].fillna(0).sum())
                    if girls_itn_col and girls_itn_col in chiefdom_data.columns:
                        chiefdom_stats['itn_girls'] += int(chiefdom_data[girls_itn_col].fillna(0).sum())
                
                # Total ITNs distributed
                chiefdom_stats['itn'] = chiefdom_stats['itn_boys'] + chiefdom_stats['itn_girls']
                
                # Calculate coverage
                chiefdom_stats['coverage'] = (chiefdom_stats['itn'] / chiefdom_stats['enrollment'] * 100) if chiefdom_stats['enrollment'] > 0 else 0
                chiefdom_stats['itn_remaining'] = chiefdom_stats['enrollment'] - chiefdom_stats['itn']
                
                chiefdom_summary.append(chiefdom_stats)
    
    summaries['chiefdom'] = chiefdom_summary
    
    return summaries


### PART 2: MAIN APPLICATION LOGIC ###

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

    # Add debugging to see actual column names and data
    st.subheader("🔍 Debug Information")
    
    # Show first few rows of key columns
    st.write("**Sample of key columns:**")
    debug_cols = []
    
    # Find columns that might contain the data
    for col in extracted_df.columns:
        col_str = str(col).lower()
        if any(keyword in col_str for keyword in ['class', 'boys', 'girls', 'pupils', 'enrolled']):
            debug_cols.append(col)
    
    if debug_cols:
        st.write("**Found columns with class/gender keywords:**")
        for i, col in enumerate(debug_cols[:15]):  # Show first 15 relevant columns
            sample_values = extracted_df[col].dropna().head(3).tolist()
            st.write(f"{i+1}. **{col}** → Sample values: {sample_values}")
    
    # Check if our mapped columns exist and have data
    st.write("**Column Mapping Check:**")
    mapping_results = {}
    
    for class_num in range(1, 6):
        st.write(f"**Class {class_num}:**")
        
        # Check enrollment column
        enrollment_col = COLUMN_MAPPING['enrollment_cols'].get(class_num)
        if enrollment_col in extracted_df.columns:
            total_enrollment = extracted_df[enrollment_col].fillna(0).sum()
            st.write(f"  ✅ Enrollment: {enrollment_col} → Total: {total_enrollment}")
            mapping_results[f'enrollment_{class_num}'] = total_enrollment
        else:
            st.write(f"  ❌ Enrollment column not found: {enrollment_col}")
        
        # Check boys column
        boys_col = COLUMN_MAPPING['boys_cols'].get(class_num)
        if boys_col in extracted_df.columns:
            total_boys = extracted_df[boys_col].fillna(0).sum()
            st.write(f"  ✅ Boys: {boys_col} → Total: {total_boys}")
            mapping_results[f'boys_{class_num}'] = total_boys
        else:
            st.write(f"  ❌ Boys column not found: {boys_col}")
        
        # Check girls column
        girls_col = COLUMN_MAPPING['girls_cols'].get(class_num)
        if girls_col in extracted_df.columns:
            total_girls = extracted_df[girls_col].fillna(0).sum()
            st.write(f"  ✅ Girls: {girls_col} → Total: {total_girls}")
            mapping_results[f'girls_{class_num}'] = total_girls
        else:
            st.write(f"  ❌ Girls column not found: {girls_col}")
    
    # Show which columns actually exist in the dataframe
    st.write("**All columns in your dataframe:**")
    with st.expander("Click to see all column names"):
        for i, col in enumerate(extracted_df.columns):
            st.write(f"{i+1}. `{col}`")
    
    # Try alternative approach - search for patterns in actual column names
    st.write("**Alternative Pattern Search:**")
    found_patterns = {}
    
    for col in extracted_df.columns:
        col_str = str(col)
        # Look for class patterns
        for class_num in range(1, 6):
            if f"Class {class_num}" in col_str:
                if "pupils" in col_str and "enrolled" in col_str:
                    found_patterns[f'enrollment_{class_num}'] = col
                elif "boys" in col_str and "received" not in col_str:
                    found_patterns[f'boys_{class_num}'] = col
                elif "girls" in col_str and "received" not in col_str:
                    found_patterns[f'girls_{class_num}'] = col
    
    if found_patterns:
        st.write("**Found alternative patterns:**")
        for key, col in found_patterns.items():
            sample_sum = extracted_df[col].fillna(0).sum()
            st.write(f"  - {key}: `{col}` → Sum: {sample_sum}")
    else:
        st.warning("No alternative patterns found")

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
            gps_col = COLUMN_MAPPING['other_cols']['gps_location']
            if gps_col in extracted_df.columns:
                all_gps_data = extracted_df[gps_col].dropna()
                
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



### PART 3: ANALYSIS AND CHARTS ###

    try:
        summaries = generate_summaries(extracted_df)
        
        # Debug: Show summary calculation results
        st.write("**Summary Calculation Results:**")
        debug_col1, debug_col2, debug_col3 = st.columns(3)
        
        with debug_col1:
            st.write(f"**Boys Data:**")
            st.write(f"Total Boys: {summaries['overall']['total_boys']:,}")
            st.write(f"Boys ITNs: {summaries['overall']['total_itn_boys']:,}")
        
        with debug_col2:
            st.write(f"**Girls Data:**")
            st.write(f"Total Girls: {summaries['overall']['total_girls']:,}")
            st.write(f"Girls ITNs: {summaries['overall']['total_itn_girls']:,}")
        
        with debug_col3:
            st.write(f"**Overall Data:**")
            st.write(f"Total Enrollment: {summaries['overall']['total_enrollment']:,}")
            st.write(f"Total ITNs: {summaries['overall']['total_itn']:,}")
        
        # Display Overall Summary
        st.subheader("📊 Overall Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Schools", f"{summaries['overall']['total_schools']:,}")
        with col2:
            st.metric("Total Students", f"{summaries['overall']['total_enrollment']:,}")
        with col3:
            st.metric("Total ITNs Distributed", f"{summaries['overall']['total_itn']:,}")
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

        # Enhanced Gender Analysis with better debugging
        st.subheader("👫 Gender Analysis")
        
        # Check if we have valid gender data
        total_boys = summaries['overall']['total_boys']
        total_girls = summaries['overall']['total_girls']
        
        st.write(f"**Debug Info:** Boys: {total_boys:,}, Girls: {total_girls:,}")
        
        # Alternative method: Check for any gender-related columns
        alternative_boys_total = 0
        alternative_girls_total = 0
        
        # Try to find gender data using broader search
        for col in extracted_df.columns:
            col_lower = str(col).lower()
            if 'boys' in col_lower and 'class' in col_lower and 'received' not in col_lower:
                col_sum = extracted_df[col].fillna(0).sum()
                alternative_boys_total += col_sum
                st.write(f"Found boys column: {col} → {col_sum:,}")
            elif 'girls' in col_lower and 'class' in col_lower and 'received' not in col_lower:
                col_sum = extracted_df[col].fillna(0).sum()
                alternative_girls_total += col_sum
                st.write(f"Found girls column: {col} → {col_sum:,}")
        
        st.write(f"**Alternative calculation:** Boys: {alternative_boys_total:,}, Girls: {alternative_girls_total:,}")
        
        # Use alternative totals if main calculation failed
        if total_boys == 0 and total_girls == 0 and (alternative_boys_total > 0 or alternative_girls_total > 0):
            total_boys = alternative_boys_total
            total_girls = alternative_girls_total
            st.info("Using alternative gender calculation method")
        
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
                
                # Show gender ratio
                if total_boys > 0:
                    gender_ratio = (total_girls / total_boys) * 100
                    st.write(f"**Gender Ratio:** {gender_ratio:.1f} girls per 100 boys")
                
            else:
                st.warning("No valid gender data for pie chart")
        else:
            st.warning("No gender data found in any columns")
            
            # Show what columns we checked
            st.write("**Columns checked for gender data:**")
            gender_cols = [col for col in extracted_df.columns if 'boys' in str(col).lower() or 'girls' in str(col).lower()]
            for col in gender_cols[:10]:  # Show first 10 gender-related columns
                st.write(f"- {col}")
            if len(gender_cols) > 10:
                st.write(f"... and {len(gender_cols) - 10} more gender-related columns")

        # Class-wise Enrollment Analysis
        st.subheader("📚 Class-wise Enrollment Analysis")
        
        class_analysis = []
        for class_num in range(1, 6):
            enrollment_col = COLUMN_MAPPING['enrollment_cols'].get(class_num)
            boys_col = COLUMN_MAPPING['boys_cols'].get(class_num)
            girls_col = COLUMN_MAPPING['girls_cols'].get(class_num)
            
            total_enrollment = 0
            total_boys = 0
            total_girls = 0
            
            if enrollment_col and enrollment_col in extracted_df.columns:
                total_enrollment = int(extracted_df[enrollment_col].fillna(0).sum())
            if boys_col and boys_col in extracted_df.columns:
                total_boys = int(extracted_df[boys_col].fillna(0).sum())
            if girls_col and girls_col in extracted_df.columns:
                total_girls = int(extracted_df[girls_col].fillna(0).sum())
            
            class_analysis.append({
                'Class': f'Class {class_num}',
                'Total_Enrollment': total_enrollment,
                'Boys': total_boys,
                'Girls': total_girls
            })
        
        if class_analysis:
            class_df = pd.DataFrame(class_analysis)
            
            # Display class analysis table
            st.write("**Enrollment by Class:**")
            st.dataframe(class_df, use_container_width=True)
            
            # Create class-wise chart if we have data
            total_class_enrollment = class_df['Total_Enrollment'].sum()
            if total_class_enrollment > 0:
                fig_class, ax_class = plt.subplots(figsize=(12, 8))
                
                x = np.arange(len(class_df['Class']))
                width = 0.25
                
                bars1 = ax_class.bar(x - width, class_df['Total_Enrollment'], width, 
                                    label='Total Enrollment', color='#8e44ad', edgecolor='purple', linewidth=1)
                bars2 = ax_class.bar(x, class_df['Boys'], width, 
                                    label='Boys', color='#3498db', edgecolor='navy', linewidth=1)
                bars3 = ax_class.bar(x + width, class_df['Girls'], width, 
                                    label='Girls', color='#e91e63', edgecolor='darkred', linewidth=1)
                
                ax_class.set_title('Class-wise Enrollment Distribution', fontsize=16, fontweight='bold', pad=20)
                ax_class.set_xlabel('Classes', fontsize=12, fontweight='bold')
                ax_class.set_ylabel('Number of Students', fontsize=12, fontweight='bold')
                ax_class.set_xticks(x)
                ax_class.set_xticklabels(class_df['Class'])
                ax_class.legend(fontsize=12)
                ax_class.grid(axis='y', alpha=0.3, linestyle='--')
                
                # Add value labels on bars
                for bars in [bars1, bars2, bars3]:
                    for bar in bars:
                        height = bar.get_height()
                        if height > 0:
                            ax_class.annotate(f'{int(height):,}',
                                            xy=(bar.get_x() + bar.get_width() / 2, height),
                                            xytext=(0, 3),
                                            textcoords="offset points",
                                            ha='center', va='bottom', fontsize=10, fontweight='bold')
                
                plt.tight_layout()
                st.pyplot(fig_class)
                
                # Save class chart
                class_buffer = save_map_as_png(fig_class, "Class_wise_Enrollment")
                if class_buffer:
                    map_images['class_enrollment'] = class_buffer
                plt.close(fig_class)

        # Display Summary Tables
        st.subheader("📈 District Summary Table")
        if summaries['district']:
            district_summary_df = pd.DataFrame(summaries['district'])
            st.dataframe(district_summary_df, use_container_width=True)
        else:
            st.info("No district data available")

        st.subheader("📈 Chiefdom Summary Table")
        if summaries['chiefdom']:
            chiefdom_summary_df = pd.DataFrame(summaries['chiefdom'])
            st.dataframe(chiefdom_summary_df, use_container_width=True)
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
                    
                    # Add summary sheets
                    if summaries['district']:
                        district_summary_df = pd.DataFrame(summaries['district'])
                        district_summary_df.to_excel(writer, sheet_name='District Summary', index=False)
                    
                    if summaries['chiefdom']:
                        chiefdom_summary_df = pd.DataFrame(summaries['chiefdom'])
                        chiefdom_summary_df.to_excel(writer, sheet_name='Chiefdom Summary', index=False)
                
                excel_data = excel_buffer.getvalue()
                
                st.download_button(
                    label="📊 Download Complete Data as Excel",
                    data=excel_data,
                    file_name="complete_extracted_data_with_summaries.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download all extracted data with summary sheets in Excel format"
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
                    
        # Performance insights
        if summaries['district']:
            st.subheader("🎯 Key Performance Insights")
            
            # Best and worst performing districts
            district_performance = sorted(summaries['district'], key=lambda x: x['coverage'], reverse=True)
            if district_performance:
                best_district = district_performance[0]
                worst_district = district_performance[-1]
                
                insight_col1, insight_col2 = st.columns(2)
                
                with insight_col1:
                    st.success(f"""
                    **🏆 Best Performing District**
                    
                    **{best_district['district']}**
                    - Coverage: {best_district['coverage']:.1f}%
                    - Schools: {best_district['schools']}
                    - Students: {best_district['enrollment']:,}
                    - ITNs Distributed: {best_district['itn']:,}
                    """)
                
                with insight_col2:
                    st.warning(f"""
                    **⚠️ District Needing Support**
                    
                    **{worst_district['district']}**
                    - Coverage: {worst_district['coverage']:.1f}%
                    - Schools: {worst_district['schools']}
                    - Students: {worst_district['enrollment']:,}
                    - ITNs Distributed: {worst_district['itn']:,}
                    """)
                
                # Overall program insights
                total_coverage = summaries['overall']['coverage']
                if total_coverage >= 80:
                    st.success(f"🎉 **Excellent Overall Performance**: {total_coverage:.1f}% coverage achieved!")
                elif total_coverage >= 60:
                    st.info(f"📈 **Good Progress**: {total_coverage:.1f}% coverage - approaching target!")
                else:
                    st.warning(f"🎯 **Room for Improvement**: {total_coverage:.1f}% coverage - additional efforts needed")

    except Exception as e:
        st.error(f"❌ Error generating summaries: {e}")
        st.info("Please check your data format and try again.")

if __name__ == "__main__":
    main()
