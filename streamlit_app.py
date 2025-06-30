import streamlit as st
import pandas as pd
import re
import numpy as np
import matplotlib.pyplot as plt
import geopandas as gpd
from io import BytesIO
import base64
import seaborn as sns
from datetime import datetime

# Custom CSS with enhanced blue and white theme
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
    
    /* Sidebar styling */
    .css-1d391kg {
        padding-left: 1rem !important;
        margin-left: 0 !important;
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
    
    /* Report export button styling */
    .report-button {
        background: linear-gradient(45deg, #e74c3c, #c0392b) !important;
        color: white !important;
        border: none !important;
        border-radius: 25px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        box-shadow: 0 4px 15px rgba(231, 76, 60, 0.3) !important;
        transition: all 0.3s ease !important;
        cursor: pointer !important;
        width: 100% !important;
        margin: 1rem 0 !important;
    }
    
    .report-button:hover {
        background: linear-gradient(45deg, #c0392b, #a93226) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(231, 76, 60, 0.4) !important;
    }
</style>
""", unsafe_allow_html=True)

# Function to save maps as PNG and return BytesIO object
def save_map_as_png(fig, filename_prefix):
    """Save matplotlib figure as PNG and return BytesIO object"""
    buffer = BytesIO()
    fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    buffer.seek(0)
    
    # Also save to disk for reference
    fig.savefig(f"{filename_prefix}.png", format='png', dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    
    return buffer

# Function to generate comprehensive summaries with direct column calculations
def generate_summaries(df):
    """Generate District, Chiefdom, and Gender summaries with direct column references"""
    summaries = {}
    
    # Overall Summary
    overall_summary = {
        'total_schools': len(df),
        'total_districts': len(df['District'].dropna().unique()),
        'total_chiefdoms': len(df['Chiefdom'].dropna().unique()),
        'total_boys_itn': 0,
        'total_girls_itn': 0,
        'total_enrollment_2025': 0,
        'total_itn_distributed': 0,
        'total_itn_with_reserve': 0
    }
    
    # Calculate totals directly from columns
    # 2025 Enrollment
    overall_summary['total_enrollment_2025'] = (
        df['How many pupils are enrolled in Class 1?'].fillna(0).sum() +
        df['How many pupils are enrolled in Class 2?'].fillna(0).sum() +
        df['How many pupils are enrolled in Class 3?'].fillna(0).sum() +
        df['How many pupils are enrolled in Class 4?'].fillna(0).sum() +
        df['How many pupils are enrolled in Class 5?'].fillna(0).sum()
    )
    
    # Boys who received ITNs
    overall_summary['total_boys_itn'] = (
        df['How many boys in Class 1 received ITNs?'].fillna(0).sum() +
        df['How many boys in Class 2 received ITNs?'].fillna(0).sum() +
        df['How many boys in Class 3 received ITNs?'].fillna(0).sum() +
        df['How many boys in Class 4 received ITNs?'].fillna(0).sum() +
        df['How many boys in Class 5 received ITNs?'].fillna(0).sum()
    )
    
    # Girls who received ITNs
    overall_summary['total_girls_itn'] = (
        df['How many girls in Class 1 received ITNs?'].fillna(0).sum() +
        df['How many girls in Class 2 received ITNs?'].fillna(0).sum() +
        df['How many girls in Class 3 received ITNs?'].fillna(0).sum() +
        df['How many girls in Class 4 received ITNs?'].fillna(0).sum() +
        df['How many girls in Class 5 received ITNs?'].fillna(0).sum()
    )
    
    # Total ITNs distributed
    overall_summary['total_itn_distributed'] = df['Total ITNs distributed'].fillna(0).sum()
    
    # Total ITNs with reserve
    overall_summary['total_itn_with_reserve'] = (
        df['Total ITNs distributed'].fillna(0).sum() + 
        df['ITNs left at the school for pupils who were absent.'].fillna(0).sum()
    )
    
    # Calculate coverage
    overall_summary['total_beneficiaries'] = overall_summary['total_boys_itn'] + overall_summary['total_girls_itn']
    overall_summary['coverage'] = (overall_summary['total_beneficiaries'] / overall_summary['total_enrollment_2025'] * 100) if overall_summary['total_enrollment_2025'] > 0 else 0
    overall_summary['itn_remaining'] = overall_summary['total_enrollment_2025'] - overall_summary['total_beneficiaries']
    overall_summary['gender_ratio'] = (overall_summary['total_girls_itn'] / overall_summary['total_boys_itn'] * 100) if overall_summary['total_boys_itn'] > 0 else 0
    
    summaries['overall'] = overall_summary
    
    # District Summary
    district_summary = []
    for district in df['District'].dropna().unique():
        district_data = df[df['District'] == district]
        
        # Calculate enrollment 2025
        enrollment_2025 = (
            district_data['How many pupils are enrolled in Class 1?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 2?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 3?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 4?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 5?'].fillna(0).sum()
        )
        
        # Boys who received ITNs
        boys_itn = (
            district_data['How many boys in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        # Girls who received ITNs
        girls_itn = (
            district_data['How many girls in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        # ITNs distributed
        itn_distributed = district_data['Total ITNs distributed'].fillna(0).sum()
        itn_with_reserve = (
            district_data['Total ITNs distributed'].fillna(0).sum() +
            district_data['ITNs left at the school for pupils who were absent.'].fillna(0).sum()
        )
        
        total_beneficiaries = boys_itn + girls_itn
        coverage = (total_beneficiaries / enrollment_2025 * 100) if enrollment_2025 > 0 else 0
        
        district_stats = {
            'district': district,
            'schools': len(district_data),
            'chiefdoms': len(district_data['Chiefdom'].dropna().unique()),
            'boys_itn': int(boys_itn),
            'girls_itn': int(girls_itn),
            'enrollment_2025': int(enrollment_2025),
            'itn_distributed': int(itn_distributed),
            'itn_with_reserve': int(itn_with_reserve),
            'total_beneficiaries': int(total_beneficiaries),
            'coverage': coverage,
            'itn_remaining': int(enrollment_2025 - total_beneficiaries)
        }
        
        district_summary.append(district_stats)
    
    summaries['district'] = district_summary
    
    # Chiefdom Summary
    chiefdom_summary = []
    for district in df['District'].dropna().unique():
        district_data = df[df['District'] == district]
        for chiefdom in district_data['Chiefdom'].dropna().unique():
            chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
            
            # Calculate enrollment 2025
            enrollment_2025 = (
                chiefdom_data['How many pupils are enrolled in Class 1?'].fillna(0).sum() +
                chiefdom_data['How many pupils are enrolled in Class 2?'].fillna(0).sum() +
                chiefdom_data['How many pupils are enrolled in Class 3?'].fillna(0).sum() +
                chiefdom_data['How many pupils are enrolled in Class 4?'].fillna(0).sum() +
                chiefdom_data['How many pupils are enrolled in Class 5?'].fillna(0).sum()
            )
            
            # Boys who received ITNs
            boys_itn = (
                chiefdom_data['How many boys in Class 1 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many boys in Class 2 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many boys in Class 3 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many boys in Class 4 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many boys in Class 5 received ITNs?'].fillna(0).sum()
            )
            
            # Girls who received ITNs
            girls_itn = (
                chiefdom_data['How many girls in Class 1 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many girls in Class 2 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many girls in Class 3 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many girls in Class 4 received ITNs?'].fillna(0).sum() +
                chiefdom_data['How many girls in Class 5 received ITNs?'].fillna(0).sum()
            )
            
            # ITNs distributed
            itn_distributed = chiefdom_data['Total ITNs distributed'].fillna(0).sum()
            itn_with_reserve = (
                chiefdom_data['Total ITNs distributed'].fillna(0).sum() +
                chiefdom_data['ITNs left at the school for pupils who were absent.'].fillna(0).sum()
            )
            
            total_beneficiaries = boys_itn + girls_itn
            coverage = (total_beneficiaries / enrollment_2025 * 100) if enrollment_2025 > 0 else 0
            
            chiefdom_stats = {
                'district': district,
                'chiefdom': chiefdom,
                'schools': len(chiefdom_data),
                'boys_itn': int(boys_itn),
                'girls_itn': int(girls_itn),
                'enrollment_2025': int(enrollment_2025),
                'itn_distributed': int(itn_distributed),
                'itn_with_reserve': int(itn_with_reserve),
                'total_beneficiaries': int(total_beneficiaries),
                'coverage': coverage,
                'itn_remaining': int(enrollment_2025 - total_beneficiaries)
            }
            
            chiefdom_summary.append(chiefdom_stats)
    
    summaries['chiefdom'] = chiefdom_summary
    
    return summaries

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
st.title("📊 School Based Distribution of ITNs in Sierra Leone 2025")

# Upload file
uploaded_file = "sbd first_submission_clean.xlsx"
if uploaded_file:
    # Read the uploaded Excel file
    df_original = pd.read_excel(uploaded_file)
    
    # Convert numeric columns to ensure they're not text
    numeric_columns = [
        'How many pupils are enrolled in Class 1?',
        'How many pupils are enrolled in Class 2?',
        'How many pupils are enrolled in Class 3?',
        'How many pupils are enrolled in Class 4?',
        'How many pupils are enrolled in Class 5?',
        'How many boys are in Class 1?',
        'How many boys are in Class 2?',
        'How many boys are in Class 3?',
        'How many boys are in Class 4?',
        'How many boys are in Class 5?',
        'How many girls are in Class 1?',
        'How many girls are in Class 2?',
        'How many girls are in Class 3?',
        'How many girls are in Class 4?',
        'How many girls are in Class 5?',
        'How many boys in Class 1 received ITNs?',
        'How many boys in Class 2 received ITNs?',
        'How many boys in Class 3 received ITNs?',
        'How many boys in Class 4 received ITNs?',
        'How many boys in Class 5 received ITNs?',
        'How many girls in Class 1 received ITNs?',
        'How many girls in Class 2 received ITNs?',
        'How many girls in Class 3 received ITNs?',
        'How many girls in Class 4 received ITNs?',
        'How many girls in Class 5 received ITNs?',
        'Total ITNs distributed',
        'ITNs left at the school for pupils who were absent.'
    ]
    
    # Convert to numeric, replacing any non-numeric values with 0
    for col in numeric_columns:
        if col in df_original.columns:
            df_original[col] = pd.to_numeric(df_original[col], errors='coerce').fillna(0)
    
    # Add calculated columns with direct column references
    # Enrollment in 2024
    if 'Enrollment' in df_original.columns:
        df_original['Enrollment_2024'] = df_original['Enrollment']
    else:
        df_original['Enrollment_2024'] = 0
    
    # Enrollment in 2025 - sum of all class enrollments
    df_original['Enrollment_2025'] = (
        df_original['How many pupils are enrolled in Class 1?'].fillna(0) +
        df_original['How many pupils are enrolled in Class 2?'].fillna(0) +
        df_original['How many pupils are enrolled in Class 3?'].fillna(0) +
        df_original['How many pupils are enrolled in Class 4?'].fillna(0) +
        df_original['How many pupils are enrolled in Class 5?'].fillna(0)
    )
    
    # Total boys
    df_original['Total_Boys'] = (
        df_original['How many boys are in Class 1?'].fillna(0) +
        df_original['How many boys are in Class 2?'].fillna(0) +
        df_original['How many boys are in Class 3?'].fillna(0) +
        df_original['How many boys are in Class 4?'].fillna(0) +
        df_original['How many boys are in Class 5?'].fillna(0)
    )
    
    # Total girls
    df_original['Total_Girls'] = (
        df_original['How many girls are in Class 1?'].fillna(0) +
        df_original['How many girls are in Class 2?'].fillna(0) +
        df_original['How many girls are in Class 3?'].fillna(0) +
        df_original['How many girls are in Class 4?'].fillna(0) +
        df_original['How many girls are in Class 5?'].fillna(0)
    )
    
    # Boys who received ITNs
    df_original['Boys_Received_ITNs'] = (
        df_original['How many boys in Class 1 received ITNs?'].fillna(0) +
        df_original['How many boys in Class 2 received ITNs?'].fillna(0) +
        df_original['How many boys in Class 3 received ITNs?'].fillna(0) +
        df_original['How many boys in Class 4 received ITNs?'].fillna(0) +
        df_original['How many boys in Class 5 received ITNs?'].fillna(0)
    )
    
    # Girls who received ITNs
    df_original['Girls_Received_ITNs'] = (
        df_original['How many girls in Class 1 received ITNs?'].fillna(0) +
        df_original['How many girls in Class 2 received ITNs?'].fillna(0) +
        df_original['How many girls in Class 3 received ITNs?'].fillna(0) +
        df_original['How many girls in Class 4 received ITNs?'].fillna(0) +
        df_original['How many girls in Class 5 received ITNs?'].fillna(0)
    )
    
    # ITNs distributed
    df_original['ITNs_Distributed_Without_Reserve'] = df_original['Total ITNs distributed'].fillna(0)
    df_original['ITNs_Distributed_With_Reserve'] = (
        df_original['Total ITNs distributed'].fillna(0) + 
        df_original['ITNs left at the school for pupils who were absent.'].fillna(0)
    )
    
    # Load shapefile
    try:
        gdf = gpd.read_file("Chiefdom2021.shp")
        st.success("✅ Shapefile loaded successfully!")
    except Exception as e:
        st.error(f"❌ Could not load shapefile: {e}")
        gdf = None
    
    # Create empty lists to store extracted data
    districts, chiefdoms, phu_names, community_names, school_names, enrollment = [], [], [], [], [], []
    
    # Process each row in the "Scan QR code" column
    for qr_text in df_original["Scan QR code"]:
        if pd.isna(qr_text):
            districts.append(None)
            chiefdoms.append(None)
            phu_names.append(None)
            community_names.append(None)
            school_names.append(None)
            enrollment.append(None)
            continue
            
        # Extract values using regex patterns
        district_match = re.search(r"District:\s*([^\n]+)", str(qr_text))
        districts.append(district_match.group(1).strip() if district_match else None)
        
        chiefdom_match = re.search(r"Chiefdom:\s*([^\n]+)", str(qr_text))
        chiefdoms.append(chiefdom_match.group(1).strip() if chiefdom_match else None)
        
        phu_match = re.search(r"PHU name:\s*([^\n]+)", str(qr_text))
        phu_names.append(phu_match.group(1).strip() if phu_match else None)
        
        community_match = re.search(r"Community name:\s*([^\n]+)", str(qr_text))
        community_names.append(community_match.group(1).strip() if community_match else None)
        
        school_match = re.search(r"Name of school:\s*([^\n]+)", str(qr_text))
        school_names.append(school_match.group(1).strip() if school_match else None)

        enrollment_match = re.search(r"Enrollment:\s*([^\n]+)", str(qr_text))
        enrollment_names.append(enrollment_match.group(1).strip() if school_match else None)
    
    # Create a new DataFrame with extracted values
    extracted_df = pd.DataFrame({
        "District": districts,
        "Chiefdom": chiefdoms,
        "PHU Name": phu_names,
        "Community Name": community_names,
        "School Name": school_names,
        "Enrollment": enrollment
    })
    
    # Add all other columns from the original DataFrame including calculated columns
    for column in df_original.columns:
        if column != "Scan QR code":  # Skip the QR code column since we've already processed it
            extracted_df[column] = df_original[column]
    
    # Create sidebar filters early so they're available for all sections
    st.sidebar.header("🔍 Filter Options")
    
    # Create radio buttons to select which level to group by
    grouping_selection = st.sidebar.radio(
        "Select the level for grouping:",
        ["District", "Chiefdom", "PHU Name", "Community Name", "School Name"],
        index=0  # Default to 'District'
    )
    
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
    for level in hierarchy[grouping_selection]:
        # Filter out None/NaN values and get sorted unique values
        level_values = sorted(filtered_df[level].dropna().unique())
        
        if level_values:
            # Create selectbox for this level
            selected_value = st.sidebar.selectbox(f"Select {level}", level_values)
            selected_values[level] = selected_value
            
            # Apply filter to the dataframe
            filtered_df = filtered_df[filtered_df[level] == selected_value]
    
    # Store map images for report
    map_images = {}
    
    # Display Dual Maps at the top
    st.subheader("🗺️ Geographic Distribution Maps")
    
    if gdf is not None:
        # OVERALL SIERRA LEONE MAP FIRST
        st.write("**Sierra Leone - All Districts Overview**")
        
        # Create overall Sierra Leone map
        fig_overall, ax_overall = plt.subplots(figsize=(16, 10))
        
        # Plot all chiefdoms with gray edges (base layer)
        gdf.plot(ax=ax_overall, color='white', edgecolor='gray', alpha=0.8, linewidth=0.5)
        
        # Plot district boundaries with thick black lines
        # Get district boundaries by dissolving chiefdoms by FIRST_DNAM
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
        
        # Extract and plot ALL GPS coordinates from entire dataset
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
            
            st.write(f"**Total valid coordinates for overall map: {len(all_coords_extracted)}**")
        
        # Plot GPS points on the overall map
        if all_coords_extracted:
            lats, lons = zip(*all_coords_extracted)
            
            # Plot GPS points with #47B5FF color
            scatter = ax_overall.scatter(
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
        
        # Add grid for reference
        ax_overall.grid(True, alpha=0.3, linestyle='--')
        
        # Set axis limits to show full country
        ax_overall.set_xlim(gdf.total_bounds[0] - 0.1, gdf.total_bounds[2] + 0.1)
        ax_overall.set_ylim(gdf.total_bounds[1] - 0.1, gdf.total_bounds[3] + 0.1)
        
        plt.tight_layout()
        st.pyplot(fig_overall)
        
        # Save overall map
        map_images['sierra_leone_overall'] = save_map_as_png(fig_overall, "Sierra_Leone_Overall_Map")
        
        st.divider()
    else:
        st.error("Shapefile not loaded. Cannot display map.")
    
    # Display Original Data Sample
    st.subheader("📄 Original Data Sample")
    st.dataframe(df_original.head())
    
    # Debug section to check data
    st.subheader("🔍 Data Verification")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Sample ITN Recipients Data (Boys)**")
        for i in range(1, 6):
            col_name = f'How many boys in Class {i} received ITNs?'
            if col_name in extracted_df.columns:
                non_zero = extracted_df[extracted_df[col_name] > 0][col_name].count()
                total_sum = extracted_df[col_name].fillna(0).sum()
                st.write(f"Class {i}: Sum={total_sum}, Non-zero entries={non_zero}")
    
    with col2:
        st.write("**Sample ITN Recipients Data (Girls)**")
        for i in range(1, 6):
            col_name = f'How many girls in Class {i} received ITNs?'
            if col_name in extracted_df.columns:
                non_zero = extracted_df[extracted_df[col_name] > 0][col_name].count()
                total_sum = extracted_df[col_name].fillna(0).sum()
                st.write(f"Class {i}: Sum={total_sum}, Non-zero entries={non_zero}")
    
    # Check enrollment data
    st.write("**Enrollment Data Check**")
    total_enrollment_check = 0
    for i in range(1, 6):
        col_name = f'How many pupils are enrolled in Class {i}?'
        if col_name in extracted_df.columns:
            class_sum = extracted_df[col_name].fillna(0).sum()
            total_enrollment_check += class_sum
            st.write(f"Class {i} enrollment: {class_sum}")
    st.write(f"**Total enrollment across all classes: {total_enrollment_check}**")
    
    # Display Extracted Data with new calculated columns
    st.subheader("📋 Enhanced Extracted Data with 2025 Calculations")
    display_cols = ['District', 'Chiefdom', 'School Name', 'Enrollment_2025', 'Total_Boys', 'Total_Girls', 
                   'Boys_Received_ITNs', 'Girls_Received_ITNs', 'ITNs_Distributed_With_Reserve']
    st.dataframe(extracted_df[display_cols].head(10))
    
    # Add download button for CSV
    csv = extracted_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Enhanced Data with All Calculations as CSV",
        data=csv,
        file_name="enhanced_school_data_2025.csv",
        mime="text/csv"
    )
    
    # Generate comprehensive summaries with updated calculations
    summaries = generate_summaries(extracted_df)
    
    # Add Data Quality Check
    st.sidebar.subheader("📊 Data Quality Check")
    st.sidebar.write(f"Total Records: {len(extracted_df)}")
    st.sidebar.write(f"Districts with data: {summaries['overall']['total_districts']}")
    st.sidebar.write(f"Total ITN Recipients: {summaries['overall']['total_beneficiaries']:,}")
    
    if summaries['overall']['total_beneficiaries'] == 0:
        st.sidebar.warning("⚠️ No ITN recipient data found!")
        st.sidebar.write("Check if ITN columns have data")
    else:
        st.sidebar.success(f"✅ Found {summaries['overall']['total_beneficiaries']:,} ITN recipients")
    
    # Display Enhanced Overall Summary
    st.subheader("📊 Enhanced Overall Summary - 2025 Analysis")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Schools", f"{summaries['overall']['total_schools']:,}")
    with col2:
        st.metric("2025 Enrollment", f"{summaries['overall']['total_enrollment_2025']:,}")
    with col3:
        st.metric("Total Beneficiaries", f"{summaries['overall']['total_beneficiaries']:,}")
    with col4:
        st.metric("Coverage", f"{summaries['overall']['coverage']:.1f}%")
    
    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.metric("Districts", f"{summaries['overall']['total_districts']}")
    with col6:
        st.metric("Chiefdoms", f"{summaries['overall']['total_chiefdoms']}")
    with col7:
        st.metric("Boys ITN", f"{summaries['overall']['total_boys_itn']:,}")
    with col8:
        st.metric("Girls ITN", f"{summaries['overall']['total_girls_itn']:,}")
    
    # Additional metrics row
    col9, col10, col11, col12 = st.columns(4)
    with col9:
        st.metric("ITNs Distributed", f"{summaries['overall']['total_itn_distributed']:,}")
    with col10:
        st.metric("ITNs + Reserve", f"{summaries['overall']['total_itn_with_reserve']:,}")
    with col11:
        st.metric("ITNs Remaining", f"{summaries['overall']['itn_remaining']:,}")
    with col12:
        st.metric("Gender Ratio", f"{summaries['overall']['gender_ratio']:.1f}%")
    
    # Enhanced Gender Analysis with ITN Recipients
    st.subheader("👫 Enhanced Gender Analysis - ITN Recipients 2025")
    
    # Check if there's data to display
    if summaries['overall']['total_boys_itn'] > 0 or summaries['overall']['total_girls_itn'] > 0:
        # Overall gender distribution pie chart for ITN recipients
        fig_gender, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 8))
        
        # Left pie chart - Boys vs Girls who received ITNs
        labels = ['Boys ITN', 'Girls ITN']
        sizes = [summaries['overall']['total_boys_itn'], summaries['overall']['total_girls_itn']]
        colors = ['#4A90E2', '#F39C12']
        
        wedges, texts, autotexts = ax1.pie(sizes, labels=labels, autopct='%1.1f%%', 
                                            colors=colors, startangle=90)
        ax1.set_title('ITN Recipients by Gender', fontsize=16, fontweight='bold', pad=20)
        plt.setp(autotexts, size=14, weight="bold")
        plt.setp(texts, size=12, weight="bold")
        
        # Right pie chart - Coverage vs Remaining
        labels2 = ['Covered', 'Not Covered']
        sizes2 = [summaries['overall']['total_beneficiaries'], summaries['overall']['itn_remaining']]
        colors2 = ['#27AE60', '#E74C3C']
        
        if summaries['overall']['total_enrollment_2025'] > 0:
            wedges2, texts2, autotexts2 = ax2.pie(sizes2, labels=labels2, autopct='%1.1f%%',
                                                  colors=colors2, startangle=90)
            ax2.set_title('Overall ITN Coverage Status', fontsize=16, fontweight='bold', pad=20)
            plt.setp(autotexts2, size=14, weight="bold")
            plt.setp(texts2, size=12, weight="bold")
        else:
            ax2.text(0.5, 0.5, 'No enrollment data available', 
                    ha='center', va='center', transform=ax2.transAxes, fontsize=14)
            ax2.set_xlim(-1, 1)
            ax2.set_ylim(-1, 1)
        
        plt.tight_layout()
        st.pyplot(fig_gender)
        
        # Save gender charts
        map_images['gender_overall'] = save_map_as_png(fig_gender, "Enhanced_Gender_Distribution")
    else:
        st.warning("No ITN recipient data available. The data shows 0 boys and 0 girls have received ITNs.")
        
        # Show enrollment gender distribution instead
        st.subheader("📊 Enrollment Gender Distribution (Alternative View)")
        
        # Calculate total boys and girls enrolled
        total_boys_enrolled = 0
        total_girls_enrolled = 0
        
        for i in range(1, 6):
            boys_col = f'How many boys are in Class {i}?'
            girls_col = f'How many girls are in Class {i}?'
            
            if boys_col in extracted_df.columns:
                total_boys_enrolled += extracted_df[boys_col].fillna(0).sum()
            if girls_col in extracted_df.columns:
                total_girls_enrolled += extracted_df[girls_col].fillna(0).sum()
        
        if total_boys_enrolled > 0 or total_girls_enrolled > 0:
            fig_enrollment_gender, ax = plt.subplots(figsize=(10, 8))
            labels = ['Boys Enrolled', 'Girls Enrolled']
            sizes = [total_boys_enrolled, total_girls_enrolled]
            colors = ['#4A90E2', '#F39C12']
            
            wedges, texts, autotexts = ax.pie(sizes, labels=labels, autopct='%1.1f%%',
                                              colors=colors, startangle=90)
            ax.set_title('Student Enrollment by Gender', fontsize=16, fontweight='bold', pad=20)
            plt.setp(autotexts, size=14, weight="bold")
            plt.setp(texts, size=12, weight="bold")
            
            plt.tight_layout()
            st.pyplot(fig_enrollment_gender)
            
            # Display metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Boys Enrolled", f"{int(total_boys_enrolled):,}")
            with col2:
                st.metric("Total Girls Enrolled", f"{int(total_girls_enrolled):,}")
            with col3:
                gender_ratio = (total_girls_enrolled / total_boys_enrolled * 100) if total_boys_enrolled > 0 else 0
                st.metric("Gender Ratio (G:B)", f"{gender_ratio:.1f}%")
    
    # Enhanced District Analysis with 2025 data
    st.subheader("📊 Enhanced District Analysis - 2025 Enrollment vs ITN Distribution")
    
    # Create comparative bar chart for enrollment vs ITN distribution
    districts = [d['district'] for d in summaries['district']]
    enrollment_2025 = [d['enrollment_2025'] for d in summaries['district']]
    beneficiaries = [d['total_beneficiaries'] for d in summaries['district']]
    itn_distributed = [d['itn_distributed'] for d in summaries['district']]
    
    fig_district, ax = plt.subplots(figsize=(16, 10))
    x = np.arange(len(districts))
    width = 0.25
    
    bars1 = ax.bar(x - width, enrollment_2025, width, label='2025 Enrollment', 
                   color='#3498DB', edgecolor='navy', linewidth=1)
    bars2 = ax.bar(x, beneficiaries, width, label='ITN Recipients', 
                   color='#2ECC71', edgecolor='darkgreen', linewidth=1)
    bars3 = ax.bar(x + width, itn_distributed, width, label='ITNs Distributed', 
                   color='#E74C3C', edgecolor='darkred', linewidth=1)
    
    ax.set_title('District Analysis: 2025 Enrollment vs ITN Distribution', 
                fontsize=18, fontweight='bold', pad=20)
    ax.set_xlabel('Districts', fontsize=14, fontweight='bold')
    ax.set_ylabel('Count', fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(districts, rotation=45, ha='right')
    ax.legend(fontsize=12)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    # Add value labels on bars
    for bars in [bars1, bars2, bars3]:
        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'{int(height):,}',
                       xy=(bar.get_x() + bar.get_width() / 2, height),
                       xytext=(0, 3),
                       textcoords="offset points",
                       ha='center', va='bottom', fontsize=9, fontweight='bold')
    
    plt.tight_layout()
    st.pyplot(fig_district)
    
    # Save district chart
    map_images['district_analysis'] = save_map_as_png(fig_district, "District_Analysis_2025")
    
    # New: Enrollment Growth Analysis (2024 vs 2025)
    st.subheader("📈 Enrollment Growth Analysis: 2024 vs 2025")
    
    if 'Enrollment' in extracted_df.columns:
        # Calculate enrollment growth by district using direct column calculations
        enrollment_growth = []
        for district in extracted_df['District'].dropna().unique():
            district_data = extracted_df[extracted_df['District'] == district]
            
            # 2024 enrollment
            enrollment_2024 = district_data['Enrollment'].fillna(0).sum()
            
            # 2025 enrollment - direct calculation
            enrollment_2025 = (
                district_data['How many pupils are enrolled in Class 1?'].fillna(0).sum() +
                district_data['How many pupils are enrolled in Class 2?'].fillna(0).sum() +
                district_data['How many pupils are enrolled in Class 3?'].fillna(0).sum() +
                district_data['How many pupils are enrolled in Class 4?'].fillna(0).sum() +
                district_data['How many pupils are enrolled in Class 5?'].fillna(0).sum()
            )
            
            growth = ((enrollment_2025 - enrollment_2024) / enrollment_2024 * 100) if enrollment_2024 > 0 else 0
            
            enrollment_growth.append({
                'District': district,
                'Enrollment_2024': enrollment_2024,
                'Enrollment_2025': enrollment_2025,
                'Growth_Percentage': growth
            })
        
        growth_df = pd.DataFrame(enrollment_growth)
        growth_df = growth_df.sort_values('Growth_Percentage', ascending=False)
        
        # Create growth chart
        fig_growth, ax = plt.subplots(figsize=(14, 8))
        colors = ['#2ECC71' if x >= 0 else '#E74C3C' for x in growth_df['Growth_Percentage']]
        bars = ax.bar(growth_df['District'], growth_df['Growth_Percentage'], color=colors, 
                      edgecolor='black', linewidth=1)
        
        ax.set_title('Enrollment Growth by District (2024 vs 2025)', fontsize=16, fontweight='bold')
        ax.set_xlabel('District', fontsize=12, fontweight='bold')
        ax.set_ylabel('Growth Percentage (%)', fontsize=12, fontweight='bold')
        ax.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'{height:.1f}%',
                       xy=(bar.get_x() + bar.get_width() / 2, height),
                       xytext=(0, 3 if height >= 0 else -15),
                       textcoords="offset points",
                       ha='center', va='bottom' if height >= 0 else 'top', 
                       fontsize=10, fontweight='bold')
        
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        st.pyplot(fig_growth)
        
        # Save growth chart
        map_images['enrollment_growth'] = save_map_as_png(fig_growth, "Enrollment_Growth_Analysis")
    
    # New: ITN Distribution Efficiency Analysis
    st.subheader("🎯 ITN Distribution Efficiency Analysis")
    
    # Calculate efficiency metrics with direct column calculations
    efficiency_data = []
    for district in extracted_df['District'].dropna().unique():
        district_data = extracted_df[extracted_df['District'] == district]
        
        total_distributed = district_data['Total ITNs distributed'].fillna(0).sum()
        
        # Calculate total received using direct column sums
        boys_received = (
            district_data['How many boys in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        girls_received = (
            district_data['How many girls in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        total_received = boys_received + girls_received
        efficiency = (total_received / total_distributed * 100) if total_distributed > 0 else 0
        
        efficiency_data.append({
            'District': district,
            'ITNs_Distributed': total_distributed,
            'ITNs_Received': total_received,
            'Efficiency': efficiency
        })
    
    efficiency_df = pd.DataFrame(efficiency_data)
    efficiency_df = efficiency_df.sort_values('Efficiency', ascending=True)
    
    # Create efficiency chart
    fig_efficiency, ax = plt.subplots(figsize=(14, 10))
    
    # Use color gradient based on efficiency
    norm = plt.Normalize(efficiency_df['Efficiency'].min(), efficiency_df['Efficiency'].max())
    colors = plt.cm.RdYlGn(norm(efficiency_df['Efficiency']))
    
    bars = ax.barh(efficiency_df['District'], efficiency_df['Efficiency'], color=colors, 
                   edgecolor='black', linewidth=1)
    
    ax.set_title('ITN Distribution Efficiency by District', fontsize=16, fontweight='bold')
    ax.set_xlabel('Efficiency (%)', fontsize=12, fontweight='bold')
    ax.set_ylabel('District', fontsize=12, fontweight='bold')
    ax.grid(axis='x', alpha=0.3, linestyle='--')
    
    # Add value labels
    for i, (idx, row) in enumerate(efficiency_df.iterrows()):
        ax.text(row['Efficiency'] + 1, i, f"{row['Efficiency']:.1f}%", 
               va='center', fontweight='bold')
    
    # Add color bar legend
    sm = plt.cm.ScalarMappable(cmap=plt.cm.RdYlGn, norm=norm)
    sm.set_array([])
    cbar = plt.colorbar(sm, ax=ax)
    cbar.set_label('Efficiency %', fontsize=12)
    
    plt.tight_layout()
    st.pyplot(fig_efficiency)
    
    # Save efficiency chart
    map_images['efficiency_analysis'] = save_map_as_png(fig_efficiency, "ITN_Efficiency_Analysis")
    
    # New: Gender Equity Score Analysis
    st.subheader("⚖️ Gender Equity Score Analysis")
    
    # Calculate gender equity scores with direct column calculations
    gender_equity_data = []
    for district in extracted_df['District'].dropna().unique():
        district_data = extracted_df[extracted_df['District'] == district]
        
        # Direct calculation of boys and girls who received ITNs
        boys_itn = (
            district_data['How many boys in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        girls_itn = (
            district_data['How many girls in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        # Direct calculation of total boys and girls
        total_boys = (
            district_data['How many boys are in Class 1?'].fillna(0).sum() +
            district_data['How many boys are in Class 2?'].fillna(0).sum() +
            district_data['How many boys are in Class 3?'].fillna(0).sum() +
            district_data['How many boys are in Class 4?'].fillna(0).sum() +
            district_data['How many boys are in Class 5?'].fillna(0).sum()
        )
        
        total_girls = (
            district_data['How many girls are in Class 1?'].fillna(0).sum() +
            district_data['How many girls are in Class 2?'].fillna(0).sum() +
            district_data['How many girls are in Class 3?'].fillna(0).sum() +
            district_data['How many girls are in Class 4?'].fillna(0).sum() +
            district_data['How many girls are in Class 5?'].fillna(0).sum()
        )
        
        boys_coverage = (boys_itn / total_boys * 100) if total_boys > 0 else 0
        girls_coverage = (girls_itn / total_girls * 100) if total_girls > 0 else 0
        
        # Gender equity score: closer to 100 means more equitable
        equity_score = 100 - abs(boys_coverage - girls_coverage)
        
        gender_equity_data.append({
            'District': district,
            'Boys_Coverage': boys_coverage,
            'Girls_Coverage': girls_coverage,
            'Equity_Score': equity_score
        })
    
    equity_df = pd.DataFrame(gender_equity_data)
    equity_df = equity_df.sort_values('Equity_Score', ascending=False)
    
    # Create dual-axis plot for gender coverage
    fig_equity, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 12))
    
    # Top chart: Boys vs Girls Coverage
    x = np.arange(len(equity_df))
    width = 0.35
    
    bars1 = ax1.bar(x - width/2, equity_df['Boys_Coverage'], width, label='Boys Coverage',
                    color='#3498DB', edgecolor='navy', linewidth=1)
    bars2 = ax1.bar(x + width/2, equity_df['Girls_Coverage'], width, label='Girls Coverage',
                    color='#E91E63', edgecolor='darkred', linewidth=1)
    
    ax1.set_title('Gender Coverage Comparison by District', fontsize=16, fontweight='bold')
    ax1.set_ylabel('Coverage (%)', fontsize=12, fontweight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(equity_df['District'], rotation=45, ha='right')
    ax1.legend(fontsize=12)
    ax1.grid(axis='y', alpha=0.3, linestyle='--')
    
    # Bottom chart: Equity Score
    colors = ['#2ECC71' if score >= 90 else '#F39C12' if score >= 80 else '#E74C3C' 
             for score in equity_df['Equity_Score']]
    bars3 = ax2.bar(equity_df['District'], equity_df['Equity_Score'], color=colors,
                   edgecolor='black', linewidth=1)
    
    ax2.set_title('Gender Equity Score by District', fontsize=16, fontweight='bold')
    ax2.set_xlabel('District', fontsize=12, fontweight='bold')
    ax2.set_ylabel('Equity Score (0-100)', fontsize=12, fontweight='bold')
    ax2.set_ylim(0, 100)
    ax2.axhline(y=90, color='green', linestyle='--', alpha=0.5, label='High Equity (>90)')
    ax2.axhline(y=80, color='orange', linestyle='--', alpha=0.5, label='Medium Equity (80-90)')
    ax2.legend(fontsize=10)
    ax2.grid(axis='y', alpha=0.3, linestyle='--')
    
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    st.pyplot(fig_equity)
    
    # Save equity chart
    map_images['gender_equity'] = save_map_as_png(fig_equity, "Gender_Equity_Analysis")
    
    # Enhanced Summary Tables with new metrics
    st.subheader("📈 Enhanced District Summary Table")
    district_summary_df = pd.DataFrame(summaries['district'])
    # Add coverage percentage column
    district_summary_df['Coverage %'] = district_summary_df['coverage'].round(1)
    # Reorder columns for better display
    display_columns = ['district', 'schools', 'chiefdoms', 'enrollment_2025', 'total_beneficiaries', 
                      'boys_itn', 'girls_itn', 'itn_distributed', 'Coverage %']
    st.dataframe(district_summary_df[display_columns])
    
    # Download enhanced district summary
    district_csv = district_summary_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Enhanced District Summary",
        data=district_csv,
        file_name="enhanced_district_summary_2025.csv",
        mime="text/csv"
    )
    
    # Chiefdom-level enhanced analysis
    st.subheader("📊 Enhanced Chiefdom Analysis by District")
    
    # Get all unique districts that have chiefdom data
    districts_with_chiefdoms = extracted_df[extracted_df['Chiefdom'].notna()]['District'].unique()
    
    for district in districts_with_chiefdoms[:3]:  # Show first 3 districts to avoid clutter
        st.write(f"### {district} District - Enhanced Chiefdom Analysis")
        
        # Filter data for this district
        district_data = extracted_df[extracted_df['District'] == district]
        district_chiefdoms = district_data['Chiefdom'].dropna().unique()
        
        if len(district_chiefdoms) > 0:
            # Calculate enhanced metrics by chiefdom with direct column calculations
            chiefdom_analysis = []
            
            for chiefdom in district_chiefdoms:
                chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
                
                # Direct calculation of enrollment 2025
                enrollment_2025 = (
                    chiefdom_data['How many pupils are enrolled in Class 1?'].fillna(0).sum() +
                    chiefdom_data['How many pupils are enrolled in Class 2?'].fillna(0).sum() +
                    chiefdom_data['How many pupils are enrolled in Class 3?'].fillna(0).sum() +
                    chiefdom_data['How many pupils are enrolled in Class 4?'].fillna(0).sum() +
                    chiefdom_data['How many pupils are enrolled in Class 5?'].fillna(0).sum()
                )
                
                # Direct calculation of boys and girls who received ITNs
                boys_itn = (
                    chiefdom_data['How many boys in Class 1 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many boys in Class 2 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many boys in Class 3 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many boys in Class 4 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many boys in Class 5 received ITNs?'].fillna(0).sum()
                )
                
                girls_itn = (
                    chiefdom_data['How many girls in Class 1 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many girls in Class 2 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many girls in Class 3 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many girls in Class 4 received ITNs?'].fillna(0).sum() +
                    chiefdom_data['How many girls in Class 5 received ITNs?'].fillna(0).sum()
                )
                
                total_beneficiaries = boys_itn + girls_itn
                coverage = (total_beneficiaries / enrollment_2025 * 100) if enrollment_2025 > 0 else 0
                
                chiefdom_analysis.append({
                    'Chiefdom': chiefdom,
                    'Schools': len(chiefdom_data),
                    'Enrollment_2025': enrollment_2025,
                    'Boys_ITN': boys_itn,
                    'Girls_ITN': girls_itn,
                    'Total_Beneficiaries': total_beneficiaries,
                    'Coverage': coverage
                })
            
            chiefdom_df = pd.DataFrame(chiefdom_analysis)
            chiefdom_df = chiefdom_df.sort_values('Coverage', ascending=False)
            
            # Create enhanced visualization
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
            
            # Left chart: Enrollment vs Beneficiaries
            x = np.arange(len(chiefdom_df))
            width = 0.35
            
            bars1 = ax1.bar(x - width/2, chiefdom_df['Enrollment_2025'], width, 
                           label='2025 Enrollment', color='#3498DB', edgecolor='navy', linewidth=1)
            bars2 = ax1.bar(x + width/2, chiefdom_df['Total_Beneficiaries'], width,
                           label='ITN Recipients', color='#2ECC71', edgecolor='darkgreen', linewidth=1)
            
            ax1.set_title(f'{district} - Enrollment vs ITN Recipients', fontsize=14, fontweight='bold')
            ax1.set_xlabel('Chiefdom', fontsize=12)
            ax1.set_ylabel('Count', fontsize=12)
            ax1.set_xticks(x)
            ax1.set_xticklabels(chiefdom_df['Chiefdom'], rotation=45, ha='right')
            ax1.legend()
            ax1.grid(axis='y', alpha=0.3, linestyle='--')
            
            # Right chart: Coverage percentage
            colors = ['#2ECC71' if c >= 80 else '#F39C12' if c >= 60 else '#E74C3C' 
                     for c in chiefdom_df['Coverage']]
            bars3 = ax2.bar(chiefdom_df['Chiefdom'], chiefdom_df['Coverage'], 
                           color=colors, edgecolor='black', linewidth=1)
            
            ax2.set_title(f'{district} - ITN Coverage by Chiefdom', fontsize=14, fontweight='bold')
            ax2.set_xlabel('Chiefdom', fontsize=12)
            ax2.set_ylabel('Coverage (%)', fontsize=12)
            ax2.set_ylim(0, 100)
            ax2.axhline(y=80, color='green', linestyle='--', alpha=0.5)
            ax2.axhline(y=60, color='orange', linestyle='--', alpha=0.5)
            ax2.grid(axis='y', alpha=0.3, linestyle='--')
            
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig)
            
            # Save chiefdom charts
            map_images[f'{district}_chiefdom_analysis'] = save_map_as_png(fig, f"{district}_Chiefdom_Analysis")
            
            # Display summary metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Chiefdoms", len(chiefdom_df))
            with col2:
                st.metric("Total Students", int(chiefdom_df['Enrollment_2025'].sum()))
            with col3:
                st.metric("Total Recipients", int(chiefdom_df['Total_Beneficiaries'].sum()))
            with col4:
                avg_coverage = chiefdom_df['Coverage'].mean()
                st.metric("Average Coverage", f"{avg_coverage:.1f}%")
            
            st.divider()
    
    # New: Class-wise Analysis
    st.subheader("📚 Class-wise ITN Distribution Analysis")
    
    # Calculate class-wise totals
    class_data = []
    for i in range(1, 6):
        enrollment = extracted_df[f'How many pupils are enrolled in Class {i}?'].fillna(0).sum()
        boys_itn = extracted_df[f'How many boys in Class {i} received ITNs?'].fillna(0).sum()
        girls_itn = extracted_df[f'How many girls in Class {i} received ITNs?'].fillna(0).sum()
        total_itn = boys_itn + girls_itn
        coverage = (total_itn / enrollment * 100) if enrollment > 0 else 0
        
        class_data.append({
            'Class': f'Class {i}',
            'Enrollment': enrollment,
            'Boys_ITN': boys_itn,
            'Girls_ITN': girls_itn,
            'Total_ITN': total_itn,
            'Coverage': coverage
        })
    
    class_df = pd.DataFrame(class_data)
    
    # Create class-wise visualization
    fig_class, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    
    # Left chart: Enrollment vs ITN Recipients by Class
    x = np.arange(len(class_df))
    width = 0.35
    
    bars1 = ax1.bar(x - width/2, class_df['Enrollment'], width, label='Enrollment', 
                    color='#3498DB', edgecolor='navy', linewidth=1)
    bars2 = ax1.bar(x + width/2, class_df['Total_ITN'], width, label='ITN Recipients',
                    color='#2ECC71', edgecolor='darkgreen', linewidth=1)
    
    ax1.set_title('Class-wise Enrollment vs ITN Recipients', fontsize=16, fontweight='bold')
    ax1.set_xlabel('Class', fontsize=12, fontweight='bold')
    ax1.set_ylabel('Number of Students', fontsize=12, fontweight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(class_df['Class'])
    ax1.legend(fontsize=12)
    ax1.grid(axis='y', alpha=0.3, linestyle='--')
    
    # Add value labels
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax1.annotate(f'{int(height):,}',
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    # Right chart: Gender distribution by class
    x2 = np.arange(len(class_df))
    bars3 = ax2.bar(x2 - width/2, class_df['Boys_ITN'], width, label='Boys',
                    color='#4A90E2', edgecolor='navy', linewidth=1)
    bars4 = ax2.bar(x2 + width/2, class_df['Girls_ITN'], width, label='Girls',
                    color='#E91E63', edgecolor='darkred', linewidth=1)
    
    ax2.set_title('Gender Distribution of ITN Recipients by Class', fontsize=16, fontweight='bold')
    ax2.set_xlabel('Class', fontsize=12, fontweight='bold')
    ax2.set_ylabel('Number of ITN Recipients', fontsize=12, fontweight='bold')
    ax2.set_xticks(x2)
    ax2.set_xticklabels(class_df['Class'])
    ax2.legend(fontsize=12)
    ax2.grid(axis='y', alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    st.pyplot(fig_class)
    
    # Save class analysis chart
    map_images['class_analysis'] = save_map_as_png(fig_class, "Class_wise_Analysis")
    
    # New: Top/Bottom Performing Analysis
    st.subheader("🏆 Top and Bottom Performing Analysis")
    
    # Create performance dataframe
    performance_data = []
    for district in extracted_df['District'].dropna().unique():
        district_data = extracted_df[extracted_df['District'] == district]
        
        # Calculate metrics
        enrollment_2025 = (
            district_data['How many pupils are enrolled in Class 1?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 2?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 3?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 4?'].fillna(0).sum() +
            district_data['How many pupils are enrolled in Class 5?'].fillna(0).sum()
        )
        
        boys_itn = (
            district_data['How many boys in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many boys in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        girls_itn = (
            district_data['How many girls in Class 1 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 2 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 3 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 4 received ITNs?'].fillna(0).sum() +
            district_data['How many girls in Class 5 received ITNs?'].fillna(0).sum()
        )
        
        total_beneficiaries = boys_itn + girls_itn
        coverage = (total_beneficiaries / enrollment_2025 * 100) if enrollment_2025 > 0 else 0
        
        performance_data.append({
            'District': district,
            'Coverage': coverage,
            'Total_Beneficiaries': total_beneficiaries,
            'Enrollment': enrollment_2025
        })
    
    performance_df = pd.DataFrame(performance_data)
    performance_df = performance_df.sort_values('Coverage', ascending=False)
    
    # Show top 5 and bottom 5
    top_5 = performance_df.head(5)
    bottom_5 = performance_df.tail(5)
    
    fig_performance, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    
    # Top 5 districts
    bars1 = ax1.bar(top_5['District'], top_5['Coverage'], color='#2ECC71', 
                    edgecolor='darkgreen', linewidth=2)
    ax1.set_title('Top 5 Performing Districts', fontsize=16, fontweight='bold')
    ax1.set_xlabel('District', fontsize=12, fontweight='bold')
    ax1.set_ylabel('Coverage (%)', fontsize=12, fontweight='bold')
    ax1.set_ylim(0, 100)
    ax1.grid(axis='y', alpha=0.3, linestyle='--')
    
    # Add value labels
    for bar in bars1:
        height = bar.get_height()
        ax1.annotate(f'{height:.1f}%',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=12, fontweight='bold')
    
    # Bottom 5 districts
    bars2 = ax2.bar(bottom_5['District'], bottom_5['Coverage'], color='#E74C3C',
                    edgecolor='darkred', linewidth=2)
    ax2.set_title('Bottom 5 Performing Districts', fontsize=16, fontweight='bold')
    ax2.set_xlabel('District', fontsize=12, fontweight='bold')
    ax2.set_ylabel('Coverage (%)', fontsize=12, fontweight='bold')
    ax2.set_ylim(0, max(bottom_5['Coverage']) * 1.2 if bottom_5['Coverage'].max() > 0 else 10)
    ax2.grid(axis='y', alpha=0.3, linestyle='--')
    
    # Add value labels
    for bar in bars2:
        height = bar.get_height()
        ax2.annotate(f'{height:.1f}%',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=12, fontweight='bold')
    
    plt.tight_layout()
    st.pyplot(fig_performance)
    
    # Save performance chart
    map_images['performance_analysis'] = save_map_as_png(fig_performance, "Performance_Analysis")
    
    # Final comprehensive report export
    st.subheader("📥 Export Enhanced Reports")
    st.write("Download comprehensive analysis reports with all enhanced visualizations and metrics:")
    
    # Create enhanced Excel report
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        # Write multiple sheets
        extracted_df.to_excel(writer, sheet_name='Complete Data', index=False)
        district_summary_df.to_excel(writer, sheet_name='District Summary', index=False)
        pd.DataFrame(summaries['chiefdom']).to_excel(writer, sheet_name='Chiefdom Summary', index=False)
        
        # Add efficiency analysis if available
        if 'efficiency_df' in locals():
            efficiency_df.to_excel(writer, sheet_name='Efficiency Analysis', index=False)
        
        # Add gender equity analysis if available
        if 'equity_df' in locals():
            equity_df.to_excel(writer, sheet_name='Gender Equity', index=False)
        
        # Add class analysis
        if 'class_df' in locals():
            class_df.to_excel(writer, sheet_name='Class Analysis', index=False)
        
        # Add performance analysis
        if 'performance_df' in locals():
            performance_df.to_excel(writer, sheet_name='Performance Analysis', index=False)
    
    excel_data = excel_buffer.getvalue()
    
    st.download_button(
        label="📊 Download Enhanced Excel Report (Multiple Sheets)",
        data=excel_data,
        file_name=f"SBD_Enhanced_Report_2025_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download comprehensive report with multiple analysis sheets"
    )
    
    # Display final summary metrics
    st.info(f"""
    📋 **Enhanced Dataset Summary**: 
    - Total Records: {len(extracted_df):,}
    - 2025 Enrollment: {summaries['overall']['total_enrollment_2025']:,}
    - ITN Recipients: {summaries['overall']['total_beneficiaries']:,}
    - Overall Coverage: {summaries['overall']['coverage']:.1f}%
    - Gender Equity Ratio: {summaries['overall']['gender_ratio']:.1f}%
    - ITNs Distributed: {summaries['overall']['total_itn_distributed']:,}
    - ITNs with Reserve: {summaries['overall']['total_itn_with_reserve']:,}
    """)
    
    # Display saved visualizations count
    if map_images:
        st.success(f"✅ **Visualizations Saved**: {len(map_images)} enhanced charts and maps have been generated")
        
        with st.expander("📁 View All Generated Visualizations"):
            for map_name in map_images.keys():
                st.write(f"• {map_name.replace('_', ' ').title()}")
else:
    st.error("error")
