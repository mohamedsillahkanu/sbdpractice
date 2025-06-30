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
    buffer = BytesIO()
    fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    buffer.seek(0)
    
    # Also save to disk for reference
    try:
        fig.savefig(f"{filename_prefix}.png", format='png', dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    except:
        pass  # Don't fail if can't save to disk
    
    return buffer

# Function to extract QR code data with improved error handling
def extract_qr_data(df_original):
    """Extract QR data with improved error handling and debugging"""
    districts, chiefdoms, phu_names, community_names, school_names, enrollment = [], [], [], [], [], []
    
    extraction_stats = {
        'total_rows': len(df_original),
        'successful_extractions': 0,
        'failed_extractions': 0
    }
    
    if "Scan QR code" not in df_original.columns:
        st.error("❌ 'Scan QR code' column not found in the data!")
        return None, None
    
    for idx, qr_text in enumerate(df_original["Scan QR code"]):
        if pd.isna(qr_text):
            districts.append(None)
            chiefdoms.append(None)
            phu_names.append(None)
            community_names.append(None)
            school_names.append(None)
            enrollment.append(None)
            extraction_stats['failed_extractions'] += 1
            continue
            
        try:
            qr_str = str(qr_text).strip()
            
            # Extract values using regex patterns
            district_match = re.search(r"District:\s*([^\n\r]+)", qr_str)
            districts.append(district_match.group(1).strip() if district_match else None)
            
            chiefdom_match = re.search(r"Chiefdom:\s*([^\n\r]+)", qr_str)
            chiefdoms.append(chiefdom_match.group(1).strip() if chiefdom_match else None)
            
            phu_match = re.search(r"PHU name:\s*([^\n\r]+)", qr_str)
            phu_names.append(phu_match.group(1).strip() if phu_match else None)
            
            community_match = re.search(r"Community name:\s*([^\n\r]+)", qr_str)
            community_names.append(community_match.group(1).strip() if community_match else None)
            
            school_match = re.search(r"Name of school:\s*([^\n\r]+)", qr_str)
            school_names.append(school_match.group(1).strip() if school_match else None)

            # Fixed enrollment extraction
            enrollment_match = re.search(r"Enrollment:\s*([^\n\r]+)", qr_str)
            enrollment.append(enrollment_match.group(1).strip() if enrollment_match else None)
            
            extraction_stats['successful_extractions'] += 1
            
        except Exception as e:
            districts.append(None)
            chiefdoms.append(None)
            phu_names.append(None)
            community_names.append(None)
            school_names.append(None)
            enrollment.append(None)
            extraction_stats['failed_extractions'] += 1
    
    # Create extracted dataframe
    extracted_df = pd.DataFrame({
        "District": districts,
        "Chiefdom": chiefdoms,
        "PHU Name": phu_names,
        "Community Name": community_names,
        "Name of school": school_names,
        "Enrollment": enrollment
    })
    
    return extracted_df, extraction_stats

# Function to clean and convert numeric columns
def clean_numeric_data(df):
    """Clean and convert numeric columns with better error handling"""
    
    # Define all possible numeric columns
    numeric_columns = []
    
    # Add enrollment columns
    for i in range(1, 6):
        numeric_columns.extend([
            f'How many pupils are enrolled in Class {i}?',
            f'How many boys are in Class {i}?',
            f'How many girls are in Class {i}?',
            f'How many boys in Class {i} received ITNs?',
            f'How many girls in Class {i} received ITNs?'
        ])
    
    # Add other numeric columns
    numeric_columns.extend([
        'Total ITNs distributed',
        'ITNs left at the school for pupils who were absent.',
        'Enrollment'  # From QR code extraction
    ])
    
    # Clean and convert each column
    conversion_log = {}
    
    for col in numeric_columns:
        if col in df.columns:
            original_dtype = df[col].dtype
            non_null_before = df[col].count()
            
            # Convert to string first, then clean
            df[col] = df[col].astype(str)
            
            # Remove any non-numeric characters except decimal points
            df[col] = df[col].str.replace(r'[^\d.-]', '', regex=True)
            
            # Replace empty strings with NaN
            df[col] = df[col].replace('', np.nan)
            
            # Convert to numeric
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # Fill NaN with 0
            df[col] = df[col].fillna(0)
            
            # Ensure non-negative values (ITNs can't be negative)
            df[col] = df[col].abs()
            
            non_null_after = df[col].count()
            total_sum = df[col].sum()
            
            conversion_log[col] = {
                'original_dtype': original_dtype,
                'non_null_before': non_null_before,
                'non_null_after': non_null_after,
                'total_sum': total_sum
            }
    
    return df, conversion_log

# Function to calculate derived columns
def calculate_derived_columns(df):
    """Calculate derived columns with proper error handling"""
    
    # Initialize columns with zeros
    df['Enrollment_2024'] = 0
    df['Enrollment_2025'] = 0
    df['Total_Boys'] = 0
    df['Total_Girls'] = 0
    df['Boys_Received_ITNs'] = 0
    df['Girls_Received_ITNs'] = 0
    df['ITNs_Distributed_Without_Reserve'] = 0
    df['ITNs_Distributed_With_Reserve'] = 0
    
    # Enrollment in 2024 (from QR code)
    if 'Enrollment' in df.columns:
        df['Enrollment_2024'] = pd.to_numeric(df['Enrollment'], errors='coerce').fillna(0)
    
    # Enrollment in 2025 - sum of all class enrollments
    for i in range(1, 6):
        col_name = f'How many pupils are enrolled in Class {i}?'
        if col_name in df.columns:
            df['Enrollment_2025'] += pd.to_numeric(df[col_name], errors='coerce').fillna(0)
    
    # Total boys and girls
    for i in range(1, 6):
        boys_col = f'How many boys are in Class {i}?'
        girls_col = f'How many girls are in Class {i}?'
        
        if boys_col in df.columns:
            df['Total_Boys'] += pd.to_numeric(df[boys_col], errors='coerce').fillna(0)
        if girls_col in df.columns:
            df['Total_Girls'] += pd.to_numeric(df[girls_col], errors='coerce').fillna(0)
    
    # Boys and girls who received ITNs
    for i in range(1, 6):
        boys_itn_col = f'How many boys in Class {i} received ITNs?'
        girls_itn_col = f'How many girls in Class {i} received ITNs?'
        
        if boys_itn_col in df.columns:
            df['Boys_Received_ITNs'] += pd.to_numeric(df[boys_itn_col], errors='coerce').fillna(0)
        if girls_itn_col in df.columns:
            df['Girls_Received_ITNs'] += pd.to_numeric(df[girls_itn_col], errors='coerce').fillna(0)
    
    # ITNs distributed
    if 'Total ITNs distributed' in df.columns:
        df['ITNs_Distributed_Without_Reserve'] = pd.to_numeric(df['Total ITNs distributed'], errors='coerce').fillna(0)
    
    # ITNs distributed with reserve
    itn_reserve = 0
    if 'ITNs left at the school for pupils who were absent.' in df.columns:
        itn_reserve = pd.to_numeric(df['ITNs left at the school for pupils who were absent.'], errors='coerce').fillna(0)
    
    df['ITNs_Distributed_With_Reserve'] = df['ITNs_Distributed_Without_Reserve'] + itn_reserve
    
    return df

# Function to generate comprehensive summaries
def generate_summaries(df):
    """Generate District, Chiefdom, and Gender summaries with proper error handling"""
    summaries = {}
    
    # Ensure we have the required columns
    if 'District' not in df.columns:
        st.error("❌ 'District' column not found!")
        return None
    
    # Overall Summary
    total_boys_itn = df['Boys_Received_ITNs'].sum() if 'Boys_Received_ITNs' in df.columns else 0
    total_girls_itn = df['Girls_Received_ITNs'].sum() if 'Girls_Received_ITNs' in df.columns else 0
    total_enrollment_2025 = df['Enrollment_2025'].sum() if 'Enrollment_2025' in df.columns else 0
    total_itn_distributed = df['ITNs_Distributed_Without_Reserve'].sum() if 'ITNs_Distributed_Without_Reserve' in df.columns else 0
    total_itn_with_reserve = df['ITNs_Distributed_With_Reserve'].sum() if 'ITNs_Distributed_With_Reserve' in df.columns else 0
    
    overall_summary = {
        'total_schools': len(df),
        'total_districts': len(df['District'].dropna().unique()),
        'total_chiefdoms': len(df['Chiefdom'].dropna().unique()) if 'Chiefdom' in df.columns else 0,
        'total_boys_itn': int(total_boys_itn),
        'total_girls_itn': int(total_girls_itn),
        'total_enrollment_2025': int(total_enrollment_2025),
        'total_itn_distributed': int(total_itn_distributed),
        'total_itn_with_reserve': int(total_itn_with_reserve),
        'total_beneficiaries': int(total_boys_itn + total_girls_itn),
        'coverage': ((total_boys_itn + total_girls_itn) / total_enrollment_2025 * 100) if total_enrollment_2025 > 0 else 0,
        'itn_remaining': int(total_enrollment_2025 - (total_boys_itn + total_girls_itn)),
        'gender_ratio': (total_girls_itn / total_boys_itn * 100) if total_boys_itn > 0 else 0
    }
    
    summaries['overall'] = overall_summary
    
    # District Summary
    district_summary = []
    for district in df['District'].dropna().unique():
        district_data = df[df['District'] == district]
        
        enrollment_2025 = district_data['Enrollment_2025'].sum() if 'Enrollment_2025' in district_data.columns else 0
        boys_itn = district_data['Boys_Received_ITNs'].sum() if 'Boys_Received_ITNs' in district_data.columns else 0
        girls_itn = district_data['Girls_Received_ITNs'].sum() if 'Girls_Received_ITNs' in district_data.columns else 0
        itn_distributed = district_data['ITNs_Distributed_Without_Reserve'].sum() if 'ITNs_Distributed_Without_Reserve' in district_data.columns else 0
        itn_with_reserve = district_data['ITNs_Distributed_With_Reserve'].sum() if 'ITNs_Distributed_With_Reserve' in district_data.columns else 0
        
        total_beneficiaries = boys_itn + girls_itn
        coverage = (total_beneficiaries / enrollment_2025 * 100) if enrollment_2025 > 0 else 0
        
        district_stats = {
            'district': district,
            'schools': len(district_data),
            'chiefdoms': len(district_data['Chiefdom'].dropna().unique()) if 'Chiefdom' in district_data.columns else 0,
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
    
    # Chiefdom Summary (if available)
    if 'Chiefdom' in df.columns:
        chiefdom_summary = []
        for district in df['District'].dropna().unique():
            district_data = df[df['District'] == district]
            for chiefdom in district_data['Chiefdom'].dropna().unique():
                chiefdom_data = district_data[district_data['Chiefdom'] == chiefdom]
                
                enrollment_2025 = chiefdom_data['Enrollment_2025'].sum()
                boys_itn = chiefdom_data['Boys_Received_ITNs'].sum()
                girls_itn = chiefdom_data['Girls_Received_ITNs'].sum()
                itn_distributed = chiefdom_data['ITNs_Distributed_Without_Reserve'].sum()
                itn_with_reserve = chiefdom_data['ITNs_Distributed_With_Reserve'].sum()
                
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
    else:
        summaries['chiefdom'] = []
    
    return summaries

# Logo Section - Clean 4 Logo Layout
st.markdown("### Partner Organizations")
col1, col2, col3, col4 = st.columns(4)

logo_configs = [
    ("NMCP.png", "National Malaria Control Program"),
    ("icf_sl.png", "ICF Sierra Leone"),
    ("pmi.png", "PMI Evolve"),
    ("abt.png", "Abt Associates")
]

for col, (logo_file, org_name) in zip([col1, col2, col3, col4], logo_configs):
    with col:
        try:
            st.image(logo_file, width=230)
            st.markdown(f'<p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">{org_name}</p>', unsafe_allow_html=True)
        except:
            st.markdown(f"""
            <div style="width: 230px; height: 160px; border: 2px dashed #3498db; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #f8f9fd, #e3f2fd); border-radius: 10px; margin: 0 auto;">
                <div style="text-align: center; color: #666; font-size: 11px;">
                    {logo_file}<br>Not Found
                </div>
            </div>
            <p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">{org_name}</p>
            """, unsafe_allow_html=True)

st.markdown("---")

# Main App Title
st.title("📊 School Based Distribution of ITNs in Sierra Leone 2025")

# File upload
uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])

# If no file uploaded, use default
if not uploaded_file:
    uploaded_file = "sbd first_submission_clean.xlsx"
    if st.button("Use Default File (sbd first_submission_clean.xlsx)"):
        st.info("Using default file...")

if uploaded_file:
    try:
        # Read the uploaded Excel file
        with st.spinner("Loading and processing data..."):
            df_original = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File loaded successfully! Shape: {df_original.shape}")
            
            # Display basic info about the data
            st.subheader("📊 Data Overview")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(df_original))
            with col2:
                st.metric("Total Columns", len(df_original.columns))
            with col3:
                st.metric("Memory Usage", f"{df_original.memory_usage(deep=True).sum() / 1024 / 1024:.1f} MB")
            
            # Show column names for debugging
            with st.expander("🔍 View Column Names"):
                st.write("**Available Columns:**")
                for i, col in enumerate(df_original.columns, 1):
                    st.write(f"{i}. {col}")
            
            # Extract QR code data
            st.subheader("📱 QR Code Data Extraction")
            extracted_df, extraction_stats = extract_qr_data(df_original)
            
            if extracted_df is not None:
                st.write(f"**Extraction Results:**")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Rows", extraction_stats['total_rows'])
                with col2:
                    st.metric("Successful", extraction_stats['successful_extractions'])
                with col3:
                    st.metric("Failed", extraction_stats['failed_extractions'])
                
                # Add all other columns from original DataFrame
                for column in df_original.columns:
                    if column not in extracted_df.columns and column != "Scan QR code":
                        extracted_df[column] = df_original[column]
                
                # Clean and convert numeric data
                st.subheader("🔧 Data Cleaning and Processing")
                with st.spinner("Cleaning numeric data..."):
                    extracted_df, conversion_log = clean_numeric_data(extracted_df)
                
                # Show conversion results
                with st.expander("📋 Data Conversion Log"):
                    for col, log in conversion_log.items():
                        if log['total_sum'] > 0:  # Only show columns with data
                            st.write(f"**{col}**: Sum = {log['total_sum']:,.0f}")
                
                # Calculate derived columns
                extracted_df = calculate_derived_columns(extracted_df)
                
                # Generate summaries
                summaries = generate_summaries(extracted_df)
                
                if summaries:
                    # Display Enhanced Overall Summary
                    st.subheader("📊 Overall Summary - 2025 Analysis")
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
                    
                    # Check if we have meaningful data for visualizations
                    if summaries['overall']['total_beneficiaries'] > 0:
                        # Gender Analysis
                        st.subheader("👫 Gender Analysis - ITN Recipients")
                        
                        fig_gender, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
                        
                        # Left pie chart - Boys vs Girls who received ITNs
                        if summaries['overall']['total_boys_itn'] > 0 or summaries['overall']['total_girls_itn'] > 0:
                            labels = ['Boys ITN', 'Girls ITN']
                            sizes = [summaries['overall']['total_boys_itn'], summaries['overall']['total_girls_itn']]
                            colors = ['#4A90E2', '#F39C12']
                            
                            # Remove zero values
                            non_zero_data = [(label, size, color) for label, size, color in zip(labels, sizes, colors) if size > 0]
                            if non_zero_data:
                                labels, sizes, colors = zip(*non_zero_data)
                                
                                wedges, texts, autotexts = ax1.pie(sizes, labels=labels, autopct='%1.1f%%', 
                                                                    colors=colors, startangle=90)
                                ax1.set_title('ITN Recipients by Gender', fontsize=14, fontweight='bold')
                                plt.setp(autotexts, size=12, weight="bold")
                                plt.setp(texts, size=11, weight="bold")
                            else:
                                ax1.text(0.5, 0.5, 'No ITN recipient data', ha='center', va='center', 
                                        transform=ax1.transAxes, fontsize=12)
                        
                        # Right pie chart - Coverage vs Remaining
                        if summaries['overall']['total_enrollment_2025'] > 0:
                            labels2 = ['ITN Recipients', 'Not Covered']
                            sizes2 = [summaries['overall']['total_beneficiaries'], summaries['overall']['itn_remaining']]
                            colors2 = ['#27AE60', '#E74C3C']
                            
                            # Only show if there's data
                            if sum(sizes2) > 0:
                                wedges2, texts2, autotexts2 = ax2.pie(sizes2, labels=labels2, autopct='%1.1f%%',
                                                                      colors=colors2, startangle=90)
                                ax2.set_title('Overall ITN Coverage Status', fontsize=14, fontweight='bold')
                                plt.setp(autotexts2, size=12, weight="bold")
                                plt.setp(texts2, size=11, weight="bold")
                            else:
                                ax2.text(0.5, 0.5, 'No enrollment data', ha='center', va='center',
                                        transform=ax2.transAxes, fontsize=12)
                        
                        plt.tight_layout()
                        st.pyplot(fig_gender)
                        
                        # District Analysis
                        st.subheader("📊 District Analysis - Enrollment vs ITN Distribution")
                        
                        if summaries['district']:
                            districts = [d['district'] for d in summaries['district']]
                            enrollment_2025 = [d['enrollment_2025'] for d in summaries['district']]
                            beneficiaries = [d['total_beneficiaries'] for d in summaries['district']]
                            
                            fig_district, ax = plt.subplots(figsize=(14, 8))
                            x = np.arange(len(districts))
                            width = 0.35
                            
                            bars1 = ax.bar(x - width/2, enrollment_2025, width, label='2025 Enrollment', 
                                           color='#3498DB', edgecolor='navy', linewidth=1)
                            bars2 = ax.bar(x + width/2, beneficiaries, width, label='ITN Recipients', 
                                           color='#2ECC71', edgecolor='darkgreen', linewidth=1)
                            
                            ax.set_title('District Analysis: 2025 Enrollment vs ITN Recipients', 
                                        fontsize=16, fontweight='bold')
                            ax.set_xlabel('Districts', fontsize=12, fontweight='bold')
                            ax.set_ylabel('Count', fontsize=12, fontweight='bold')
                            ax.set_xticks(x)
                            ax.set_xticklabels(districts, rotation=45, ha='right')
                            ax.legend(fontsize=11)
                            ax.grid(axis='y', alpha=0.3, linestyle='--')
                            
                            # Add value labels on bars
                            for bars in [bars1, bars2]:
                                for bar in bars:
                                    height = bar.get_height()
                                    if height > 0:  # Only label non-zero bars
                                        ax.annotate(f'{int(height):,}',
                                                   xy=(bar.get_x() + bar.get_width() / 2, height),
                                                   xytext=(0, 3),
                                                   textcoords="offset points",
                                                   ha='center', va='bottom', fontsize=8, fontweight='bold')
                            
                            plt.tight_layout()
                            st.pyplot(fig_district)
                        
                        # District Summary Table
                        st.subheader("📈 District Summary Table")
                        district_summary_df = pd.DataFrame(summaries['district'])
                        if not district_summary_df.empty:
                            district_summary_df['Coverage %'] = district_summary_df['coverage'].round(1)
                            display_columns = ['district', 'schools', 'enrollment_2025', 'total_beneficiaries', 
                                              'boys_itn', 'girls_itn', 'Coverage %']
                            st.dataframe(district_summary_df[display_columns])
                            
                            # Download district summary
                            district_csv = district_summary_df.to_csv(index=False)
                            st.download_button(
                                label="📥 Download District Summary",
                                data=district_csv,
                                file_name="district_summary_2025.csv",
                                mime="text/csv"
                            )
                    else:
                        st.warning("⚠️ No ITN recipient data found. Please check if the ITN columns contain valid numeric data.")
                        
                        # Show alternative analysis with enrollment data
                        if summaries['overall']['total_enrollment_2025'] > 0:
                            st.subheader("📊 Enrollment Analysis (Alternative View)")
                            
                            # Show enrollment by district
                            if summaries['district']:
                                districts = [d['district'] for d in summaries['district']]
                                enrollment_2025 = [d['enrollment_2025'] for d in summaries['district']]
                                
                                fig_enrollment, ax = plt.subplots(figsize=(12, 6))
                                bars = ax.bar(districts, enrollment_2025, color='#3498DB', 
                                             edgecolor='navy', linewidth=1)
                                
                                ax.set_title('2025 Enrollment by District', fontsize=16, fontweight='bold')
                                ax.set_xlabel('Districts', fontsize=12, fontweight='bold')
                                ax.set_ylabel('Enrollment Count', fontsize=12, fontweight='bold')
                                ax.grid(axis='y', alpha=0.3, linestyle='--')
                                
                                # Add value labels
                                for bar in bars:
                                    height = bar.get_height()
                                    if height > 0:
                                        ax.annotate(f'{int(height):,}',
                                                   xy=(bar.get_x() + bar.get_width() / 2, height),
                                                   xytext=(0, 3),
                                                   textcoords="offset points",
                                                   ha='center', va='bottom', fontsize=10, fontweight='bold')
                                
                                plt.xticks(rotation=45, ha='right')
                                plt.tight_layout()
                                st.pyplot(fig_enrollment)
                    
                    # Enhanced Data Export
                    st.subheader("📥 Export Enhanced Data")
                    
                    # Add calculated columns for export
                    display_cols = ['District', 'Chiefdom', 'Name of school', 'Enrollment_2025', 'Total_Boys', 'Total_Girls', 
                                   'Boys_Received_ITNs', 'Girls_Received_ITNs', 'ITNs_Distributed_With_Reserve']
                    
                    available_cols = [col for col in display_cols if col in extracted_df.columns]
                    
                    st.write("**Enhanced Data Preview:**")
                    st.dataframe(extracted_df[available_cols].head(10))
                    
                    # CSV download
                    csv_data = extracted_df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Complete Enhanced Data as CSV",
                        data=csv_data,
                        file_name=f"enhanced_school_data_2025_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv"
                    )
                    
                    # Excel download with multiple sheets
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        extracted_df.to_excel(writer, sheet_name='Complete Data', index=False)
                        if summaries['district']:
                            pd.DataFrame(summaries['district']).to_excel(writer, sheet_name='District Summary', index=False)
                        if summaries['chiefdom']:
                            pd.DataFrame(summaries['chiefdom']).to_excel(writer, sheet_name='Chiefdom Summary', index=False)
                    
                    excel_data = excel_buffer.getvalue()
                    st.download_button(
                        label="📊 Download Excel Report (Multiple Sheets)",
                        data=excel_data,
                        file_name=f"SBD_Report_2025_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Final summary
                    st.info(f"""
                    📋 **Final Dataset Summary**: 
                    - Total Records: {len(extracted_df):,}
                    - 2025 Enrollment: {summaries['overall']['total_enrollment_2025']:,}
                    - ITN Recipients: {summaries['overall']['total_beneficiaries']:,}
                    - Overall Coverage: {summaries['overall']['coverage']:.1f}%
                    - Districts: {summaries['overall']['total_districts']}
                    - Chiefdoms: {summaries['overall']['total_chiefdoms']}
                    """)
                
                else:
                    st.error("❌ Failed to generate summaries. Please check your data format.")
            
            else:
                st.error("❌ Failed to extract QR code data. Please check if the 'Scan QR code' column exists and contains valid data.")
    
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.write("**Error Details:**")
        st.code(str(e))
        
        # Show some debugging info
        try:
            df_debug = pd.read_excel(uploaded_file)
            st.write(f"File shape: {df_debug.shape}")
            st.write("First few column names:")
            st.write(df_debug.columns[:10].tolist())
        except:
            st.write("Could not read file for debugging")

else:
    st.info(" error")

    """)
