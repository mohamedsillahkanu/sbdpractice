import streamlit as st
import pandas as pd
import re
import numpy as np
import matplotlib.pyplot as plt
import geopandas as gpd
from io import BytesIO
import base64
import warnings
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
</style>
""", unsafe_allow_html=True)

# Function to save maps as PNG and return BytesIO object
def save_map_as_png(fig, filename_prefix):
    """Save matplotlib figure as PNG and return BytesIO object"""
    try:
        buffer = BytesIO()
        fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight', 
                   facecolor='white', edgecolor='none', pad_inches=0.2)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Error saving map {filename_prefix}: {str(e)}")
        return None

# Function to parse GPS coordinates
def parse_gps_coordinates(gps_string):
    """Parse GPS coordinates from string format"""
    if pd.isna(gps_string):
        return None, None
    
    try:
        gps_str = str(gps_string).strip()
        
        # Handle comma-separated format: lat,lon
        if ',' in gps_str:
            parts = gps_str.split(',')
            if len(parts) == 2:
                lat = float(parts[0].strip())
                lon = float(parts[1].strip())
                
                # Validate coordinates for Sierra Leone
                if 6.0 <= lat <= 11.0 and -14.0 <= lon <= -10.0:
                    return lat, lon
        
        return None, None
    except (ValueError, TypeError):
        return None, None

# Function to generate comprehensive summaries
def generate_summaries(df):
    """Generate District, Chiefdom, and Gender summaries"""
    summaries = {}
    
    # Overall Summary
    overall_summary = {
        'total_schools': len(df),
        'total_districts': len(df['District'].dropna().unique()),
        'total_chiefdoms': len(df['Chiefdom'].dropna().unique()),
        'total_boys': 0,
        'total_girls': 0,
        'total_enrollment': 0,
        'total_itn': 0
    }
    
    # Calculate totals using the correct columns
    for class_num in range(1, 6):
        # Total enrollment
        enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
        if enrollment_col in df.columns:
            overall_summary['total_enrollment'] += int(df[enrollment_col].fillna(0).sum())
        
        # Boys and girls for ITN calculation
        boys_col = f"How many boys in Class {class_num} received ITNs?"
        girls_col = f"How many girls in Class {class_num} received ITNs?"
        if boys_col in df.columns:
            overall_summary['total_boys'] += int(df[boys_col].fillna(0).sum())
        if girls_col in df.columns:
            overall_summary['total_girls'] += int(df[girls_col].fillna(0).sum())
    
    # Total ITNs = boys + girls
    overall_summary['total_itn'] = overall_summary['total_boys'] + overall_summary['total_girls']
    
    # Calculate coverage
    overall_summary['coverage'] = (overall_summary['total_itn'] / overall_summary['total_enrollment'] * 100) if overall_summary['total_enrollment'] > 0 else 0
    overall_summary['itn_remaining'] = overall_summary['total_enrollment'] - overall_summary['total_itn']
    
    summaries['overall'] = overall_summary
    
    # District Summary
    district_summary = []
    for district in df['District'].dropna().unique():
        district_data = df[df['District'] == district]
        district_stats = {
            'district': district,
            'schools': len(district_data),
            'chiefdoms': len(district_data['Chiefdom'].dropna().unique()),
            'boys': 0,
            'girls': 0,
            'enrollment': 0,
            'itn': 0
        }
        
        for class_num in range(1, 6):
            enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
            if enrollment_col in district_data.columns:
                district_stats['enrollment'] += int(district_data[enrollment_col].fillna(0).sum())
            
            boys_col = f"How many boys in Class {class_num} received ITNs?"
            girls_col = f"How many girls in Class {class_num} received ITNs?"
            if boys_col in district_data.columns:
                district_stats['boys'] += int(district_data[boys_col].fillna(0).sum())
            if girls_col in district_data.columns:
                district_stats['girls'] += int(district_data[girls_col].fillna(0).sum())
        
        district_stats['itn'] = district_stats['boys'] + district_stats['girls']
        district_stats['coverage'] = (district_stats['itn'] / district_stats['enrollment'] * 100) if district_stats['enrollment'] > 0 else 0
        district_stats['itn_remaining'] = district_stats['enrollment'] - district_stats['itn']
        
        district_summary.append(district_stats)
    
    summaries['district'] = district_summary
    
    # Chiefdom Summary
    chiefdom_summary = []
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
                enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
                if enrollment_col in chiefdom_data.columns:
                    chiefdom_stats['enrollment'] += int(chiefdom_data[enrollment_col].fillna(0).sum())
                
                boys_col = f"How many boys in Class {class_num} received ITNs?"
                girls_col = f"How many girls in Class {class_num} received ITNs?"
                if boys_col in chiefdom_data.columns:
                    chiefdom_stats['boys'] += int(chiefdom_data[boys_col].fillna(0).sum())
                if girls_col in chiefdom_data.columns:
                    chiefdom_stats['girls'] += int(chiefdom_data[girls_col].fillna(0).sum())
            
            chiefdom_stats['itn'] = chiefdom_stats['boys'] + chiefdom_stats['girls']
            chiefdom_stats['coverage'] = (chiefdom_stats['itn'] / chiefdom_stats['enrollment'] * 100) if chiefdom_stats['enrollment'] > 0 else 0
            chiefdom_stats['itn_remaining'] = chiefdom_stats['enrollment'] - chiefdom_stats['itn']
            
            chiefdom_summary.append(chiefdom_stats)
    
    summaries['chiefdom'] = chiefdom_summary
    
    return summaries

# Function to create beautiful maps
def create_district_map(gdf, district_name, gps_data, title_suffix=""):
    """Create a beautiful district map with GPS points"""
    try:
        # Filter shapefile for this district
        district_gdf = gdf[gdf['FIRST_DNAM'] == district_name].copy()
        
        if len(district_gdf) == 0:
            st.warning(f"No shapefile data found for {district_name} district")
            return None
        
        # Create figure with high quality
        fig, ax = plt.subplots(figsize=(16, 12))
        
        # Plot chiefdom boundaries
        district_gdf.plot(ax=ax, color='lightblue', edgecolor='navy', alpha=0.7, linewidth=2)
        
        # Extract GPS coordinates
        coords = []
        if gps_data is not None and "GPS Location" in gps_data.columns:
            for _, row in gps_data.iterrows():
                lat, lon = parse_gps_coordinates(row["GPS Location"])
                if lat is not None and lon is not None:
                    coords.append([lat, lon])
        
        # Plot GPS points
        if coords:
            lats, lons = zip(*coords)
            
            # Create beautiful scatter plot
            scatter = ax.scatter(
                lons, lats,
                c='red',
                s=200,
                alpha=0.9,
                edgecolors='white',
                linewidth=3,
                zorder=100,
                label=f'Schools ({len(coords)})',
                marker='o'
            )
            
            # Add numbered labels for each point
            for i, (lat, lon) in enumerate(coords):
                ax.annotate(f'{i+1}', 
                           (lon, lat),
                           xytext=(0, 0), 
                           textcoords='offset points',
                           fontsize=12,
                           fontweight='bold',
                           color='white',
                           ha='center',
                           va='center',
                           bbox=dict(boxstyle='circle,pad=0.1', facecolor='red', alpha=0.8))
            
            # Set map extent to show all points with padding
            margin = 0.02
            ax.set_xlim(min(lons) - margin, max(lons) + margin)
            ax.set_ylim(min(lats) - margin, max(lats) + margin)
            
        else:
            # If no GPS coordinates, show the district boundaries
            bounds = district_gdf.total_bounds
            margin = 0.01
            ax.set_xlim(bounds[0] - margin, bounds[2] + margin)
            ax.set_ylim(bounds[1] - margin, bounds[3] + margin)
        
        # Add chiefdom labels
        for idx, row in district_gdf.iterrows():
            if 'FIRST_CHIE' in row and pd.notna(row['FIRST_CHIE']):
                centroid = row.geometry.centroid
                ax.annotate(
                    row['FIRST_CHIE'], 
                    (centroid.x, centroid.y),
                    fontsize=11,
                    ha='center',
                    va='center',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8, edgecolor='navy'),
                    fontweight='bold'
                )
        
        # Customize plot
        title = f'{district_name} District - Geographic Distribution{title_suffix}'
        if coords:
            title += f' | {len(coords)} GPS Points | {len(district_gdf)} Chiefdoms'
        
        ax.set_title(title, fontsize=18, fontweight='bold', pad=20)
        ax.set_xlabel('Longitude', fontsize=14, fontweight='bold')
        ax.set_ylabel('Latitude', fontsize=14, fontweight='bold')
        
        # Add legend
        if coords:
            ax.legend(fontsize=14, loc='best')
        
        # Add grid
        ax.grid(True, alpha=0.3, linestyle='--')
        
        # Improve layout
        plt.tight_layout()
        
        return fig
        
    except Exception as e:
        st.error(f"Error creating map for {district_name}: {str(e)}")
        return None

# Logo Section
st.markdown("### Partner Organizations")
col1, col2, col3, col4 = st.columns(4)

logo_names = ["NMCP", "ICF Sierra Leone", "PMI Evolve", "Abt Associates"]
logo_files = ["NMCP.png", "icf_sl.png", "pmi.png", "abt.png"]

for i, (col, name, file) in enumerate(zip([col1, col2, col3, col4], logo_names, logo_files)):
    with col:
        try:
            st.image(file, width=230)
            st.markdown(f'<p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">{name}</p>', unsafe_allow_html=True)
        except:
            st.markdown(f"""
            <div style="width: 230px; height: 160px; border: 2px dashed #3498db; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #f8f9fd, #e3f2fd); border-radius: 10px; margin: 0 auto;">
                <div style="text-align: center; color: #666; font-size: 11px;">
                    {file}<br>Not Found
                </div>
            </div>
            <p style="text-align: center; font-size: 12px; font-weight: 600; color: #2c3e50; margin-top: 5px;">{name}</p>
            """, unsafe_allow_html=True)

st.markdown("---")

# Main Title
st.title("📊 School Based Distribution of ITNs in Sierra Leone")

# Initialize session state for storing maps
if 'map_images' not in st.session_state:
    st.session_state.map_images = {}

# File upload section
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])

# Use default file if no upload
if uploaded_file is None:
    uploaded_file = "sbd first_submission_clean.xlsx"
    st.info("Using default file: sbd first_submission_clean.xlsx")

try:
    # Read the Excel file
    if isinstance(uploaded_file, str):
        df_original = pd.read_excel(uploaded_file)
    else:
        df_original = pd.read_excel(uploaded_file)
    
    st.success(f"✅ Data loaded successfully! {len(df_original)} records found.")
    
    # Load shapefile
    try:
        gdf = gpd.read_file("Chiefdom2021.shp")
        st.success(f"✅ Shapefile loaded successfully! {len(gdf)} chiefdoms found.")
    except Exception as e:
        st.error(f"❌ Could not load shapefile: {e}")
        gdf = None
    
    # Extract data from QR codes
    st.subheader("📋 Data Extraction from QR Codes")
    
    with st.spinner("Extracting data from QR codes..."):
        # Create empty lists to store extracted data
        districts, chiefdoms, phu_names, community_names, school_names = [], [], [], [], []
        
        # Process each row in the "Scan QR code" column
        for qr_text in df_original["Scan QR code"]:
            if pd.isna(qr_text):
                districts.append(None)
                chiefdoms.append(None)
                phu_names.append(None)
                community_names.append(None)
                school_names.append(None)
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
        
        # Create extracted DataFrame
        extracted_df = pd.DataFrame({
            "District": districts,
            "Chiefdom": chiefdoms,
            "PHU Name": phu_names,
            "Community Name": community_names,
            "School Name": school_names
        })
        
        # Add all other columns from the original DataFrame
        for column in df_original.columns:
            if column != "Scan QR code":
                extracted_df[column] = df_original[column]
    
    st.success(f"✅ Data extraction completed! Extracted information for {len(extracted_df)} records.")
    
    # Generate summaries
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
    
    # Geographic Distribution Maps
    st.subheader("🗺️ Geographic Distribution Maps")
    
    if gdf is not None:
        # Create overall Sierra Leone map
        st.write("### Sierra Leone - National Overview")
        
        with st.spinner("Creating national overview map..."):
            fig_overall, ax_overall = plt.subplots(figsize=(16, 12))
            
            # Plot all chiefdoms
            gdf.plot(ax=ax_overall, color='lightgray', edgecolor='black', alpha=0.6, linewidth=0.5)
            
            # Plot district boundaries
            if 'FIRST_DNAM' in gdf.columns:
                district_boundaries = gdf.dissolve(by='FIRST_DNAM')
                district_boundaries.plot(ax=ax_overall, facecolor='none', edgecolor='blue', linewidth=3)
                
                # Add district labels
                for idx, row in district_boundaries.iterrows():
                    centroid = row.geometry.centroid
                    ax_overall.annotate(
                        idx,
                        (centroid.x, centroid.y),
                        fontsize=12,
                        fontweight='bold',
                        ha='center',
                        va='center',
                        color='black',
                        bbox=dict(boxstyle='round,pad=0.3', facecolor='yellow', alpha=0.8)
                    )
            
            # Plot all GPS points
            all_coords = []
            if "GPS Location" in extracted_df.columns:
                for _, row in extracted_df.iterrows():
                    lat, lon = parse_gps_coordinates(row["GPS Location"])
                    if lat is not None and lon is not None:
                        all_coords.append([lat, lon])
            
            if all_coords:
                lats, lons = zip(*all_coords)
                ax_overall.scatter(
                    lons, lats,
                    c='red',
                    s=80,
                    alpha=0.8,
                    edgecolors='white',
                    linewidth=1,
                    zorder=100,
                    label=f'Schools ({len(all_coords)})'
                )
                ax_overall.legend(fontsize=14, loc='best')
            
            ax_overall.set_title(f'Sierra Leone - School Distribution Overview | {len(all_coords)} Schools across {summaries["overall"]["total_districts"]} Districts', 
                               fontsize=18, fontweight='bold', pad=20)
            ax_overall.set_xlabel('Longitude', fontsize=14)
            ax_overall.set_ylabel('Latitude', fontsize=14)
            ax_overall.grid(True, alpha=0.3)
            
            plt.tight_layout()
            st.pyplot(fig_overall)
            
            # Save overall map
            overall_buffer = save_map_as_png(fig_overall, "Sierra_Leone_Overall")
            if overall_buffer:
                st.session_state.map_images['sierra_leone_overall'] = overall_buffer
        
        # Create district-specific maps
        st.write("### District-Specific Maps with GPS Coordinates")
        
        # Get all unique districts
        districts = extracted_df['District'].dropna().unique()
        
        for district in sorted(districts):
            st.write(f"#### {district} District")
            
            with st.spinner(f"Creating map for {district} district..."):
                # Filter data for this district
                district_data = extracted_df[extracted_df['District'] == district]
                
                # Create the map
                fig = create_district_map(gdf, district, district_data)
                
                if fig is not None:
                    st.pyplot(fig)
                    
                    # Save the map
                    map_buffer = save_map_as_png(fig, f"{district}_District_Map")
                    if map_buffer:
                        st.session_state.map_images[f'{district.lower()}_district'] = map_buffer
                    
                    plt.close(fig)  # Close to free memory
                else:
                    st.warning(f"Could not create map for {district} district")
    
    # Enhanced Analysis Charts
    st.subheader("📊 Enhanced Analysis Charts")
    
    # Gender Analysis
    st.write("### Gender Distribution Analysis")
    
    # Overall gender pie chart
    fig_gender, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
    
    # Overall gender distribution
    labels = ['Boys', 'Girls']
    sizes = [summaries['overall']['total_boys'], summaries['overall']['total_girls']]
    colors = ['#4A90E2', '#F39C12']
    
    ax1.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
    ax1.set_title('Overall Gender Distribution', fontsize=16, fontweight='bold')
    
    # Gender by district
    districts = [d['district'] for d in summaries['district']]
    boys_counts = [d['boys'] for d in summaries['district']]
    girls_counts = [d['girls'] for d in summaries['district']]
    
    x = np.arange(len(districts))
    width = 0.35
    
    ax2.bar(x - width/2, boys_counts, width, label='Boys', color='#4A90E2')
    ax2.bar(x + width/2, girls_counts, width, label='Girls', color='#F39C12')
    
    ax2.set_title('Gender Distribution by District', fontsize=16, fontweight='bold')
    ax2.set_xlabel('Districts')
    ax2.set_ylabel('Number of Students')
    ax2.set_xticks(x)
    ax2.set_xticklabels(districts, rotation=45, ha='right')
    ax2.legend()
    ax2.grid(axis='y', alpha=0.3)
    
    plt.tight_layout()
    st.pyplot(fig_gender)
    
    # Save gender chart
    gender_buffer = save_map_as_png(fig_gender, "Gender_Analysis")
    if gender_buffer:
        st.session_state.map_images['gender_analysis'] = gender_buffer
    
    # Enrollment vs ITN Distribution Analysis
    st.write("### Enrollment vs ITN Distribution Analysis")
    
    # Calculate district analysis
    district_analysis = []
    for district in extracted_df['District'].dropna().unique():
        district_data = extracted_df[extracted_df['District'] == district]
        
        total_enrollment = 0
        total_boys = 0
        total_girls = 0
        
        for class_num in range(1, 6):
            enrollment_col = f"How many pupils are enrolled in Class {class_num}?"
            boys_col = f"How many boys in Class {class_num} received ITNs?"
            girls_col = f"How many girls in Class {class_num} received ITNs?"
            
            if enrollment_col in district_data.columns:
                total_enrollment += int(district_data[enrollment_col].fillna(0).sum())
            if boys_col in district_data.columns:
                total_boys += int(district_data[boys_col].fillna(0).sum())
            if girls_col in district_data.columns:
                total_girls += int(district_data[girls_col].fillna(0).sum())
        
        total_itn = total_boys + total_girls
        itn_remaining = total_enrollment - total_itn
        coverage = (total_itn / total_enrollment * 100) if total_enrollment > 0 else 0
        
        district_analysis.append({
            'District': district,
            'Total_Enrollment': total_enrollment,
            'Total_ITN': total_itn,
            'ITN_Remaining': itn_remaining,
            'Coverage': coverage
        })
    
    district_df = pd.DataFrame(district_analysis)
    
    # Create enhanced bar chart
    fig_enhanced, ax_enhanced = plt.subplots(figsize=(16, 10))
    
    x = np.arange(len(district_df['District']))
    width = 0.25
    
    bars1 = ax_enhanced.bar(x - width, district_df['Total_Enrollment'], width, 
                           label='Total Enrollment', color='#47B5FF')
    bars2 = ax_enhanced.bar(x, district_df['Total_ITN'], width, 
                           label='ITNs Distributed', color='lightcoral')
    bars3 = ax_enhanced.bar(x + width, district_df['ITN_Remaining'], width, 
                           label='ITNs Remaining', color='hotpink')
    
    ax_enhanced.set_title('District Analysis: Enrollment vs ITN Distribution', fontsize=18, fontweight='bold')
    ax_enhanced.set_xlabel('Districts', fontsize=14)
    ax_enhanced.set_ylabel('Number of Students/ITNs', fontsize=14)
    ax_enhanced.set_xticks(x)
    ax_enhanced.set_xticklabels(district_df['District'], rotation=45, ha='right')
    ax_enhanced.legend(fontsize=12)
    ax_enhanced.grid(axis='y', alpha=0.3)
    
    # Add value labels
    def add_value_labels(bars):
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax_enhanced.annotate(f'{int(height):,}',
                                   xy=(bar.get_x() + bar.get_width() / 2, height),
                                   xytext=(0, 3),
                                   textcoords="offset points",
                                   ha='center', va='bottom', fontsize=9, fontweight='bold')
    
    add_value_labels(bars1)
    add_value_labels(bars2)
    add_value_labels(bars3)
    
    plt.tight_layout()
    st.pyplot(fig_enhanced)
    
    # Save enhanced chart
    enhanced_buffer = save_map_as_png(fig_enhanced, "Enhanced_Analysis")
    if enhanced_buffer:
        st.session_state.map_images['enhanced_analysis'] = enhanced_buffer
    
    # Display Summary Tables
    st.subheader("📈 Summary Tables")
    
    # District Summary Table
    st.write("### District Summary")
    district_summary_df = pd.DataFrame(summaries['district'])
    district_summary_df.columns = ['District', 'Schools', 'Chiefdoms', 'Boys', 'Girls', 'Enrollment', 'ITNs', 'Coverage (%)', 'ITNs Remaining']
    st.dataframe(district_summary_df, use_container_width=True)
    
    # Chiefdom Summary Table
    st.write("### Chiefdom Summary")
    chiefdom_summary_df = pd.DataFrame(summaries['chiefdom'])
    chiefdom_summary_df.columns = ['District', 'Chiefdom', 'Schools', 'Boys', 'Girls', 'Enrollment', 'ITNs', 'Coverage (%)', 'ITNs Remaining']
    st.dataframe(chiefdom_summary_df, use_container_width=True)
    
    # Download Section
    st.subheader("📥 Download Options")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        # CSV Download
        csv_data = extracted_df.to_csv(index=False)
        st.download_button(
            label="📄 Download Data (CSV)",
            data=csv_data,
            file_name="sbd_extracted_data.csv",
            mime="text/csv"
        )
    
    with col2:
        # District Summary CSV
        district_csv = district_summary_df.to_csv(index=False)
        st.download_button(
            label="📊 District Summary (CSV)",
            data=district_csv,
            file_name="district_summary.csv",
            mime="text/csv"
        )
    
    with col3:
        # Chiefdom Summary CSV
        chiefdom_csv = chiefdom_summary_df.to_csv(index=False)
        st.download_button(
            label="📋 Chiefdom Summary (CSV)",
            data=chiefdom_csv,
            file_name="chiefdom_summary.csv",
            mime="text/csv"
        )
    
    with col4:
        # Excel Download
        if st.button("📊 Generate Excel Report"):
            with st.spinner("Generating Excel report..."):
                try:
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        extracted_df.to_excel(writer, sheet_name='Raw Data', index=False)
                        district_summary_df.to_excel(writer, sheet_name='District Summary', index=False)
                        chiefdom_summary_df.to_excel(writer, sheet_name='Chiefdom Summary', index=False)
                    
                    excel_data = excel_buffer.getvalue()
                    
                    st.download_button(
                        label="💾 Download Excel Report",
                        data=excel_data,
                        file_name="sbd_complete_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("✅ Excel report generated successfully!")
                except Exception as e:
                    st.error(f"Error generating Excel report: {str(e)}")
    
    # Display saved maps information
    if st.session_state.map_images:
        st.subheader("🗺️ Generated Maps")
        st.success(f"✅ {len(st.session_state.map_images)} maps have been generated and saved!")
        
        with st.expander("View Generated Maps List"):
            for map_name in st.session_state.map_images.keys():
                st.write(f"• {map_name}")

except Exception as e:
    st.error(f"❌ Error loading data: {str(e)}")
    st.info("Please ensure you have the correct Excel file and shapefile in the working directory.")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 12px; margin-top: 2rem;">
    <p>School-Based Distribution (SBD) Analysis Dashboard</p>
    <p>Developed for Sierra Leone Malaria Control Program</p>
</div>
""", unsafe_allow_html=True)
