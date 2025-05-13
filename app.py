import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import time, datetime
import io
import base64

# Set page config with JABIL colors
st.set_page_config(page_title="Loss Time Analysis", layout="wide", 
                  initial_sidebar_state="expanded")

# Define JABIL colors
jabil_blue = "#00517C"
jabil_green = "#28A745"

# Custom CSS to apply JABIL colors
st.markdown("""
<style>
    .main {
        background-color: #FFFFFF;
    }
    .stButton>button {
        background-color: #00517C;
        color: white;
    }
    .stHeader {
        color: #00517C;
    }
    .stDataFrame {
        border: 1px solid #00517C;
    }
    h1, h2, h3 {
        color: #00517C;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: white;
        border-radius: 4px 4px 0 0;
        color: #00517C;
        border-left: 1px solid #00517C;
        border-right: 1px solid #00517C;
        border-top: 1px solid #00517C;
    }
    .stTabs [aria-selected="true"] {
        background-color: #00517C;
        color: white;
    }
    .css-1kyxreq {
        justify-content: center;
    }
    .css-1544g2n {
        padding-top: 4rem;
    }
</style>
""", unsafe_allow_html=True)

# Display the JABIL logo
def add_logo():
    # JABIL logo base64 encoded
    logo_base64 = """
    iVBORw0KGgoAAAANSUhEUgAAAZQAAABOCAYAAAAjsVSOAAAACXBIWXMAAAsTAAALEwEAmpwYAAAF4GlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS42LWMxNDggNzkuMTY0MDM2LCAyMDE5LzA4LzEzLTAxOjA2OjU3ICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxOSAoTWFjaW50b3NoKSIgeG1wOkNyZWF0ZURhdGU9IjIwMjMtMDUtMTZUMDk6NDE6MDYrMDI6MDAiIHhtcDpNb2RpZnlEYXRlPSIyMDIzLTA1LTE2VDA5OjQ0OjIzKzAyOjAwIiB4bXA6TWV0YWRhdGFEYXRlPSIyMDIzLTA1LTE2VDA5OjQ0OjIzKzAyOjAwIiBkYzpmb3JtYXQ9ImltYWdlL3BuZyIgcGhvdG9zaG9wOkNvbG9yTW9kZT0iMyIgcGhvdG9zaG9wOklDQ1Byb2ZpbGU9InNSR0IgSUVDNjE5NjYtMi4xIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjI0MmU3M2MwLWRk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
    """
    st.markdown(
        f"""
        <div style="display: flex; justify-content: center; margin-bottom: 20px;">
            <img src="data:image/png;base64,{logo_base64}" alt="JABIL Logo" width="300">
        </div>
        """,
        unsafe_allow_html=True,
    )

add_logo()

st.title("Loss Time Analysis Dashboard")

st.markdown("""
This application analyzes access logs to calculate time spent outside designated areas.
Upload an Excel file with access logs containing entry/exit records.
""")

# Upload file
uploaded_file = st.file_uploader("Upload the Excel access log file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Main processing function
    @st.cache_data
    def process_data(file):
        try:
            df = pd.read_excel(file)
            st.success("File uploaded successfully!")
            
            # Display sample data
            st.subheader("Sample Data")
            st.dataframe(df.head())
            
            # Step 1: Extract ID from 'Who' column
            df['ID'] = df['Who'].str.extract(r'(\d+)', expand=False)
            
            # Step 2: List of IDs to exclude (as in your notebook)
            excluded_ids = [
                "100022883", "100028130", "100033411", "100034976", "100037620", "100038232", "100039869", "100040212",
                "100041020", "100041781", "100041693", "100041696", "100043374", "100043704", "100044690", "100047163",
                "100048938", "100049561", "100052278", "100054726", "100055362", "100057600", "100048754", "100062079",
                "100064653", "100064655", "2434673", "2588319", "2946384", "3013011", "2436060", "3227019", "3229196",
                "3242216", "3252881", "3251966", "3294315", "3301049", "3301547", "3334845", "3328895", "3442257",
                "3524553", "3568419", "3616448", "3616442", "3632147", "3641947", "3650221", "3667379", "3644405",
                "3677808", "3631443", "3706789", "3793827", "3808739", "3810788", "3812296", "3810787", "3883004",
                "3929368", "3988422", "4002219", "4039413", "4094250", "4117222", "3293057", "3645325", "3937753",
                "3957764", "3808729", "4033286", "4115553", "4116558", "4090423", "4090422", "100046032", "100046033",
                "2279149", "3329836", "3329835", "3329837", "3451556", "3510265", "3644227", "4048557", "4091122",
                "4036324"
            ]
            
            # Step 3: Filter out rows with IDs in the exclusion list
            df = df[~df['ID'].isin(excluded_ids)]
            
            # Keep only rows with PEDESTRIAN IN/OUT
            df = df[df['Where'].str.contains('PEDESTRIAN IN|PEDESTRIAN OUT', na=False)]
            
            # Drop entries with generic names
            df = df[~df['Who'].str.contains('Temporary|Kartu|Costumer|999999', na=False, case=False)]
            
            # Drop unnamed column if exists
            if 'Unnamed: 5' in df.columns:
                df = df.drop(columns=['Unnamed: 5'])
            
            # Parse datetime column
            df['When'] = pd.to_datetime(df['When'], format='%d/%m/%Y %H:%M', errors='coerce')
            
            # Drop rows with failed date parsing
            df = df.dropna(subset=['When'])
            
            # Assign IN / OUT
            df['Direction'] = df['Where'].apply(lambda x: 'IN' if 'PEDESTRIAN IN' in x else 'OUT')
            
            # Count IN and OUT per Who
            in_count = df[df['Direction'] == 'IN'].groupby('Who').size()
            out_count = df[df['Direction'] == 'OUT'].groupby('Who').size()
            
            # Combine into one DataFrame
            direction_counts = pd.DataFrame({'IN': in_count, 'OUT': out_count}).fillna(0)
            
            # Keep only those with equal IN = OUT
            valid_whos = direction_counts[direction_counts['IN'] == direction_counts['OUT']].index
            df = df[df['Who'].isin(valid_whos)]
            
            # Extract name
            df['Nama'] = df['Who'].str.extract(r',\s*(.*)$')[0]
            
            # Clean leading/trailing spaces
            df['Nama'] = df['Nama'].str.strip()
            
            # Filter specific names
            df = df[df['Nama'] != 'Dari A011']
            
            # Process pairing
            pairings = []
            
            # Iterate for each person
            for name, group in df.groupby('Nama'):
                group = group.reset_index(drop=True)
                used_idx = set()
            
                for i, row in group.iterrows():
                    if row['Direction'] == 'OUT' and i not in used_idx:
                        out_time = row['When']
                        out_idx = i
            
                        # Find first IN after OUT on the same day
                        for j in range(i+1, len(group)):
                            if group.loc[j, 'Direction'] == 'IN' and group.loc[j, 'When'].date() == out_time.date() and j not in used_idx:
                                in_time = group.loc[j, 'When']
                                duration = (in_time - out_time).total_seconds() / 60
            
                                pairings.append({
                                    'Name': name,
                                    'Date': out_time.date(),
                                    'OUT': out_time,
                                    'IN': in_time,
                                    'Duration (minutes)': round(duration)
                                })
            
                                used_idx.update([out_idx, j])
                                break  # stop after finding first pair
            
            # Create paired DataFrame
            paired_df = pd.DataFrame(pairings)
            
            # Filter: remove 0 duration, skip duration > 210 minutes
            paired_df = paired_df[(paired_df['Duration (minutes)'] > 0) & (paired_df['Duration (minutes)'] <= 210)]
            
            # Define shift slots
            shift_slots = {
                'Shift 1': [
                    {'start': time(6,30), 'end': time(9,30), 'jenis': '-'},
                    {'start': time(9,30), 'end': time(11,0), 'jenis': 'Tea Break', 'durasi': 20},
                    {'start': time(11,0), 'end': time(13,30), 'jenis': 'Lunch', 'durasi': 45},
                    {'start': time(13,30), 'end': time(14,0), 'jenis': '-'}
                ],
                'Shift 2': [
                    {'start': time(14,30), 'end': time(15,0), 'jenis': '-'},
                    {'start': time(15,0), 'end': time(16,40), 'jenis': 'Afternoon Prayer', 'durasi': 20},
                    {'start': time(16,40), 'end': time(18,0), 'jenis': '-'},
                    {'start': time(18,0), 'end': time(19,0), 'jenis': 'Dinner', 'durasi': 45},
                    {'start': time(19,0), 'end': time(20,30), 'jenis': '-'},
                    {'start': time(20,30), 'end': time(21,0), 'jenis': 'Evening Prayer', 'durasi': 20},
                    {'start': time(21,0), 'end': time(22,0), 'jenis': '-'}
                ],
                'Shift 3': [
                    {'start': time(0,0), 'end': time(1,30), 'jenis': 'Meal', 'durasi': 30},
                    {'start': time(1,30), 'end': time(4,0), 'jenis': '-'},
                    {'start': time(4,0), 'end': time(5,20), 'jenis': 'Morning Prayer', 'durasi': 30},
                    {'start': time(5,20), 'end': time(6,0), 'jenis': '-'},
                    {'start': time(22,30), 'end': time(23,59,59), 'jenis': '-'}
                ]
            }
            
            # Determine shift based on OUT time
            def determine_shift(out_time):
                hour = out_time.hour
                if 6 <= hour < 14: return 'Shift 1'
                elif 14 <= hour < 22: return 'Shift 2'
                else: return 'Shift 3'
            
            # Calculate overlap minutes between two time ranges
            def overlap_minutes(start1, end1, start2, end2):
                latest_start = max(start1, start2)
                earliest_end = min(end1, end2)
                overlap = (earliest_end - latest_start).total_seconds() / 60
                return max(0, overlap)
            
            # Calculate category and loss time
            def calculate_loss_time(row):
                shift = determine_shift(row['OUT'])
                slots = shift_slots[shift]
                total_break = 0
                category = 'Work'
            
                for slot in slots:
                    slot_start = datetime.combine(row['OUT'].date(), slot['start'])
                    slot_end = datetime.combine(row['OUT'].date(), slot['end'])
                    leave = row['OUT']
                    enter = row['IN']
            
                    overlap = overlap_minutes(leave, enter, slot_start, slot_end)
                    if overlap > 0 and slot['jenis'] != '-':
                        total_break += min(overlap, slot.get('durasi', 0))
                        category = slot['jenis']  # Take one matching category
            
                loss = row['Duration (minutes)'] - total_break
                return pd.Series([category, max(0, loss)], index=['Category', 'Loss Time (minutes)'])
            
            # Apply calculations
            paired_df[['Category', 'Loss Time (minutes)']] = paired_df.apply(calculate_loss_time, axis=1)
            
            return paired_df
        
        except Exception as e:
            st.error(f"Error processing file: {e}")
            return None
    
    # Process the data
    result_df = process_data(uploaded_file)
    
    if result_df is not None:
        # Visualizations and analysis
        st.subheader("Data Analysis Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write(f"Total records: {len(result_df)}")
            st.write(f"Total employees: {result_df['Name'].nunique()}")
            st.write(f"Date range: {result_df['Date'].min()} to {result_df['Date'].max()}")
        
        with col2:
            st.write(f"Average loss time: {result_df['Loss Time (minutes)'].mean():.2f} minutes")
            st.write(f"Minimum loss time: {result_df['Loss Time (minutes)'].min():.2f} minutes")
            st.write(f"Maximum loss time: {result_df['Loss Time (minutes)'].max():.2f} minutes")
        
        # Display processed data
        st.subheader("Processed Data")
        st.dataframe(result_df)
        
        # Create tabs for different visualizations
        tab1, tab2, tab3 = st.tabs(["Top Employees", "Category Analysis", "Time Distribution"])
        
        with tab1:
            # Top employees with highest loss time
            st.subheader("Top 10 Employees with Highest Loss Time")
            
            top10 = result_df.groupby('Name')['Loss Time (minutes)'].sum().sort_values(ascending=False).head(10) / 60
            
            fig, ax = plt.subplots(figsize=(12, 6))
            top10.plot(kind='bar', color=jabil_blue, ax=ax)
            plt.title('Top 10 Employees with Highest Loss Time (Hours)', color=jabil_blue)
            plt.xlabel('Name')
            plt.ylabel('Loss Time (Hours)')
            plt.xticks(rotation=45)
            plt.grid(axis='y', linestyle='--', alpha=0.6)
            plt.tight_layout()
            st.pyplot(fig)
        
        with tab2:
            # Loss time by category
            st.subheader("Loss Time by Category")
            
            category_loss = result_df.groupby('Category')['Loss Time (minutes)'].sum().sort_values(ascending=False) / 60
            
            fig, ax = plt.subplots(figsize=(10, 6))
            bars = category_loss.plot(kind='bar', colormap='Blues', ax=ax)
            
            # Add a green color for one bar to represent the JABIL green
            if len(bars.patches) > 1:
                bars.patches[1].set_facecolor(jabil_green)
                
            plt.title('Total Loss Time by Category (Hours)', color=jabil_blue)
            plt.xlabel('Category')
            plt.ylabel('Loss Time (Hours)')
            plt.xticks(rotation=45)
            plt.grid(axis='y', linestyle='--', alpha=0.6)
            plt.tight_layout()
            st.pyplot(fig)
        
        with tab3:
            # Time distribution
            st.subheader("Time Distribution Analysis")
            
            fig, ax = plt.subplots(figsize=(12, 6))
            result_df['Hour'] = result_df['OUT'].dt.hour
            sns.histplot(data=result_df, x='Hour', bins=24, kde=True, ax=ax, color=jabil_blue)
            plt.title('Distribution of OUT Times by Hour', color=jabil_blue)
            plt.xlabel('Hour of Day')
            plt.ylabel('Count')
            plt.xticks(range(0, 24))
            plt.grid(axis='y', linestyle='--', alpha=0.6)
            st.pyplot(fig)
        
        # Summary metrics in boxes
        st.subheader("Key Metrics")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div style='background-color:{jabil_blue}; padding:10px; border-radius:10px; text-align:center; color:white;'>
                <h3>Total Loss Time</h3>
                <h2>{result_df['Loss Time (minutes)'].sum() / 60:.1f} Hours</h2>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div style='background-color:{jabil_green}; padding:10px; border-radius:10px; text-align:center; color:white;'>
                <h3>Total Employees</h3>
                <h2>{result_df['Name'].nunique()}</h2>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div style='background-color:{jabil_blue}; padding:10px; border-radius:10px; text-align:center; color:white;'>
                <h3>Avg. Loss Time per Entry</h3>
                <h2>{result_df['Loss Time (minutes)'].mean():.1f} Minutes</h2>
            </div>
            """, unsafe_allow_html=True)
        
        # Download processed data
        st.subheader("Download Results")
        col1, col2 = st.columns(2)
        
        with col1:
            csv = result_df.to_csv(index=False)
            st.download_button(
                label="Download as CSV",
                data=csv,
                file_name="loss_time_analysis.csv",
                mime="text/csv",
            )
        
        with col2:
            # Excel download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, sheet_name='Loss Time Analysis', index=False)
            
            st.download_button(
                label="Download as Excel",
                data=buffer.getvalue(),
                file_name="loss_time_analysis.xlsx",
                mime="application/vnd.ms-excel",
            )

else:
    st.info("Please upload an Excel file to begin analysis.")

# Add information about the application
st.sidebar.header("About")
st.sidebar.info("""
This application analyzes access logs to calculate loss time.
It processes entry/exit records to identify when employees leave designated areas.
The application can identify breaks, prayer times, and unauthorized absences.
""")

st.sidebar.header("How to Use")
st.sidebar.info("""
1. Upload your Excel file containing access logs
2. Review the processed data
3. Analyze visualizations across different tabs
4. Download the results for further analysis
""")

# Add JABIL-themed footer
st.markdown("""
<div style='text-align: center; color: #00517C; padding: 20px;'>
    Powered by JABIL Loss Time Analysis System | &copy; 2023
</div>
""", unsafe_allow_html=True)
