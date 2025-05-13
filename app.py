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
