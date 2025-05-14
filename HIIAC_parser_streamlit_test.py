import streamlit as st
import PyPDF2
import re
import openpyxl
import pandas as pd
from io import BytesIO

def extract_text_from_pdf(uploaded_file):
    try:
        text = ""
        reader = PyPDF2.PdfReader(uploaded_file)
        number_of_pages = len(reader.pages)
        # Extract text from each page
        for page_number in range(number_of_pages):
            page = reader.pages[page_number]
            text += page.extract_text() + "\n"
        return text
    except Exception as e:
        st.error(f"An error occurred while reading the PDF: {e}")
        return None

def extract_info(text):
    try:
        # Debugging output to see the text being processed
        #st.write("Text being processed:", text)

        # Adjusted regex patterns to handle both formats
        material_match = re.search(r'Material\s*:\s*(\S+)|Material\s*(\S+)', text)
        operator_name_match = re.search(r'Operator Name\s*:\s*(.*?)(?=\s|,|/)', text)
        sample_date_time_match = re.search(r'Sample Date\s*:\s*(.*?)(?=\s|$)', text)
        sensor_serial_number_match = re.search(r'Sensor Serial Number\s*:\s*(\S+)|SensorSerialNumber\s*:\s*(\S+)', text)
        charge_match = re.search(r'Charge\s*:\s*(.*?)(?=(?:Sensor\s*Serial\s*Number|SensorSerialNumber))', text, re.DOTALL | re.IGNORECASE)

        # Check if matches are found and extract the groups
        material = material_match.group(1) if material_match else "N/A"
        operator_name = operator_name_match.group(1) if operator_name_match else "N/A"
        sample_date_time = sample_date_time_match.group(1) if sample_date_time_match else "N/A"
        sensor_serial_number = sensor_serial_number_match.group(1) if sensor_serial_number_match else "N/A"
        charge = charge_match.group(1).strip() if charge_match else "N/A"

        average_counts = {}
        # Extract the section starting with "Average" and below
        average_section = re.search(r'Average.*', text, re.DOTALL)
        if average_section:
            average_text = average_section.group(0)
            for size in [1.5, 2, 5, 10, 15, 25]:
                if size == 1.5:
                    match = re.search(r'1[,.]500\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)', average_text)
                else:
                    match = re.search(rf'{size}[,.]000\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)', average_text)
                if match:
                    value = match.group(3).replace(',', '.')
                    average_counts[size] = value

        data_dict = {
            "Material": material,
            "Operator Name": operator_name,
            "Sample Date and Time": sample_date_time,
            "Sensor Serial Number": sensor_serial_number,
            "Charge": charge,
            "Average Cumulative Counts": average_counts
        }
        
        # Debugging output
        #st.write("Extracted Info:", data_dict)
        
        return data_dict
    except AttributeError as e:
        st.error(f"An error occurred while extracting information: {e}")
        return None

def extract_info_from_pages(text):
    try:
        pages = text.split('\n\n')
        data_dicts = []
        for page in pages:
            data_dict = extract_info(page)
            if data_dict:
                data_dicts.append(data_dict)
        return data_dicts
    except Exception as e:
        st.error(f"An error occurred while extracting information from pages: {e}")
        return None

def create_excel(data_dicts):
    try:
        df = pd.DataFrame(columns=[
            "Instrument Sample Id", "Sample Name", "Particle Concentration (>= 2 mcm)", 
            "Particle Concentration (>= 2 mcm) UoM", "Particle Concentration (>= 10 mcm)", 
            "Particle Concentration (>= 10 mcm) UoM", "Particle Concentration (>=25 mcm)", 
            "Particle Concentration (>=25 mcm) UoM", "Comment"
        ])
        for data_dict in data_dicts:
            concentration_2 = sum([float(data_dict["Average Cumulative Counts"].get(size, '0').replace(',', '.')) for size in [2, 5]])
            concentration_10 = sum([float(data_dict["Average Cumulative Counts"].get(size, '0').replace(',', '.')) for size in [10, 15]])
            concentration_25 = sum([float(data_dict["Average Cumulative Counts"].get(size, '0').replace(',', '.')) for size in [25]])
            
            df.loc[len(df)] = [
                data_dict["Material"], data_dict["Charge"], concentration_2, "Particles/ml", 
                concentration_10, "Particles/ml", concentration_25, "Particles/ml", 
                f"{data_dict['Operator Name']}"
            ]
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        return output.getvalue()
    except Exception as e:
        st.error(f"An error occurred while creating the Excel file: {e}")
        return None

def create_raw_data_excel(data_dicts):
    try:
        df = pd.DataFrame(columns=[
            "Instrument Sample Id", "Sample Name", 
            "Particle Concentration (>= 1.5 mcm)", "Particle Concentration (>= 1.5 mcm) UoM",
            "Particle Concentration (>= 2 mcm)", "Particle Concentration (>= 2 mcm) UoM",
            "Particle Concentration (>= 5 mcm)", "Particle Concentration (>= 5 mcm) UoM",
            "Particle Concentration (>= 10 mcm)", "Particle Concentration (>= 10 mcm) UoM",
            "Particle Concentration (>= 15 mcm)", "Particle Concentration (>= 15 mcm) UoM",
            "Particle Concentration (>= 25 mcm)", "Particle Concentration (>= 25 mcm) UoM"
        ])

        for data_dict in data_dicts:
            concentration_1_5 = float(data_dict["Average Cumulative Counts"].get(1.5, '0').replace(',', '.'))
            concentration_2 = float(data_dict["Average Cumulative Counts"].get(2, '0').replace(',', '.'))
            concentration_5 = float(data_dict["Average Cumulative Counts"].get(5, '0').replace(',', '.'))
            concentration_10 = float(data_dict["Average Cumulative Counts"].get(10, '0').replace(',', '.'))
            concentration_15 = float(data_dict["Average Cumulative Counts"].get(15, '0').replace(',', '.'))
            concentration_25 = float(data_dict["Average Cumulative Counts"].get(25, '0').replace(',', '.'))

            df.loc[len(df)] = [
                data_dict["Material"], data_dict["Charge"], 
                concentration_1_5, "Particles/mL", 
                concentration_2, "Particles/mL", 
                concentration_5, "Particles/mL", 
                concentration_10, "Particles/mL", 
                concentration_15, "Particles/mL", 
                concentration_25, "Particles/mL"
            ]

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        return output.getvalue()
    except Exception as e:
        st.error(f"An error occurred while creating the Excel file: {e}")
        return None

def main():
    st.title("PDF Reader and Info Extractor App")
    st.markdown("""
    This "app" allows you to upload one or multiple HIAC PDF files, extract the necessary information, 
    and save it into an Excel file using the Harmonised DAA template format. 
    The uploaded PDFs files will appear on the excel file in the order they were uploaded.
    New uploads will be added to the existing Excel file (you'll need to download it again). 
    To start fresh, reload the app or delete the uploaded PDFs.
    For any questions or suggestions, contact Nicholas Michelarakis. :)
    
    Mit dieser "App" können Sie eine oder mehrere HIAC-PDF-Dateien hochladen, um die Informationen in einem harmonisierten DAA-Template verfügbar zu machen. 
    Die hochgeladenen PDF-Dateien erscheinen in der Excel-Datei in der Reihenfolge, in der sie hochgeladen wurden. 
    Neue Uploads werden der bestehenden Excel-Datei hinzugefügt (Sie müssen sie erneut herunterladen). 
    Um neu zu beginnen, laden Sie die Anwendung neu oder löschen Sie die hochgeladenen PDFs.
    Bei Fragen oder Anregungen wenden Sie sich bitte an Nicholas Michelarakis :)
    """)
    
    # Initialize session state variables if they don't exist
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = []
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'last_upload_id' not in st.session_state:
        st.session_state.last_upload_id = None
    
    uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)
    
    if uploaded_files:
        # Generate a unique ID for this batch of uploads
        current_upload_id = hash(tuple(f.name for f in uploaded_files))
        
        # Reset processed files if new files are uploaded
        if st.session_state.last_upload_id != current_upload_id:
            st.session_state.processed_files = []
            st.session_state.last_upload_id = current_upload_id
        
        # Show file summary
        st.write(f"📁 Number of files uploaded: **{len(uploaded_files)}**")
        
        # Create an expandable section to show uploaded files
        with st.expander("View uploaded files"):
            for file in uploaded_files:
                st.write(f"- {file.name}")
        
        data_dicts = []
        
        # Only process files if they haven't been processed yet
        if len(st.session_state.processed_files) != len(uploaded_files):
            # Create a progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Process each file with progress updates
            for idx, uploaded_file in enumerate(uploaded_files):
                try:
                    # Skip if already processed
                    if uploaded_file.name in st.session_state.processed_files:
                        continue
                        
                    status_text.text(f"Processing {uploaded_file.name}...")
                    text = extract_text_from_pdf(uploaded_file)
                    
                    if text:
                        page_data_dicts = extract_info_from_pages(text)
                        if page_data_dicts:
                            data_dicts.extend(page_data_dicts)
                            st.session_state.processed_files.append(uploaded_file.name)
                        else:
                            st.warning(f"⚠️ No data could be extracted from {uploaded_file.name}")
                    
                    # Update progress
                    progress = (idx + 1) / len(uploaded_files)
                    progress_bar.progress(progress)
                    
                except Exception as e:
                    st.error(f"❌ Error processing {uploaded_file.name}: {str(e)}")
                    continue
            
            # Clear progress bar and status when complete
            progress_bar.empty()
            status_text.empty()
        else:
            # If all files are already processed, just load the data
            for uploaded_file in uploaded_files:
                text = extract_text_from_pdf(uploaded_file)
                if text:
                    page_data_dicts = extract_info_from_pages(text)
                    if page_data_dicts:
                        data_dicts.extend(page_data_dicts)
        
        if data_dicts:
            st.success(f"✅ Successfully processed {len(st.session_state.processed_files)} files")
            
            excel_data = create_raw_data_excel(data_dicts)
            
            # File name input with validation
            if "file_name" not in st.session_state:
                st.session_state.file_name = "HIAC_DAA_format"
            
            def update_file_name():
                # Remove invalid characters from filename
                valid_name = re.sub(r'[<>:"/\\|?*]', '', st.session_state.file_name_input)
                st.session_state.file_name = valid_name
            
            file_name = st.text_input(
                "Enter the desired file name (without extension):", 
                st.session_state.file_name, 
                key="file_name_input", 
                on_change=update_file_name,
                help="The file name will be automatically sanitized to remove invalid characters"
            )
            
            if excel_data:
                if st.download_button(
                    label="📥 Download Excel file",
                    data=excel_data,
                    file_name=f"{st.session_state.file_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=update_file_name
                ):
                    st.balloons()
                    st.success(f"✅ Excel file '{st.session_state.file_name}.xlsx' was created successfully.")
        
        else:
            st.error("❌ No data could be extracted from any of the uploaded files. Please check if the files are in the correct format.")
    
    # Show helpful message when no files are uploaded
    else:
        st.info("👆 Please upload one or more PDF files to begin")
        # Clear processed files when no files are uploaded
        st.session_state.processed_files = []
        st.session_state.last_upload_id = None

if __name__ == "__main__":
    main()