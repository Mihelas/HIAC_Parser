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
    This "app" allows you to upload one or multiple HIAC PDF files, extract the necessary information for iLab, 
    and save it into an Excel iLab upload template. New uploads will be added to the existing Excel file (you'll need to download it again). 
    To start fresh, reload the app or delete the uploaded PDFs.
    The uploaded PDFs files will appear on the excel file in the order they were uploaded.
    Please tick the check-box if you also need the Harmonised DAA template and it will create a second excel file.  
    For any questions or suggestions, contact Nicholas Michelarakis. :)
    
    Mit dieser "App" können Sie eine oder mehrere HIAC-PDF-Dateien hochladen, um die Informationen in einem für iLab passenden Template verfügbar zu machen. 
    Die hochgeladenen PDF-Dateien erscheinen in der Excel-Datei in der Reihenfolge, in der sie hochgeladen wurden. 
    Bitte aktivieren Sie das Kontrollkästchen, wenn Sie das Format für das harmonisierte Upload-Template benötigen, und es wird eine zweite Excel-Datei erstellt. 
    Neue Uploads werden der bestehenden Excel-Datei hinzugefügt (Sie müssen sie erneut herunterladen). 
    Um neu zu beginnen, laden Sie die Anwendung neu oder löschen Sie die hochgeladenen PDFs.  Bei Fragen oder Anregungen wenden Sie sich bitte an Nicholas Michelarakis :)
    """)
    
    uploaded_files = st.file_uploader("Choose PDF files", type="pdf", accept_multiple_files=True)
    
    if uploaded_files:
        st.session_state['uploaded_files'] = []
        st.session_state['data_dicts'] = []
        data_dicts = []
        
        for uploaded_file in uploaded_files:
            text = extract_text_from_pdf(uploaded_file)
            if text:
                page_data_dicts = extract_info_from_pages(text)
                if page_data_dicts:
                    data_dicts.extend(page_data_dicts)
        
        if data_dicts:
            st.success("PDFs were parsed successfully.")
            
            raw_data_checkbox = st.checkbox("Check this for the DAA template")
            excel_data = create_excel(data_dicts)
            if raw_data_checkbox:
                raw_data_excel = create_raw_data_excel(data_dicts)
            
            if "file_name" not in st.session_state:
                st.session_state.file_name = "extracted_info"
            
            def update_file_name():
                st.session_state.file_name = st.session_state.file_name_input
            
            file_name = st.text_input(
                "Enter the desired file name (without extension):", 
                st.session_state.file_name, 
                key="file_name_input", 
                on_change=update_file_name
            )
            
            if excel_data:
                if st.download_button(
                    label="Download Excel file",
                    data=excel_data,
                    file_name=f"{st.session_state.file_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=update_file_name
                ):
                    st.success(f"Excel file '{st.session_state.file_name}.xlsx' was created successfully.")
            
            if raw_data_checkbox and raw_data_excel:
                if st.download_button(
                    label="Download Harmonised DAA format Excel file",
                    data=raw_data_excel,
                    file_name="Harmonised_DAA_format.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ):
                    st.success("DAA formatted Excel file 'Harmonised_DAA_format.xlsx' was created successfully.")

if __name__ == "__main__":
    main()