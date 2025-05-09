# HIAC Data Processor v1.0

## Overview
HIAC Data Processor is a Streamlit web application designed to simplify the processing of HIAC PDF reports into standardized Excel formats. It automatically extracts particle count data and generates Excel files in the Harmonised DAA template format.

## Features
- Upload multiple PDF files simultaneously
- Real-time processing status with progress indicators
- Automatic data extraction of key parameters:
  - Material information
  - Operator details
  - Sample date and time
  - Sensor serial number
  - Charge information
  - Particle count data for various sizes (1.5, 2, 5, 10, 15, 25 mcm)
- Customizable output filename
- Excel output in Harmonised DAA format
- Bilingual interface (English/German)

## Requirements
```python
streamlit>=1.0.0
PyPDF2
pandas
openpyxl

Data Processing
The application processes the following particle size measurements:

≥ 1.5 mcm
≥ 2.0 mcm
≥ 5.0 mcm
≥ 10.0 mcm
≥ 15.0 mcm
≥ 25.0 mcm

Output Format
The generated Excel file follows the Harmonised DAA template with columns:

Instrument Sample Id
Sample Name
Particle Concentration measurements for each size
Units of Measurement (Particles/mL)
Error Handling
Validates PDF files before processing
Provides clear error messages for invalid or unreadable files
Continues processing remaining files if one fails
Limitations
Only processes HIAC PDF reports in the standard format
Requires proper PDF text extraction capability
Excel output is limited to the Harmonised DAA template format
Support
For questions, suggestions, or issues, please contact Nicholas Michelarakis.

Version History
v1.0 (Current)
Initial release
Multiple file processing
Progress tracking
Harmonised DAA template output
Bilingual interface
License
Internal Sanofi use only

Developed by Nicholas Michelarakis
Last Updated: May 2025

