import os
import pandas as pd
from ansys_report_extractor import ANSYSReportExtractor
import tempfile
import shutil
from reportlab.pdfgen import canvas
from io import BytesIO

# Create a sample ANSYS report as PDF
def create_sample_pdf_report(filepath):
    """Create a sample ANSYS report PDF file."""
    sample_content = """
    ANSYS Mechanical Analysis Report
    ===============================
    
    Model Information:
    -----------------
    Bridge Length: 1.016 m
    Bridge Width: 0.254 m
    Bridge Height: 0.254 m
    Cross-Sectional Diameter: 1.20E-02 m
    Cross-Sectional Area: 1.13E-04 m²
    Cross-Sectional Moment of Inertia: 1.02E-09 m⁴
    
    Material Properties:
    ------------------
    Young's Modulus: 2.75E+09 Pa
    Poisson's Ratio: 0.26
    Density: 1470 kg/m³
    Tensile Yield Strength: 9.50E+06 Pa
    Shear Modulus: 1.09E+09 Pa
    
    Loading Conditions:
    -----------------
    Applied Force: -1703 N at position [0.508, 0.0, 0.127]
    
    Mesh Information:
    ---------------
    Total Elements: 164
    Mesh Density: 7.30E+04 elements/m³
    
    Analysis Results:
    ---------------
    Maximum Equivalent Stress: 6.25E+06 Pa
    Maximum Principal Stress: 5.97E+06 Pa
    Deformation Range: [0.0, 0.00637] m
    Safety Factor: 1.519
    Reaction Forces: [[0.0, 850.0, 0.0], [0.0, 850.0, 0.0]]
    Strain Energy: 4.82 J
    """
    
    pdf = canvas.Canvas(filepath)
    y_position = 800
    line_height = 14
    
    # Split content by lines and write each line to the PDF
    for line in sample_content.split('\n'):
        pdf.drawString(40, y_position, line)
        y_position -= line_height
    
    pdf.save()
    
    return filepath

# Create a test Excel template
def create_test_excel():
    # Create DataFrame with the same structure as your template
    columns = [
        'Bridge ID', 'Bridge Length (m)', 'Bridge Width (m)', 'Bridge Height (m)',
        'Cross-Sectional Diameter (m)', 'Cross-Sectional Area (m²)',
        'Cross-Sectional Moment of Inertia (m⁴)', 'Number of Strands', 'Number of Beams',
        'Angle of Inclination (°)', 'Angle of Declination (°)', 'Bridge Type',
        'Symmetry', 'Joint Design', 'Young\'s Modulus (Pa)', 'Poisson\'s Ratio',
        'Density (kg/m³)', 'Tensile Yield Strength (Pa)', 'Shear Modulus (Pa)',
        'Applied Force (N)', 'Load Location (x, y, z)', 'Load Type', 'Support Type',
        'Support Locations', 'Mesh Elements', 'Mesh Density (elements/m³)',
        'Element Type', 'Max Equivalent Stress (Pa)', 'Max Principal Stress (Pa)',
        'Min/Max Deformation (m)', 'Safety Factor', 'Reaction Forces (N)',
        'Strain Energy (J)', 'Work Done (J)', 'Stress Equilibrium Residual',
        'Constitutive Law Residual', 'Energy Residual (J)', 'Yield Constraint Residual',
        'Max Failure Load (N)'
    ]
    
    # Create sample data (similar to your example)
    data = [
        [0, 1.016, 0.254, 0.254, 1.20E-02, 1.13E-04, 1.02E-09, 6, 2, 0, 0, 
         'Truss', 1, 'Bonded', 2.75E+09, 0.26, 1470, 9.50E+06, 1.09E+09, -1703,
         '[0.508, 0.0, 0.127]', 'Point', 'Fixed', '[[0.0,0.0,0.0],[1.016,0.0,0.0]]',
         164, 7.30E+04, 'Beam', 6.25E+06, 5.97E+06, '[0.0, 0.00637]', 1.519,
         '[[0.0, 850.0, 0.0], [0.0, 850.0, 0.0]]', 4.82, 5.43E+00, 'N/A', 'N/A',
         0.61, 0, 173.8]
    ]
    
    df = pd.DataFrame(data, columns=columns)
    
    # Create a temporary Excel file
    temp_file = os.path.join(tempfile.gettempdir(), 'test_template.xlsx')
    df.to_excel(temp_file, index=False)
    
    return temp_file

# Test function
def test_extractor(api_key):
    # Create a temporary directory for test reports
    test_dir = os.path.join(tempfile.gettempdir(), 'test_ansys_reports')
    if os.path.exists(test_dir):
        shutil.rmtree(test_dir)
    os.makedirs(test_dir)
    
    # Create sample PDF report files
    report_paths = []
    for i in range(2):
        report_path = os.path.join(test_dir, f'sample_report_{i}.pdf')
        create_sample_pdf_report(report_path)
        report_paths.append(report_path)
    
    # Create test Excel template
    excel_template = create_test_excel()
    
    # Create output path
    output_path = os.path.join(tempfile.gettempdir(), 'test_output.xlsx')
    
    # Create extractor and process reports
    extractor = ANSYSReportExtractor(
        api_key=api_key,
        excel_template_path=excel_template,
        reports_dir=test_dir,
        output_path=output_path
    )
    
    # Mock the extract_parameters method to avoid actual API calls
    def mock_extract(self, report_content):
        return {
            'Bridge Length (m)': 1.016,
            'Bridge Width (m)': 0.254,
            'Bridge Height (m)': 0.254,
            'Cross-Sectional Diameter (m)': 1.20E-02,
            'Cross-Sectional Area (m²)': 1.13E-04,
            'Cross-Sectional Moment of Inertia (m⁴)': 1.02E-09,
            'Number of Strands': 6,
            'Number of Beams': 2,
            'Angle of Inclination (°)': 0,
            'Angle of Declination (°)': 0,
            'Young\'s Modulus (Pa)': 2.75E+09,
            'Poisson\'s Ratio': 0.26,
            'Density (kg/m³)': 1470,
            'Tensile Yield Strength (Pa)': 9.50E+06,
            'Shear Modulus (Pa)': 1.09E+09,
            'Applied Force (N)': -1703,
            'Mesh Elements': 164,
            'Mesh Density (elements/m³)': 7.30E+04,
            'Max Equivalent Stress (Pa)': 6.25E+06,
            'Max Principal Stress (Pa)': 5.97E+06,
            'Min/Max Deformation (m)': [0.0, 0.00637],
            'Safety Factor': 1.519,
            'Reaction Forces (N)': [[0.0, 850.0, 0.0], [0.0, 850.0, 0.0]],
            'Strain Energy (J)': 4.82
        }
    
    # Replace the method with our mock
    import types
    extractor.extract_parameters = types.MethodType(mock_extract, extractor)
    
    # Process the reports
    extractor.process_all_reports()
    
    print(f"Test completed successfully! Output saved to {output_path}")
    
    # Clean up
    shutil.rmtree(test_dir)
    os.remove(excel_template)
    
    return output_path

if __name__ == "__main__":
    # Replace with your actual OpenRouter API key
    api_key = "your_openrouter_api_key_here"
    output_file = test_extractor(api_key)
    print(f"Check the output Excel file: {output_file}")