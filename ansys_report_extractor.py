import os
import re
import pandas as pd
import json
import requests
import numpy as np
import pdfplumber  # Added for PDF processing
from typing import Dict, List, Union, Optional, Any
import logging
import argparse

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger('ANSYS_Report_Extractor')

class ANSYSReportExtractor:
    def __init__(self, api_key: str, excel_template_path: str, reports_dir: str, output_path: str):
        """
        Initialize the ANSYS Report Extractor with required parameters.
        
        Args:
            api_key: OpenRouter API key
            excel_template_path: Path to the Excel template file
            reports_dir: Directory containing ANSYS reports
            output_path: Path where the populated Excel file will be saved
        """
        self.api_key = api_key
        self.excel_template_path = excel_template_path
        self.reports_dir = reports_dir
        self.output_path = output_path
        self.df = None
        
        # Load the Excel template
        self.load_excel_template()
    
    def load_excel_template(self):
        """Load the Excel template file into a pandas DataFrame."""
        try:
            self.df = pd.read_excel(self.excel_template_path)
            logger.info(f"Successfully loaded Excel template from {self.excel_template_path}")
        except Exception as e:
            logger.error(f"Failed to load Excel template: {e}")
            raise
    
    def find_ansys_reports(self) -> List[str]:
        """
        Find all ANSYS report PDF files in the specified directory.
        
        Returns:
            List of paths to ANSYS report PDF files.
        """
        report_files = []
        
        for root, _, files in os.walk(self.reports_dir):
            for file in files:
                # Looking specifically for PDF files that might be ANSYS reports
                if file.endswith('.pdf'):
                    report_files.append(os.path.join(root, file))
        
        logger.info(f"Found {len(report_files)} ANSYS PDF report files")
        return report_files
    
    def read_report_file(self, file_path: str) -> str:
        """
        Read the content of an ANSYS report PDF file.
        
        Args:
            file_path: Path to the ANSYS report PDF file
            
        Returns:
            Content of the file as a string
        """
        try:
            text_content = ""
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    extracted_text = page.extract_text()
                    if extracted_text:
                        text_content += extracted_text + "\n"
            
            logger.info(f"Successfully extracted text from PDF file: {file_path}")
            return text_content
        except Exception as e:
            logger.error(f"Error reading PDF file {file_path}: {e}")
            raise
    
    def extract_parameters(self, report_content: str) -> Dict[str, Any]:
        """
        Extract relevant parameters from the ANSYS report using OpenRouter API.
        
        Args:
            report_content: Content of the ANSYS report
            
        Returns:
            Dictionary of extracted parameters
        """
        # Prepare the prompt for the API
        prompt = f"""
        Task: Extract specific engineering parameters from the following ANSYS report and format them into values only.

        Extract values for:
        Bridge Length (m), Bridge Width (m), Bridge Height (m), Cross-Sectional Diameter (m), Cross-Sectional Area (m²), 
        Cross-Sectional Moment of Inertia (m⁴), Number of Strands, Number of Beams, Angle of Inclination (°), 
        Angle of Declination (°), Young's Modulus (Pa), Poisson's Ratio, Density (kg/m³), Tensile Yield Strength (Pa), 
        Shear Modulus (Pa), Applied Force (N), Mesh Elements, Mesh Density (elements/m³), Max Equivalent Stress (Pa), 
        Max Principal Stress (Pa), Min/Max Deformation (m), Safety Factor, Reaction Forces (N), Strain Energy (J), 
        Work Done (J).

        Formatting Rules:
        - Return the results as a JSON object with parameter names as keys and values as floating-point numbers or strings.
        - Use scientific notation when appropriate (e.g., 1.017e-9).
        - Mark missing data as null.
        - Apply these manual overrides:
            - Number of Strands = 6
            - Number of Beams = 2
            - Angle of Inclination/Declination = 0
        - Format Reaction Forces as a nested array, e.g., [[0,850,0],[0,850,0]].

        ANSYS REPORT (extracted from PDF):
        {report_content[:4000]}  # Truncate to avoid token limits

        Output JSON format only, no explanations or additional text:
        """
        
        # Call the OpenRouter API to extract parameters
        try:
            response = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": "anthropic/claude-3-opus", # Using a capable model for technical extraction
                    "messages": [{"role": "user", "content": prompt}]
                }
            )
            
            # Check for successful response
            response.raise_for_status()
            result = response.json()
            
            # Extract the JSON from the response
            extracted_text = result['choices'][0]['message']['content']
            
            # Sometimes the API might return the JSON with markdown code block formatting
            # We need to clean this up
            json_pattern = r'```(?:json)?\s*(.*?)```'
            json_match = re.search(json_pattern, extracted_text, re.DOTALL)
            
            if json_match:
                extracted_text = json_match.group(1).strip()
            
            # Parse the JSON data
            extracted_params = json.loads(extracted_text)
            logger.info("Successfully extracted parameters from report")
            
            # Apply manual overrides
            extracted_params['Number of Strands'] = 6
            extracted_params['Number of Beams'] = 2
            extracted_params['Angle of Inclination (°)'] = 0
            extracted_params['Angle of Declination (°)'] = 0
            
            return extracted_params
            
        except Exception as e:
            logger.error(f"Error calling OpenRouter API: {e}")
            raise
    
    def calculate_derived_parameters(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        Calculate derived parameters based on extracted values.
        
        Args:
            params: Dictionary of extracted parameters
            
        Returns:
            Dictionary with added derived parameters
        """
        # Create a copy of the input parameters
        result = params.copy()
        
        # Extract required values for calculations
        applied_force = float(params.get('Applied Force (N)', 0))
        max_deformation = 0
        if params.get('Min/Max Deformation (m)'):
            if isinstance(params['Min/Max Deformation (m)'], list):
                max_deformation = abs(float(params['Min/Max Deformation (m)'][1]))
            else:
                # Try to parse string representation like "[0.0, 0.00637]"
                deformation_str = params['Min/Max Deformation (m)']
                match = re.search(r'\[(.*?), (.*?)\]', str(deformation_str))
                if match:
                    max_deformation = abs(float(match.group(2)))
        
        max_equiv_stress = float(params.get('Max Equivalent Stress (Pa)', 0))
        tensile_yield = float(params.get('Tensile Yield Strength (Pa)', 0))
        strain_energy = float(params.get('Strain Energy (J)', 0))
        
        # Calculate derived parameters
        # Work Done (J) = 0.5 * |Applied Force| * Max Deformation
        work_done = 0.5 * abs(applied_force) * max_deformation
        result['Work Done (J)'] = round(work_done, 5)
        
        # Energy Residual (J) = |Strain Energy - Work Done|
        energy_residual = abs(strain_energy - work_done)
        result['Energy Residual (J)'] = round(energy_residual, 5)
        
        # Yield Constraint Residual = ReLU(Max Equivalent Stress - Tensile Yield Strength)
        yield_constraint = max(0, max_equiv_stress - tensile_yield)
        result['Yield Constraint Residual'] = round(yield_constraint, 5)
        
        # Max Failure Load (kg) = Applied Force (N) / 9.8
        max_failure_load = abs(applied_force) / 9.8
        result['Max Failure Load (kg)'] = round(max_failure_load, 5)
        
        return result
    
    def update_excel_with_parameters(self, params: Dict[str, Any], bridge_id: int):
        """
        Update the Excel DataFrame with extracted and calculated parameters.
        
        Args:
            params: Dictionary containing parameters
            bridge_id: ID of the bridge to use in the DataFrame
        """
        # Map of parameter names to column names
        param_to_column = {
            'Bridge Length (m)': 'Bridge Length (m)',
            'Bridge Width (m)': 'Bridge Width (m)',
            'Bridge Height (m)': 'Bridge Height (m)',
            'Cross-Sectional Diameter (m)': 'Cross-Sectional Diameter (m)',
            'Cross-Sectional Area (m²)': 'Cross-Sectional Area (m²)',
            'Cross-Sectional Moment of Inertia (m⁴)': 'Cross-Sectional Moment of Inertia (m⁴)',
            'Number of Strands': 'Number of Strands',
            'Number of Beams': 'Number of Beams',
            'Angle of Inclination (°)': 'Angle of Inclination (°)',
            'Angle of Declination (°)': 'Angle of Declination (°)',
            'Young\'s Modulus (Pa)': 'Young\'s Modulus (Pa)',
            'Poisson\'s Ratio': 'Poisson\'s Ratio',
            'Density (kg/m³)': 'Density (kg/m³)',
            'Tensile Yield Strength (Pa)': 'Tensile Yield Strength (Pa)',
            'Shear Modulus (Pa)': 'Shear Modulus (Pa)',
            'Applied Force (N)': 'Applied Force (N)',
            'Mesh Elements': 'Mesh Elements',
            'Mesh Density (elements/m³)': 'Mesh Density (elements/m³)',
            'Max Equivalent Stress (Pa)': 'Max Equivalent Stress (Pa)',
            'Max Principal Stress (Pa)': 'Max Principal Stress (Pa)',
            'Min/Max Deformation (m)': 'Min/Max Deformation (m)',
            'Safety Factor': 'Safety Factor',
            'Reaction Forces (N)': 'Reaction Forces (N)',
            'Strain Energy (J)': 'Strain Energy (J)',
            'Work Done (J)': 'Work Done (J)',
            'Energy Residual (J)': 'Energy Residual (J)',
            'Yield Constraint Residual': 'Yield Constraint Residual',
            'Max Failure Load (kg)': 'Max Failure Load (N)'  # Note: Despite column name confusion, this is kg
        }
        
        # Create a new row
        new_row = {'Bridge ID': bridge_id}
        
        # Fill in values from parameters
        for param, column in param_to_column.items():
            if param in params and params[param] is not None:
                new_row[column] = params[param]
            else:
                new_row[column] = 'N/A'
        
        # Set default values for categorical columns
        new_row['Bridge Type'] = 'Truss'
        new_row['Symmetry'] = 1
        new_row['Joint Design'] = 'Bonded'
        new_row['Load Type'] = 'Point'
        new_row['Support Type'] = 'Fixed'
        
        # Add the row to the DataFrame
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        logger.info(f"Added bridge with ID {bridge_id} to the DataFrame")
    
    def process_all_reports(self):
        """Process all ANSYS reports and update the Excel file."""
        report_files = self.find_ansys_reports()
        next_bridge_id = self.get_next_bridge_id()
        
        for i, report_file in enumerate(report_files):
            try:
                logger.info(f"Processing report {i+1}/{len(report_files)}: {report_file}")
                
                # Read report content
                report_content = self.read_report_file(report_file)
                
                # Extract parameters
                params = self.extract_parameters(report_content)
                
                # Calculate derived parameters
                params = self.calculate_derived_parameters(params)
                
                # Update Excel with the parameters
                self.update_excel_with_parameters(params, next_bridge_id)
                
                # Increment bridge ID
                next_bridge_id += 1
                
            except Exception as e:
                logger.error(f"Error processing report {report_file}: {e}")
                continue
        
        # Save the updated Excel file
        self.save_excel()
    
    def get_next_bridge_id(self) -> int:
        """
        Get the next available Bridge ID.
        
        Returns:
            Next available Bridge ID
        """
        if self.df.empty:
            return 0
        return self.df['Bridge ID'].max() + 1
    
    def save_excel(self):
        """Save the DataFrame to Excel."""
        try:
            self.df.to_excel(self.output_path, index=False)
            logger.info(f"Successfully saved Excel file to {self.output_path}")
        except Exception as e:
            logger.error(f"Error saving Excel file: {e}")
            raise

def main():
    """Main function to run the script."""
    parser = argparse.ArgumentParser(description='Extract ANSYS report parameters from PDF files and update Excel file.')
    parser.add_argument('--api_key', required=True, help='OpenRouter API key')
    parser.add_argument('--excel_template', required=True, help='Path to Excel template file')
    parser.add_argument('--reports_dir', required=True, help='Directory containing ANSYS report PDFs')
    parser.add_argument('--output', required=True, help='Path to save the output Excel file')
    
    args = parser.parse_args()
    
    extractor = ANSYSReportExtractor(
        api_key=args.api_key,
        excel_template_path=args.excel_template,
        reports_dir=args.reports_dir,
        output_path=args.output
    )
    
    extractor.process_all_reports()
    
    logger.info("Extraction process completed successfully!")

if __name__ == "__main__":
    main()