import csv
from docx import Document
from docx.shared import Pt, RGBColor
import os
import re

# Function to read parameters from the CSV file
def read_parameters_from_csv(csv_file):
    params = []
    with open(csv_file, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            params.append(row)
    return params

# Function to map color names to RGBColor objects
def get_color(color_name):
    colors = {
        'red': RGBColor(255, 0, 0),
        'blue': RGBColor(0, 0, 255),
        'green': RGBColor(0, 128, 0),
        'orange': RGBColor(255, 165, 0),
        # Add more colors as needed
    }
    return colors.get(color_name.lower(), RGBColor(0, 0, 0))  # Default to black if color not found

# Function to replace placeholders and apply formatting in a .docx file
def create_docx_with_replacements(template_docx_path, params, output_folder):
    for param in params:
        # Load the template document
        doc = Document(template_docx_path)
        
        # Identify and replace placeholders dynamically
        for paragraph in doc.paragraphs:
            # Iterate over all keys that match the pattern 'xxx#_text'
            for key in param:
                match = re.match(r'xxx(\d+)_text', key)
                if match:
                    placeholder_number = match.group(1)
                    placeholder = f"xxx{placeholder_number}"
                    text_key = key
                    font_size_key = f"xxx{placeholder_number}_font_size"
                    color_key = f"xxx{placeholder_number}_color"

                    # Skip this placeholder if any of the parameters are None
                    if not param[text_key] or param[text_key].lower() == 'none':
                        continue

                    # Only attempt replacement if the placeholder is found in the text
                    if placeholder in paragraph.text:
                        replace_text_with_formatting(
                            paragraph,
                            placeholder,
                            param[text_key],
                            int(param.get(font_size_key, 12)),  # Default font size is 12 if not specified
                            get_color(param.get(color_key, 'black'))  # Default color is black if not specified
                        )

        # Save the new document in the specified output folder
        output_filename = os.path.join(output_folder, param["output_file"])
        doc.save(output_filename)
        print(f"{output_filename} created successfully.")

# Function to replace text and apply formatting
def replace_text_with_formatting(paragraph, placeholder, new_text, font_size, font_color):
    # Replace the placeholder while preserving formatting
    if placeholder in paragraph.text:
        inline = paragraph.runs
        for run in inline:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, new_text)
                run.font.size = Pt(font_size)
                run.font.color.rgb = font_color

# Read parameters from CSV
params = read_parameters_from_csv('parameters.csv')

# Define the template path and output folder
template_path = os.path.join('docs', 'template.docx')
output_folder = 'docs/generated'

# Use the function with your template DOCX and parameters
create_docx_with_replacements(template_path, params, output_folder)
