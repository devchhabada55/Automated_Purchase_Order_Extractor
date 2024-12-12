import pandas as pd
import os
import PyPDF2
import pytesseract
from PIL import Image
import openpyxl

# Function to extract text from a PDF
def extract_text_from_pdf(file_path):
    text = ""
    try:
        with open(file_path, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            for page in reader.pages:
                text += page.extract_text()
    except Exception as e:
        text = f"Error reading PDF: {e}"
    return text

# Function to extract text from an Excel file
def extract_text_from_excel(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        return df.to_string(index=False)
    except Exception as e:
        return f"Error reading Excel: {e}"

# Function to extract text from an image
def extract_text_from_image(file_path):
    try:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img)
        return text
    except Exception as e:
        return f"Error reading image: {e}"

# Function to extract text from a .py file
def extract_text_from_py(file_path):
    try:
        with open(file_path, 'r') as file:
            return file.read()
    except Exception as e:
        return f"Error reading Python file: {e}"

# Function to handle the attachment extraction based on file type
def extract_attachment_content(file_path):
    if file_path.endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    elif file_path.endswith(('.xls', '.xlsx')):
        return extract_text_from_excel(file_path)
    elif file_path.endswith(('.png', '.jpg', '.jpeg')):
        return extract_text_from_image(file_path)
    elif file_path.endswith('.py'):
        return extract_text_from_py(file_path)
    else:
        return f"Unsupported file type: {file_path}"

# Main processing function
def process_excel(input_excel_path, output_excel_path):
    # Read the input Excel file
    df = pd.read_excel(input_excel_path)
    
    # Initialize a list to store processed rows
    extracted_data = []
    
    # Iterate over rows in the DataFrame
    for _, row in df.iterrows():
        # Get the attachment link and ensure it's a string
        attachment_link = str(row.get("Attachment Link", "")).strip()
        
        # Convert backslashes to forward slashes for consistent path handling
        attachment_link = attachment_link.replace("\\", "/")
        
        # Check if the attachment link is valid and exists
        if attachment_link and os.path.exists(attachment_link):
            print(f"Processing file: {attachment_link}")  # Debugging line
            extracted_content = extract_attachment_content(attachment_link)
        else:
            extracted_content = "File not found or invalid attachment link."
        
        # Add the extracted content to the row
        row["extracted information from attachment"] = extracted_content
        extracted_data.append(row)
    
    # Save the updated data to a new Excel file
    output_df = pd.DataFrame(extracted_data)
    output_df.to_excel(output_excel_path, index=False, engine='openpyxl')
    print(f"Processed data saved to {output_excel_path}")

# Example usage
input_excel = "emails_data-testcase.xlsx"  # Replace with the path to your input Excel file
output_excel = "output-test-case.xlsx"  # Replace with the desired path for the output Excel file
process_excel(input_excel, output_excel)
