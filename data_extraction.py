import os
import pandas as pd
import pytesseract
from PIL import Image
from PyPDF2 import PdfReader
import pdfplumber
import re
from docx import Document
import traceback
import logging
import requests
import json

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,  # Set to DEBUG for more detailed output
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('po_extraction-json.log'),
        logging.StreamHandler()
    ]
)

class POExtractor:
    def __init__(self, input_excel, attachments_folder, output_json):
        """
        Initialize POExtractor with input excel, attachments folder, and output JSON file path.
        
        :param input_excel: Path to the input Excel file containing email data
        :param attachments_folder: Folder containing attachment files
        :param output_json: Path to save the extracted PO details JSON
        """
        self.input_excel = input_excel
        self.attachments_folder = attachments_folder
        self.output_json = output_json

        # Define the columns for the output
        self.output_columns = [
            "Email Sender", 
            "Customer PO Number", 
            "Item Name", 
            "Quantity", 
            "Rate per unit", 
            "Unit of measurement", 
            "Item wise Delivery Dates", 
            "Customer Name", 
            "Customer details", 
            "Applicable Taxes", 
            "Terms of Payment", 
            "Discount", 
            "Other remarks/instructions"
        ]

    def normalize_path(self, file_path):
        """
        Normalize the file path to ensure correct file access.
        
        :param file_path: Original file path
        :return: Normalized full path to the file
        """
        try:
            file_path = str(file_path).strip()
            # Remove 'attachments/' or 'attachments\' prefix if present
            if file_path.lower().startswith(('attachments/', 'attachments\\')):
                file_path = file_path.split('attachments', 1)[1].lstrip('/').lstrip('\\')
            
            # Join with attachments folder to get full path
            full_path = os.path.normpath(os.path.join(self.attachments_folder, file_path))
            logging.info(f"Normalized path: {full_path}")
            return full_path
        except Exception as e:
            logging.error(f"Path normalization error: {e}")
            return None

    def extract_text_from_pdf(self, file_path):
        """
        Extract text from a PDF file using multiple methods.
        
        :param file_path: Path to the PDF file
        :return: Extracted text from the PDF
        """
        texts = []
        try:
            # Try extracting with pdfplumber first
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        texts.append(page_text)
            
            # If pdfplumber fails, use PyPDF2
            if not texts:
                with open(file_path, 'rb') as file:
                    pdf_reader = PdfReader(file)
                    for page in pdf_reader.pages:
                        page_text = page.extract_text()
                        if page_text:
                            texts.append(page_text)
            
            full_text = "\n".join(texts)
            logging.info(f"PDF text extraction successful: {len(full_text)} characters")
            return full_text
        except Exception as e:
            logging.error(f"PDF text extraction error: {e}")
            logging.error(traceback.format_exc())
            return ""

    def extract_text_from_image(self, file_path):
        """
        Extract text from an image file using Tesseract OCR.
        
        :param file_path: Path to the image file
        :return: Extracted text from the image
        """
        try:
            img = Image.open(file_path)
            text = pytesseract.image_to_string(img, config='--psm 6')
            logging.info(f"Image OCR successful: {len(text)} characters")
            return text
        except Exception as e:
            logging.error(f"Image text extraction error: {e}")
            return ""

    def extract_text_from_excel(self, file_path):
        """
        Extract text from an Excel file.
        
        :param file_path: Path to the Excel file
        :return: Text representation of the Excel data
        """
        try:
            df = pd.read_excel(file_path)
            text = df.to_string()
            logging.info(f"Excel text extraction successful: {len(text)} characters")
            return text
        except Exception as e:
            logging.error(f"Excel text extraction error: {e}")
            return ""

    def extract_text_from_word(self, file_path):
        """
        Extract text from a Word document.
        
        :param file_path: Path to the Word document
        :return: Extracted text from the document
        """
        try:
            doc = Document(file_path)
            text = "\n".join([para.text for para in doc.paragraphs])
            logging.info(f"Word document text extraction successful: {len(text)} characters")
            return text
        except Exception as e:
            logging.error(f"Word text extraction error: {e}")
            return ""

    def extract_po_details(self, text):
        """
        Extract PO details from text using Ollama API with Llama3.1 model.
        
        :param text: Input text to extract PO details from
        :return: Dictionary of extracted PO details
        """
        try:
            ollama_url = "http://localhost:11434/api/generate"
            headers = {"Content-Type": "application/json"}
            prompt = (
                "Extract the following details as a valid JSON object with these keys only:\n"
                "Customer PO Number, Item Name, Quantity, Rate per unit, Unit of measurement, "
                "Item wise Delivery Dates, Customer Name, Customer details, Applicable Taxes, "
                "Terms of Payment, Discount, Other remarks/instructions\n\n"
                "If a detail is not found, use 'N/A' as the value. Ensure the JSON is properly formatted.\n\n"
                f"Text:\n{text}"
            )
            payload = {
                "model": "llama3.1:latest", 
                "prompt": prompt,
                "stream": False,
                "format": "json"
            }
            response = requests.post(ollama_url, headers=headers, data=json.dumps(payload))
            
            if response.status_code != 200:
                logging.error(f"API request failed with status {response.status_code}")
                logging.error(f"Response content: {response.text}")
                return {key: "N/A" for key in self.output_columns[1:]}

            response_text = response.json().get('response', '')
            logging.debug(f"Raw API response: {response_text}")

            try:
                # Attempt to parse the JSON response
                po_details = json.loads(response_text)
            except json.JSONDecodeError:
                # If direct parsing fails, try extracting JSON using regex
                json_match = re.search(r'\{.*\}', response_text, re.DOTALL | re.MULTILINE)
                if json_match:
                    try:
                        po_details = json.loads(json_match.group(0))
                    except json.JSONDecodeError:
                        logging.error("Failed to parse extracted JSON")
                        return {key: "N/A" for key in self.output_columns[1:]}
                else:
                    logging.error("No JSON found in the response")
                    return {key: "N/A" for key in self.output_columns[1:]}

            # Helper function to flatten or stringify complex values
            def process_value(value):
                if value is None:
                    return "N/A"
                if isinstance(value, list):
                    # If it's a list of strings, join them
                    if all(isinstance(v, (str, int, float)) for v in value):
                        return ', '.join(str(v) for v in value)
                    # If it's a list of dicts or complex objects, convert to string
                    return str(value)
                if isinstance(value, dict):
                    # For dictionaries, convert to a readable string
                    return ', '.join(f"{k}: {v}" for k, v in value.items())
                return str(value).strip() or "N/A"

            # Prepare the result dictionary with processed values
            cleaned_details = {}
            for key in self.output_columns[1:]:
                cleaned_details[key] = process_value(po_details.get(key, "N/A"))

            logging.info(f"PO details extracted successfully: {cleaned_details}")
            return cleaned_details

        except Exception as e:
            logging.error(f"Error in extracting PO details with Llama: {e}")
            logging.error(traceback.format_exc())
            return {key: "N/A" for key in self.output_columns[1:]}

    def extract_attachment_content(self, attachment_path):
        """
        Extract content from an attachment based on its file type.
        
        :param attachment_path: Path to the attachment file
        :return: Extracted PO details
        """
        full_path = self.normalize_path(attachment_path)
        if not full_path or not os.path.exists(full_path):
            logging.warning(f"Attachment not found: {attachment_path}")
            return {key: "N/A" for key in self.output_columns[1:]}

        try:
            # Determine file type and extract text accordingly
            if full_path.lower().endswith('.pdf'):
                text = self.extract_text_from_pdf(full_path)
            elif full_path.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp')):
                text = self.extract_text_from_image(full_path)
            elif full_path.lower().endswith(('.xls', '.xlsx')):
                text = self.extract_text_from_excel(full_path)
            elif full_path.lower().endswith('.docx'):
                text = self.extract_text_from_word(full_path)
            else:
                logging.error(f"Unsupported file type: {full_path}")
                return {key: "N/A" for key in self.output_columns[1:]}

            if text:
                po_details = self.extract_po_details(text)
                logging.debug(f"Extracted PO details: {po_details}")
                return po_details
            else:
                logging.warning(f"No text extracted from {full_path}")
                return {key: "N/A" for key in self.output_columns[1:]}
        except Exception as e:
            logging.error(f"Error extracting content from {full_path}: {e}")
            logging.error(traceback.format_exc())
            return {key: "N/A" for key in self.output_columns[1:]}

    def process_emails(self):
        """
        Process emails from the input Excel file and extract PO details.
        Save results to a JSON file.
        """
        try:
            # Read the input Excel file
            df = pd.read_excel(self.input_excel)
            
            # Filter rows classified as PO
            po_df = df[df['Classification'] == 'PO']
            
            # List to store results
            results = []

            # Process each PO row
            for index, row in po_df.iterrows():
                logging.info(f"Processing row {index}")
                
                # Initialize result row with default 'N/A' values
                result_row = {col: "N/A" for col in self.output_columns}
                
                # Set email sender
                result_row["Email Sender"] = row.get("Email Sender", "N/A")
                
                # Get attachment link and extract content
                attachment_link = str(row.get("Attachment Link", "")).strip()
                extracted_content = self.extract_attachment_content(attachment_link)
                
                # Update result row with extracted content
                for key in self.output_columns[1:]:
                    result_row[key] = extracted_content.get(key, "N/A")
                
                logging.debug(f"Result row: {result_row}")
                results.append(result_row)

            # Save results to JSON
            with open(self.output_json, 'w', encoding='utf-8') as json_file:
                json.dump(results, json_file, indent=4, ensure_ascii=False)
            
            logging.info(f"Process completed. Results saved to {self.output_json}")
        
        except Exception as e:
            logging.error(f"Error processing emails: {e}")
            logging.error(traceback.format_exc())

'''def main():
    """
    Main function to run the PO extraction process.
    """
    try:
        # Initialize POExtractor with input and output paths
        po_extractor = POExtractor(
            input_excel='classified_emails-test-case.xlsx', 
            attachments_folder='attachments', 
            output_json='po_extracted-testcase.json'
        )
        
        # Process emails and extract PO details
        po_extractor.process_emails()
    
    except Exception as e:
        logging.error(f"Unhandled error in main: {e}")
        logging.error(traceback.format_exc())

if __name__ == "__main__":
    main()'''