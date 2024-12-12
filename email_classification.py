import pandas as pd
import os
import json
from openai import OpenAI

def classify_emails_in_file(input_file, output_file, api_key):
    """
    Classifies emails in an Excel file as "PO" or "Not PO" using OpenAI's API.

    Parameters:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to save the classified output Excel file.
        api_key (str): API key for OpenAI.

    Returns:
        None
    """
    # Initialize OpenAI client
    client = OpenAI(api_key=api_key)

    # Define model parameters
    MODEL = "gpt-4o-mini-2024-07-18"
    TEMPERATURE = 0
    MAX_TOKENS = 3000

    # Define the prompt template
    def create_prompt(subject, body, filename, attachment_link, attachment_info):
        return f"""
        Classify the following email as PO or Not PO based on the provided information:

    - Subject: {subject}  
    - Body: {body}  
    - Filename: {filename}  
    - Attachment Link: {attachment_link}  
    - Attachment Information:{attachment_info}  

    ---

    Definition of PO:
    An email is classified as PO if it contains information related to a Purchase Order (a formal request to purchase goods or services). A Purchase Order typically includes at least some of the following details:  
    - Customer PO Number  
    - Item Name(s)  
    - Quantity  
    - Rate per Unit  
    - Unit of Measurement  
    - Delivery Dates  
    - Customer Name  

    Examples of Relevant Keywords:
    - "Purchase Order"  
    - "PO Number"  
    - "Order Confirmation"  
    - "Item Quantity"  
    - "Delivery Schedule"  

    Instructions for Classification:
    1. Analyze the Subject and Body:
       - Look for keywords or phrases indicating a Purchase Order.
       - Check for mentions of specific items, quantities, delivery details, or rates.

    2. Examine Attachments (if present):
       - Determine if the attachment contains any PO-related information (e.g., PO Number, item details).
       - Prioritize files labeled with terms like "PO" or "Order."

    3. Contextual Relevance:
       - If the email references a PO like pending order or has context of order is been given then classify as PO.
       - If sufficient PO-related information is found in the email body or attachments, classify it as PO.
       - Mark all the other general or specific business related email as Not PO whenever it does not implies anything about PO or associated with it.

    Classification Output:
    - PO: If the email contains PO-related information or references a formal request for items.  
    - Not PO: If the email lacks sufficient details or relevance to a Purchase Order.

    Strictly Classify as PO or Not PO do not mention any other information. 

        """

    # Function to classify a single email using OpenAI
    def classify_email(subject, body, filename, attachment_link, attachment_info):
        prompt = create_prompt(subject, body, filename, attachment_link, attachment_info)
        try:
            # Call OpenAI's chat completion API
            response = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": "You are a helpful assistant trained to classify emails as PO or Not PO."},
                    {"role": "user", "content": prompt},
                ],
                temperature=TEMPERATURE,
                max_tokens=MAX_TOKENS,
            )
            # Access the `content` attribute directly
            return response.choices[0].message.content.strip()
        except Exception as e:
            print(f"Error during classification: {e}")
            return "Error: Classification Failed"

    # Read data from Excel
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"Input file {input_file} not found.")
        return

    # Add a new column for Classification
    try:
        df['Classification'] = df.apply(
            lambda row: classify_email(
                row.get('Subject', ''),
                row.get('Body', ''),
                row.get('Filename', ''),
                row.get('Attachment Link', ''),
                row.get('extracted information from attachment', '')
            ),
            axis=1
        )
    except Exception as e:
        print(f"Error during classification: {e}")
        return

    # Write results back to Excel
    try:
        df.to_excel(output_file, index=False)
        print(f"Classification completed. The results are saved in {output_file}.")
    except Exception as e:
        print(f"Error saving output file: {e}")

# Example usage
'''if __name__ == "__main__":
    # Replace with your input file, output file, and API key
    input_file = 'output-test-case.xlsx'
    output_file = 'classified_emails-test-case.xlsx'
    api_key = os.environ.get("OPENAI_API_KEY", "Put your key")

    classify_emails_in_file(input_file, output_file, api_key)'''
