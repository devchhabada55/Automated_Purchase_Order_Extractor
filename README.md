# AI-Powered Purchase Order Extraction Pipeline

## Project Overview
This automated pipeline extracts, classifies, and processes purchase order (PO) information from emails using advanced AI and machine learning techniques. The pipeline integrates robust natural language processing (NLP) capabilities to streamline PO handling with high accuracy and efficiency.

## 🚀 Features
- **Automated Email Extraction:** Seamlessly retrieves email data from Gmail accounts using the Gmail API.
- **AI-Powered Email Classification:** Leverages state-of-the-art AI models to classify emails as "PO" or "Not PO."
- **Purchase Order Details Extraction:** Extracts relevant details like PO numbers, items, quantities, and delivery schedules from email content and attachments.
- **Flexible Date Range Processing:** Process emails within user-defined date ranges.
- **Comprehensive Logging and Error Handling:** Detailed logs for debugging and pipeline performance tracking.

## 📋 Prerequisites
- Python 3.8+
- Google Cloud Project with Gmail API enabled
- OpenAI API Key

## 🔧 Setup Instructions

### 1. Clone the Repository
```bash
[git clone https://github.com/devchhabada55/po-extraction-pipeline.git](https://github.com/devchhabada55/Automated_Purchase_Order_Extractor.git)
cd po-extraction-pipeline
```

### 2. Create Virtual Environment
```bash
python -m venv po_extraction_env
source po_extraction_env/bin/activate  # Windows: po_extraction_env\Scripts\activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Credentials
1. **Google Cloud Console Setup:**
   - Create a new project
   - Enable Gmail API
   - Generate OAuth 2.0 credentials
   - Download and place the credentials file in the project directory

2. **Set Environment Variables**
```bash
export OPENAI_API_KEY='your_openai_api_key'
export GOOGLE_APPLICATION_CREDENTIALS='path/to/credentials.json'
```

## 🖥️ Usage

### Command Line Options
```bash
python pipeline.py --start-date YYYY-MM-DD --end-date YYYY-MM-DD [optional_arguments]

Optional Arguments:
--input-file     Specify a custom input Excel file
--model          Choose classification model (openai/local, default: openai)
```

### Example Command
```bash
python pipeline.py --start-date 2024-11-30 --end-date 2024-12-02
```

## 🧠 AI Models Used
### Default Model:
- **OpenAI GPT-4:** Provides high accuracy for email classification and PO detail extraction. Tested extensively with various scenarios and proved reliable for most real-world use cases.

### Alternative Models Tried:
- **OpenAI GPT-3.5:** Offers a cost-effective solution with good classification performance, though slightly less accurate compared to GPT-4 for edge cases.
- **Open Source Models:** 
  - **Llama 3 (Meta):** Fine-tuned for text classification tasks. Provides comparable performance to OpenAI models for email classification when properly fine-tuned.
  - **Flan-T5 (Google):** Efficient for small-scale deployments with minimal resource usage.
  - **Hugging Face Models:** Experimented with DistilBERT and Roberta-based models. While they required fine-tuning, they showed promise for budget-friendly setups.

### Recommendations:
- If budget constraints are a concern, try fine-tuning open-source models like Llama 2 or Flan-T5 using your dataset.
- For scalability and ease of use, OpenAI's GPT-4 remains the most robust choice.

## 📂 Project Structure
- `pipeline.py`: Main pipeline execution script
- `gmailreader.py`: Email extraction module
- `email_classification.py`: AI-based email classification
- `data_extraction.py`: Purchase Order details extraction
- `attachments/`: Folder for downloaded email attachments
- `logs/`: Logging output directory

## 🔍 Output Files
- `emails_data-testcase.xlsx`: Extracted email metadata
- `classified_emails-test-case.xlsx`: Classified emails
- `po_extracted-testcase.json`: Extracted Purchase Order details
- `po_extraction_pipeline.log`: Detailed execution logs

### Deployment Considerations
- **OpenAI Models:** Seamless integration, higher costs, no fine-tuning required.
- **Open Source Models:** Cost-effective for high-volume processing, requires fine-tuning for optimal performance.

## 🛠️ Troubleshooting
- **Dependency Issues:** Ensure all dependencies are installed via `pip install -r requirements.txt`.
- **API Key Errors:** Verify that the `OPENAI_API_KEY` and `GOOGLE_APPLICATION_CREDENTIALS` environment variables are set correctly.
- **Log File Analysis:** Check `po_extraction_pipeline.log` for detailed error information and pipeline progress.

## 📜 Logging
Extensive logging is implemented to track pipeline execution:
- Console output
- Detailed log file (`po_extraction_pipeline.log`)
- Captures each pipeline stage's status and potential errors

## 🔒 Security Considerations
- Use environment variables for sensitive credentials
- Do not commit API keys or credentials to version control
- Implement proper access controls for sensitive data

## 📊 Performance Monitoring
- Tracks execution time for each pipeline stage
- Provides size verification of output files
- Detailed error reporting mechanism

## 🤝 Contributing
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request


## 📧 Contact
Dev Chhabada/ chhabadadev@gmail.com
