# OCR Tax Details Extraction

A Python-based solution for extracting tax-related information from documents using Optical Character Recognition (OCR) technology.

## Overview

This project automates the extraction of tax details from various document formats including PDF files, images, and scanned documents. The system uses OCR to convert text from images and applies intelligent parsing to identify and extract specific tax-related information.

## Features

- **Multi-format Support**: Process PDF files, images (PNG, JPEG, TIFF), and scanned documents
- **Tax Information Extraction**: Automatically identifies and extracts key tax details including:
  - Tax ID numbers
  - Tax amounts
  - Tax rates
  - Invoice numbers
  - Date information
  - Vendor/company details
- **Data Validation**: Validates extracted data using predefined patterns and rules
- **Export Options**: Output results in JSON, CSV, or Excel formats
- **Batch Processing**: Handle multiple documents simultaneously
- **Error Handling**: Robust error handling for various document qualities and formats

## Prerequisites

- Python 3.7 or higher
- Tesseract OCR engine installed on your system

## Installation

1. Clone the repository:
```bash
git clone https://github.com/Stanley17932/ocr-taxdetailsextraction.git
cd ocr-taxdetailsextraction
```

2. Install required Python packages:
```bash
pip install -r requirements.txt
```

3. Install Tesseract OCR:

**Windows:**
```bash
# Download and install from: https://github.com/UB-Mannheim/tesseract/wiki
```

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

## Dependencies

```
pytesseract==0.3.10
opencv-python==4.8.0.74
Pillow==9.5.0
pandas==2.0.3
numpy==1.24.3
PyPDF2==3.0.1
regex==2023.6.3
openpyxl==3.1.2
```

## Usage

### Basic Usage

```python
from ocr_tax_extractor import TaxDetailsExtractor

# Initialize the extractor
extractor = TaxDetailsExtractor()

# Extract tax details from a single document
result = extractor.extract_from_file('path/to/tax_document.pdf')
print(result)
```

### Command Line Interface

```bash
# Process a single file
python main.py --input document.pdf --output results.json

# Process multiple files
python main.py --input-dir /path/to/documents/ --output-dir /path/to/results/

# Specify output format
python main.py --input document.pdf --output results.csv --format csv
```

### Configuration

Create a `config.json` file to customize extraction parameters:

```json
{
  "ocr_config": {
    "language": "eng",
    "psm": 6,
    "oem": 3
  },
  "extraction_patterns": {
    "tax_id": "\\b\\d{2}-\\d{7}\\b",
    "amount": "\\$?\\d+\\.\\d{2}",
    "date": "\\d{1,2}/\\d{1,2}/\\d{4}"
  },
  "output_format": "json"
}
```

## Project Structure

```
ocr-taxdetailsextraction/
├── src/
│   ├── __init__.py
│   ├── ocr_engine.py          # OCR processing functionality
│   ├── text_parser.py         # Text parsing and pattern matching
│   ├── data_extractor.py      # Tax details extraction logic
│   └── utils.py               # Utility functions
├── tests/
│   ├── test_ocr_engine.py
│   ├── test_parser.py
│   └── sample_documents/
├── config/
│   └── config.json
├── requirements.txt
├── main.py
└── README.md
```

## Output Format

The extracted tax details are returned in the following structure:

```json
{
  "document_name": "tax_document.pdf",
  "extraction_timestamp": "2024-01-15T10:30:00Z",
  "tax_details": {
    "tax_id": "12-3456789",
    "company_name": "ABC Corporation",
    "invoice_number": "INV-2024-001",
    "tax_amount": 150.75,
    "tax_rate": 0.15,
    "subtotal": 1005.00,
    "total_amount": 1155.75,
    "date": "2024-01-15",
    "currency": "USD"
  },
  "confidence_score": 0.92,
  "processing_time": 3.45
}
```

## Error Handling

The system includes comprehensive error handling for:
- Invalid file formats
- Poor image quality
- Missing or corrupted files
- OCR processing failures
- Pattern matching errors

## Performance Optimization

- **Image Preprocessing**: Automatic image enhancement for better OCR accuracy
- **Multi-threading**: Parallel processing for batch operations
- **Caching**: Results caching to avoid reprocessing identical documents
- **Memory Management**: Efficient memory usage for large document batches

## Testing

Run the test suite:

```bash
python -m pytest tests/
```

Run specific tests:

```bash
python -m pytest tests/test_ocr_engine.py -v
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## Known Issues

- OCR accuracy may vary depending on document quality
- Some handwritten text may not be recognized accurately
- Complex document layouts may require manual review

## Future Enhancements

- Support for additional languages
- Machine learning model integration for improved accuracy
- Web-based interface
- API endpoint development
- Database integration for result storage

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

Stanley Otieno - [LinkedIn](https://www.linkedin.com/in/stanley-otieno-838707142/)

Project Link: [https://github.com/Stanley17932/ocr-taxdetailsextraction](https://github.com/Stanley17932/ocr-taxdetailsextraction)

## Acknowledgments

- Tesseract OCR for text recognition capabilities
- OpenCV for image processing
- Python community for excellent libraries and documentation
