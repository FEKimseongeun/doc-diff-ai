# Document Comparison Service

An AI-powered service that compares original and revised documents (Word, PDF, Excel) and identifies changes including text modifications, formatting differences, table changes, and image alterations.

## Features

- **Multi-format Support**: Compare DOCX, PDF, and XLSX files
- **Comprehensive Change Detection**:
  - Text content changes (additions, deletions, modifications)
  - Formatting changes (fonts, colors, styles, borders)
  - Table structure and content changes
  - Image content and size changes
  - Structural changes (page count, sheet count, etc.)
- **Visual Reports**: Generate detailed HTML reports with highlighted changes
- **Web Interface**: User-friendly web application for easy document comparison
- **Export Functionality**: Download detailed comparison reports

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

### Setup

1. **Clone or download the project**:
   ```bash
   git clone <repository-url>
   cd document-comparison-services
   ```

2. **Create a virtual environment** (recommended):
   ```bash
   python -m venv venv
   
   # On Windows:
   venv\Scripts\activate
   
   # On macOS/Linux:
   source venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**:
   ```bash
   python app.py
   ```

5. **Access the web interface**:
   Open your browser and navigate to `http://localhost:5000`

## Quick Start

### Windows Users
```bash
# Double-click start_service.bat
# Or run manually:
start_service.bat
```

### Unix/Linux/Mac Users
```bash
# Make the script executable:
chmod +x start_service.sh

# Run the startup script:
./start_service.sh
```

## Usage

### Web Interface

1. **Upload Documents**: Select your original and revised documents using the file upload interface
2. **Compare**: Click "Compare Documents" to start the analysis
3. **View Results**: Review the summary of changes detected
4. **Download Report**: Get a detailed HTML report with all changes highlighted

### Supported File Types

- **Microsoft Word (.docx)**: Full text, formatting, tables, and images
- **PDF (.pdf)**: Text content and page structure
- **Excel (.xlsx)**: Cell values, formatting, and sheet structure

### File Size Limits

- Maximum file size: 50MB per file
- Recommended: Keep files under 10MB for optimal performance

## API Endpoints

### POST /compare
Compare two documents and return change analysis.

**Request**:
- `original`: Original document file
- `revised`: Revised document file

**Response**:
```json
{
  "success": true,
  "changes": {
    "text_changes": [...],
    "formatting_changes": [...],
    "table_changes": [...],
    "image_changes": [...],
    "structural_changes": [...],
    "summary": {
      "total_changes": 15,
      "text_changes_count": 8,
      "formatting_changes_count": 3,
      "table_changes_count": 2,
      "image_changes_count": 1,
      "structural_changes_count": 1
    }
  },
  "report_path": "comparison_report_20231201_143022.html"
}
```

### GET /download_report/<filename>
Download a generated comparison report.

## Architecture

### Core Components

1. **DocumentParser**: Extracts content from various document formats
2. **ChangeDetector**: Identifies differences between documents
3. **ReportGenerator**: Creates detailed HTML and JSON reports
4. **Web Interface**: Flask-based web application

### Document Processing Pipeline

1. **Parse**: Extract text, formatting, tables, and images from documents
2. **Compare**: Analyze differences using various algorithms
3. **Report**: Generate comprehensive change reports
4. **Visualize**: Present results in user-friendly format

## Change Detection Algorithms

### Text Changes
- Uses Python's `difflib` for line-by-line comparison
- Identifies additions, deletions, and modifications
- Preserves context and positioning information

### Formatting Changes
- Compares font properties (name, size, color, style)
- Analyzes paragraph and cell formatting
- Detects border and fill changes in spreadsheets

### Table Changes
- Compares table structure (rows, columns)
- Identifies cell content modifications
- Tracks table additions and deletions

### Image Changes
- Uses structural similarity index (SSIM) for content comparison
- Compares image dimensions and properties
- Detects image additions, deletions, and modifications

## Configuration

### Environment Variables

- `MAX_CONTENT_LENGTH`: Maximum file size (default: 50MB)
- `UPLOAD_FOLDER`: Directory for temporary file storage
- `REPORTS_FOLDER`: Directory for generated reports

### Customization

You can modify the following parameters in `change_detector.py`:

- `similarity_threshold`: Text similarity threshold (default: 0.8)
- `image_similarity_threshold`: Image similarity threshold (default: 0.95)

## Troubleshooting

### Common Issues

1. **File Upload Errors**:
   - Ensure file size is under 50MB
   - Check file format is supported (.docx, .pdf, .xlsx)
   - Verify file is not corrupted

2. **Installation Issues**:
   - Make sure Python 3.8+ is installed
   - Use virtual environment to avoid dependency conflicts
   - On Windows, you may need to install Visual C++ Build Tools for some packages

3. **Performance Issues**:
   - Large files may take longer to process
   - Consider splitting very large documents
   - Ensure sufficient RAM for image processing

### Error Messages

- **"Unsupported file format"**: File type not supported
- **"File too large"**: Exceeds 50MB limit
- **"Network error"**: Connection issues with the service
- **"Parsing error"**: Document may be corrupted or password-protected

## Development

### Project Structure

```
document-comparison-service/
├── app.py                 # Flask web application
├── document_parser.py     # Document parsing logic
├── change_detector.py     # Change detection algorithms
├── report_generator.py    # Report generation
├── templates/
│   └── index.html        # Web interface
├── static/               # Static assets
├── uploads/              # Temporary file storage
├── reports/              # Generated reports
├── requirements.txt      # Python dependencies
├── start_service.bat     # Windows startup script
├── start_service.sh      # Unix/Linux/Mac startup script
└── README.md            # This file
```

### Adding New File Formats

1. Extend `DocumentParser` class with new parsing method
2. Update `supported_formats` dictionary
3. Add format-specific change detection logic
4. Update web interface to accept new file types

### Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review the error messages
3. Create an issue with detailed information about your problem

## Future Enhancements

- Support for additional file formats (PPTX, RTF, etc.)
- Batch processing capabilities
- API authentication and rate limiting
- Cloud storage integration
- Real-time collaboration features
- Advanced change visualization
- Machine learning-based change classification
