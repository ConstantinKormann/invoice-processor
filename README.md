# Invoice Processor

Batch process scanned invoices using OpenAI's GPT-4o Vision API and generate Excel spreadsheets with extracted data.

## 🎯 Features

- **Multi-format support**: Process JPG, PNG, TIFF, BMP, GIF, and PDF files
- **PDF handling**: Automatically converts PDF pages to images for processing
- **AI-powered extraction**: Uses GPT-4o Vision to extract structured data from invoices
- **Excel output**: Generates comprehensive Excel reports with all extracted data
- **Error tracking**: Color-coded status indicators (green for success, yellow for partial, red for errors)
- **Batch processing**: Process multiple invoices in one go
- **User-friendly GUI**: Easy-to-use graphical interface (or use CLI)
- **Persistent settings**: Remembers your API key and folder preferences

## 📋 Extracted Invoice Fields

The tool extracts the following information from each invoice:

- Invoice Number
- Invoice Date (YYYY-MM-DD format)
- Vendor/Supplier Name
- Total Amount
- Currency (EUR, USD, GBP, etc.)
- Tax/VAT Amount
- Payment Due Date
- Line Items (with description, quantity, unit price, and total)

## 🚀 Installation

### Prerequisites

- Python 3.8 or higher
- OpenAI API key (get one from https://platform.openai.com/api-keys)

### Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/ConstantinKormann/invoice-processor.git
   cd invoice-processor
   ```

2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up your OpenAI API key:
   
   **Option 1: Environment variable**
   ```bash
   export OPENAI_API_KEY='your-api-key-here'
   ```
   
   **Option 2: .env file**
   ```bash
   cp .env.example .env
   # Edit .env and add your API key
   ```

## 💻 Usage

### GUI Mode (Recommended)

1. Run the GUI application:
   ```bash
   python gui_processor.py
   ```

2. In the GUI:
   - Enter your OpenAI API key
   - Select the input folder containing your invoices
   - Select the output folder for the Excel report
   - Click "🚀 Process Invoices"

### CLI Mode

1. Place your invoice files (JPG, PNG, PDF, etc.) in the `INPUT` folder

2. Run the processor:
   ```bash
   python invoice_processor.py
   ```

3. Find the generated Excel report in the `OUTPUT` folder

## 📊 Output Format

The tool generates an Excel spreadsheet with the following columns:

| Column | Description |
|--------|-------------|
| Source File | Original filename |
| Invoice Number | Extracted invoice number/ID |
| Invoice Date | Invoice date (YYYY-MM-DD) |
| Vendor | Vendor/supplier name |
| Total Amount | Total amount (numeric only) |
| Currency | Currency code (EUR, USD, etc.) |
| Tax Amount | Tax/VAT amount if present |
| Due Date | Payment due date if present |
| Line Items | Formatted list of line items |
| Processing Status | ✅ OK / ⚠️ PARTIAL DATA / ❌ PROCESSING ERROR |
| Notes | Error messages or missing field details |

### Status Indicators

- **✅ OK** (white background): All fields successfully extracted
- **⚠️ PARTIAL DATA** (yellow background): Some fields could not be extracted
- **❌ PROCESSING ERROR** (red background): Invoice could not be processed

## 🔍 Error Handling

The tool is designed to be resilient:

- **Never crashes on individual failures**: If one invoice fails, processing continues with the rest
- **Detailed error reporting**: Each failed invoice is clearly marked in the Excel output
- **Warning summary**: At the end of processing, you'll see a summary of all problematic invoices
- **Partial data handling**: Even if some fields are missing, successfully extracted data is still saved

## 🛠️ Building Standalone Executables

### Windows

```bash
python build_windows.py
```

The executable will be created in the `dist` folder as `InvoiceProcessor.exe`.

### macOS

```bash
python build_macos.py
```

The application bundle will be created in the `dist` folder as `InvoiceProcessor.app`.

You can then drag it to your Applications folder.

## 📦 Dependencies

- `openai>=1.0.0` - OpenAI API client
- `openpyxl>=3.1.0` - Excel file generation
- `PyMuPDF>=1.23.0` - PDF processing
- `python-dotenv>=1.0.0` - Environment variable management
- `pyinstaller>=6.0.0` - Executable building (optional)

## 📁 Project Structure

```
invoice-processor/
├── invoice_processor.py    # Core processing logic
├── gui_processor.py         # GUI application
├── config.py                # Configuration management
├── build_windows.py         # Windows build script
├── build_macos.py           # macOS build script
├── requirements.txt         # Python dependencies
├── .env.example             # Environment variable template
├── .gitignore               # Git ignore rules
├── README.md                # This file
├── INPUT/                   # Place invoices here (CLI mode)
├── OUTPUT/                  # Generated reports appear here
└── assets/                  # Application icons (optional)
```

## 🐛 Troubleshooting

### "No module named 'openai'" error

Make sure you've installed the dependencies:
```bash
pip install -r requirements.txt
```

### "OPENAI_API_KEY not set" error

Set your API key using one of these methods:
- Set the `OPENAI_API_KEY` environment variable
- Create a `.env` file with your API key
- Enter your API key in the GUI

### PDF processing errors

Ensure PyMuPDF is properly installed:
```bash
pip install --upgrade PyMuPDF
```

### Excel file won't open

Make sure openpyxl is installed:
```bash
pip install --upgrade openpyxl
```

### API rate limits

If you're processing many invoices, you might hit OpenAI's rate limits. The tool will show an error for those invoices. Wait a moment and try again, or upgrade your OpenAI plan.

## 🔐 Security Note

Your OpenAI API key is stored locally on your computer in:
- **Windows**: `%APPDATA%\InvoiceProcessor\config.json`
- **macOS/Linux**: `~/.config/InvoiceProcessor/config.json`

Never share this file or commit it to version control.

## 📄 License

This project is open source and available under the MIT License.

## 🙏 Acknowledgments

Powered by OpenAI's GPT-4o Vision API.

## 📌 Version

v1.0.0

---

**Need help?** Open an issue on GitHub or contact the maintainer.