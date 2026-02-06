"""
Invoice Processor
Batch processes scanned invoices using OpenAI's GPT-4o Vision API
and generates Excel spreadsheets with extracted data.
"""

import os
import base64
import sys
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional

from openai import OpenAI
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


# Supported image formats
SUPPORTED_IMAGE_FORMATS = {'.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp', '.gif'}
SUPPORTED_FORMATS = SUPPORTED_IMAGE_FORMATS | {'.pdf'}


# System prompt for OpenAI Vision API
SYSTEM_PROMPT = """You are an expert at reading and extracting data from invoices (both printed and scanned).
Your task is to extract the following fields from the invoice image:
- invoice_number: The invoice number/ID
- invoice_date: The date of the invoice (format: YYYY-MM-DD if possible)
- vendor_name: The vendor/supplier name
- total_amount: The total amount (numeric value only, no currency symbols)
- currency: The currency (e.g., EUR, USD, GBP)
- tax_amount: The tax/VAT amount if present (numeric value only), or null
- due_date: The payment due date if present (format: YYYY-MM-DD if possible), or null
- line_items: A list of line items, each with: description, quantity, unit_price, total

Return ONLY valid JSON with these fields. If you cannot confidently extract a field, set it to null.
Do NOT include any markdown formatting, code fences, or explanation — just the raw JSON object."""


def encode_image_to_base64(image_path: str) -> str:
    """
    Convert an image file to base64 encoding for API transmission.
    
    Args:
        image_path: Path to the image file
        
    Returns:
        Base64 encoded string of the image
    """
    with open(image_path, 'rb') as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')


def pdf_to_images(pdf_path: str) -> List[bytes]:
    """
    Convert PDF pages to images using PyMuPDF.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        List of image bytes (PNG format) for each page
    """
    images = []
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # Render page to image at 2x resolution for better quality
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_bytes = pix.tobytes("png")
            images.append(img_bytes)
        doc.close()
    except Exception as e:
        raise Exception(f"Error converting PDF to images: {str(e)}")
    
    return images


def extract_invoice_data(image_data: bytes, api_key: str, filename: str, 
                        error_callback=None) -> Dict:
    """
    Call OpenAI Vision API to extract invoice data from an image.
    
    Args:
        image_data: Image data as bytes (for PDFs) or None (for image files)
        api_key: OpenAI API key
        filename: Source filename for error messages
        error_callback: Optional callback function for errors
        
    Returns:
        Dictionary with extraction results including status, data, and error_message
    """
    result = {
        'filename': filename,
        'status': 'error',
        'error_message': '',
        'data': {}
    }
    
    try:
        client = OpenAI(api_key=api_key)
        
        # Encode image to base64
        if isinstance(image_data, bytes):
            # PDF page already in bytes
            base64_image = base64.b64encode(image_data).decode('utf-8')
            media_type = 'image/png'
        else:
            # Regular image file
            base64_image = encode_image_to_base64(image_data)
            ext = Path(image_data).suffix.lower()
            media_type_map = {
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.png': 'image/png',
                '.gif': 'image/gif',
                '.bmp': 'image/bmp',
                '.tiff': 'image/tiff',
                '.tif': 'image/tiff'
            }
            media_type = media_type_map.get(ext, 'image/jpeg')
        
        # Call OpenAI API
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": SYSTEM_PROMPT
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{media_type};base64,{base64_image}"
                            }
                        },
                        {
                            "type": "text",
                            "text": "Please extract all invoice data from this image."
                        }
                    ]
                }
            ],
            max_tokens=16384
        )
        
        # Parse the response
        response_text = response.choices[0].message.content.strip()
        
        # Try to clean up the response if it contains markdown code fences
        if response_text.startswith('```'):
            # Remove markdown code fences
            lines = response_text.split('\n')
            # Remove first line (```json or ```)
            if lines[0].strip().startswith('```'):
                lines = lines[1:]
            # Remove last line (```)
            if lines and lines[-1].strip() == '```':
                lines = lines[:-1]
            response_text = '\n'.join(lines).strip()
        
        # Parse JSON
        try:
            data = json.loads(response_text)
            
            # Check if we got any data (only check for None, empty string is valid)
            null_count = sum(1 for v in data.values() if v is None)
            total_fields = len(data)
            
            if null_count == total_fields:
                # All fields are null
                result['status'] = 'error'
                result['error_message'] = 'Could not extract any data from invoice'
                result['data'] = data
            elif null_count > 0:
                # Some fields are null
                result['status'] = 'partial'
                missing_fields = [k for k, v in data.items() if v is None]
                result['error_message'] = f"Missing fields: {', '.join(missing_fields)}"
                result['data'] = data
            else:
                # All fields extracted
                result['status'] = 'ok'
                result['data'] = data
                
        except json.JSONDecodeError as e:
            result['status'] = 'error'
            result['error_message'] = f'Invalid JSON response: {str(e)}'
            result['data'] = {}
            if error_callback:
                error_callback(f"JSON parse error for {filename}: {str(e)}")
        
    except Exception as e:
        result['status'] = 'error'
        result['error_message'] = f'API error: {str(e)}'
        result['data'] = {}
        if error_callback:
            error_callback(f"Error processing {filename}: {str(e)}")
    
    return result


def process_all_invoices(input_folder: str, api_key: str, progress_callback=None, 
                         status_callback=None, cancel_event=None) -> List[Dict]:
    """
    Batch process all invoices in the INPUT folder.
    
    Args:
        input_folder: Path to the INPUT folder
        api_key: OpenAI API key
        progress_callback: Optional callback(current, total) for progress updates
        status_callback: Optional callback(message) for status updates
        cancel_event: Optional threading.Event to check for cancellation
        
    Returns:
        List of dictionaries containing processing results
    """
    input_path = Path(input_folder)
    
    if not input_path.exists():
        raise FileNotFoundError(f"INPUT folder not found: {input_folder}")
    
    # Get all supported files and sort them
    invoice_files = []
    for ext in SUPPORTED_FORMATS:
        invoice_files.extend(input_path.glob(f"*{ext}"))
        invoice_files.extend(input_path.glob(f"*{ext.upper()}"))
    
    invoice_files = sorted(set(invoice_files))
    
    if not invoice_files:
        raise ValueError(f"No invoice files found in {input_folder}")
    
    if status_callback:
        status_callback(f"Found {len(invoice_files)} invoices to process")
    else:
        print(f"Found {len(invoice_files)} invoices to process")
    
    results = []
    for idx, invoice_file in enumerate(invoice_files, 1):
        # Check for cancellation
        if cancel_event and cancel_event.is_set():
            if status_callback:
                status_callback("Processing cancelled by user")
            raise InterruptedError("Processing cancelled by user")
        
        if status_callback:
            status_callback(f"Processing invoice {idx}/{len(invoice_files)}: {invoice_file.name}")
        else:
            print(f"Processing {idx}/{len(invoice_files)}: {invoice_file.name}")
        
        if progress_callback:
            progress_callback(idx, len(invoice_files))
        
        # Create error callback wrapper if status_callback is provided
        def error_cb(msg):
            if status_callback:
                status_callback(msg)
        
        # Handle PDFs differently
        if invoice_file.suffix.lower() == '.pdf':
            try:
                # Convert PDF pages to images
                images = pdf_to_images(str(invoice_file))
                
                if not images:
                    results.append({
                        'filename': invoice_file.name,
                        'status': 'error',
                        'error_message': 'PDF has no pages',
                        'data': {}
                    })
                    continue
                
                # Process each page and combine results
                # For multi-page PDFs, we'll process the first page only
                # (you could modify this to handle multi-page differently)
                page_result = extract_invoice_data(
                    images[0],
                    api_key,
                    invoice_file.name,
                    error_callback=error_cb if status_callback else None
                )
                
                if len(images) > 1 and status_callback:
                    status_callback(f"  Note: PDF has {len(images)} pages, processing first page only")
                
                results.append(page_result)
                
            except Exception as e:
                results.append({
                    'filename': invoice_file.name,
                    'status': 'error',
                    'error_message': f'PDF processing error: {str(e)}',
                    'data': {}
                })
                if error_cb:
                    error_cb(f"Error processing PDF {invoice_file.name}: {str(e)}")
        else:
            # Process regular image file
            result = extract_invoice_data(
                str(invoice_file),
                api_key,
                invoice_file.name,
                error_callback=error_cb if status_callback else None
            )
            results.append(result)
    
    # Log warning summary for failed/partial invoices
    failed_invoices = [r['filename'] for r in results if r['status'] in ['error', 'partial']]
    if failed_invoices:
        warning_msg = f"Could not fully process the following invoices: {', '.join(failed_invoices)}"
        logger.warning(warning_msg)
        if status_callback:
            status_callback(f"\n⚠️ WARNING: {warning_msg}")
    
    return results


def format_line_items(line_items) -> str:
    """
    Format line items as a readable string.
    
    Args:
        line_items: List of line item dictionaries
        
    Returns:
        Formatted string representation
    """
    if not line_items or not isinstance(line_items, list):
        return ''
    
    formatted = []
    for item in line_items:
        if isinstance(item, dict):
            desc = item.get('description', '')
            qty = item.get('quantity', '')
            unit_price = item.get('unit_price', '')
            total = item.get('total', '')
            formatted.append(f"{qty}x {desc} @{unit_price} = {total}")
    
    return '; '.join(formatted)


def create_excel_report(results: List[Dict], output_path: str, status_callback=None):
    """
    Generate an Excel report from processing results.
    
    Args:
        results: List of processing result dictionaries
        output_path: Path to save the Excel file
        status_callback: Optional callback(message) for status updates
    """
    if status_callback:
        status_callback("Creating Excel report...")
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Data"
    
    # Define headers
    headers = [
        'Source File',
        'Invoice Number',
        'Invoice Date',
        'Vendor',
        'Total Amount',
        'Currency',
        'Tax Amount',
        'Due Date',
        'Line Items',
        'Processing Status',
        'Notes'
    ]
    
    # Write headers with formatting
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True, size=11)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Write data rows
    for row_num, result in enumerate(results, 2):
        data = result.get('data', {})
        status = result.get('status', 'error')
        error_message = result.get('error_message', '')
        
        # Source File
        ws.cell(row=row_num, column=1).value = result.get('filename', '')
        
        # Invoice fields
        ws.cell(row=row_num, column=2).value = data.get('invoice_number', '')
        ws.cell(row=row_num, column=3).value = data.get('invoice_date', '')
        ws.cell(row=row_num, column=4).value = data.get('vendor_name', '')
        ws.cell(row=row_num, column=5).value = data.get('total_amount', '')
        ws.cell(row=row_num, column=6).value = data.get('currency', '')
        ws.cell(row=row_num, column=7).value = data.get('tax_amount', '')
        ws.cell(row=row_num, column=8).value = data.get('due_date', '')
        
        # Line items (formatted)
        line_items = data.get('line_items', [])
        ws.cell(row=row_num, column=9).value = format_line_items(line_items)
        
        # Processing Status
        if status == 'ok':
            ws.cell(row=row_num, column=10).value = '✅ OK'
        elif status == 'partial':
            ws.cell(row=row_num, column=10).value = '⚠️ PARTIAL DATA'
            # Yellow fill for partial data
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for col in range(1, 12):
                ws.cell(row=row_num, column=col).fill = yellow_fill
        else:  # error
            ws.cell(row=row_num, column=10).value = '❌ PROCESSING ERROR'
            # Red fill for errors
            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            for col in range(1, 12):
                cell = ws.cell(row=row_num, column=col)
                cell.fill = red_fill
                cell.font = Font(color="FFFFFF")
        
        # Notes (error message)
        ws.cell(row=row_num, column=11).value = error_message
    
    # Auto-size columns
    for col_num in range(1, len(headers) + 1):
        column_letter = get_column_letter(col_num)
        max_length = 0
        for cell in ws[column_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except (TypeError, AttributeError):
                pass
        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save workbook
    wb.save(output_path)
    
    if status_callback:
        status_callback(f"Excel report saved to: {output_path}")
    else:
        print(f"Excel report saved to: {output_path}")


def main():
    """
    Main function to orchestrate the workflow.
    """
    print("=" * 60)
    print("Invoice Processor v1.0.0")
    print("=" * 60)
    print()
    
    # Get API key from environment
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        print("Error: OPENAI_API_KEY environment variable not set")
        print("Please set it in a .env file or export it in your shell")
        sys.exit(1)
    
    # Define paths
    script_dir = Path(__file__).parent
    input_folder = script_dir / 'INPUT'
    output_folder = script_dir / 'OUTPUT'
    
    # Create OUTPUT folder if it doesn't exist
    output_folder.mkdir(exist_ok=True)
    
    try:
        # Process all invoices
        print("Processing invoices from INPUT folder...")
        results = process_all_invoices(str(input_folder), api_key)
        print(f"\nProcessed {len(results)} invoices")
        print()
        
        # Generate timestamp for filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Create Excel report
        print("Generating Excel report...")
        excel_path = output_folder / f'invoices_{timestamp}.xlsx'
        create_excel_report(results, str(excel_path))
        print()
        
        print("=" * 60)
        print("Processing complete!")
        print("=" * 60)
        print(f"Output file created: {excel_path}")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    # Load environment variables from .env file if it exists
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass
    
    main()
