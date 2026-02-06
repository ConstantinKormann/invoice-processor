"""
Student Notes Processor
Batch processes scanned student notes using OpenAI's GPT-4o Vision API
and generates compiled PDF and Word documents.
"""

import os
import base64
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Dict

from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF


# Supported image formats
SUPPORTED_FORMATS = {'.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp', '.gif'}

# System prompt for OpenAI
SYSTEM_PROMPT = """
You are an expert at reading handwritten student notes. 
Your task is to:
1. Extract all text accurately from the handwritten notes
2. Preserve the structure of questions and answers
3. Format questions clearly with Q1:, Q2:, Q3: etc.
4. Maintain the original layout and organization as much as possible

Please transcribe all visible text from the image, organizing it clearly.
"""

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


def extract_text_from_image(image_path: str, api_key: str, error_callback=None) -> str:
    """
    Call OpenAI Vision API to extract handwritten text from an image.
    
    Args:
        image_path: Path to the image file
        api_key: OpenAI API key
        error_callback: Optional callback function for errors (receives error message)
        
    Returns:
        Extracted text from the image
    """
    try:
        client = OpenAI(api_key=api_key)
        
        # Encode image to base64
        base64_image = encode_image_to_base64(image_path)
        
        # Get file extension to determine media type
        ext = Path(image_path).suffix.lower()
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
                            "text": "Please transcribe all the handwritten text from this image."
                        }
                    ]
                }
            ],
            max_tokens=16384
        )
        
        return response.choices[0].message.content
        
    except Exception as e:
        error_msg = f"Error processing {Path(image_path).name}: {str(e)}"
        if error_callback:
            error_callback(error_msg)
        else:
            print(error_msg)
        return f"[Error extracting text from {Path(image_path).name}: {str(e)}]"


def process_all_images(input_folder: str, api_key: str, progress_callback=None, 
                      status_callback=None, cancel_event=None) -> List[Dict[str, str]]:
    """
    Batch process all images in the INPUT folder.
    
    Args:
        input_folder: Path to the INPUT folder
        api_key: OpenAI API key
        progress_callback: Optional callback(current, total) for progress updates
        status_callback: Optional callback(message) for status updates
        cancel_event: Optional threading.Event to check for cancellation
        
    Returns:
        List of dictionaries containing filename and extracted text
    """
    input_path = Path(input_folder)
    
    if not input_path.exists():
        raise FileNotFoundError(f"INPUT folder not found: {input_folder}")
    
    # Get all image files and sort them
    image_files = []
    for ext in SUPPORTED_FORMATS:
        image_files.extend(input_path.glob(f"*{ext}"))
        image_files.extend(input_path.glob(f"*{ext.upper()}"))
    
    image_files = sorted(set(image_files))
    
    if not image_files:
        raise ValueError(f"No image files found in {input_folder}")
    
    if status_callback:
        status_callback(f"Found {len(image_files)} images to process")
    else:
        print(f"Found {len(image_files)} images to process")
    
    results = []
    for idx, image_file in enumerate(image_files, 1):
        # Check for cancellation
        if cancel_event and cancel_event.is_set():
            if status_callback:
                status_callback("Processing cancelled by user")
            raise InterruptedError("Processing cancelled by user")
        
        if status_callback:
            status_callback(f"Processing page {idx}/{len(image_files)}: {image_file.name}")
        else:
            print(f"Processing {idx}/{len(image_files)}: {image_file.name}")
        
        if progress_callback:
            progress_callback(idx, len(image_files))
        
        # Create error callback wrapper if status_callback is provided
        def error_cb(msg):
            if status_callback:
                status_callback(msg)
        
        extracted_text = extract_text_from_image(
            str(image_file), 
            api_key,
            error_callback=error_cb if status_callback else None
        )
        
        results.append({
            'filename': image_file.name,
            'page_number': idx,
            'text': extracted_text
        })
    
    return results


def create_word_document(results: List[Dict[str, str]], output_path: str, status_callback=None):
    """
    Generate a Word document from extracted text.
    
    Args:
        results: List of dictionaries with extracted text and metadata
        output_path: Path to save the Word document
        status_callback: Optional callback(message) for status updates
    """
    if status_callback:
        status_callback("Generating Word document...")
    
    doc = Document()
    
    # Add title
    title = doc.add_heading('Extracted Student Notes', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add timestamp
    timestamp = doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    timestamp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    # Add each page
    for result in results:
        # Add page heading
        heading = doc.add_heading(
            f"Page {result['page_number']} - {result['filename']}", 
            level=1
        )
        
        # Add extracted text
        text_para = doc.add_paragraph(result['text'])
        text_para.paragraph_format.space_after = Pt(12)
        
        # Add page break (except for last page)
        if result != results[-1]:
            doc.add_page_break()
    
    # Save document
    doc.save(output_path)
    
    if status_callback:
        status_callback(f"Word document saved to: {output_path}")
    else:
        print(f"Word document saved to: {output_path}")


def create_pdf_document(results: List[Dict[str, str]], output_path: str, status_callback=None):
    """
    Generate a PDF document from extracted text.
    
    Args:
        results: List of dictionaries with extracted text and metadata
        output_path: Path to save the PDF document
        status_callback: Optional callback(message) for status updates
    """
    if status_callback:
        status_callback("Generating PDF document...")
    
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Add first page with title
    pdf.add_page()
    pdf.set_font('Arial', 'B', 24)
    pdf.cell(0, 20, 'Extracted Student Notes', align='C', ln=True)
    
    pdf.set_font('Arial', '', 12)
    pdf.cell(0, 10, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", 
             align='C', ln=True)
    pdf.ln(10)
    
    # Add each page
    for result in results:
        pdf.add_page()
        
        # Add page heading
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, f"Page {result['page_number']} - {result['filename']}", ln=True)
        pdf.ln(5)
        
        # Add extracted text
        pdf.set_font('Arial', '', 11)
        
        # Handle text encoding for PDF - use UTF-8 compatible approach
        text = result['text']
        # FPDF2 supports UTF-8, so we can use the text directly
        # Only replace truly problematic characters if any encoding issues occur
        try:
            pdf.multi_cell(0, 6, text)
        except Exception:
            # Fallback: remove non-ASCII characters only if necessary
            text = ''.join(char if ord(char) < 128 else '?' for char in text)
            pdf.multi_cell(0, 6, text)
        pdf.ln(5)
    
    # Save PDF
    pdf.output(output_path)
    
    if status_callback:
        status_callback(f"PDF document saved to: {output_path}")
    else:
        print(f"PDF document saved to: {output_path}")


def main():
    """
    Main function to orchestrate the workflow.
    """
    print("=" * 60)
    print("Student Notes Processor v1")
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
        # Process all images
        print("Step 1: Processing images from INPUT folder...")
        results = process_all_images(str(input_folder), api_key)
        print(f"\nSuccessfully processed {len(results)} images")
        print()        
        # Generate timestamp for filenames
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Create Word document
        print("Step 2: Generating Word document...")
        word_path = output_folder / f'student_notes_{timestamp}.docx'
        create_word_document(results, str(word_path))
        print()        
        # Create PDF document
        print("Step 3: Generating PDF document...")
        pdf_path = output_folder / f'student_notes_{timestamp}.pdf'
        create_pdf_document(results, str(pdf_path))
        print()        
        print("=" * 60)
        print("Processing complete!")
        print("=" * 60)
        print(f"Output files created in: {output_folder}")
        print(f"  - {word_path.name}")
        print(f"  - {pdf_path.name}")
        
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