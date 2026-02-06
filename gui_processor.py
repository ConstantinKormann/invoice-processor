"""
Invoice Processor - GUI Application
Batch processes scanned invoices using OpenAI's GPT-4o Vision API
and generates Excel spreadsheets with extracted data.
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import os
from pathlib import Path
from datetime import datetime

from config import Config
import invoice_processor


class InvoiceProcessorGUI:
    """Main GUI application for Invoice Processor."""
    
    def __init__(self, root):
        """Initialize the GUI."""
        self.root = root
        self.root.title("🧾 Invoice Processor")
        self.root.geometry("900x700")
        
        # Configuration manager
        self.config = Config()
        
        # Processing state
        self.processing = False
        self.cancel_event = threading.Event()
        
        # Build the UI
        self._create_widgets()
        
        # Load saved settings
        self._load_settings()
    
    def _create_widgets(self):
        """Create all GUI widgets."""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text="🧾 Invoice Processor",
            font=('Arial', 20, 'bold')
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # API Key section
        ttk.Label(main_frame, text="OpenAI API Key:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        
        self.api_key_var = tk.StringVar()
        api_key_entry = ttk.Entry(main_frame, textvariable=self.api_key_var, width=50, show="*")
        api_key_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Show/Hide API key button
        self.show_api_key = False
        self.toggle_api_btn = ttk.Button(
            main_frame,
            text="Show",
            command=self._toggle_api_key_visibility,
            width=8
        )
        self.toggle_api_btn.grid(row=1, column=2, padx=(5, 0), pady=5)
        self.api_key_entry = api_key_entry
        
        # API Key help text
        api_help = ttk.Label(
            main_frame,
            text="Get your API key from: https://platform.openai.com/api-keys",
            font=('Arial', 8),
            foreground='gray'
        )
        api_help.grid(row=2, column=1, sticky=tk.W, padx=(5, 0))
        
        # Input folder section
        ttk.Label(main_frame, text="Input Folder:", font=('Arial', 10, 'bold')).grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        
        self.input_folder_var = tk.StringVar()
        input_entry = ttk.Entry(main_frame, textvariable=self.input_folder_var, width=50)
        input_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        input_btn = ttk.Button(
            main_frame,
            text="Browse...",
            command=self._browse_input_folder,
            width=12
        )
        input_btn.grid(row=3, column=2, padx=(5, 0), pady=5)
        
        # Output folder section
        ttk.Label(main_frame, text="Output Folder:", font=('Arial', 10, 'bold')).grid(
            row=4, column=0, sticky=tk.W, pady=5
        )
        
        self.output_folder_var = tk.StringVar()
        output_entry = ttk.Entry(main_frame, textvariable=self.output_folder_var, width=50)
        output_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        output_btn = ttk.Button(
            main_frame,
            text="Browse...",
            command=self._browse_output_folder,
            width=12
        )
        output_btn.grid(row=4, column=2, padx=(5, 0), pady=5)
        
        # Process button
        self.process_btn = ttk.Button(
            main_frame,
            text="🚀 Process Invoices",
            command=self._start_processing,
            style='Accent.TButton'
        )
        self.process_btn.grid(row=5, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate',
            length=300
        )
        self.progress_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Status log
        ttk.Label(main_frame, text="Status Log:", font=('Arial', 10, 'bold')).grid(
            row=7, column=0, sticky=tk.W, pady=(10, 5)
        )
        
        self.status_text = scrolledtext.ScrolledText(
            main_frame,
            height=15,
            width=80,
            wrap=tk.WORD,
            font=('Courier', 9)
        )
        self.status_text.grid(
            row=8, column=0, columnspan=3,
            sticky=(tk.W, tk.E, tk.N, tk.S),
            pady=(0, 10)
        )
        
        # Cancel button (initially hidden)
        self.cancel_btn = ttk.Button(
            main_frame,
            text="Cancel Processing",
            command=self._cancel_processing,
            state='disabled'
        )
        self.cancel_btn.grid(row=9, column=0, columnspan=3, pady=(0, 10))
        
        # Footer
        footer = ttk.Label(
            main_frame,
            text="Invoice Processor v1.0.0 | Powered by OpenAI GPT-4o Vision",
            font=('Arial', 8),
            foreground='gray'
        )
        footer.grid(row=10, column=0, columnspan=3, pady=(10, 0))
    
    def _toggle_api_key_visibility(self):
        """Toggle API key visibility."""
        self.show_api_key = not self.show_api_key
        if self.show_api_key:
            self.api_key_entry.config(show="")
            self.toggle_api_btn.config(text="Hide")
        else:
            self.api_key_entry.config(show="*")
            self.toggle_api_btn.config(text="Show")
    
    def _browse_input_folder(self):
        """Open folder browser for input folder."""
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_folder_var.set(folder)
    
    def _browse_output_folder(self):
        """Open folder browser for output folder."""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder_var.set(folder)
    
    def _load_settings(self):
        """Load saved settings from config."""
        api_key = self.config.get_api_key()
        input_folder = self.config.get_input_folder()
        output_folder = self.config.get_output_folder()
        
        if api_key:
            self.api_key_var.set(api_key)
        
        if input_folder:
            self.input_folder_var.set(input_folder)
        elif Path('INPUT').exists():
            # Default to INPUT folder in current directory
            self.input_folder_var.set(str(Path('INPUT').absolute()))
        
        if output_folder:
            self.output_folder_var.set(output_folder)
        elif Path('OUTPUT').exists():
            # Default to OUTPUT folder in current directory
            self.output_folder_var.set(str(Path('OUTPUT').absolute()))
    
    def _save_settings(self):
        """Save current settings to config."""
        self.config.save_all_settings(
            self.api_key_var.get(),
            self.input_folder_var.get(),
            self.output_folder_var.get()
        )
    
    def _log_status(self, message):
        """Add a message to the status log."""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def _update_progress(self, current, total):
        """Update the progress bar."""
        if total > 0:
            progress = (current / total) * 100
            self.progress_var.set(progress)
        self.root.update_idletasks()
    
    def _start_processing(self):
        """Start the invoice processing in a background thread."""
        # Validate inputs
        api_key = self.api_key_var.get().strip()
        input_folder = self.input_folder_var.get().strip()
        output_folder = self.output_folder_var.get().strip()
        
        if not api_key:
            messagebox.showerror("Error", "Please enter your OpenAI API key")
            return
        
        if not input_folder:
            messagebox.showerror("Error", "Please select an input folder")
            return
        
        if not output_folder:
            messagebox.showerror("Error", "Please select an output folder")
            return
        
        if not Path(input_folder).exists():
            messagebox.showerror("Error", f"Input folder does not exist: {input_folder}")
            return
        
        # Create output folder if it doesn't exist
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        
        # Save settings
        self._save_settings()
        
        # Clear previous status and reset progress
        self.status_text.delete(1.0, tk.END)
        self.progress_var.set(0)
        self.cancel_event.clear()
        
        # Disable process button and enable cancel button
        self.process_btn.config(state='disabled')
        self.cancel_btn.config(state='normal')
        self.processing = True
        
        # Start processing in background thread
        thread = threading.Thread(
            target=self._process_invoices,
            args=(api_key, input_folder, output_folder),
            daemon=True
        )
        thread.start()
    
    def _cancel_processing(self):
        """Cancel the current processing."""
        self.cancel_event.set()
        self._log_status("\n⚠️ Cancellation requested... waiting for current invoice to finish...")
        self.cancel_btn.config(state='disabled')
    
    def _process_invoices(self, api_key, input_folder, output_folder):
        """Process invoices (runs in background thread)."""
        try:
            self._log_status("=" * 80)
            self._log_status("🧾 Invoice Processor Starting...")
            self._log_status("=" * 80)
            self._log_status("")
            
            # Process all invoices
            results = invoice_processor.process_all_invoices(
                input_folder,
                api_key,
                progress_callback=self._update_progress,
                status_callback=self._log_status,
                cancel_event=self.cancel_event
            )
            
            if not results:
                self._log_status("\n⚠️ No invoices were processed")
                return
            
            # Generate Excel output
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = Path(output_folder) / f'invoices_{timestamp}.xlsx'
            
            self._log_status(f"\n📊 Generating Excel report...")
            invoice_processor.create_excel_report(
                results,
                str(output_file),
                status_callback=self._log_status
            )
            
            self._log_status("")
            self._log_status("=" * 80)
            self._log_status("✅ Processing Complete!")
            self._log_status("=" * 80)
            self._log_status(f"Output file: {output_file}")
            self._log_status("")
            
            # Show success message
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Success",
                    f"Processing complete!\n\nOutput saved to:\n{output_file}"
                )
            )
            
        except InterruptedError:
            self._log_status("\n❌ Processing cancelled by user")
            self.root.after(
                0,
                lambda: messagebox.showwarning("Cancelled", "Processing was cancelled")
            )
        
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self._log_status(f"\n❌ {error_msg}")
            self.root.after(
                0,
                lambda: messagebox.showerror("Error", error_msg)
            )
        
        finally:
            # Re-enable process button and disable cancel button
            self.processing = False
            self.root.after(0, self._reset_ui)
    
    def _reset_ui(self):
        """Reset UI state after processing."""
        self.process_btn.config(state='normal')
        self.cancel_btn.config(state='disabled')
        self.progress_var.set(0)


def main():
    """Main entry point for the GUI application."""
    # Try to load environment variables from .env file
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass
    
    # Create and run the GUI
    root = tk.Tk()
    app = InvoiceProcessorGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
