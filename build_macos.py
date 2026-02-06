"""
Build script for creating macOS application bundle of InvoiceProcessor
Requires PyInstaller to be installed: pip install pyinstaller
"""

import PyInstaller.__main__
import sys
from pathlib import Path

def build_macos_app():
    """Build macOS .app using PyInstaller."""
    
    print("=" * 60)
    print("Building InvoiceProcessor for macOS")
    print("=" * 60)
    print()
    
    # Get the directory containing this script
    script_dir = Path(__file__).parent
    
    # Define paths
    gui_script = script_dir / 'gui_processor.py'
    icon_path = script_dir / 'assets' / 'icon.icns'
    
    # Check if icon exists
    icon_arg = []
    if icon_path.exists():
        icon_arg = ['--icon', str(icon_path)]
        print(f"Using icon: {icon_path}")
    else:
        print("No icon found, building without custom icon")
    
    # PyInstaller arguments
    args = [
        str(gui_script),
        '--name=InvoiceProcessor',
        '--onefile',
        '--windowed',
        '--clean',
        '--osx-bundle-identifier=com.invoiceprocessor.app',
        # Hidden imports that PyInstaller might miss
        '--hidden-import=openpyxl',
        '--hidden-import=openpyxl.cell',
        '--hidden-import=openpyxl.styles',
        '--hidden-import=fitz',
        '--hidden-import=invoice_processor',
        '--hidden-import=config',
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.ttk',
        '--hidden-import=tkinter.filedialog',
        '--hidden-import=tkinter.scrolledtext',
    ] + icon_arg
    
    print()
    print("Running PyInstaller with the following configuration:")
    print(f"  Script: {gui_script}")
    print(f"  Output name: InvoiceProcessor.app")
    print(f"  Bundle ID: com.invoiceprocessor.app")
    print(f"  Mode: Windowed (no console)")
    print()
    
    try:
        PyInstaller.__main__.run(args)
        print()
        print("=" * 60)
        print("Build complete!")
        print("=" * 60)
        print()
        print("The application bundle can be found in the 'dist' folder:")
        print("  dist/InvoiceProcessor.app")
        print()
        print("You can now drag InvoiceProcessor.app to your Applications folder.")
        print()
    except Exception as e:
        print()
        print("=" * 60)
        print("Build failed!")
        print("=" * 60)
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    build_macos_app()
