"""
Configuration Manager for Invoice Processor
Handles persistent settings and API key management.
"""

import os
import json
from pathlib import Path
from typing import Optional


class Config:
    """Manages application configuration and settings."""
    
    def __init__(self):
        """Initialize configuration manager."""
        # Determine config directory based on OS
        if os.name == 'nt':  # Windows
            config_dir = Path(os.getenv('APPDATA', '~')) / 'InvoiceProcessor'
        else:  # macOS/Linux
            config_dir = Path.home() / '.config' / 'InvoiceProcessor'
        
        self.config_dir = config_dir
        self.config_file = config_dir / 'config.json'
        
        # Create config directory if it doesn't exist
        self.config_dir.mkdir(parents=True, exist_ok=True)
        
        # Load existing config or create default
        self.settings = self._load_config()
    
    def _load_config(self) -> dict:
        """Load configuration from file or create default."""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            except Exception:
                pass
        
        # Default configuration
        return {
            'api_key': '',
            'input_folder': '',
            'output_folder': ''
        }
    
    def _save_config(self):
        """Save current configuration to file."""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save config: {e}")
    
    def get_api_key(self) -> str:
        """Get the stored API key."""
        return self.settings.get('api_key', '')
    
    def set_api_key(self, api_key: str):
        """Store the API key."""
        self.settings['api_key'] = api_key
        self._save_config()
    
    def get_input_folder(self) -> str:
        """Get the stored input folder path."""
        return self.settings.get('input_folder', '')
    
    def set_input_folder(self, folder: str):
        """Store the input folder path."""
        self.settings['input_folder'] = folder
        self._save_config()
    
    def get_output_folder(self) -> str:
        """Get the stored output folder path."""
        return self.settings.get('output_folder', '')
    
    def set_output_folder(self, folder: str):
        """Store the output folder path."""
        self.settings['output_folder'] = folder
        self._save_config()
    
    def save_all_settings(self, api_key: str, input_folder: str, output_folder: str):
        """Save all settings at once."""
        self.settings['api_key'] = api_key
        self.settings['input_folder'] = input_folder
        self.settings['output_folder'] = output_folder
        self._save_config()
    
    def clear_api_key(self):
        """Remove the stored API key."""
        self.settings['api_key'] = ''
        self._save_config()
