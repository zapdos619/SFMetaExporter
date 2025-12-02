"""
Configuration constants for Salesforce Picklist Exporter
"""
import os

# API Configuration
API_VERSION = '65.0'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# GUI Configuration
WINDOW_TITLE = "Salesforce Metadata Exporter"
WINDOW_GEOMETRY = "1200x800"
APPEARANCE_MODE = "System"
COLOR_THEME = "blue"

# Export Configuration
DEFAULT_PICKLIST_FILENAME = 'Picklist_Export_{timestamp}.xlsx'
DEFAULT_METADATA_FILENAME = 'Object_Metadata_{timestamp}.csv'
DEFAULT_CONTENTDOCUMENT_FILENAME = 'ContentDocument_Export_{timestamp}.csv'