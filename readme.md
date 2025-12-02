# Salesforce Picklist & Metadata Exporter

A Python GUI application for exporting Salesforce picklist values and object metadata.

## Features

- **Dual Export Functionality:**
  - Export picklist values with active/inactive status to Excel (.xlsx)
  - Export object metadata (fields, types, attributes) to CSV (.csv)
- **User-Friendly GUI:** Modern interface built with CustomTkinter
- **Object Selection:** Search, filter, and select multiple Salesforce objects
- **Real-time Status Updates:** Monitor export progress in real-time
- **Comprehensive Statistics:** Detailed export statistics and error reporting

## Project Structure

```
salesforce-exporter/
│
├── main.py                  # Application entry point
├── gui.py                   # Main GUI application
├── config.py                # Configuration constants
├── models.py                # Data models
├── salesforce_client.py     # Salesforce connection handler
├── picklist_exporter.py     # Picklist export logic
├── metadata_exporter.py     # Metadata export logic
├── utils.py                 # Utility functions
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Installation

### Prerequisites

- Python 3.7 or higher
- Salesforce account with API access
- Security Token (Settings → My Personal Information → Reset Security Token)

### Install Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Running the Application

```bash
python main.py
```

### Login

1. Enter your Salesforce credentials:
   - Username
   - Password
   - Security Token
2. Select org type (Production or Sandbox)
3. Click "Login to Salesforce"

### Exporting Data

1. **Select Objects:**
   - Use the search box to filter objects
   - Select one or more objects from "Available Objects"
   - Click ">> Add Selected >>" to add to export list
   - Or use "Select All" to select all visible objects

2. **Export Picklist Data:**
   - Click "Export Picklist Data" button
   - Choose save location
   - Excel file will contain: Object, Field Label, Field API, Picklist Value Label, Picklist Value API, Status

3. **Export Metadata:**
   - Click "Export Metadata" button
   - Choose save location
   - CSV file will contain: Object, Field Label, API Name, Type, Help Text, Formula, Attributes, Field Usage

### Logout

Click the "Logout" button in the top-right corner to disconnect and return to the login screen.

## File Descriptions

### `config.py`
Contains all configuration constants including API version, window settings, and default filenames.

### `models.py`
Defines data models:
- `FieldInfo`: Picklist field metadata
- `PicklistValueDetail`: Individual picklist value
- `ProcessingResult`: Object processing results
- `MetadataField`: Field metadata for export

### `salesforce_client.py`
Handles Salesforce authentication and connection:
- Establishes connection to Salesforce
- Fetches all queryable objects
- Manages session and headers

### `picklist_exporter.py`
Exports picklist values:
- Retrieves picklist fields from objects
- Queries picklist values using multiple fallback methods
- Creates formatted Excel output

### `metadata_exporter.py`
Exports object metadata:
- Retrieves all field metadata for objects
- Formats field types with details
- Determines field attributes
- Creates CSV output

### `utils.py`
Utility functions:
- Runtime formatting
- Statistics printing for console output

### `gui.py`
Main GUI application:
- Login screen
- Object selection interface
- Export controls
- Status monitoring

## Export Formats

### Picklist Export (Excel)

| Object | Field Label | Field API | Picklist Value Label | Picklist Value API | Status |
|--------|-------------|-----------|---------------------|-------------------|--------|
| Account | Type | Type | Prospect | Prospect | Active |
| Account | Type | Type | Customer | Customer | Active |

### Metadata Export (CSV)

| Object | Field Label | API Name | Type | Help Text | Formula | Attributes | Field Usage |
|--------|-------------|----------|------|-----------|---------|------------|-------------|
| Account | Account Name | Name | String (255) | | | Required | |
| Account | Account Owner | OwnerId | Lookup (User) | | | Required | |

## Troubleshooting

### Connection Issues
- Verify username, password, and security token
- Ensure correct org type (Production vs Sandbox)
- Check firewall settings

### Export Errors
- Some standard objects may have restricted access
- Custom objects require appropriate permissions
- Large exports may take time - do not close the window

### Empty Picklist Results
- Object may not have any picklist fields
- This is expected behavior, not an error

## Technical Details

### API Version
- Uses Salesforce API version 65.0
- Compatible with most Salesforce orgs

### Picklist Query Methods
The application uses multiple fallback methods to query picklist values:
1. FieldDefinition Tooling API
2. CustomField with EntityDefinition ID
3. CustomField with TableEnumOrId
4. REST Describe API

### Metadata Fields
Extracted metadata includes:
- Field labels and API names
- Field types with precision/length
- Help text and formulas
- Attributes (Required, Unique, External ID, etc.)

## Dependencies

- `requests`: HTTP library for API calls
- `simple-salesforce`: Salesforce REST API wrapper
- `openpyxl`: Excel file creation and formatting
- `customtkinter`: Modern GUI framework

## License

This project is provided as-is for educational and business purposes.

## Support

For issues or questions:
1. Check the status log in the application
2. Review console output for detailed error messages
3. Verify Salesforce permissions and API access

---

**Version:** 2.0  
**Last Updated:** October 2024

---
# Make executable File

Great! You can create a portable executable (.exe) file from your Python project using **PyInstaller**. This will bundle your Python application and all its dependencies into a single executable file that can run on Windows without requiring Python to be installed.

## Here's how to do it:

### 1. Install PyInstaller
First, add PyInstaller to your project:
```bash
pip install pyinstaller
```

### 2. Create the EXE
Navigate to your project directory and run:

**For a single file executable (recommended for portability):**
```bash
pyinstaller --onefile --windowed main.py
```

**For a folder with separate files (faster startup):**
```bash
pyinstaller --windowed main.py
```

### 3. Important flags for your project:

- `--onefile` - Packages everything into a single .exe file (more portable)
- `--windowed` or `-w` - Hides the console window (important for GUI apps with customtkinter)
- `--name YourAppName` - Give your exe a custom name
- `--icon=icon.ico` - Add a custom icon (if you have one)

### Complete command example:
```bash
pyinstaller --onefile --windowed --name "SF Meta Exporter" --icon=app_icon.ico main.py
```

### 4. Find your EXE
After PyInstaller finishes, you'll find your executable in:
- `dist/` folder - Your final .exe file will be here
- `build/` folder - Temporary build files (can be deleted)
- `.spec` file - Build configuration (can be modified for advanced options)

### 5. Test it
Copy the .exe from the `dist/` folder to another location or computer and test it to make sure it runs properly.

**Note:** The first run might be slow as it extracts files to a temp directory. Also, make sure any external files your app needs (config files, images, etc.) are either bundled or included alongside the .exe.

Would you like help with any specific issues or customizing the build process?
