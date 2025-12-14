"""
Object metadata export functionality
"""
import csv
from typing import List, Dict, Tuple
from openpyxl import Workbook
import zipfile
import os
from datetime import datetime

from models import MetadataField
from salesforce_client import SalesforceClient
from field_usage_tracker import FieldUsageTracker
from excel_style_helper import ExcelStyleHelper  # ✅ NEW IMPORT


class MetadataExporter:
    """Handles object metadata export from Salesforce"""
    
    def __init__(self, sf_client: SalesforceClient):
        """Initialize with Salesforce client"""
        self.sf_client = sf_client
        self.sf = sf_client.sf
        # Field Usage Track
        self.usage_tracker = FieldUsageTracker(self.sf, sf_client.status_callback)
    
    def export_metadata(self, object_names: List[str], output_path: str) -> Tuple[str, Dict]:
        """
        Export metadata for specified objects
        
        ⚠️ LEGACY METHOD: This uses CSV format.
        New code should use export_metadata_excel() instead.
        
        Args:
            object_names: List of object API names
            output_path: Output CSV file path
            
        Returns:
            Tuple of (output_path, statistics_dict)
        """
        self._log_status("=== Starting Metadata Export ===")
        self._log_status(f"Total objects to process: {len(object_names)}")
        
        stats = {
            'total_objects': len(object_names),
            'successful_objects': 0,
            'failed_objects': 0,
            'total_fields': 0,
            'failed_object_details': []
        }
        
        all_metadata_fields: List[MetadataField] = []
        
        for i, obj_name in enumerate(object_names, 1):
            self._log_status(f"[{i}/{len(object_names)}] Processing object: {obj_name}")
            try:
                fields = self._get_object_metadata(obj_name)
                all_metadata_fields.extend(fields)
                stats['successful_objects'] += 1
                stats['total_fields'] += len(fields)
                self._log_status(f"  ✅ Retrieved {len(fields)} fields")
            except Exception as e:
                error_msg = str(e)
                self._log_status(f"  ❌ ERROR: {error_msg}")
                stats['failed_objects'] += 1
                stats['failed_object_details'].append({'name': obj_name, 'reason': error_msg})
            self._log_status("")
        
        self._log_status("=== Creating CSV File ===")
        final_output_path = self._create_csv_file(all_metadata_fields, output_path)
        return final_output_path, stats
    
  
    def export_metadata_excel(self, object_names: List[str], output_path: str, 
                            export_mode: str = "single_tab") -> Tuple[str, Dict]:
        """
        Export metadata for specified objects to Excel with styling
        
        Args:
            object_names: List of object API names
            output_path: Path for output file(s)
            export_mode: "single_tab", "multi_tab", or "individual_files"
            
        Returns:
            Tuple of (output_path, statistics_dict)
        """
        self._log_status("=== Starting Metadata Export (Excel Format) ===")
        self._log_status(f"Export Mode: {export_mode}")
        self._log_status(f"Total objects to process: {len(object_names)}")
        
        stats = {
            'total_objects': len(object_names),
            'successful_objects': 0,
            'failed_objects': 0,
            'total_fields': 0,
            'failed_object_details': [],
            'export_mode': export_mode
        }
        
        # Dispatch to appropriate export method based on mode
        if export_mode == "single_tab":
            return self._export_single_tab(object_names, output_path, stats)
        elif export_mode == "multi_tab":
            return self._export_multi_tab(object_names, output_path, stats)
        elif export_mode == "individual_files":
            return self._export_individual_files(object_names, output_path, stats)
        else:
            raise ValueError(f"Invalid export mode: {export_mode}")    
    
    
    def _export_single_tab(self, object_names: List[str], output_path: str, 
                        stats: Dict) -> Tuple[str, Dict]:
        """
        Export all objects to a single Excel sheet with styling
        
        Args:
            object_names: List of object API names
            output_path: Output Excel file path
            stats: Statistics dictionary
            
        Returns:
            Tuple of (output_path, stats)
        """
        self._log_status("📄 Single Tab Mode: All objects in one sheet")
        
        # Collect all metadata fields
        all_metadata_fields: List[MetadataField] = []
        
        for i, obj_name in enumerate(sorted(object_names), 1):  # ✅ Sort alphabetically
            self._log_status(f"[{i}/{len(object_names)}] Processing object: {obj_name}")
            try:
                fields = self._get_object_metadata(obj_name)
                all_metadata_fields.extend(fields)
                stats['successful_objects'] += 1
                stats['total_fields'] += len(fields)
                self._log_status(f"  ✅ Retrieved {len(fields)} fields")
            except Exception as e:
                error_msg = str(e)
                self._log_status(f"  ❌ ERROR: {error_msg}")
                stats['failed_objects'] += 1
                stats['failed_object_details'].append({'name': obj_name, 'reason': error_msg})
            self._log_status("")
        
        # Create Excel file with styling
        self._log_status("=== Creating Excel File ===")
        final_output_path = self._create_single_tab_excel(all_metadata_fields, output_path, stats)
        
        return final_output_path, stats    
    
    
    def _create_single_tab_excel(self, metadata_fields: List[MetadataField], 
                                output_path: str, stats: Dict) -> str:
        """
        Create Excel file with single tab and styling
        
        Args:
            metadata_fields: List of MetadataField objects
            output_path: Output file path
            stats: Statistics for header info
            
        Returns:
            Final output path
        """
        wb = Workbook()
        ws = wb.active
        ws.title = "Metadata Export"
        
        # Define headers (same as current CSV columns)
        headers = [
            'Object',
            'Field Label',
            'API Name',
            'Type',
            'Help Text',
            'Formula',
            'Attributes',
            'Field Usage'
        ]
        
        num_cols = len(headers)
        
        # ✅ Add Title Row (Row 1)
        ExcelStyleHelper.add_title_row(
            ws,
            title="Salesforce Metadata Export",
            num_cols=num_cols,
            row_num=1
        )
        
        # ✅ Add Info Row (Row 2)
        total_objects = stats['successful_objects']
        total_fields = len(metadata_fields)
        
        # For single tab, show summary info
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        info_cell = ws.cell(row=2, column=1)
        export_date = datetime.now().strftime('%Y-%m-%d %H:%M')
        info_cell.value = (
            f"Objects: {total_objects} | "
            f"Total Fields: {total_fields} | "
            f"Export Date: {export_date}"
        )
        
        # Apply info style
        info_style = ExcelStyleHelper.get_info_style()
        ExcelStyleHelper.apply_style_to_cell(info_cell, info_style)
        for col_num in range(2, num_cols + 1):
            cell = ws.cell(row=2, column=col_num)
            ExcelStyleHelper.apply_style_to_cell(cell, info_style)
        ws.row_dimensions[2].height = 20
        
        # ✅ Add Header Row (Row 3)
        ExcelStyleHelper.add_header_row(ws, headers, row_num=3)
        
        # ✅ Add Data Rows (Starting from Row 4)
        data_style = ExcelStyleHelper.get_data_style()
        
        for row_idx, field in enumerate(metadata_fields, start=4):
            row_data = field.to_row()
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                ExcelStyleHelper.apply_style_to_cell(cell, data_style)
        
        # ✅ Auto-adjust column widths
        ExcelStyleHelper.auto_adjust_column_widths(ws, headers)
        
        # ✅ Freeze header rows (title, info, headers)
        ExcelStyleHelper.freeze_header_rows(ws, num_rows=3)
        
        # Save workbook
        wb.save(output_path)
        
        self._log_status(f"✅ Excel file created: {output_path}")
        self._log_status(f"✅ Total fields exported: {len(metadata_fields)}")
        
        return output_path   
    
 
    def _export_multi_tab(self, object_names: List[str], output_path: str, 
                        stats: Dict) -> Tuple[str, Dict]:
        """
        Export each object to its own sheet in a single Excel file
        
        Args:
            object_names: List of object API names
            output_path: Output Excel file path
            stats: Statistics dictionary
            
        Returns:
            Tuple of (output_path, stats)
        """
        self._log_status("📑 Multiple Tabs Mode: One sheet per object")
        
        # Create workbook
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        # Sort objects alphabetically
        sorted_objects = sorted(object_names)
        
        # Process each object and create a sheet
        for i, obj_name in enumerate(sorted_objects, 1):
            self._log_status(f"[{i}/{len(sorted_objects)}] Processing object: {obj_name}")
            
            try:
                fields = self._get_object_metadata(obj_name)
                stats['successful_objects'] += 1
                stats['total_fields'] += len(fields)
                
                # ✅ Get object label for tab name
                object_label = ExcelStyleHelper.get_object_label(self.sf_client, obj_name)
                
                # ✅ Create sheet for this object
                self._create_sheet_for_object(
                    wb,
                    obj_name,
                    object_label,
                    fields
                )
                
                self._log_status(f"  ✅ Retrieved {len(fields)} fields")
                
            except Exception as e:
                error_msg = str(e)
                self._log_status(f"  ❌ ERROR: {error_msg}")
                stats['failed_objects'] += 1
                stats['failed_object_details'].append({'name': obj_name, 'reason': error_msg})
            
            self._log_status("")
        
        # Save workbook
        if len(wb.worksheets) == 0:
            # No sheets created - create a placeholder
            ws = wb.create_sheet("No Data")
            ws['A1'] = "No metadata found for selected objects"
            self._log_status("⚠️ No data exported - created placeholder sheet")
        
        wb.save(output_path)
        
        self._log_status(f"✅ Excel file created: {output_path}")
        self._log_status(f"✅ Total sheets: {len(wb.worksheets)}")
        
        return output_path, stats    
    
    
    def _create_sheet_for_object(self, wb: Workbook, object_api: str, 
                                object_label: str, fields: List[MetadataField]):
        """
        Create a single sheet for an object with styling
        
        Args:
            wb: Workbook instance
            object_api: Object API name
            object_label: Object label (for display)
            fields: List of MetadataField objects for this object
        """
        # ✅ Sanitize sheet name (Excel 31-char limit, no special chars)
        sheet_name = ExcelStyleHelper.sanitize_sheet_name(object_label)
        
        # Create sheet
        ws = wb.create_sheet(sheet_name)
        
        # Define headers (same as current CSV columns)
        headers = [
            'Object',
            'Field Label',
            'API Name',
            'Type',
            'Help Text',
            'Formula',
            'Attributes',
            'Field Usage'
        ]
        
        num_cols = len(headers)
        
        # ✅ Add Title Row (Row 1)
        ExcelStyleHelper.add_title_row(
            ws,
            title=f"Salesforce Metadata Export - {object_label}",
            num_cols=num_cols,
            row_num=1
        )
        
        # ✅ Add Info Row (Row 2) with object details
        ExcelStyleHelper.add_info_row(
            ws,
            object_label=object_label,
            object_api=object_api,
            record_count=len(fields),
            num_cols=num_cols,
            row_num=2
        )
        
        # ✅ Add Header Row (Row 3)
        ExcelStyleHelper.add_header_row(ws, headers, row_num=3)
        
        # ✅ Add Data Rows (Starting from Row 4)
        data_style = ExcelStyleHelper.get_data_style()
        
        for row_idx, field in enumerate(fields, start=4):
            row_data = field.to_row()
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                ExcelStyleHelper.apply_style_to_cell(cell, data_style)
        
        # ✅ Auto-adjust column widths
        ExcelStyleHelper.auto_adjust_column_widths(ws, headers)
        
        # ✅ Freeze header rows (title, info, headers)
        ExcelStyleHelper.freeze_header_rows(ws, num_rows=3)   
    

    def _export_individual_files(self, object_names: List[str], output_path: str, 
                                stats: Dict) -> Tuple[str, Dict]:
        """
        Export each object to its own Excel file, auto-zip if multiple objects
        
        Args:
            object_names: List of object API names
            output_path: Base output path (will be modified for zip if needed)
            stats: Statistics dictionary
            
        Returns:
            Tuple of (final_output_path, stats)
                - If 1 object: returns .xlsx path
                - If 2+ objects: returns .zip path
        """
        self._log_status("📦 Individual Files Mode: Separate .xlsx per object")
        
        # Sort objects alphabetically
        sorted_objects = sorted(object_names)
        
        # Determine output strategy
        export_type = "Metadata"  # Will be used in filenames
        
        # Get base directory and filename
        output_dir = os.path.dirname(output_path)
        base_filename = os.path.splitext(os.path.basename(output_path))[0]
        
        # List to store created file paths
        created_files = []
        
        # Process each object
        for i, obj_name in enumerate(sorted_objects, 1):
            self._log_status(f"[{i}/{len(sorted_objects)}] Processing object: {obj_name}")
            
            try:
                fields = self._get_object_metadata(obj_name)
                stats['successful_objects'] += 1
                stats['total_fields'] += len(fields)
                
                # ✅ Get object label
                object_label = ExcelStyleHelper.get_object_label(self.sf_client, obj_name)
                
                # ✅ Create individual Excel file
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"{obj_name}_{export_type}_{timestamp}.xlsx"
                file_path = os.path.join(output_dir, filename)
                
                self._create_individual_excel_file(
                    file_path,
                    obj_name,
                    object_label,
                    fields
                )
                
                created_files.append(file_path)
                
                self._log_status(
                    f"  ✅ Created: {filename} | Fields: {len(fields)}"
                )
                
            except Exception as e:
                error_msg = str(e)
                self._log_status(f"  ❌ ERROR: {error_msg}")
                stats['failed_objects'] += 1
                stats['failed_object_details'].append({'name': obj_name, 'reason': error_msg})
            
            self._log_status("")
        
        # ✅ Decide: Single file or ZIP?
        if len(created_files) == 0:
            # No files created
            self._log_status("⚠️ No files created - no data to export")
            raise Exception("No metadata found for any selected objects")
            
        elif len(created_files) == 1:
            # ✅ Single file - return as-is (no zip)
            final_path = created_files[0]
            self._log_status(f"✅ Single file created: {final_path}")
            return final_path, stats
            
        else:
            # ✅ Multiple files - create ZIP
            self._log_status(f"📦 Creating ZIP archive for {len(created_files)} files...")
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"Salesforce_{export_type}_Export_{timestamp}.zip"
            zip_path = os.path.join(output_dir, zip_filename)
            
            final_path = self._create_zip_archive(created_files, zip_path)
            
            return final_path, stats
    

    def _create_individual_excel_file(self, file_path: str, object_api: str, 
                                    object_label: str, fields: List[MetadataField]):
        """
        Create a single Excel file for one object with styling
        
        Args:
            file_path: Full path for output Excel file
            object_api: Object API name
            object_label: Object label (for display)
            fields: List of MetadataField objects for this object
        """
        # Create new workbook
        wb = Workbook()
        ws = wb.active
        ws.title = ExcelStyleHelper.sanitize_sheet_name(object_label)
        
        # Define headers (same as current CSV columns)
        headers = [
            'Object',
            'Field Label',
            'API Name',
            'Type',
            'Help Text',
            'Formula',
            'Attributes',
            'Field Usage'
        ]
        
        num_cols = len(headers)
        
        # ✅ Add Title Row (Row 1)
        ExcelStyleHelper.add_title_row(
            ws,
            title=f"Salesforce Metadata Export - {object_label}",
            num_cols=num_cols,
            row_num=1
        )
        
        # ✅ Add Info Row (Row 2) with object details
        ExcelStyleHelper.add_info_row(
            ws,
            object_label=object_label,
            object_api=object_api,
            record_count=len(fields),
            num_cols=num_cols,
            row_num=2
        )
        
        # ✅ Add Header Row (Row 3)
        ExcelStyleHelper.add_header_row(ws, headers, row_num=3)
        
        # ✅ Add Data Rows (Starting from Row 4)
        data_style = ExcelStyleHelper.get_data_style()
        
        for row_idx, field in enumerate(fields, start=4):
            row_data = field.to_row()
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                ExcelStyleHelper.apply_style_to_cell(cell, data_style)
        
        # ✅ Auto-adjust column widths
        ExcelStyleHelper.auto_adjust_column_widths(ws, headers)
        
        # ✅ Freeze header rows (title, info, headers)
        ExcelStyleHelper.freeze_header_rows(ws, num_rows=3)
        
        # Save workbook
        wb.save(file_path)   
        
    
    def _create_zip_archive(self, file_paths: List[str], zip_path: str) -> str:
        """
        Create ZIP archive containing multiple Excel files
        
        Args:
            file_paths: List of file paths to include in ZIP
            zip_path: Output ZIP file path
            
        Returns:
            Path to created ZIP file
        """
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in file_paths:
                    # Add file to ZIP with just the filename (no directory structure)
                    arcname = os.path.basename(file_path)
                    zipf.write(file_path, arcname=arcname)
                    self._log_status(f"  📄 Added to ZIP: {arcname}")
            
            # ✅ Delete individual files after zipping
            self._log_status("🧹 Cleaning up individual files...")
            for file_path in file_paths:
                try:
                    os.remove(file_path)
                except Exception as e:
                    self._log_status(f"  ⚠️ Could not delete {file_path}: {e}")
            
            self._log_status(f"✅ ZIP archive created: {zip_path}")
            self._log_status(f"✅ Total files in ZIP: {len(file_paths)}")
            
            return zip_path
            
        except Exception as e:
            self._log_status(f"❌ Error creating ZIP: {str(e)}")
            raise   
        
    
    
    
    
    
    def _get_object_metadata(self, object_name: str) -> List[MetadataField]:
        """Get metadata for all fields of an object"""
        metadata_fields = []
        
        try:
            obj_describe = getattr(self.sf, object_name).describe()
            
            for field in obj_describe['fields']:
                # Extract field information
                field_label = field.get('label', '')
                api_name = field.get('name', '')
                field_type = self._format_field_type(field)
                help_text = field.get('inlineHelpText', '')
                formula = field.get('calculatedFormula', '')
                
                # Determine attributes
                attributes = self._get_field_attributes(field)
                
                # Field usage would require additional queries to find references
                # For now, we'll leave it empty or you can implement usage tracking
                # field_usage = ""
                # GET FIELD USAGE
                field_usage = self.usage_tracker.get_field_usage(object_name, api_name)
                
                metadata_field = MetadataField(
                    object_name=object_name,
                    field_label=field_label,
                    api_name=api_name,
                    field_type=field_type,
                    help_text=help_text,
                    formula=formula,
                    attributes=attributes,
                    field_usage=field_usage
                )
                
                metadata_fields.append(metadata_field)
        
        except Exception as e:
            self._log_status(f"  ERROR in _get_object_metadata: {str(e)}")
            raise
        
        return metadata_fields
    
    def _format_field_type(self, field: dict) -> str:
        """Format field type with additional details like length, precision"""
        field_type = field.get('type', '')
        
        # Add length for text fields
        if field_type in ['string', 'textarea', 'url', 'email', 'phone']:
            length = field.get('length', 0)
            if length:
                return f"{field_type.capitalize()} ({length})"
        
        # Add precision and scale for number/currency fields
        elif field_type in ['double', 'currency', 'percent']:
            precision = field.get('precision', 0)
            scale = field.get('scale', 0)
            if precision:
                return f"Number ({precision}, {scale})"
        
        # Add reference type for lookups
        elif field_type == 'reference':
            ref_to = field.get('referenceTo', [])
            if ref_to:
                ref_str = ', '.join(ref_to)
                return f"Lookup ({ref_str})"
        
        # Add picklist values count
        elif field_type in ['picklist', 'multipicklist']:
            picklist_values = field.get('picklistValues', [])
            values_preview = '; '.join([pv.get('value', '') for pv in picklist_values[:3]])
            if len(picklist_values) > 3:
                values_preview += f"; ...and {len(picklist_values) - 3} more"
            return f"{field_type.capitalize()} ({values_preview})"
        
        return field_type.capitalize()
    
    def _get_field_attributes(self, field: dict) -> str:
        """Get field attributes like Required, Unique, External ID, etc."""
        attributes = []
        
        if not field.get('nillable', True) and not field.get('defaultedOnCreate', False):
            attributes.append('Required')
        
        if field.get('unique', False):
            attributes.append('Unique')
        
        if field.get('externalId', False):
            attributes.append('External ID')
        
        if field.get('autoNumber', False):
            attributes.append('Auto Number')
        
        if field.get('calculated', False):
            attributes.append('Formula')
        
        if field.get('cascadeDelete', False):
            attributes.append('Cascade Delete')
        
        if field.get('restrictedPicklist', False):
            attributes.append('Restricted Picklist')
        
        if field.get('encrypted', False):
            attributes.append('Encrypted')
        
        return ', '.join(attributes) if attributes else ''
    
    def _create_csv_file(self, metadata_fields: List[MetadataField], output_path: str) -> str:
        """Create CSV file with metadata"""
        headers = [
            'Object', 'Field Label', 'API Name', 'Type', 
            'Help Text', 'Formula', 'Attributes', 'Field Usage'
        ]
        
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            
            for field in metadata_fields:
                writer.writerow(field.to_row())
        
        self._log_status(f"✅ CSV file created: {output_path}")
        self._log_status(f"✅ Total fields exported: {len(metadata_fields)}")
        return output_path
    
    def _log_status(self, message: str):
        """Log status message"""
        if self.sf_client.status_callback:
            self.sf_client.status_callback(message, verbose=True)
