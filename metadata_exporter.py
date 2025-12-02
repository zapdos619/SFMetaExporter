"""
Object metadata export functionality
"""
import csv
from typing import List, Dict, Tuple
from models import MetadataField
from salesforce_client import SalesforceClient
from field_usage_tracker import FieldUsageTracker


class MetadataExporter:
    """Handles object metadata export from Salesforce"""
    
    def __init__(self, sf_client: SalesforceClient):
        """Initialize with Salesforce client"""
        self.sf_client = sf_client
        self.sf = sf_client.sf
        # Field Usage Track
        self.usage_tracker = FieldUsageTracker(self.sf, sf_client.status_callback)
    
    def export_metadata(self, object_names: List[str], output_path: str) -> Tuple[str, Dict]:
        """Export metadata for specified objects"""
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
