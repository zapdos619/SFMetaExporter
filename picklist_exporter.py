"""
Picklist export functionality
"""
import requests
from typing import List, Dict, Optional, Tuple
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from config import API_VERSION
from models import FieldInfo, PicklistValueDetail, ProcessingResult
from salesforce_client import SalesforceClient


class PicklistExporter:
    """Handles picklist data export from Salesforce"""
    
    def __init__(self, sf_client: SalesforceClient):
        """Initialize with Salesforce client"""
        self.sf_client = sf_client
        self.sf = sf_client.sf
        self.base_url = sf_client.base_url
        self.headers = sf_client.headers
        self.api_version = sf_client.api_version
    
    def export_picklists(self, object_names: List[str], output_path: str) -> Tuple[str, Dict]:
        """Export picklist values for specified objects"""
        self._log_status("=== Starting Picklist Export ===")
        self._log_status(f"Total objects to process: {len(object_names)}")
        
        stats = {
            'total_objects': len(object_names), 'successful_objects': 0, 'failed_objects': 0, 
            'objects_not_found': 0, 'objects_with_zero_picklists': 0, 'objects_with_picklists': 0, 
            'total_picklist_fields': 0, 'total_values': 0, 'total_active_values': 0, 
            'total_inactive_values': 0, 'failed_object_details': [], 'objects_without_picklists': [], 
            'objects_not_found_list': []
        }
        
        all_rows = [['Object', 'Field Label', 'Field API', 'Picklist Value Label', 'Picklist Value API', 'Status']]
        
        for i, obj_name in enumerate(object_names, 1):
            self._log_status(f"[{i}/{len(object_names)}] Processing object: {obj_name}")
            try:
                result = self._process_object(obj_name)
                
                if not result.object_exists:
                    stats['objects_not_found'] += 1
                    stats['objects_not_found_list'].append(obj_name)
                    stats['failed_object_details'].append({'name': obj_name, 'reason': 'Object does not exist in org'})
                    self._log_status(f"  ⚠️  Object not found in org")
                elif result.picklist_fields_count == 0:
                    stats['objects_with_zero_picklists'] += 1
                    stats['objects_without_picklists'].append(obj_name)
                    stats['successful_objects'] += 1
                    self._log_status(f"  ℹ️  No picklist fields found")
                else:
                    stats['objects_with_picklists'] += 1
                    stats['successful_objects'] += 1
                    stats['total_picklist_fields'] += result.picklist_fields_count
                    all_rows.extend(result.rows)
                    stats['total_values'] += result.values_processed
                    stats['total_inactive_values'] += result.inactive_values
                    stats['total_active_values'] += (result.values_processed - result.inactive_values)
                    self._log_status(f"  ✅ Fields: {result.picklist_fields_count}, Active: {result.values_processed - result.inactive_values}, Inactive: {result.inactive_values}")
            except Exception as e:
                error_msg = str(e)
                self._log_status(f"  ❌ ERROR: {error_msg}")
                stats['failed_objects'] += 1
                stats['failed_object_details'].append({'name': obj_name, 'reason': error_msg})
            self._log_status("")
        
        self._log_status("=== Creating Excel File ===")
        final_output_path = self._create_excel_file(all_rows, output_path)
        return final_output_path, stats
    
    def _process_object(self, obj_name: str) -> ProcessingResult:
        """Process a single object for picklist fields"""
        result = ProcessingResult()
        try:
            getattr(self.sf, obj_name).describe()
        except Exception as e:
            if 'NOT_FOUND' in str(e) or 'INVALID_TYPE' in str(e):
                result.object_exists = False
                return result
            raise
        
        picklist_fields = self._get_picklist_fields(obj_name)
        result.picklist_fields_count = len(picklist_fields)
        if not picklist_fields:
            return result
        
        self._log_status(f"  Found {len(picklist_fields)} picklist fields")
        entity_def_id = self._resolve_entity_definition_id(obj_name)
        if entity_def_id:
            self._log_status(f"  EntityDefinition.Id: {entity_def_id}")
        
        for field_api, field_info in picklist_fields.items():
            values = self._query_picklist_values_with_fallback(obj_name, entity_def_id, field_api)
            if not values:
                continue
            
            self._log_status(f"    Field: {field_api} - {len(values)} values")
            for value in values:
                is_active = value.is_active if value.is_active is not None else True
                status = 'Active' if is_active else 'Inactive'
                if not is_active:
                    result.inactive_values += 1
                
                row = [obj_name, field_info.label, field_api, value.label, value.value, status]
                result.rows.append(row)
                result.values_processed += 1
        
        return result
    
    def _get_picklist_fields(self, object_name: str) -> Dict[str, FieldInfo]:
        """Get all picklist fields for an object"""
        fields_dict = {}
        try:
            obj_describe = getattr(self.sf, object_name).describe()
            for field in obj_describe['fields']:
                if field['type'] in ['picklist', 'multipicklist']:
                    fields_dict[field['name']] = FieldInfo(
                        api_name=field['name'], 
                        label=field['label']
                    )
        except Exception as e:
            self._log_status(f"  ERROR in _get_picklist_fields: {str(e)}")
        return fields_dict
    
    def _resolve_entity_definition_id(self, object_name: str) -> Optional[str]:
        """Resolve EntityDefinition ID for an object"""
        try:
            query = f"SELECT Id FROM EntityDefinition WHERE QualifiedApiName = '{object_name}'"
            url = f"{self.base_url}/services/data/v{self.api_version}/tooling/query/"
            response = requests.get(url, headers=self.headers, params={'q': query}, timeout=60)
            if response.status_code == 200:
                records = response.json().get('records', [])
                if records:
                    return records[0]['Id']
        except Exception as e:
            self._log_status(f"  ERROR resolveEntityDefinitionId: {str(e)}")
        return None
    
    def _query_picklist_values_with_fallback(self, object_name: str, entity_def_id: Optional[str], 
                                            field_name: str) -> List[PicklistValueDetail]:
        """Query picklist values using multiple fallback methods"""
        values = self._query_field_definition_tooling(object_name, field_name)
        if values:
            return values
        
        if entity_def_id:
            values = self._query_custom_field_tooling(entity_def_id, field_name)
            if values:
                return values
        
        values = self._query_custom_field_tooling_table_enum(object_name, field_name)
        if values:
            return values
        
        values = self._query_rest_describe_for_picklist(object_name, field_name)
        if values:
            return values
        
        return []
    
    def _query_field_definition_tooling(self, object_name: str, field_name: str) -> List[PicklistValueDetail]:
        """Query using FieldDefinition"""
        try:
            query = f"SELECT Metadata FROM FieldDefinition WHERE EntityDefinition.QualifiedApiName = '{object_name}' AND QualifiedApiName = '{field_name}'"
            url = f"{self.base_url}/services/data/v{self.api_version}/tooling/query/"
            response = requests.get(url, headers=self.headers, params={'q': query}, timeout=60)
            if response.status_code == 200:
                records = response.json().get('records', [])
                if records:
                    return self._parse_value_set(records[0].get('Metadata', {}))
        except Exception as e:
            self._log_status(f"      ERROR queryFieldDefinitionTooling: {str(e)}")
        return []
    
    def _query_custom_field_tooling(self, entity_def_id: str, field_name: str) -> List[PicklistValueDetail]:
        """Query using CustomField with entity ID"""
        try:
            dev_name = field_name[:-3] if field_name.endswith('__c') else field_name
            query = f"SELECT Metadata FROM CustomField WHERE TableEnumOrId = '{entity_def_id}' AND DeveloperName = '{dev_name}'"
            url = f"{self.base_url}/services/data/v{self.api_version}/tooling/query/"
            response = requests.get(url, headers=self.headers, params={'q': query}, timeout=60)
            if response.status_code == 200:
                records = response.json().get('records', [])
                if records:
                    return self._parse_value_set(records[0].get('Metadata', {}))
        except Exception as e:
            self._log_status(f"      ERROR queryCustomFieldTooling: {str(e)}")
        return []
    
    def _query_custom_field_tooling_table_enum(self, object_name: str, field_name: str) -> List[PicklistValueDetail]:
        """Query using CustomField with object name"""
        try:
            dev_name = field_name[:-3] if field_name.endswith('__c') else field_name
            query = f"SELECT Metadata FROM CustomField WHERE TableEnumOrId = '{object_name}' AND DeveloperName = '{dev_name}'"
            url = f"{self.base_url}/services/data/v{self.api_version}/tooling/query/"
            response = requests.get(url, headers=self.headers, params={'q': query}, timeout=60)
            if response.status_code == 200:
                records = response.json().get('records', [])
                if records:
                    return self._parse_value_set(records[0].get('Metadata', {}))
        except Exception as e:
            self._log_status(f"      ERROR queryCustomFieldToolingTableEnum: {str(e)}")
        return []
    
    def _query_rest_describe_for_picklist(self, object_name: str, field_name: str) -> List[PicklistValueDetail]:
        """Query using REST describe endpoint"""
        try:
            url = f"{self.base_url}/services/data/v{self.api_version}/sobjects/{object_name}/describe"
            response = requests.get(url, headers=self.headers, timeout=60)
            if response.status_code == 200:
                for field in response.json().get('fields', []):
                    if field['name'].lower() == field_name.lower():
                        return [
                            PicklistValueDetail(
                                label=pv.get('label', ''), 
                                value=pv.get('value', ''), 
                                is_active=pv.get('active', True)
                            ) for pv in field.get('picklistValues', [])
                        ]
        except Exception as e:
            self._log_status(f"      ERROR queryRestDescribeForPicklist: {str(e)}")
        return []
    
    def _parse_value_set(self, metadata: dict) -> List[PicklistValueDetail]:
        """Parse picklist values from metadata"""
        results = []
        try:
            value_set = metadata.get('valueSet', {})
            if not value_set:
                return results
            
            values = value_set.get('valueSetDefinition', {}).get('value', []) or value_set.get('value', [])
            for v in values:
                # is_active = bool(v.get('isActive', True))
                is_active_raw = v.get('isActive')
                if is_active_raw is None:
                    is_active = True  # Default for missing key
                else:
                    is_active = bool(is_active_raw)  # Convert any truthy/falsy value
                results.append(PicklistValueDetail(
                    label=v.get('label', ''), 
                    value=v.get('valueName') or v.get('value', ''), 
                    is_active=is_active
                ))
        except Exception as e:
            self._log_status(f"      ERROR parseValueSet: {str(e)}")
        return results
    
    def _create_excel_file(self, rows: List[List[str]], output_path: str) -> str:
        """Create Excel file with formatted data"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Picklist Export"
        
        for row in rows:
            ws.append(row)
        
        # Format header
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_length + 2, 50)
        
        ws.freeze_panes = "A2"
        wb.save(output_path)
        
        self._log_status(f"✅ Excel file created: {output_path}")
        self._log_status(f"✅ Total data rows: {len(rows) - 1}")
        return output_path
    
    def _log_status(self, message: str):
        """Log status message"""
        if self.sf_client.status_callback:
            self.sf_client.status_callback(message, verbose=True)
