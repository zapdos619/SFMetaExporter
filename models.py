"""
Data models for Salesforce field and picklist information
"""
from typing import List, Dict, Optional


class FieldInfo:
    """Represents picklist field metadata"""
    def __init__(self, api_name: str, label: str):
        self.api_name = api_name
        self.label = label


class PicklistValueDetail:
    """Represents a single picklist value"""
    def __init__(self, label: str, value: str, is_active: bool = True):
        self.label = label
        self.value = value
        self.is_active = is_active


class ProcessingResult:
    """Stores processing results for an object"""
    def __init__(self):
        self.values_processed = 0
        self.inactive_values = 0
        self.rows: List[List[str]] = []
        self.picklist_fields_count = 0
        self.object_exists = True
        self.error_message = None


class MetadataField:
    """Represents metadata for a single field"""
    def __init__(self, object_name: str, field_label: str, api_name: str, 
                 field_type: str, help_text: str = "", formula: str = "", 
                 attributes: str = "", field_usage: str = ""):
        self.object_name = object_name
        self.field_label = field_label
        self.api_name = api_name
        self.field_type = field_type
        self.help_text = help_text
        self.formula = formula
        self.attributes = attributes
        self.field_usage = field_usage
    
    def to_row(self) -> List[str]:
        """Convert to CSV row format"""
        return [
            self.object_name,
            self.field_label,
            self.api_name,
            self.field_type,
            self.help_text,
            self.formula,
            self.attributes,
            self.field_usage
        ]
