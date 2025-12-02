"""
Salesforce connection and authentication handler
"""
from typing import List, Optional, Callable
from simple_salesforce import Salesforce
from config import API_VERSION


class SalesforceClient:
    """Handles Salesforce authentication and connection"""
    
    def __init__(self, username: str, password: str, security_token: str, 
                 domain: str = 'login', status_callback: Optional[Callable] = None):
        """Initialize Salesforce connection"""
        self.status_callback = status_callback
        self.all_org_objects: List[str] = []
        
        self._log_status("Initializing Salesforce Connection...")
        
        try:
            self.sf = Salesforce(
                username=username,
                password=password,
                security_token=security_token,
                domain=domain
            )
            self.base_url = f"https://{self.sf.sf_instance}"
            self.session_id = self.sf.session_id
            self.api_version = self.sf.sf_version
            self.headers = {
                'Authorization': f'Bearer {self.session_id}',
                'Content-Type': 'application/json'
            }
            self._log_status(f"✅ Connected to: {self.base_url}")
            
            self._fetch_all_org_objects()

        except Exception as e:
            self._log_status(f"❌ Connection failed: {str(e)}")
            raise

    def _fetch_all_org_objects(self):
        """Fetches all SObjects (Standard and Custom) from the org"""
        self._log_status("Fetching all available SObjects from the organization...")
        try:
            response = self.sf.describe()
            self.all_org_objects = sorted([
                obj['name'] for obj in response['sobjects'] 
                if obj.get('queryable', False) and not obj.get('deprecatedAndHidden', False)
            ])
            self._log_status(f"✅ Found {len(self.all_org_objects)} queryable objects.")
        except Exception as e:
            self._log_status(f"❌ Failed to fetch all SObjects: {str(e)}")
            self.all_org_objects = []
    
    def get_all_objects(self) -> List[str]:
        """Accessor for the fetched object list"""
        return self.all_org_objects
    
    def _log_status(self, message: str):
        """Internal helper to send log messages back to the GUI"""
        if self.status_callback:
            self.status_callback(message, verbose=True)