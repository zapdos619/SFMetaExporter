"""
Salesforce connection and authentication handler
"""
from typing import List, Optional, Callable
from simple_salesforce import Salesforce
from config import API_VERSION


class SalesforceClient:
    """Handles Salesforce authentication and connection"""
    

    # salesforce_client.py - REPLACE __init__ method (lines ~28-58)

    def __init__(self, username: str, password: str, security_token: str, 
                domain: str = 'login', status_callback: Optional[Callable] = None):
        """
        Initialize Salesforce connection
        
        Args:
            username: Salesforce username
            password: Salesforce password
            security_token: Security token (can be empty if IP whitelisted)
            domain: Either 'login', 'test', or custom domain WITHOUT .salesforce.com suffix
                    Example: 'mycompany.my' (NOT 'mycompany.my.salesforce.com')
            status_callback: Optional callback for status updates
        """
        self.status_callback = status_callback
        self.all_org_objects: List[str] = []
        
        self._log_status("Initializing Salesforce Connection...")
        
        try:
            # ✅ CRITICAL: Detect domain type
            is_custom_domain = domain not in ['login', 'test']
            
            if is_custom_domain:
                # ✅ Custom domain - use directly (gui.py already cleaned it)
                # Domain should be like 'mycompany.my' (WITHOUT .salesforce.com)
                self._log_status(f"🌐 Using custom domain: {domain}")
                
                # ✅ REMOVED: Redundant suffix stripping
                # The suffix was already removed in gui.py - no need to do it again
                
                self.sf = Salesforce(
                    username=username,
                    password=password,
                    security_token=security_token,
                    domain=domain  # Pass cleaned domain directly
                )
            else:
                # ✅ Standard domain (login or test)
                org_type = "Production" if domain == 'login' else "Sandbox"
                self._log_status(f"🏢 Connecting to {org_type} org...")
                
                self.sf = Salesforce(
                    username=username,
                    password=password,
                    security_token=security_token,
                    domain=domain
                )
            
            # ✅ Connection successful - extract session info
            self.base_url = f"https://{self.sf.sf_instance}"
            self.session_id = self.sf.session_id
            self.api_version = self.sf.sf_version
            self.headers = {
                'Authorization': f'Bearer {self.session_id}',
                'Content-Type': 'application/json'
            }
            
            # ✅ Log connection details
            if is_custom_domain:
                self._log_status(f"✅ Connected to custom domain: {self.base_url}")
            else:
                self._log_status(f"✅ Connected to: {self.base_url}")
            
            # ✅ Fetch available objects
            self._fetch_all_org_objects()

        except Exception as e:
            # ✅ Let the exception propagate to gui.py with full details
            # gui.py's _infer_login_error will analyze the raw error
            error_msg = str(e)
            
            self._log_status(f"❌ Connection failed: {error_msg}")
            raise  # Re-raise with original error intact
 

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