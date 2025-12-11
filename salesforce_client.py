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
        """
        Initialize Salesforce connection
        
        Args:
            username: Salesforce username
            password: Salesforce password
            security_token: Security token (can be None if IP whitelisted)
            domain: Either 'login', 'test', or custom domain WITHOUT .salesforce.com suffix
                    Example: 'mycompany.my' (NOT 'mycompany.my.salesforce.com')
            status_callback: Optional callback for status updates
        """
        self.status_callback = status_callback
        self.all_org_objects: List[str] = []
        
        # ✅ FIX 1: Initialize these BEFORE connection attempt
        self.sf = None
        self.base_url = None
        self.session_id = None
        self.api_version = API_VERSION  # ✅ Use constant, not sf.sf_version
        self.headers = None
        
        self._log_status("Initializing Salesforce Connection...")
        
        try:
            # ✅ CRITICAL: Detect domain type
            is_custom_domain = domain not in ['login', 'test']
            
            if is_custom_domain:
                # ✅ Custom domain - use directly (gui.py already cleaned it)
                self._log_status(f"🌐 Using custom domain: {domain}")
                
                # ✅ FIX 2: Pass security_token correctly (None if not provided)
                self.sf = Salesforce(
                    username=username,
                    password=password,
                    security_token=security_token if security_token else None,
                    domain=domain
                )
            else:
                # ✅ Standard domain (login or test)
                org_type = "Production" if domain == 'login' else "Sandbox"
                self._log_status(f"🏢 Connecting to {org_type} org...")
                
                self.sf = Salesforce(
                    username=username,
                    password=password,
                    security_token=security_token if security_token else None,
                    domain=domain
                )
            
            # ✅ FIX 3: Verify connection was successful BEFORE extracting data
            if not self.sf:
                raise Exception("Salesforce connection object is None")
            
            if not hasattr(self.sf, 'session_id') or not self.sf.session_id:
                raise Exception("No session_id - authentication failed")
            
            # ✅ FIX 4: Extract session info with validation
            self.session_id = self.sf.session_id
            
            # Get instance URL
            if hasattr(self.sf, 'sf_instance'):
                self.base_url = f"https://{self.sf.sf_instance}"
            else:
                # Fallback for older simple-salesforce versions
                self.base_url = f"https://{domain}.salesforce.com"
            
            # ✅ FIX 5: Use API_VERSION constant (from config.py)
            # Don't rely on self.sf.sf_version as it may be formatted differently
            self.api_version = API_VERSION
            
            # Setup headers
            self.headers = {
                'Authorization': f'Bearer {self.session_id}',
                'Content-Type': 'application/json'
            }
            
            # ✅ Log connection details
            if is_custom_domain:
                self._log_status(f"✅ Connected to custom domain: {self.base_url}")
            else:
                self._log_status(f"✅ Connected to: {self.base_url}")
            
            self._log_status(f"📡 API Version: v{self.api_version}")
            self._log_status(f"🔑 Session ID: {self.session_id[:20]}...")
            
            # ✅ FIX 6: Fetch objects AFTER connection is fully initialized
            self._fetch_all_org_objects()
            
        except Exception as e:
            # ✅ Clean up on error
            self.sf = None
            self.all_org_objects = []
            
            error_msg = str(e)
            self._log_status(f"❌ Connection failed: {error_msg}")
            
            # Re-raise with original error intact
            raise

    def _fetch_all_org_objects(self):
        """
        Fetches all SObjects (Standard and Custom) from the org.
        
        ✅ FIXED: Better validation and error messages
        """
        self._log_status("Fetching all available SObjects from the organization...")
        
        # ✅ FIX 7: Verify we have a valid connection BEFORE calling describe()
        if not self.sf:
            raise Exception("Cannot fetch objects - Salesforce connection not initialized")
        
        if not self.session_id:
            raise Exception("Cannot fetch objects - No valid session")
        
        try:
            # ✅ Call describe() API
            self._log_status("📞 Calling sf.describe() API...")
            response = self.sf.describe()
            
            # ✅ Validate response structure
            if not response:
                self._log_status("⚠️ describe() returned empty response")
                raise Exception(
                    "Salesforce describe() returned empty response.\n\n"
                    "Possible causes:\n"
                    "• 'View All Data' permission is missing\n"
                    "• API access is disabled\n"
                    "• Session expired during connection"
                )
            
            if not isinstance(response, dict):
                raise TypeError(f"Invalid response type: {type(response)}")
            
            if 'sobjects' not in response:
                raise KeyError("Response missing 'sobjects' field")
            
            sobjects = response.get('sobjects', [])
            
            if not isinstance(sobjects, list):
                raise TypeError(f"Invalid sobjects type: {type(sobjects)}")
            
            # ✅ Log raw count for debugging
            self._log_status(f"📊 Total objects returned: {len(sobjects)}")
            
            # ✅ Filter queryable objects
            all_objects = [
                obj['name'] for obj in sobjects 
                if obj.get('queryable', False) and not obj.get('deprecatedAndHidden', False)
            ]
            
            self._log_status(f"🔍 Queryable objects after filtering: {len(all_objects)}")
            
            # ✅ FIX 8: Set all_org_objects even if empty (so GUI can show proper message)
            self.all_org_objects = sorted(all_objects)
            
            if not self.all_org_objects:
                # ⚠️ No objects found - log detailed info
                self._log_status("⚠️ No queryable objects found!")
                
                # Count different object types for diagnosis
                total = len(sobjects)
                queryable_count = sum(1 for obj in sobjects if obj.get('queryable', False))
                hidden_count = sum(1 for obj in sobjects if obj.get('deprecatedAndHidden', False))
                
                self._log_status(f"  Total objects: {total}")
                self._log_status(f"  Queryable: {queryable_count}")
                self._log_status(f"  Hidden/Deprecated: {hidden_count}")
                
                # Don't raise error - let GUI handle empty list
                self._log_status("⚠️ Continuing with empty object list...")
            else:
                # ✅ SUCCESS
                self._log_status(f"✅ Found {len(self.all_org_objects)} queryable objects.")
                
                # Log first 5 objects for verification
                sample = ', '.join(self.all_org_objects[:5])
                self._log_status(f"📦 Sample: {sample}...")
        
        except Exception as e:
            error_msg = str(e)
            self._log_status(f"❌ Failed to fetch SObjects: {error_msg}")
            
            # ✅ FIX 9: Set empty list instead of raising (GUI will show warning)
            self.all_org_objects = []
            
            # Log full error details
            import traceback
            self._log_status(f"🔍 Full error:\n{traceback.format_exc()}")
    
    def get_all_objects(self) -> List[str]:
        """Accessor for the fetched object list"""
        return self.all_org_objects
    
    def _log_status(self, message: str):
        """Internal helper to send log messages back to the GUI"""
        if self.status_callback:
            self.status_callback(message, verbose=True)