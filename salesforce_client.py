"""
Salesforce connection and authentication handler
✅ FIXED: Detects expired passwords and shows clear error messages
"""
from typing import List, Optional, Callable
from simple_salesforce import Salesforce
from simple_salesforce.exceptions import SalesforceExpiredSession
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
        
        # ✅ Initialize these BEFORE connection attempt
        self.sf = None
        self.base_url = None
        self.session_id = None
        self.api_version = API_VERSION
        self.headers = None
        
        self._log_status("Initializing Salesforce Connection...")
        
        try:
            # ✅ Detect domain type
            is_custom_domain = domain not in ['login', 'test']
            
            if is_custom_domain:
                self._log_status(f"🌐 Using custom domain: {domain}")
                
                self.sf = Salesforce(
                    username=username,
                    password=password,
                    security_token=security_token if security_token else None,
                    domain=domain
                )
            else:
                org_type = "Production" if domain == 'login' else "Sandbox"
                self._log_status(f"🏢 Connecting to {org_type} org...")
                
                self.sf = Salesforce(
                    username=username,
                    password=password,
                    security_token=security_token if security_token else None,
                    domain=domain
                )
            
            # ✅ Verify connection was successful
            if not self.sf:
                raise Exception("Salesforce connection object is None")
            
            if not hasattr(self.sf, 'session_id') or not self.sf.session_id:
                raise Exception("No session_id - authentication failed")
            
            # ✅ Extract session info with validation
            self.session_id = self.sf.session_id
            
            # Get instance URL
            if hasattr(self.sf, 'sf_instance'):
                self.base_url = f"https://{self.sf.sf_instance}"
            else:
                # Fallback for older simple-salesforce versions
                self.base_url = f"https://{domain}.salesforce.com"
            
            # ✅ Use API_VERSION constant
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
            self._log_status(f"🔑 Session established successfully")
            
            # ✅ Fetch objects AFTER connection is fully initialized
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
        ✅ FIXED: Detects expired passwords and other auth issues
        """
        self._log_status("Fetching all available SObjects from the organization...")
        
        try:
            # Verify connection first
            if not self.sf:
                raise Exception("Salesforce connection not initialized")
            
            if not self.session_id:
                raise Exception("No valid Salesforce session")
            
            # Call describe() API
            self._log_status("📞 Calling Salesforce describe() API...")
            
            try:
                response = self.sf.describe()
            except SalesforceExpiredSession as e:
                # ✅ DETECT EXPIRED PASSWORD
                error_str = str(e).lower()
                
                if 'password has expired' in error_str or 'expired_password' in error_str:
                    self._log_status("❌ PASSWORD EXPIRED")
                    raise Exception(
                        "🔐 Your Salesforce password has expired!\n\n"
                        "To fix this:\n"
                        "1. Go to your Salesforce org\n"
                        "2. Reset your password (Setup → Change Password)\n"
                        "3. Get new security token (Setup → My Personal Info → Reset Security Token)\n"
                        "4. Use the new password + token to log in again\n\n"
                        f"Technical error: {str(e)}"
                    )
                else:
                    # Other expired session errors
                    raise Exception(f"Session expired: {str(e)}")
            
            if not response or not isinstance(response, dict):
                raise Exception("Invalid response from Salesforce describe()")
            
            sobjects = response.get('sobjects', [])
            
            if not sobjects:
                self._log_status("⚠️ No SObjects returned from describe()")
                self.all_org_objects = []
                return
            
            self._log_status(f"📊 Received {len(sobjects)} total objects from Salesforce")
            
            # Filter for queryable, non-deprecated objects
            queryable_objects = [
                obj['name'] for obj in sobjects 
                if obj.get('queryable', False) and not obj.get('deprecatedAndHidden', False)
            ]
            
            # Sort and store
            self.all_org_objects = sorted(queryable_objects)
            
            if not self.all_org_objects:
                self._log_status("⚠️ No queryable objects found after filtering")
                self._log_status("⚠️ This usually means insufficient permissions")
                
                # Show breakdown
                queryable_count = sum(1 for obj in sobjects if obj.get('queryable', False))
                deprecated_count = sum(1 for obj in sobjects if obj.get('deprecatedAndHidden', False))
                
                self._log_status(f"📊 Breakdown:")
                self._log_status(f"  - Total objects: {len(sobjects)}")
                self._log_status(f"  - Queryable: {queryable_count}")
                self._log_status(f"  - Deprecated: {deprecated_count}")
                self._log_status(f"  - Final (queryable + not deprecated): {len(self.all_org_objects)}")
            else:
                self._log_status(f"✅ Found {len(self.all_org_objects)} queryable objects")
                
                # Show sample
                sample = ', '.join(self.all_org_objects[:5])
                if len(self.all_org_objects) > 5:
                    sample += f", ... (+{len(self.all_org_objects) - 5} more)"
                self._log_status(f"📦 Sample objects: {sample}")
        
        except Exception as e:
            error_msg = str(e)
            self._log_status(f"❌ Failed to fetch SObjects: {error_msg}")
            
            # Set empty list on error
            self.all_org_objects = []
            
            # Log detailed error for debugging
            import traceback
            traceback_str = traceback.format_exc()
            print(f"❌ DETAILED ERROR in _fetch_all_org_objects:\n{traceback_str}")
            
            # ✅ RE-RAISE if it's a password expiry error (so GUI can show it)
            if 'password has expired' in error_msg.lower() or 'expired_password' in error_msg.lower():
                raise
            else:
                # Other errors - log but don't crash
                self._log_status(f"🔍 Technical details logged to console")
    
    def get_all_objects(self) -> List[str]:
        """Accessor for the fetched object list"""
        return self.all_org_objects
    
    def _log_status(self, message: str):
        """Internal helper to send log messages back to the GUI"""
        if self.status_callback:
            self.status_callback(message, verbose=True)