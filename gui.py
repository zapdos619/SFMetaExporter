"""
Main GUI application for Salesforce Picklist & Metadata Exporter
"""
import os
import time
import tkinter as tk
from datetime import datetime
from typing import Optional, Set, List
from tkinter import messagebox, filedialog, END
from threading_helper import ThreadHelper

import customtkinter as ctk

from config import WINDOW_TITLE, WINDOW_GEOMETRY, APPEARANCE_MODE, COLOR_THEME
from config import DEFAULT_PICKLIST_FILENAME, DEFAULT_METADATA_FILENAME, DEFAULT_CONTENTDOCUMENT_FILENAME
from salesforce_client import SalesforceClient
from picklist_exporter import PicklistExporter
from metadata_exporter import MetadataExporter
from content_document_exporter import ContentDocumentExporter
from utils import format_runtime, print_picklist_statistics, print_metadata_statistics, print_content_document_statistics

from soql_runner import SOQLRunner
from soql_query_frame import SOQLQueryFrame

from metadata_switch_manager import MetadataSwitchManager
from salesforce_switch_frame import SalesforceSwitchFrame

# ✅ NEW - Report Exporter Module
from report_exporter.main_app import SalesforceExporterApp


# Set appearance mode and default color theme
ctk.set_appearance_mode(APPEARANCE_MODE)
ctk.set_default_color_theme(COLOR_THEME)


# gui.py - ADD THIS RIGHT AFTER IMPORTS

class ButtonStateManager:
    """
    Centralized manager for all operation buttons.
    Ensures only one operation can run at a time.
    """
    
    def __init__(self, gui_instance):
        self.gui = gui_instance
        self.operation_running = False
        self.current_operation = None
        self._buttons = {}
    
    def register_buttons(self, buttons_dict):
        """
        Register all operation buttons.
        
        Args:
            buttons_dict: {'picklist': btn, 'metadata': btn, ...}
        """
        self._buttons = buttons_dict
    
    def start_operation(self, operation_name: str) -> bool:
        """
        Start an operation. Returns False if another operation is running.
        """
        if self.operation_running:
            messagebox.showwarning(
                "Operation in Progress",
                f"Cannot start {operation_name}.\n\n"
                f"{self.current_operation} is currently running.\n"
                f"Please wait for it to complete."
            )
            return False
        
        self.operation_running = True
        self.current_operation = operation_name
        
        # Disable ALL buttons
        self._set_all_buttons_state("disabled")
        
        return True
    
    def end_operation(self):
        """End current operation and re-enable all buttons."""
        self.operation_running = False
        self.current_operation = None
        
        # Re-enable ALL buttons
        self._set_all_buttons_state("normal")
    
    def _set_all_buttons_state(self, state: str):
        """Set state for all registered buttons."""
        for button_name, button_widget in self._buttons.items():
            try:
                if button_widget and button_widget.winfo_exists():
                    button_widget.configure(state=state)
            except Exception as e:
                print(f"⚠️ Button state error ({button_name}): {e}")





class SalesforceExporterGUI(ctk.CTk):
    """Main GUI application class"""

    def __init__(self):
        super().__init__()

        self.title(WINDOW_TITLE)
        self.geometry(WINDOW_GEOMETRY)
        
        # ✅ CRITICAL: Initialize button_manager FIRST (before _setup_ui)
        self.button_manager = ButtonStateManager(self)

        self.sf_client: Optional[SalesforceClient] = None
        self.picklist_exporter: Optional[PicklistExporter] = None
        self.metadata_exporter: Optional[MetadataExporter] = None
        self.content_document_exporter: Optional[ContentDocumentExporter] = None
        self.all_org_objects: List[str] = []
        self.selected_objects: Set[str] = set()

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Create frames
        self.login_frame = ctk.CTkFrame(self)
        self.export_frame = ctk.CTkFrame(self)

        self.login_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.export_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

        self._setup_login_frame()
        self._setup_export_frame()

        # Initially show login frame
        self.export_frame.grid_forget()

        # Create SOQL frame
        self.soql_frame = None  # Will be created after login
        
        self.metadata_switch_manager: Optional[MetadataSwitchManager] = None
        self.switch_frame = None  # Will be created after login
        
        # ✅ NEW - Report Exporter frame
        self.report_exporter_frame = None  # Will be created after login
        


    # ==================================
    # Screen 1: Login & Authentication
    # ==================================

    def _setup_login_frame(self):
        """Setup the login screen UI - REORGANIZED with fixed Custom Domain behavior"""
        login_frame = self.login_frame
        login_frame.columnconfigure(1, weight=1)

        # Title
        ctk.CTkLabel(
            login_frame,
            text="Salesforce Login",
            font=ctk.CTkFont(size=30, weight="bold")
        ).grid(row=0, column=0, columnspan=2, pady=(50, 40))

        # ========== ROW 1: ORG TYPE ==========
        ctk.CTkLabel(
            login_frame,
            text="Org Type:",
            anchor="w",
            font=ctk.CTkFont(size=14),
            text_color=("#2c3e50", "#ecf0f1")  # Dark gray (light) / Light gray (dark)
        ).grid(row=1, column=0, padx=10, pady=15, sticky="w")

        self.org_type_var = ctk.StringVar(value="Production")
        
        org_container = ctk.CTkFrame(login_frame, fg_color="transparent")
        org_container.grid(row=1, column=1, padx=10, pady=15, sticky="w")
        
        self.radio_prod = ctk.CTkRadioButton(
            org_container,
            text="Production",
            variable=self.org_type_var,
            value="Production"
        )
        self.radio_prod.grid(row=0, column=0, padx=(0, 15), sticky="w")
        
        self.radio_sandbox = ctk.CTkRadioButton(
            org_container,
            text="Sandbox/Test",
            variable=self.org_type_var,
            value="Sandbox"
        )
        self.radio_sandbox.grid(row=0, column=1, padx=(0, 0), sticky="w")
        
        # ========== ROW 2: CUSTOM DOMAIN ==========
        ctk.CTkLabel(
            login_frame,
            text="Custom Domain:",
            anchor="w",
            font=ctk.CTkFont(size=14),
            text_color=("#2c3e50", "#ecf0f1")
        ).grid(row=2, column=0, padx=10, pady=15, sticky="w")
        
        self.custom_domain_var = ctk.BooleanVar(value=False)
        
        custom_domain_container = ctk.CTkFrame(login_frame, fg_color="transparent")
        custom_domain_container.grid(row=2, column=1, padx=10, pady=15, sticky="ew")
        custom_domain_container.grid_columnconfigure(0, weight=1)
        
        # Checkbox
        self.custom_domain_check = ctk.CTkCheckBox(
            custom_domain_container,
            text="🌐 Use Custom Domain",
            variable=self.custom_domain_var,
            command=self._on_custom_domain_toggle,
            font=ctk.CTkFont(size=12)
        )
        self.custom_domain_check.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        # ✅ Entry field - ALWAYS VISIBLE, state controlled by checkbox
        self.custom_domain_entry = ctk.CTkEntry(
            custom_domain_container,
            placeholder_text="mycompany.my.salesforce.com",
            width=350,
            state="disabled"  # ✅ Start disabled
        )
        self.custom_domain_entry.grid(row=1, column=0, sticky="ew")
        
        # Hint label
        ctk.CTkLabel(
            custom_domain_container,
            text="💡 Example: mycompany.my.salesforce.com (no https://)",
            font=ctk.CTkFont(size=10),
            text_color=("#7f8c8d", "#95a5a6"),  # Readable gray in both modes
            anchor="w"
        ).grid(row=2, column=0, sticky="w", pady=(2, 0))
        
        # ========== ROW 3: USERNAME ==========
        ctk.CTkLabel(
            login_frame,
            text="Username:",
            anchor="w",
            font=ctk.CTkFont(size=14),
            text_color=("#2c3e50", "#ecf0f1")
        ).grid(row=3, column=0, padx=10, pady=15, sticky="w")
        
        self.username_entry = ctk.CTkEntry(login_frame, width=350)
        self.username_entry.grid(row=3, column=1, padx=10, pady=15, sticky="ew")
        
        # ========== ROW 4: PASSWORD ==========
        ctk.CTkLabel(
            login_frame,
            text="Password:",
            anchor="w",
            font=ctk.CTkFont(size=14),
            text_color=("#2c3e50", "#ecf0f1")
        ).grid(row=4, column=0, padx=10, pady=15, sticky="w")
        
        self.password_entry = ctk.CTkEntry(login_frame, width=350, show="*")
        self.password_entry.grid(row=4, column=1, padx=10, pady=15, sticky="ew")
        
        # ========== ROW 5: SECURITY TOKEN ==========
        ctk.CTkLabel(
            login_frame,
            text="Security Token:",
            anchor="w",
            font=ctk.CTkFont(size=14),
            text_color=("#2c3e50", "#ecf0f1")
        ).grid(row=5, column=0, padx=10, pady=15, sticky="w")
        
        token_container = ctk.CTkFrame(login_frame, fg_color="transparent")
        token_container.grid(row=5, column=1, padx=10, pady=15, sticky="ew")
        token_container.grid_columnconfigure(0, weight=1)
        
        self.token_entry = ctk.CTkEntry(
            token_container, 
            width=350, 
            show="*",
            placeholder_text="Leave blank if IP whitelisted"
        )
        self.token_entry.grid(row=0, column=0, sticky="ew")
        
        # Hint label
        ctk.CTkLabel(
            token_container,
            text="💡 Optional - only needed if IP not whitelisted",
            font=ctk.CTkFont(size=10),
            text_color=("#7f8c8d", "#95a5a6"),
            anchor="w"
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        # ========== ROW 6: LOGIN BUTTON ==========
        self.login_button = ctk.CTkButton(
            login_frame,
            text="Login to Salesforce",
            command=self.login_action,
            width=150,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.login_button.grid(row=6, column=0, columnspan=2, pady=50, sticky="ew", padx=10)


    def _on_org_type_changed(self):
        """Handle org type radio button change"""
        # Disable custom domain when org type is selected
        if self.custom_domain_var.get():
            # User was using custom domain, keep it enabled
            pass
        else:
            # Normal behavior - radio buttons control domain
            pass
        

    def _on_custom_domain_toggle(self):
        """
        Toggle custom domain entry state (enabled/disabled).
        ✅ FIXED: Entry is always VISIBLE, only state changes.
        ✅ FIXED: Does NOT clear field when unchecking (preserves user input)
        """
        if self.custom_domain_var.get():
            # ✅ Checkbox CHECKED - Enable entry and disable radio buttons
            self.custom_domain_entry.configure(state="normal")
            
            # Disable org type radio buttons
            self.radio_prod.configure(state="disabled")
            self.radio_sandbox.configure(state="disabled")
            
            # Focus on custom domain entry
            self.custom_domain_entry.focus()
            
        else:
            # ✅ Checkbox UNCHECKED - Disable entry but DON'T clear it
            self.custom_domain_entry.configure(state="disabled")
            
            # Re-enable org type radio buttons
            self.radio_prod.configure(state="normal")
            self.radio_sandbox.configure(state="normal")
            
            # ✅ REMOVED: self.custom_domain_entry.delete(0, "end")
            # Keep the value so user doesn't lose their input
        
        # ✅ Force update to prevent state desync when moving windows
        self.custom_domain_entry.update_idletasks()


    def login_action(self):
        """Handle login button click - ENHANCED with Custom Domain support and Suffix Fix"""
        self.login_button.configure(state="disabled", text="Connecting...")

        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        token = self.token_entry.get().strip()
        
        # ✅ ENHANCED: Determine domain based on custom domain checkbox
        if self.custom_domain_var.get():
            # Custom domain mode
            domain = self.custom_domain_entry.get().strip()
            
            if not domain:
                messagebox.showerror("Input Error", "Please enter your custom Salesforce domain.")
                self.login_button.configure(state="normal", text="Login to Salesforce")
                self.custom_domain_entry.focus()
                return
            
            # ✅ Clean up domain input
            domain = domain.lower()
            
            # Remove https:// or http:// if user added it
            if domain.startswith("https://"):
                domain = domain[8:]
            elif domain.startswith("http://"):
                domain = domain[7:]
            
            # Remove trailing slashes
            domain = domain.rstrip('/')
            
            # ✅ Validate domain format
            valid_suffixes = ["salesforce.com", "force.com", "cloudforce.com"]
            is_valid = any(suffix in domain for suffix in valid_suffixes)
            
            if not is_valid:
                messagebox.showerror(
                    "Invalid Domain", 
                    "Custom domain must be a valid Salesforce domain.\n\n"
                    "Examples:\n"
                    "• mycompany.my.salesforce.com\n"
                    "• mycompany.lightning.force.com"
                )
                self.login_button.configure(state="normal", text="Login to Salesforce")
                self.custom_domain_entry.focus()
                return
            
            # ✅ CRITICAL FIX: Strip .salesforce.com suffix
            # The simple_salesforce library automatically appends '.salesforce.com' to the domain argument.
            # If we pass 'my.salesforce.com', it becomes 'my.salesforce.com.salesforce.com'.
            # We must strip it here before passing it to the client.
            if domain.endswith(".salesforce.com"):
                domain = domain[:-15]  # Removes the last 15 characters (.salesforce.com)

        else:
            # Standard mode (Production or Sandbox)
            domain = 'test' if self.org_type_var.get() == 'Sandbox' else 'login'

        # ✅ ENHANCED: Token is now optional
        if not all([username, password]):
            messagebox.showerror("Input Error", "Username and Password are required.")
            self.login_button.configure(state="normal", text="Login to Salesforce")
            return
        
        # ✅ INFO: Show message if token is empty
        if not token:
            self.update_status("ℹ️ Logging in without security token (IP must be whitelisted)")

        # Run login in background thread
        def do_login():
            try:
                self.sf_client = SalesforceClient(
                    username=username,
                    password=password,
                    security_token=token,  # Can be empty string now
                    domain=domain,
                    status_callback=self.update_status
                )

                # Initialize exporters
                self.picklist_exporter = PicklistExporter(self.sf_client)
                self.metadata_exporter = MetadataExporter(self.sf_client)
                self.content_document_exporter = ContentDocumentExporter(self.sf_client)

                self.all_org_objects = self.sf_client.get_all_objects()

                # Update UI on main thread
                self.after(0, self._on_login_success)

            except Exception as e:
                # Handle error on main thread
                error_message = str(e);
                self.after(0, lambda: self._on_login_error(error_message))

        ThreadHelper.run_in_thread(do_login)
  
        
    def _show_login_status(self, message: str, color: str = "gray"):
        """
        Show status message during login process
        
        Args:
            message: Status message to display
            color: Text color (gray, green, red, orange)
        """
        try:
            # Update the status textbox with colored message
            self.status_textbox.configure(state="normal")
            
            timestamp = datetime.now().strftime("[%H:%M:%S]")
            
            # Add colored indicator
            if color == "green":
                indicator = "✅"
            elif color == "red":
                indicator = "❌"
            elif color == "orange":
                indicator = "⚠️"
            else:
                indicator = "ℹ️"
            
            log_msg = f"{timestamp} {indicator} {message}\n"
            
            self.status_textbox.insert("end", log_msg)
            self.status_textbox.see("end")
            self.status_textbox.configure(state="disabled")
            
            print(log_msg.strip())
            
        except Exception as e:
            print(f"⚠️ Status update error: {e}")
        


    def _on_login_success(self):
        """Called after successful login - ENHANCED with detailed info"""
        
        # Determine connection type
        if self.custom_domain_var.get():
            connection_type = "Custom Domain"
            domain_used = self.custom_domain_entry.get().strip()
        else:
            connection_type = self.org_type_var.get()
            domain_used = 'login.salesforce.com' if connection_type == 'Production' else 'test.salesforce.com'
        
        # Check if token was used
        token_status = "with token" if self.token_entry.get().strip() else "without token (IP whitelisted)"
        
        # Build success message
        success_msg = (
            f"Successfully connected to Salesforce!\n\n"
            f"Connection Details:\n"
            f"• Type: {connection_type}\n"
            f"• Domain: {domain_used}\n"
            f"• Instance: {self.sf_client.base_url}\n"
            f"• API Version: v{self.sf_client.api_version}\n"
            f"• Authentication: {token_status}"
        )
        
        messagebox.showinfo("Success", success_msg)

        # Switch to Export Frame
        self.login_frame.grid_forget()
        self.export_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.populate_available_objects(self.all_org_objects)
        self.populate_selected_objects()
        self.login_button.configure(state="normal", text="Login to Salesforce")

        # Initialize SOQL Runner
        self.soql_runner = SOQLRunner(self.sf_client)
        
        # Initialize Metadata Switch Manager
        self.metadata_switch_manager = MetadataSwitchManager(
            self.sf_client.sf,
            status_callback=self.update_status
        )
        
        # Log detailed connection info
        self.update_status("=" * 60)
        self.update_status(f"✅ CONNECTED TO SALESFORCE")
        self.update_status(f"📊 Connection Type: {connection_type}")
        self.update_status(f"🌐 Domain: {domain_used}")
        self.update_status(f"🔗 Instance: {self.sf_client.base_url}")
        self.update_status(f"📡 API Version: v{self.sf_client.api_version}")
        self.update_status(f"🔐 Authentication: {token_status}")
        self.update_status(f"📦 Objects Found: {len(self.all_org_objects)}")
        self.update_status("=" * 60)


    def _make_login_error_friendly(self, error_msg: str) -> tuple[str, str]:
        """
        Convert technical error messages to user-friendly ones.
        Returns a tuple: (Short Friendly Message for Popup, Detailed Log Message).
        """
        error_lower = error_msg.lower()
        friendly_text = "❌ An unexpected error occurred." # Default short message

        # --- LOG MESSAGE PREPARATION ---
        MAX_LOG_LEN = 1000
        if len(error_msg) > MAX_LOG_LEN:
            tech_log_display = error_msg[:MAX_LOG_LEN] + f"\n... [Truncated {len(error_msg)-MAX_LOG_LEN} more chars]"
        else:
            tech_log_display = error_msg

        detailed_log_message = (
            f"❌ LOGIN FAILED\n"
            f"{'-'*40}\n"
            f"Technical Details:\n{tech_log_display}"
        )
        
        # --- SHORT POPUP MESSAGE DETERMINATION ---
        
        # ✅ ENHANCED: Better detection for DNS/Domain issues
        is_dns_issue = any(x in error_lower for x in [
            "nameresolutionerror", 
            "getaddrinfo", 
            "name or service not known",
            "nodename nor servname provided",
            "no such host",
            "name resolution"
        ])
        
        is_connection_issue = any(x in error_lower for x in [
            "max retries exceeded",
            "connectionerror",
            "failed to establish",
            "connection refused",
            "connection reset"
        ])
        
        is_timeout = "timeout" in error_lower or "timed out" in error_lower

        # 1. Custom Domain / DNS Errors (HIGHEST PRIORITY)
        if self.custom_domain_var.get() and is_dns_issue:
            friendly_text = (
                "🌐 Custom Domain Not Found\n\n"
                "The domain you entered could not be reached. "
                "Please check:\n"
                "• Spelling is correct\n"
                "• Domain is active in Salesforce\n"
                "• No typos in the domain name"
            )
        
        # 2. Custom Domain Connection Issues (but not DNS)
        elif self.custom_domain_var.get() and is_connection_issue:
            friendly_text = (
                "⚠️ Cannot Connect to Custom Domain\n\n"
                "Failed to establish connection. Please check:\n"
                "• Domain spelling\n"
                "• Your internet connection\n"
                "• Firewall/VPN settings"
            )
        
        # 3. Custom Domain Timeout
        elif self.custom_domain_var.get() and is_timeout:
            friendly_text = (
                "⏱️ Connection Timeout (Custom Domain)\n\n"
                "The domain took too long to respond. "
                "This might mean:\n"
                "• Domain is spelled incorrectly\n"
                "• Domain is not accessible\n"
                "• Network is slow"
            )
        
        # 4. IP Whitelisting / Security Token
        elif "invalid_grant" in error_lower and "security token" not in error_msg:
            friendly_text = (
                "⚠️ Login Rejected (Security Token Required)\n\n"
                "Your IP address is not whitelisted. "
                "Please add your Security Token."
            )
        
        # 5. Invalid Password/Username
        elif "invalid username" in error_lower or "invalid login" in error_lower or "authentication failure" in error_lower:
            friendly_text = "❌ Invalid username or password."
        
        # 6. Locked Account
        elif "locked" in error_lower:
            friendly_text = "🔒 Account is locked. Please contact your Salesforce administrator."
        
        # 7. Generic Timeout (non-custom domain)
        elif is_timeout:
            friendly_text = "⏱️ Connection timed out. Try again or check your internet."
        
        # 8. Generic 404/403
        elif "404" in error_msg or "not found" in error_lower:
            friendly_text = "❌ The requested resource was not found (404)."
        
        return friendly_text, detailed_log_message

    def _on_login_error(self, error_message):
        """Called when login fails - Now uses the short message for the popup and detailed for the log."""
        
        # Ensure error_message is a string and not empty
        error_message = str(error_message) if error_message else "An unknown error occurred during login."
        
        # Get the friendly message for the popup and the detailed log message
        popup_msg, detailed_log_msg = self._make_login_error_friendly(error_message)
        
        # 1. Show short error dialog for the user
        messagebox.showerror("Login Failed", popup_msg)
        
        # 2. Log full technical details to the status bar (where the user can scroll)
        self.update_status("=" * 60)
        self.update_status(detailed_log_msg) 
        self.update_status("=" * 60)
        
        # Reset state
        self.sf_client = None
        self.login_button.configure(state="normal", text="Login to Salesforce")
        
        # Smart Focus: Decide where to put cursor based on the RAW error content
        err_lower = error_message.lower()
        
        # If it looks like a domain/DNS issue and Custom Domain is checked -> Focus Domain field
        is_dns_error = any(x in err_lower for x in ["nameresolutionerror", "getaddrinfo", "max retries", "connectionerror"])
        
        if self.custom_domain_var.get() and is_dns_error:
            self.custom_domain_entry.focus()
            self.custom_domain_entry.select_range(0, 'end')
            
        # If token/IP issue -> Focus Token field
        elif "token" in err_lower or "invalid_grant" in err_lower:
            self.token_entry.focus()
            
        # Otherwise (Password/Username) -> Focus Password field
        else:
            self.password_entry.focus()
            self.password_entry.delete(0, "end")



    # ==================================
    # Screen 2: Object Selection & Export
    # ==================================

    def _setup_export_frame(self):
        """Setup the export screen UI"""
        export_frame = self.export_frame
        export_frame.grid_rowconfigure(2, weight=1)
        export_frame.grid_columnconfigure(0, weight=1)

        # Header with logout button
        header_frame = ctk.CTkFrame(export_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, pady=(10, 5), sticky="ew")
        header_frame.columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header_frame,
            text="Object Selection & Export",
            font=ctk.CTkFont(size=30, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        self.logout_button = ctk.CTkButton(
            header_frame,
            text="Logout",
            command=self.logout_action,
            width=100,
            fg_color="#CC3333"
        )
        self.logout_button.grid(row=0, column=1, sticky="e", padx=10)

        # Selection frame with three columns
        selection_frame = ctk.CTkFrame(export_frame)
        selection_frame.grid(row=1, column=0, pady=10, sticky="nsew")
        selection_frame.grid_columnconfigure(0, weight=3)
        selection_frame.grid_columnconfigure(1, weight=1)
        selection_frame.grid_columnconfigure(2, weight=2)
        selection_frame.grid_rowconfigure(0, weight=1)

        # Available Objects (Left)
        self._setup_available_objects_panel(selection_frame)

        # Action Buttons (Middle)
        self._setup_action_buttons_panel(selection_frame)

        # Selected Objects (Right)
        self._setup_selected_objects_panel(selection_frame)

        # Status textbox
        self.status_textbox = ctk.CTkTextbox(export_frame, height=150)
        self.status_textbox.grid(row=2, column=0, padx=20, pady=(10, 10), sticky="ew")
        self.status_textbox.insert("end", "Status: Ready to select objects and export.")
        self.status_textbox.configure(state="disabled")

        # Export buttons frame (NOW WITH 5 BUTTONS)
        export_buttons_frame = ctk.CTkFrame(export_frame, fg_color="transparent")
        export_buttons_frame.grid(row=3, column=0, pady=(10, 20), sticky="ew", padx=20)  
        export_buttons_frame.grid_columnconfigure(0, weight=1)
        export_buttons_frame.grid_columnconfigure(1, weight=1)
        export_buttons_frame.grid_columnconfigure(2, weight=1)
        export_buttons_frame.grid_columnconfigure(3, weight=1)

        # Configure 5 columns with equal weight
        for i in range(6):
            export_buttons_frame.grid_columnconfigure(i, weight=1)
        
        self.export_picklist_button = ctk.CTkButton(
            export_buttons_frame,
            text="Export Picklist Data",
            command=self.export_picklist_action,
            height=50,
            fg_color="green",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.export_picklist_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        self.export_metadata_button = ctk.CTkButton(
            export_buttons_frame,
            text="Export Metadata",
            command=self.export_metadata_action,
            height=50,
            fg_color="#1F538D",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.export_metadata_button.grid(row=0, column=1, sticky="ew", padx=(5, 5))
        
        self.download_files_button = ctk.CTkButton(
            export_buttons_frame,
            text="Download Files",
            command=self.download_files_action,
            height=50,
            fg_color="#FF6B35",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.download_files_button.grid(row=0, column=2, sticky="ew", padx=(5, 5))
        
        self.run_soql_button = ctk.CTkButton(
            export_buttons_frame,
            text="Run SOQL",
            command=self.run_soql_action,
            height=50,
            fg_color="#9B59B6",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.run_soql_button.grid(row=0, column=3, sticky="ew", padx=(5, 5))
        
        # NEW 5TH BUTTON - SALESFORCE SWITCH
        self.salesforce_switch_button = ctk.CTkButton(
            export_buttons_frame,
            text="Salesforce Switch",
            command=self.salesforce_switch_action,
            height=50,
            fg_color="#E74C3C",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.salesforce_switch_button.grid(row=0, column=4, sticky="ew", padx=(5, 0))
        
        # ✅ NEW 6TH BUTTON - REPORT EXPORTER
        self.report_exporter_button = ctk.CTkButton(
            export_buttons_frame,
            text="📊 Report Export",
            command=self.report_exporter_action,
            height=50,
            fg_color="#16A085",  # Teal/green color
            hover_color="#138D75",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.report_exporter_button.grid(row=0, column=5, sticky="ew", padx=(5, 0))
        
        # ✅ NEW: Register all buttons with state manager (add at END of _setup_export_frame)
        self.button_manager.register_buttons({
            'picklist': self.export_picklist_button,
            'metadata': self.export_metadata_button,
            'download': self.download_files_button,
            'soql': self.run_soql_button,
            'switch': self.salesforce_switch_button,
            'report': self.report_exporter_button
        })


    def _setup_available_objects_panel(self, parent):
        """Setup the available objects panel"""
        available_frame = ctk.CTkFrame(parent)
        available_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        available_frame.grid_rowconfigure(2, weight=1)
        available_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            available_frame,
            text="Available Objects (Org)",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=(5, 5))

        self.search_entry = ctk.CTkEntry(
            available_frame,
            placeholder_text="Search Object API Name...",
            height=35
        )
        self.search_entry.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.search_entry.bind("<KeyRelease>", self.filter_available_objects)

        self.available_listbox = tk.Listbox(
            available_frame,
            selectmode="extended",
            height=15,
            exportselection=False,
            font=("Arial", 12),
            borderwidth=0,
            highlightthickness=0,
            selectbackground="#1F538D",
            fg="white",
            background="#242424"
        )
        self.available_listbox.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")

    def _setup_action_buttons_panel(self, parent):
        """Setup the action buttons panel"""
        action_frame = ctk.CTkFrame(parent, fg_color="transparent")
        action_frame.grid(row=0, column=1, padx=5, pady=10, sticky="n")

        ctk.CTkLabel(
            action_frame,
            text="Actions",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=5)

        ctk.CTkButton(
            action_frame,
            text=">> Add Selected >>",
            command=self.add_selected_to_export,
            height=35
        ).pack(pady=5, padx=5, fill="x")

        ctk.CTkButton(
            action_frame,
            text="<< Remove Selected <<",
            command=self.remove_selected_from_export,
            height=35
        ).pack(pady=5, padx=5, fill="x")

        ctk.CTkButton(
            action_frame,
            text="Select All",
            command=self.select_all_available,
            height=35
        ).pack(pady=(20, 5), padx=5, fill="x")

        ctk.CTkButton(
            action_frame,
            text="Deselect All",
            command=self.deselect_all_available,
            height=35
        ).pack(pady=5, padx=5, fill="x")

    def _setup_selected_objects_panel(self, parent):
        """Setup the selected objects panel"""
        selected_frame = ctk.CTkFrame(parent)
        selected_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")
        selected_frame.grid_rowconfigure(1, weight=1)
        selected_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            selected_frame,
            text="Selected for Export",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=(5, 5))

        self.selected_listbox = tk.Listbox(
            selected_frame,
            selectmode="extended",
            height=15,
            exportselection=False,
            font=("Arial", 12),
            borderwidth=0,
            highlightthickness=0,
            selectbackground="#3366CC",
            fg="white",
            background="#242424"
        )
        self.selected_listbox.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

    # ==================================
    # Object List Management Methods
    # ==================================

    def populate_available_objects(self, objects: List[str]):
        """Populates the Left ListBox based on the current search filter"""
        self.available_listbox.delete(0, END)
        for obj in objects:
            self.available_listbox.insert(END, obj)
            if obj in self.selected_objects:
                idx = self.available_listbox.get(0, END).index(obj)
                self.available_listbox.itemconfig(idx, {'fg': '#87CEEB'})

    def populate_selected_objects(self):
        """Populates the Right ListBox from the internal selected_objects set"""
        self.selected_listbox.delete(0, END)
        for obj in sorted(list(self.selected_objects)):
            self.selected_listbox.insert(END, obj)

    def filter_available_objects(self, event):
        """Filters the Available ListBox based on the search entry content"""
        search_term = self.search_entry.get().lower()
        filtered_objects = [
            obj for obj in self.all_org_objects
            if search_term in obj.lower()
        ]
        self.populate_available_objects(filtered_objects)

    def add_selected_to_export(self):
        """Adds selected objects from the Available List to the Export Set"""
        selected_indices = self.available_listbox.curselection()

        if not selected_indices:
            messagebox.showwarning(
                "Selection",
                "Please select one or more objects from the 'Available Objects' list to add."
            )
            return

        added_count = 0
        for i in selected_indices:
            obj_name = self.available_listbox.get(i)
            if obj_name not in self.selected_objects:
                self.selected_objects.add(obj_name)
                added_count += 1

        if added_count > 0:
            self.populate_selected_objects()
            self.filter_available_objects(None)
            self.update_status(f"Added {added_count} object(s) to export list.")

    def remove_selected_from_export(self):
        """Removes selected objects from the Selected List"""
        selected_indices = self.selected_listbox.curselection()

        if not selected_indices:
            messagebox.showwarning(
                "Selection",
                "Please select one or more objects from the 'Selected for Export' list to remove."
            )
            return

        removed_objects = []
        for i in reversed(selected_indices):
            obj_name = self.selected_listbox.get(i)
            removed_objects.append(obj_name)

        for obj_name in removed_objects:
            self.selected_objects.discard(obj_name)

        if removed_objects:
            self.populate_selected_objects()
            self.filter_available_objects(None)
            self.update_status(f"Removed {len(removed_objects)} object(s) from export list.")

    def select_all_available(self):
        """Selects all objects currently visible in the Available ListBox"""
        self.available_listbox.select_set(0, END)

    def deselect_all_available(self):
        """Deselects all objects currently visible in the Available ListBox"""
        self.available_listbox.select_clear(0, END)

    # ==================================
    # Run SOQL Action Methods
    # ==================================
    
    def run_soql_action(self):
        """Handle Run SOQL button click"""
        if not self.sf_client or not self.soql_runner:
            messagebox.showerror("Error", "Not logged in. Please log in first.")
            return
        
        # ✅ NEW: Check if another operation is running
        if self.button_manager.operation_running:
            messagebox.showwarning(
                "Operation in Progress",
                f"{self.button_manager.current_operation} is currently running.\n\n"
                f"Please wait for it to complete before opening SOQL runner."
            )
            return
        
        # Create SOQL frame if it doesn't exist
        if self.soql_frame is None:
            self.soql_frame = SOQLQueryFrame(
                self,
                self.soql_runner,
                status_callback=self.update_status
            )
            self.soql_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
            
            # Connect back button
            self.soql_frame.back_button.configure(command=self.show_export_frame)
        
        # Hide export frame and show SOQL frame
        self.export_frame.grid_forget()
        self.soql_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
    
    def show_export_frame(self):
        """Show the export frame and hide SOQL frame"""
        if self.soql_frame:
            self.soql_frame.grid_forget()
        self.export_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        
    # ============================================
    # salesforce_switch_action
    # ============================================

    def salesforce_switch_action(self):
        """Handle Salesforce Switch button click"""
        if not self.sf_client or not self.metadata_switch_manager:
            messagebox.showerror("Error", "Not logged in. Please log in first.")
            return
        
        # ✅ NEW: Check if another operation is running
        if self.button_manager.operation_running:
            messagebox.showwarning(
                "Operation in Progress",
                f"{self.button_manager.current_operation} is currently running.\n\n"
                f"Please wait for it to complete before opening Salesforce Switch."
            )
            return
        
        # Create switch frame if it doesn't exist
        if self.switch_frame is None:
            self.switch_frame = SalesforceSwitchFrame(
                self,
                self.metadata_switch_manager,
                username=self.username_entry.get(),
                status_callback=self.update_status
            )
            self.switch_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
            
            # Connect back button
            self.switch_frame.back_button.configure(command=self.show_export_frame_from_switch)
        
        # Hide export frame and show switch frame
        self.export_frame.grid_forget()
        self.switch_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        # Load components
        self.switch_frame.load_components()


    # ============================================
    # show_export_frame_from_switch
    # ============================================



    def show_export_frame_from_switch(self):
        """Show the export frame and hide switch frame"""
        if self.switch_frame:
            self.switch_frame.grid_forget()
        self.export_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
    def _get_window_monitor_geometry(self) -> tuple:
        """
        Get the geometry of the monitor where this window is currently displayed.
        
        Returns:
            (x, y, width, height) of the monitor containing this window
        """
        try:
            # Get main window position and size
            window_x = self.winfo_x()
            window_y = self.winfo_y()
            window_width = self.winfo_width()
            window_height = self.winfo_height()
            
            # Calculate window center point
            window_center_x = window_x + (window_width // 2)
            window_center_y = window_y + (window_height // 2)
            
            # Get screen dimensions
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            # Detect which monitor the window is on
            if window_center_x > screen_width:
                # Window is on RIGHT monitor (extended display)
                monitor_x = screen_width
                monitor_y = 0
                monitor_width = screen_width
                monitor_height = screen_height
            elif window_center_x < 0:
                # Window is on LEFT monitor
                monitor_x = -screen_width
                monitor_y = 0
                monitor_width = screen_width
                monitor_height = screen_height
            elif window_center_y < 0:
                # Window is on TOP monitor (stacked setup)
                monitor_x = 0
                monitor_y = -screen_height
                monitor_width = screen_width
                monitor_height = screen_height
            elif window_center_y > screen_height:
                # Window is on BOTTOM monitor
                monitor_x = 0
                monitor_y = screen_height
                monitor_width = screen_width
                monitor_height = screen_height
            else:
                # Window is on PRIMARY monitor
                monitor_x = 0
                monitor_y = 0
                monitor_width = screen_width
                monitor_height = screen_height
            
            return (monitor_x, monitor_y, monitor_width, monitor_height)
            
        except Exception as e:
            print(f"⚠️ Error detecting monitor: {e}")
            # Fallback to primary monitor
            return (0, 0, self.winfo_screenwidth(), self.winfo_screenheight())

    def _get_window_state_info(self) -> dict:
        """
        Get current window state and geometry information.
        
        Returns:
            Dictionary with window state info
        """
        try:
            # Get window state
            state = self.state()
            
            # Check if zoomed (maximized)
            is_zoomed = (state == 'zoomed')
            
            # Check if fullscreen
            try:
                is_fullscreen = self.attributes('-fullscreen')
            except:
                is_fullscreen = False
            
            # Determine state string
            if is_fullscreen:
                state_str = 'fullscreen'
            elif is_zoomed:
                state_str = 'zoomed'
            else:
                state_str = 'normal'
            
            # Get window geometry
            width = self.winfo_width()
            height = self.winfo_height()
            x = self.winfo_x()
            y = self.winfo_y()
            
            # Get monitor geometry
            monitor_x, monitor_y, monitor_width, monitor_height = self._get_window_monitor_geometry()
            
            return {
                'state': state_str,
                'width': width,
                'height': height,
                'x': x,
                'y': y,
                'monitor_x': monitor_x,
                'monitor_y': monitor_y,
                'monitor_width': monitor_width,
                'monitor_height': monitor_height
            }
            
        except Exception as e:
            print(f"⚠️ Error getting window state: {e}")
            # Fallback to defaults
            return {
                'state': 'normal',
                'width': 1200,
                'height': 800,
                'x': 100,
                'y': 100,
                'monitor_x': 0,
                'monitor_y': 0,
                'monitor_width': self.winfo_screenwidth(),
                'monitor_height': self.winfo_screenheight()
            }

    def _center_window_on_monitor(self, window, window_width: int, window_height: int, 
                                monitor_x: int, monitor_y: int, 
                                monitor_width: int, monitor_height: int):
        """
        Center a window on a specific monitor.
        
        Args:
            window: The window to center
            window_width: Desired window width
            window_height: Desired window height
            monitor_x: Monitor X offset
            monitor_y: Monitor Y offset
            monitor_width: Monitor width
            monitor_height: Monitor height
        """
        try:
            # Calculate center position on the monitor
            center_x = monitor_x + (monitor_width - window_width) // 2
            center_y = monitor_y + (monitor_height - window_height) // 2
            
            # Set geometry
            window.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
            
        except Exception as e:
            print(f"⚠️ Error centering window: {e}")
            # Fallback to default positioning
            window.geometry(f"{window_width}x{window_height}")

    def _apply_parent_state_to_child(self, child_window, parent_state: dict):
        """
        Apply parent window's state (position, size, fullscreen) to child window.
        
        Args:
            child_window: The child Toplevel window
            parent_state: Dictionary from _get_window_state_info()
        """
        try:
            state = parent_state['state']
            
            if state == 'fullscreen':
                # Parent is fullscreen - make child fullscreen too
                try:
                    child_window.attributes('-fullscreen', True)
                    print("🖥️ Report Exporter: Fullscreen mode")
                except:
                    pass
                
            elif state == 'zoomed':
                # Parent is maximized - maximize child
                try:
                    child_window.state('zoomed')
                    print("🖥️ Report Exporter: Maximized mode")
                except:
                    pass
                
            else:
                # Parent is normal - match parent's size and center on same monitor
                width = parent_state['width']
                height = parent_state['height']
                monitor_x = parent_state['monitor_x']
                monitor_y = parent_state['monitor_y']
                monitor_width = parent_state['monitor_width']
                monitor_height = parent_state['monitor_height']
                
                # Use 90% of parent size (looks better than exact match)
                child_width = int(width * 0.9)
                child_height = int(height * 0.9)
                
                # Ensure minimum size
                child_width = max(child_width, 1000)
                child_height = max(child_height, 700)
                
                # Center on same monitor as parent
                self._center_window_on_monitor(
                    child_window,
                    child_width,
                    child_height,
                    monitor_x,
                    monitor_y,
                    monitor_width,
                    monitor_height
                )
                
                print(f"🖥️ Report Exporter: Normal mode ({child_width}x{child_height})")
            
            # Force window to update
            child_window.update_idletasks()
            
        except Exception as e:
            print(f"⚠️ Error applying parent state to child: {e}")
            # Fallback to default size and position
            try:
                child_window.geometry("1200x740")
            except:
                pass     


    # ✅ NEW METHOD 1 - Report Exporter Action
    # gui.py - REPLACE the report_exporter_action method

    def report_exporter_action(self):
        """Handle Report Exporter button click (6th button)"""
        if not self.sf_client:
            messagebox.showerror("Error", "Not logged in. Please log in first.")
            return
        
        # ✅ NEW: Get current appearance mode
        current_appearance = ctk.get_appearance_mode()  # Returns "Light" or "Dark"
        
        # Build session info for Report Exporter
        session_info = {
            "session_id": self.sf_client.session_id,
            "instance_url": self.sf_client.base_url,
            "api_version": self.sf_client.api_version,
            "user_name": self.username_entry.get(),
            "appearance_mode": current_appearance  # ✅ NEW: Pass theme to child
        }
        
        # ✅ CRITICAL FIX: Check if window exists and is alive
        window_exists = (
            self.report_exporter_frame is not None and 
            hasattr(self.report_exporter_frame, 'winfo_exists') and
            self.report_exporter_frame.winfo_exists()
        )
        
        if window_exists:
            # Window already exists - just show it
            try:
                self.report_exporter_frame.deiconify()
                self.report_exporter_frame.lift()
                self.report_exporter_frame.focus_force()
                
                # Hide main window
                self.withdraw()
                
                print("✅ Report Exporter: Restored existing window")
                return
                
            except Exception as e:
                print(f"⚠️ Error showing existing window: {e}")
                # Window is broken, recreate it
                self.report_exporter_frame = None
        
        # Create new window
        try:
            print("🔨 Creating new Report Exporter window...")
            
            self.report_exporter_frame = SalesforceExporterApp(
                master=self,
                session_info=session_info,
                on_logout=self.show_export_frame_from_report_exporter
            )
            
            # ✅ CRITICAL: Don't hide parent yet - let child initialize first
            print("⏳ Window created, initializing...")
            
            # ✅ Get parent window state AFTER child is created
            parent_state = self._get_window_state_info()
            
            # ✅ Apply parent state with delay (let child finish _setup_ui first)
            self.after(100, lambda: self._finalize_report_exporter_window(parent_state))
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            
            print(f"❌ Failed to create Report Exporter:")
            print(error_details)
            
            messagebox.showerror(
                "Error",
                f"Failed to open Report Exporter:\n\n{str(e)}"
            )
            
            self.report_exporter_frame = None
            return
    
    # gui.py - ADD this new method

    def _finalize_report_exporter_window(self, parent_state):
        """
        Finalize report exporter window after initialization.
        
        ✅ Called with delay to ensure child window is fully initialized.
        """
        try:
            # Check if window still exists
            if not self.report_exporter_frame or not self.report_exporter_frame.winfo_exists():
                print("⚠️ Report Exporter window destroyed during initialization")
                self.deiconify()  # Show parent again
                return
            
            # Apply parent state to child
            print("🎨 Applying window state...")
            self._apply_parent_state_to_child(self.report_exporter_frame, parent_state)
            
            # NOW hide parent window
            print("👁️ Hiding parent window...")
            self.withdraw()
            
            # Ensure child is visible and focused
            self.report_exporter_frame.deiconify()
            self.report_exporter_frame.lift()
            self.report_exporter_frame.focus_force()
            
            print("✅ Report Exporter window finalized successfully")
            
            # Log action
            try:
                self.update_status("📊 Opened Report Exporter")
            except:
                pass
                
        except Exception as e:
            print(f"❌ Error finalizing Report Exporter window: {e}")
            import traceback
            traceback.print_exc()
            
            # Recovery: show parent window again
            try:
                self.deiconify()
            except:
                pass    
    
    
    
    
    # ✅ NEW METHOD 2 - Back from Report Exporter
    def show_export_frame_from_report_exporter(self):
        """Show the export frame and hide report exporter frame"""
        if self.report_exporter_frame:
            try:
                # ✅ FIXED: Toplevel windows use withdraw(), not grid_forget()
                self.report_exporter_frame.withdraw()
                
                # Alternative: destroy and recreate next time
                # self.report_exporter_frame.destroy()
                # self.report_exporter_frame = None
                
            except Exception as e:
                print(f"⚠️ Error hiding report exporter: {e}")
        
        # Show main export frame
        self.export_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        # Bring main window to front
        self.deiconify()
        self.lift()
        self.focus_force()
        
        self._log("⬅️ Returned from Report Exporter")


    # ==================================
    # Export Action Methods
    # ==================================

    def export_picklist_action(self):
        """Handle export picklist button click"""
        if not self.sf_client or not self.picklist_exporter:
            messagebox.showerror("Error", "Not logged in. Please log in first.")
            return

        selected_objects_list = sorted(list(self.selected_objects))

        if not selected_objects_list:
            messagebox.showwarning(
                "Warning",
                "The 'Selected for Export' list is empty. Please add objects."
            )
            return

        # ✅ NEW: Check if another operation is running
        if not self.button_manager.start_operation("Picklist Export"):
            return

        default_filename = DEFAULT_PICKLIST_FILENAME.format(
            timestamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        )
        output_file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not output_file_path:
            # ✅ NEW: User cancelled, end operation
            self.button_manager.end_operation()
            return

        # ✅ REMOVE OLD CODE: Delete these lines if they exist:
        # self.export_picklist_button.configure(state="disabled", text="Exporting... DO NOT CLOSE")
        # self.export_metadata_button.configure(state="disabled")
        # self.download_files_button.configure(state="disabled")
        # self.run_soql_button.configure(state="disabled")
        # self.salesforce_switch_button.configure(state="disabled")

        self.update_status(
            f"Starting picklist export for {len(selected_objects_list)} objects to {output_file_path}..."
        )
        start_time = time.time()

        # Run export in background thread
        def do_export():
            try:
                output_path, stats = self.picklist_exporter.export_picklists(
                    selected_objects_list,
                    output_file_path
                )

                end_time = time.time()
                runtime_seconds = end_time - start_time
                runtime_formatted = format_runtime(runtime_seconds)

                # Update UI on main thread
                self.after(0, lambda: self._on_picklist_export_success(
                    output_path, stats, runtime_formatted
                ))

            except Exception as e:
                # Handle error on main thread
                self.after(0, lambda: self._on_picklist_export_error(str(e)))

        ThreadHelper.run_in_thread(do_export)

    def _on_picklist_export_success(self, output_path, stats, runtime_formatted):
        """Called after successful picklist export"""
        self.update_status(f"Export Complete! Total Runtime: {runtime_formatted}")
        messagebox.showinfo(
            "Export Done",
            f"Picklist data successfully exported to:\n{output_path}"
        )

        print_picklist_statistics(stats, runtime_formatted, output_path)

        # ✅ NEW: Re-enable all buttons
        self.button_manager.end_operation()

        # ✅ REMOVE OLD CODE: Delete these lines if they exist:
        # self.export_picklist_button.configure(state="normal", text="Export Picklist Data")
        # self.export_metadata_button.configure(state="normal")
        # self.download_files_button.configure(state="normal")
        # self.run_soql_button.configure(state="normal")
        # self.salesforce_switch_button.configure(state="normal")

    def _on_picklist_export_error(self, error_message):
        """Called when picklist export fails"""
        self.update_status(f"❌ FATAL EXPORT ERROR: {error_message}")
        messagebox.showerror("Export Error", f"A fatal error occurred during export: {error_message}")

        # ✅ NEW: Re-enable all buttons
        self.button_manager.end_operation()

        # ✅ REMOVE OLD CODE: Delete these lines if they exist:
        # self.export_picklist_button.configure(state="normal", text="Export Picklist Data")
        # self.export_metadata_button.configure(state="normal")
        # self.download_files_button.configure(state="normal")
        # self.salesforce_switch_button.configure(state="normal")

    def export_metadata_action(self):
        """Handle export metadata button click"""
        if not self.sf_client or not self.metadata_exporter:
            messagebox.showerror("Error", "Not logged in. Please log in first.")
            return

        selected_objects_list = sorted(list(self.selected_objects))

        if not selected_objects_list:
            messagebox.showwarning(
                "Warning",
                "The 'Selected for Export' list is empty. Please add objects."
            )
            return

        # ✅ NEW: Check if another operation is running
        if not self.button_manager.start_operation("Metadata Export"):
            return

        default_filename = DEFAULT_METADATA_FILENAME.format(
            timestamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        )
        output_file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=default_filename,
            filetypes=[("CSV files", "*.csv")]
        )

        if not output_file_path:
            # ✅ NEW: User cancelled, end operation
            self.button_manager.end_operation()
            return

        # ✅ REMOVE OLD CODE: Delete button disable lines

        self.update_status(
            f"Starting metadata export for {len(selected_objects_list)} objects to {output_file_path}..."
        )
        start_time = time.time()

        # Run export in background thread
        def do_export():
            try:
                output_path, stats = self.metadata_exporter.export_metadata(
                    selected_objects_list,
                    output_file_path
                )

                end_time = time.time()
                runtime_seconds = end_time - start_time
                runtime_formatted = format_runtime(runtime_seconds)

                # Update UI on main thread
                self.after(0, lambda: self._on_metadata_export_success(
                    output_path, stats, runtime_formatted
                ))

            except Exception as e:
                # Handle error on main thread
                self.after(0, lambda: self._on_metadata_export_error(str(e)))

        ThreadHelper.run_in_thread(do_export)

    def _on_metadata_export_success(self, output_path, stats, runtime_formatted):
        """Called after successful metadata export"""
        self.update_status(f"Export Complete! Total Runtime: {runtime_formatted}")
        messagebox.showinfo(
            "Export Done",
            f"Metadata successfully exported to:\n{output_path}"
        )

        print_metadata_statistics(stats, runtime_formatted, output_path)

        # ✅ NEW: Re-enable all buttons
        self.button_manager.end_operation()

    def _on_metadata_export_error(self, error_message):
        """Called when metadata export fails"""
        self.update_status(f"❌ FATAL EXPORT ERROR: {error_message}")
        messagebox.showerror("Export Error", f"A fatal error occurred during export: {error_message}")

        # ✅ NEW: Re-enable all buttons
        self.button_manager.end_operation()

    def download_files_action(self):
        """Handle download files button click"""
        if not self.sf_client or not self.content_document_exporter:
            messagebox.showerror("Error", "Not logged in. Please log in first.")
            return

        # ✅ NEW: Check if another operation is running
        if not self.button_manager.start_operation("File Download"):
            return

        default_filename = DEFAULT_CONTENTDOCUMENT_FILENAME.format(
            timestamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        )
        output_file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=default_filename,
            filetypes=[("CSV files", "*.csv")]
        )

        if not output_file_path:
            # ✅ NEW: User cancelled, end operation
            self.button_manager.end_operation()
            return

        # ✅ REMOVE OLD CODE: Delete these lines if they exist:
        # self.download_files_button.configure(state="disabled", text="Downloading... DO NOT CLOSE")
        # self.export_picklist_button.configure(state="disabled")
        # self.export_metadata_button.configure(state="disabled")
        # self.run_soql_button.configure(state="disabled")

        self.update_status("Starting ContentDocument export and file downloads...")
        start_time = time.time()

        # Run export in background thread
        def do_export():
            try:
                output_path, stats = self.content_document_exporter.export_content_documents(
                    output_file_path
                )

                end_time = time.time()
                runtime_seconds = end_time - start_time
                runtime_formatted = format_runtime(runtime_seconds)

                # Update UI on main thread
                self.after(0, lambda: self._on_download_files_success(
                    output_path, stats, runtime_formatted
                ))

            except Exception as e:
                # Handle error on main thread
                self.after(0, lambda: self._on_download_files_error(str(e)))

        ThreadHelper.run_in_thread(do_export)

    def _on_download_files_success(self, output_path, stats, runtime_formatted):
        """Called after successful file downloads"""
        self.update_status(f"Export Complete! Total Runtime: {runtime_formatted}")

        # Get documents folder path
        csv_dir = os.path.dirname(output_path)
        documents_folder = os.path.join(csv_dir, "Documents")

        messagebox.showinfo(
            "Export Done",
            f"ContentDocument data exported to:\n{output_path}\n\nFiles downloaded to:\n{documents_folder}"
        )

        print_content_document_statistics(stats, runtime_formatted, output_path, documents_folder)

        # ✅ NEW: Re-enable all buttons
        self.button_manager.end_operation()

        # ✅ REMOVE OLD CODE: Delete button re-enable lines

    def _on_download_files_error(self, error_message):
        """Called when file download fails"""
        self.update_status(f"❌ FATAL EXPORT ERROR: {error_message}")
        messagebox.showerror("Export Error", f"A fatal error occurred during export: {error_message}")

        # ✅ NEW: Re-enable all buttons
        self.button_manager.end_operation()

        # ✅ REMOVE OLD CODE: Delete button re-enable lines

    # ==================================
    # Utility Methods
    # ==================================

    def update_status(self, message: str, verbose: bool = False):
        """Updates the GUI status text box with new messages"""
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        display_message = f"{timestamp} {message}"

        self.status_textbox.configure(state="normal")
        self.status_textbox.insert("end", "\n" + display_message)
        self.status_textbox.see("end")

        if not verbose:
            print(display_message)

        self.status_textbox.configure(state="disabled")
        self.update_idletasks()

    def logout_action(self):
        """Clears connection, resets state, and returns to the login screen"""
        confirm = messagebox.askyesno("Logout", "Are you sure you want to log out?")
        if confirm:
            self.sf_client = None
            self.picklist_exporter = None
            self.metadata_exporter = None
            self.content_document_exporter = None
            self.selected_objects.clear()
            self.all_org_objects.clear()
            self.soql_runner = None
            
            # Clear SOQL frame (existing)
            if self.soql_frame:
                try:
                    self.soql_frame.destroy()
                except:
                    pass
                self.soql_frame = None
            
            # Clear switch frame and manager (existing)
            if self.switch_frame:
                try:
                    self.switch_frame.destroy()
                except:
                    pass
                self.switch_frame = None
            self.metadata_switch_manager = None
            
            # Clear report exporter frame
            if self.report_exporter_frame:
                try:
                    if self.report_exporter_frame.winfo_exists():
                        try:
                            self.report_exporter_frame._is_being_destroyed = True
                        except:
                            pass
                        try:
                            self.report_exporter_frame.export_cancel_event.set()
                        except:
                            pass
                        self.after(100, lambda: self._destroy_report_frame())
                    else:
                        self.report_exporter_frame = None
                except Exception as e:
                    print(f"⚠️ Error closing report exporter: {e}")
                    self.report_exporter_frame = None
            
            # # ✅ NEW: Clear all login fields EXCEPT custom domain
            # self.username_entry.delete(0, "end")
            # self.password_entry.delete(0, "end")
            # self.token_entry.delete(0, "end")
            
            # # ✅ Reset org type to Production (but keep custom domain checkbox state)
            # self.org_type_var.set("Production")
            
            # Reset the login button state and text
            self.login_button.configure(state="normal", text="Login to Salesforce")
            
            try:
                self.update_status("Logged out successfully. Please log in again.")
            except:
                print("Logged out successfully. Please log in again.")
            
            # Switch back to Login Frame
            self.export_frame.grid_forget()
            self.login_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
            
            # Show main window if it was hidden
            self.deiconify()

    def _destroy_report_frame(self):
        """Helper method to destroy report exporter frame safely"""
        if self.report_exporter_frame:
            try:
                if self.report_exporter_frame.winfo_exists():
                    self.report_exporter_frame.destroy()
            except Exception as e:
                print(f"⚠️ Error destroying report frame: {e}")
            finally:
                self.report_exporter_frame = None



def main():
    """Main entry point"""
    try:
        app = SalesforceExporterGUI()
        app.mainloop()
    except Exception as e:
        print(f"\n❌ GUI Application Failed: {str(e)}")
        import sys
        sys.exit(1)


if __name__ == "__main__":
    main()