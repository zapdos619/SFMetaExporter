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
        """Setup the login screen UI"""
        login_frame = self.login_frame
        login_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(
            login_frame,
            text="Salesforce Login",
            font=ctk.CTkFont(size=30, weight="bold")
        ).grid(row=0, column=0, columnspan=2, pady=(50, 40))

        def create_input_row(parent, row, label_text, password_mode=False):
            ctk.CTkLabel(
                parent,
                text=label_text,
                anchor="w",
                font=ctk.CTkFont(size=14)
            ).grid(row=row, column=0, padx=10, pady=15, sticky="w")
            entry = ctk.CTkEntry(parent, width=350, show="*" if password_mode else "")
            entry.grid(row=row, column=1, padx=10, pady=15, sticky="ew")
            return entry

        self.username_entry = create_input_row(login_frame, 1, "Username:")
        self.password_entry = create_input_row(login_frame, 2, "Password:", password_mode=True)
        self.token_entry = create_input_row(login_frame, 3, "Security Token:", password_mode=True)

        ctk.CTkLabel(
            login_frame,
            text="Org Type:",
            anchor="w",
            font=ctk.CTkFont(size=14)
        ).grid(row=4, column=0, padx=10, pady=15, sticky="w")

        self.org_type_var = ctk.StringVar(value="Production")
        radio_prod = ctk.CTkRadioButton(
            login_frame,
            text="Production",
            variable=self.org_type_var,
            value="Production"
        )
        radio_test = ctk.CTkRadioButton(
            login_frame,
            text="Sandbox/Test",
            variable=self.org_type_var,
            value="Sandbox"
        )

        radio_prod.grid(row=4, column=1, padx=(10, 5), pady=15, sticky="w")
        radio_test.grid(row=4, column=1, padx=(140, 10), pady=15, sticky="w")

        self.login_button = ctk.CTkButton(
            login_frame,
            text="Login to Salesforce",
            command=self.login_action,
            width=150,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.login_button.grid(row=5, column=0, columnspan=2, pady=50, sticky="ew", padx=10)

    def login_action(self):
        """Handle login button click"""
        self.login_button.configure(state="disabled", text="Connecting...")

        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        token = self.token_entry.get().strip()
        domain = 'test' if self.org_type_var.get() == 'Sandbox' else 'login'

        if not all([username, password, token]):
            messagebox.showerror("Input Error", "All fields (Username, Password, Security Token) are required.")
            self.login_button.configure(state="normal", text="Login to Salesforce")
            return

        # Run login in background thread
        def do_login():
            try:
                self.sf_client = SalesforceClient(
                    username=username,
                    password=password,
                    security_token=token,
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
                self.after(0, lambda: self._on_login_error(str(e)))

        ThreadHelper.run_in_thread(do_login)

    def _on_login_success(self):
        """Called after successful login"""
        messagebox.showinfo("Success", "Successfully connected to Salesforce!")

        # Switch to Export Frame
        self.login_frame.grid_forget()
        self.export_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.populate_available_objects(self.all_org_objects)
        self.populate_selected_objects()
        self.login_button.configure(state="normal", text="Login to Salesforce")

        # Initialize SOQL Runner
        self.soql_runner = SOQLRunner(self.sf_client)
        
        # Initialize Metadata Switch Manager (NEW)
        self.metadata_switch_manager = MetadataSwitchManager(
            self.sf_client.sf,
            status_callback=self.update_status
        )

    def _on_login_error(self, error_message):
        """Called when login fails"""
        messagebox.showerror("Login Failed", f"Connection Error: {error_message}")
        self.sf_client = None
        self.login_button.configure(state="normal", text="Login to Salesforce")

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
            # Simple heuristic for multi-monitor setups
            
            if window_center_x > screen_width:
                # Window is on RIGHT monitor (extended display)
                monitor_x = screen_width
                monitor_y = 0
                monitor_width = screen_width  # Assume same size as primary
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
            Dictionary with window state info:
            {
                'state': 'normal' | 'zoomed' | 'fullscreen',
                'width': int,
                'height': int,
                'x': int,
                'y': int,
                'monitor_x': int,
                'monitor_y': int,
                'monitor_width': int,
                'monitor_height': int
            }
        """
        try:
            # Get window state
            state = self.state()
            
            # Check if zoomed (maximized)
            is_zoomed = (state == 'zoomed')
            
            # Check if fullscreen (on some systems)
            is_fullscreen = self.attributes('-fullscreen') if hasattr(self, 'attributes') else False
            
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



    
    # ✅ NEW METHOD 1 - Report Exporter Action
    def report_exporter_action(self):
        """Handle Report Exporter button click (6th button)"""
        if not self.sf_client:
            messagebox.showerror("Error", "Not logged in. Please log in first.")
            return
        
        # Build session info for Report Exporter
        session_info = {
            "session_id": self.sf_client.session_id,
            "instance_url": self.sf_client.base_url,
            "api_version": self.sf_client.api_version,
            "user_name": self.username_entry.get()  # Get from login
        }
        
        # ✅ FIXED: Create or show the Toplevel window
        if self.report_exporter_frame is None or not self.report_exporter_frame.winfo_exists():
            # Create new window
            self.report_exporter_frame = SalesforceExporterApp(
                master=self,
                session_info=session_info,
                on_logout=self.show_export_frame_from_report_exporter
            )
        else:
            # Window already exists, just show it
            self.report_exporter_frame.deiconify()
            self.report_exporter_frame.lift()
            self.report_exporter_frame.focus_force()
        
        # ✅ FIXED: Hide main window (not grid_forget)
        self.withdraw()
        
        self._log("📊 Opened Report Exporter")
    
    
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
                child_window.attributes('-fullscreen', True)
                self._log("🖥️ Report Exporter: Fullscreen mode")
                
            elif state == 'zoomed':
                # Parent is maximized - maximize child
                child_window.state('zoomed')
                self._log("🖥️ Report Exporter: Maximized mode")
                
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
                
                self._log(f"🖥️ Report Exporter: Normal mode ({child_width}x{child_height})")
            
            # Force window to update
            child_window.update_idletasks()
            
        except Exception as e:
            print(f"⚠️ Error applying parent state to child: {e}")
            # Fallback to default size and position
            try:
                child_window.geometry("1200x740")
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
                self.soql_frame.destroy()
                self.soql_frame = None
            
            # Clear switch frame and manager (NEW)
            if self.switch_frame:
                self.switch_frame.destroy()
                self.switch_frame = None
            self.metadata_switch_manager = None
            
            # ✅ Clear report exporter frame (Toplevel window)
            if self.report_exporter_frame:
                try:
                    if self.report_exporter_frame.winfo_exists():
                        self.report_exporter_frame.destroy()
                except:
                    pass
                self.report_exporter_frame = None
            
            # Reset the login button state and text
            self.login_button.configure(state="normal", text="Login to Salesforce")
            
            self.update_status("Logged out successfully. Please log in again.")
            
            # Switch back to Login Frame
            self.export_frame.grid_forget()
            self.login_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)



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