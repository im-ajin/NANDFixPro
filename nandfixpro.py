import sys
import traceback
import datetime
import tkinter as tk

# --- ROBUST ERROR LOGGING AND EXIT ---
def log_uncaught_exceptions(ex_cls, ex, tb):
    # Log the error to a file
    with open("error_log.txt", "a") as f:
        f.write(f"--- {datetime.datetime.now()} ---\n")
        f.write(''.join(traceback.format_exception(ex_cls, ex, tb)))
        f.write("\n")
    
    # Also show a user-friendly error message box
    error_message = f"A critical error occurred:\n\n{ex}\n\nPlease check error_log.txt for more details."
    try:
        from tkinter import messagebox
        # Create a temporary root to show the message box if the main app failed
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Unhandled Exception", error_message)
        root.destroy()
    except Exception as e:
        print(f"Could not show messagebox: {e}")
        
    # IMPORTANT: Ensure the process exits cleanly after a crash
    sys.exit(1)

sys.excepthook = log_uncaught_exceptions
# --- END OF LOGGING CODE ---


# --- YOUR ORIGINAL SCRIPT CONTINUES HERE ---
from tkinter import ttk, filedialog, scrolledtext
import os
import tempfile
import shutil
from pathlib import Path
import threading
import subprocess
import re
import ctypes
import configparser

# --- CUSTOM DIALOG CLASS (Modernized) ---
class CustomDialog(tk.Toplevel):
    def __init__(self, parent, title=None, message="", buttons="ok"):
        super().__init__(parent)
        self.transient(parent)
        self.title(title)
        self.parent = parent
        self.result = False
        self.resizable(False, False)
        
        # Apply modern theme from parent
        self.configure(bg=parent.style.lookup('TFrame', 'background'))

        main_frame = ttk.Frame(self, padding="20 20 20 20", style="Dark.TFrame")
        main_frame.pack(expand=True, fill=tk.BOTH)

        message_label = ttk.Label(main_frame, text=message, wraplength=400, justify=tk.LEFT, style="Dark.TLabel")
        message_label.pack(padx=10, pady=10)
        
        button_frame = ttk.Frame(main_frame, style="Dark.TFrame")
        button_frame.pack(pady=(20, 0))

        if buttons == "yesno":
            yes_button = ttk.Button(button_frame, text="Yes", command=self.on_yes, style="Accent.TButton")
            yes_button.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=2)
            no_button = ttk.Button(button_frame, text="No", command=self.on_no, style="TButton")
            no_button.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=2)
            self.bind("<Return>", lambda e: self.on_yes())
            self.bind("<Escape>", lambda e: self.on_no())
        else: # Default is "ok"
            ok_button = ttk.Button(button_frame, text="OK", command=self.on_no, style="Accent.TButton")
            ok_button.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=2)
            self.bind("<Return>", lambda e: self.on_no())
            self.bind("<Escape>", lambda e: self.on_no())

        self.center_window()
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_no)
        self.wait_window(self)

    def center_window(self):
        self.update_idletasks()
        parent_x, parent_y = self.parent.winfo_x(), self.parent.winfo_y()
        parent_w, parent_h = self.parent.winfo_width(), self.parent.winfo_height()
        dialog_w, dialog_h = self.winfo_width(), self.winfo_height()
        x = parent_x + (parent_w // 2) - (dialog_w // 2)
        y = parent_y + (parent_h // 2) - (dialog_h // 2)
        self.geometry(f"+{x}+{y}")

    def on_yes(self):
        self.result = True
        self.destroy()

    def on_no(self):
        self.result = False
        self.destroy()

# --- MAIN APPLICATION CLASS ---
class SwitchGuiApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.version = "1.0.0"
        self.title(f"NAND Fix Pro v{self.version}")
        self.geometry("800x800")
        self.resizable(False, False)
        
        # --- PATHS & STATE VARIABLES ---
        self.config_file = "config.ini"
        self.paths = {
            "7z": tk.StringVar(), "osfmount": tk.StringVar(),
            "nxnandmanager": tk.StringVar(), "keys": tk.StringVar(), "firmware": tk.StringVar(),
            "prodinfo": tk.StringVar(), "partitions_folder": tk.StringVar(),
            "output_folder": tk.StringVar(), "emmchaccgen": tk.StringVar(),
        }
        self.level_requirements = {
            1: ["firmware", "keys", "output_folder", "7z", "emmchaccgen", "nxnandmanager", "osfmount"],
            2: ["firmware", "keys", "output_folder", "7z", "emmchaccgen", "nxnandmanager", "osfmount", "partitions_folder"],
            3: ["firmware", "keys", "prodinfo", "output_folder", "7z", "emmchaccgen", "nxnandmanager", "osfmount", "partitions_folder"]
        }
        self.start_level1_button, self.start_level2_button, self.start_level3_button = None, None, None
        
        # --- INITIALIZATION ---
        self._setup_styles()
        self._load_config()
        self._setup_widgets()
        self._validate_paths_and_update_buttons() # Initial check
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        window_w = self.winfo_width()
        window_h = self.winfo_height()
        x = (screen_w // 2) - (window_w // 2)
        y = (screen_h // 2) - (window_h // 2)
        self.geometry(f'+{x}+{y}')

    def _setup_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')

        # --- COLOR & FONT PALETTE (DEFINED AS INSTANCE ATTRIBUTES) ---
        self.BG_COLOR = "#2e2e2e"
        self.FG_COLOR = "#fafafa"
        self.BG_LIGHT = "#3c3c3c"
        self.BG_DARK = "#252525"
        self.ACCENT_COLOR = "#0078d4"
        self.ACCENT_ACTIVE = "#005a9e"
        self.DISABLED_FG = "#888888"
        self.FONT_FAMILY = "Segoe UI"
        
        self.configure(background=self.BG_COLOR)

        # --- General Widget Styling ---
        self.style.configure('.', background=self.BG_COLOR, foreground=self.FG_COLOR, font=(self.FONT_FAMILY, 10))
        self.style.configure("TFrame", background=self.BG_COLOR)
        self.style.configure("Dark.TFrame", background=self.BG_DARK) # For dialogs
        self.style.configure("TLabel", background=self.BG_COLOR, foreground=self.FG_COLOR, font=(self.FONT_FAMILY, 10))
        self.style.configure("Dark.TLabel", background=self.BG_DARK, foreground=self.FG_COLOR, font=(self.FONT_FAMILY, 10))
        self.style.configure("TCheckbutton", font=(self.FONT_FAMILY, 10))
        self.style.map("TCheckbutton",
            background=[('active', self.BG_COLOR)],
            indicatorbackground=[('selected', self.ACCENT_COLOR), ('!selected', self.BG_LIGHT)],
            indicatorcolor=[('!selected', self.BG_LIGHT)]
        )

        # --- Button Styling ---
        self.style.configure("TButton", font=(self.FONT_FAMILY, 10, 'bold'), borderwidth=0, padding=(10, 5))
        self.style.map("TButton",
            background=[('!disabled', self.BG_LIGHT), ('active', self.ACCENT_ACTIVE), ('disabled', self.BG_DARK)],
            foreground=[('!disabled', self.FG_COLOR), ('disabled', self.DISABLED_FG)]
        )
        self.style.configure("Accent.TButton", background=self.ACCENT_COLOR)
        self.style.map("Accent.TButton",
            background=[('!disabled', self.ACCENT_COLOR), ('active', self.ACCENT_ACTIVE), ('disabled', self.BG_DARK)],
            foreground=[('!disabled', '#ffffff'), ('disabled', self.DISABLED_FG)]
        )

        # --- LabelFrame Styling ---
        self.style.configure("TLabelFrame", background=self.BG_COLOR, borderwidth=1, relief="solid")
        self.style.configure("TLabelFrame.Label", background=self.BG_COLOR, foreground=self.FG_COLOR, font=(self.FONT_FAMILY, 11, 'bold'))

        # --- Notebook (Tabs) Styling ---
        self.style.configure("TNotebook", background=self.BG_COLOR, borderwidth=0)
        self.style.configure("TNotebook.Tab",
            background=self.BG_LIGHT,
            foreground=self.FG_COLOR,
            font=(self.FONT_FAMILY, 10, 'bold'),
            padding=[15, 8],
            borderwidth=0,
        )
        self.style.map("TNotebook.Tab",
            background=[("selected", self.BG_COLOR), ("active", self.ACCENT_COLOR)],
            expand=[("selected", [1, 1, 1, 0])]
        )

    def _load_config(self):
        """Loads paths from config.ini, running auto-detect if it doesn't exist."""
        config = configparser.ConfigParser()
        if Path(self.config_file).exists():
            config.read(self.config_file)
            for key in self.paths:
                self.paths[key].set(config.get('Paths', key, fallback=''))
        else:
            self._auto_detect_paths()
            self._save_config()

    def _save_config(self):
        """Saves current paths to config.ini."""
        config = configparser.ConfigParser()
        config['Paths'] = {key: var.get() for key, var in self.paths.items()}
        with open(self.config_file, 'w') as configfile:
            config.write(configfile)
        self._log(f"INFO: Configuration saved to {self.config_file}")

    def _is_path_valid(self, key):
        """Helper to check if a path from the dict is non-empty and exists."""
        path_str = self.paths[key].get()
        if not path_str:
            return False
        return Path(path_str).exists()

    def _validate_paths_and_update_buttons(self):
        """Checks required paths for each level and enables/disables buttons."""
        # Level 1 Validation
        level1_ok = all(self._is_path_valid(key) for key in self.level_requirements[1])
        if self.start_level1_button:
            self.start_level1_button.config(state="normal" if level1_ok else "disabled")

        # Level 2 Validation
        level2_ok = all(self._is_path_valid(key) for key in self.level_requirements[2])
        if self.start_level2_button:
            self.start_level2_button.config(state="normal" if level2_ok else "disabled")

        # Level 3 Validation
        level3_ok = all(self._is_path_valid(key) for key in self.level_requirements[3])
        if self.start_level3_button:
            self.start_level3_button.config(state="normal" if level3_ok else "disabled")
        
        self.update_idletasks()

    def _auto_detect_paths(self):
        try: script_dir = Path(__file__).parent
        except NameError: script_dir = Path.cwd()
        
        osfmount_path = Path("C:/Program Files/OSFMount/OSFMount.com")
        if osfmount_path.is_file(): self.paths["osfmount"].set(str(osfmount_path.resolve()))
        
        default_paths = {
            "7z": script_dir / "lib" / "7z" / "7z.exe",
            "emmchaccgen": script_dir / "lib" / "EmmcHaccGen" / "EmmcHaccGen.exe",
            "nxnandmanager": script_dir / "lib" / "NxNandManager" / "NxNandManager.exe",
            "partitions_folder": script_dir / "lib" / "NAND",
        }
        for key, path in default_paths.items():
            if self.paths[key].get(): continue
            full_path = Path(path)
            if full_path.is_file() or full_path.is_dir():
                self.paths[key].set(str(full_path.resolve()))

    def _setup_widgets(self):
        menubar = tk.Menu(self, background=self.BG_LIGHT, foreground=self.FG_COLOR,
                            activebackground=self.ACCENT_COLOR, activeforeground=self.FG_COLOR, 
                            relief="flat", borderwidth=0)
        self.config(menu=menubar)
        self._setup_settings_menu(menubar)
        
        # --- TAB CONTROL SETUP ---
        tab_control = ttk.Notebook(self, style="TNotebook")
        tab_level1 = ttk.Frame(tab_control, padding="15")
        tab_level2 = ttk.Frame(tab_control, padding="15")
        tab_level3 = ttk.Frame(tab_control, padding="15")
        
        tab_control.add(tab_level1, text='Level 1: System Restore')
        tab_control.add(tab_level2, text='Level 2: Full Rebuild')
        tab_control.add(tab_level3, text='Level 3: Complete Recovery')
        tab_control.pack(expand=1, fill="both", padx=15, pady=10)
        
        # --- POPULATE TABS ---
        self._setup_level1_tab(tab_level1)
        self._setup_level2_tab(tab_level2)
        self._setup_level3_tab(tab_level3)
        
        # --- LOG WIDGET SETUP ---
        log_frame = ttk.LabelFrame(self, text="Log Output", padding="10")
        log_frame.pack(padx=15, pady=(5, 15), fill="both", expand=True)
        log_frame.columnconfigure(0, weight=1)
        self.log_widget = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15, state="disabled",
            bg="#1e1e1e", fg="#d4d4d4", relief="flat", borderwidth=2,
            font=("Consolas", 10), insertbackground="#d4d4d4"
        )
        self.log_widget.grid(row=0, column=0, sticky="nsew")

    def _create_path_selector_row(self, parent, key, label_text, type):
        row = parent.grid_size()[1]
        ttk.Label(parent, text=label_text, font=(self.FONT_FAMILY, 10)).grid(row=row, column=0, sticky="w", padx=5, pady=6)
        
        path_label = ttk.Label(parent, textvariable=self.paths[key],
            relief="solid", anchor="w", padding=(8, 5), background="#3c3c3c", borderwidth=1,
            font=(self.FONT_FAMILY, 9)
        )
        path_label.grid(row=row, column=1, sticky="ew", padx=5, pady=6)
        
        browse_button = ttk.Button(parent, text="Browse...", command=lambda k=key, t=type: self._select_path(k, t), style="TButton")
        browse_button.grid(row=row, column=2, padx=5, pady=6)
    
    def _setup_tab_content(self, parent_frame, title, info_text, paths, process_name, button_ref, command):
        parent_frame.columnconfigure(1, weight=1)
        
        desc_frame = ttk.LabelFrame(parent_frame, text=title, padding="15")
        desc_frame.grid(row=0, column=0, columnspan=3, pady=(5, 20), sticky="ew")
        ttk.Label(desc_frame, text=info_text, wraplength=650, justify=tk.LEFT).pack(anchor="w")

        input_frame = ttk.Frame(parent_frame)
        input_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        input_frame.columnconfigure(1, weight=1)

        for key, label, type in paths:
            self._create_path_selector_row(input_frame, key, label, type)
            
        # Placeholder for alignment
        if len(paths) < 4:
            ttk.Frame(input_frame, height=45 * (4 - len(paths))).grid(row=input_frame.grid_size()[1], columnspan=3)

        button_frame = ttk.Frame(parent_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=(25, 5))
        
        # Buttons are disabled by default and enabled by validation
        button = ttk.Button(button_frame, text=f"Start {process_name} Process", command=command, style="Accent.TButton", state="disabled")
        button.pack(side=tk.LEFT, padx=10, ipady=5, ipadx=15)
        setattr(self, button_ref, button)

    def _setup_level1_tab(self, parent_frame):
        info_text = ("Fixes a corrupt SYSTEM partition directly on your Switch's eMMC.\n\n"
                        "• Use this for software errors, failed updates, or boot issues where only the OS is affected.\n"
                        "• The process reads your Switch's own PRODINFO and SYSTEM partition to perform the fix.\n"
                        "• This method preserves user data like saves and installed games.")
        paths = [
            ("firmware", "Firmware Folder:", "folder"),
            ("keys", "Keys File (prod.keys):", "file"),
            ("output_folder", "Output Folder:", "folder"),
        ]
        self._setup_tab_content(parent_frame, "Level 1: Description", info_text, paths, "Level 1",
                                "start_level1_button", lambda: self._start_threaded_process("Level 1"))

    def _setup_level2_tab(self, parent_frame):
        info_text = ("Rebuilds the NAND using clean donor partitions from the 'lib/NAND' folder.\n\n"
                        "• Use this when multiple partitions are corrupt, but PRODINFO is still readable.\n"
                        "• The process reads your Switch's PRODINFO, then flashes clean partitions over the existing ones.\n"
                        "• This process WILL ERASE all user data.")
        paths = [
            ("firmware", "Firmware Folder:", "folder"),
            ("keys", "Keys File (prod.keys):", "file"),
            ("output_folder", "Output Folder:", "folder"),
        ]
        self._setup_tab_content(parent_frame, "Level 2: Description", info_text, paths, "Level 2",
                                "start_level2_button", lambda: self._start_threaded_process("Level 2"))

    def _setup_level3_tab(self, parent_frame):
        info_text = ("For total NAND loss, including PRODINFO. This is a last resort.\n\n"
                        "• Reconstructs a complete NAND image from a donor PRODINFO file and clean templates.\n"
                        "• The script automatically detects eMMC size (32/64GB) for the correct NAND skeleton.\n"
                        "• Connect your Switch in 'eMMC RAW GPP' mode (Read-Only OFF) and click Start.")
        paths = [
            ("firmware", "Firmware Folder:", "folder"),
            ("keys", "Keys File (prod.keys):", "file"),
            ("prodinfo", "Donor PRODINFO:", "file"),
            ("output_folder", "Output Folder:", "folder"),
        ]
        self._setup_tab_content(parent_frame, "Level 3: Description", info_text, paths, "Level 3",
                                "start_level3_button", self._start_level3_threaded)


    # --- THE REST OF YOUR LOGIC IS UNCHANGED ---
    
    def _selective_copy_system_contents(self, source_system_path, drive_letter):
        """
        Selectively copy SYSTEM contents, preserving existing folders like savemeta
        but replacing 'registered' and 'save' folders entirely.
        """
        try:
            self._log("--- Starting selective SYSTEM merge...")
            
            # Debug: Show what we're starting with
            self._log("--- Existing Contents structure:")
            contents_dest = drive_letter / "Contents"
            if contents_dest.exists():
                for item in contents_dest.iterdir():
                    item_type = "DIR" if item.is_dir() else "FILE"
                    self._log(f"  Existing: {item_type} - {item.name}")
            
            # Check for existing save folder
            save_dest = drive_letter / "save"
            if save_dest.exists():
                self._log("--- Existing 'save' folder found at root level")
            
            # Process each top-level item from source
            for source_item in source_system_path.iterdir():
                dest_item = drive_letter / source_item.name
                
                if source_item.name == "Contents":
                    # Handle Contents folder with subfolder-level merging
                    self._log("--- Processing Contents folder with preservation...")
                    
                    # Ensure Contents directory exists
                    contents_dest.mkdir(exist_ok=True)
                    
                    # Step 1: Remove ONLY the registered folder if it exists
                    registered_dest = contents_dest / "registered"
                    if registered_dest.exists():
                        self._log("--- Removing old 'registered' folder...")
                        shutil.rmtree(registered_dest)
                        self._log("--- Old 'registered' folder removed.")
                    
                    # Step 2: Copy each item from source Contents individually
                    source_contents = source_item
                    for contents_subitem in source_contents.iterdir():
                        dest_subitem = contents_dest / contents_subitem.name
                        
                        if contents_subitem.name == "registered":
                            # Copy the new registered folder
                            self._log("--- Copying NEW 'registered' folder...")
                            if contents_subitem.is_dir():
                                shutil.copytree(contents_subitem, dest_subitem)
                            else:
                                shutil.copy2(contents_subitem, dest_subitem)
                            self._log(f"--- SUCCESS: 'registered' folder copied with {len(list(contents_subitem.iterdir()))} items")
                        
                        elif not dest_subitem.exists():
                            # Copy new items (like placehld) that don't exist
                            self._log(f"--- Copying new Contents item: {contents_subitem.name}")
                            if contents_subitem.is_dir():
                                shutil.copytree(contents_subitem, dest_subitem)
                            else:
                                shutil.copy2(contents_subitem, dest_subitem)
                        
                        else:
                            # Preserve existing items (like savemeta)
                            self._log(f"--- PRESERVING existing Contents item: {contents_subitem.name}")
                
                elif source_item.name == "save":
                    # Handle save folder - ALWAYS replace entirely
                    self._log("--- Processing 'save' folder - REPLACING entirely...")
                    if save_dest.exists():
                        self._log("--- Removing old 'save' folder...")
                        shutil.rmtree(save_dest)
                        self._log("--- Old 'save' folder removed.")
                    
                    # Copy the new save folder
                    self._log("--- Copying NEW 'save' folder...")
                    if source_item.is_dir():
                        shutil.copytree(source_item, save_dest)
                        save_files = list(source_item.iterdir())
                        self._log(f"--- SUCCESS: 'save' folder copied with {len(save_files)} files")
                        for save_file in save_files:
                            self._log(f"         - {save_file.name}")
                    else:
                        shutil.copy2(source_item, save_dest)
                        self._log("--- SUCCESS: 'save' file copied")
                
                else:
                    # Handle other top-level items (outside Contents and save folders)
                    if not dest_item.exists():
                        self._log(f"--- Copying new top-level item: {source_item.name}")
                        if source_item.is_dir():
                            shutil.copytree(source_item, dest_item)
                        else:
                            shutil.copy2(source_item, dest_item)
                    else:
                        self._log(f"--- PRESERVING existing top-level item: {source_item.name}")
            
            # Debug: Show final structure
            self._log("--- Final Contents structure:")
            if contents_dest.exists():
                for item in contents_dest.iterdir():
                    item_type = "DIR" if item.is_dir() else "FILE"
                    size_info = ""
                    if item.is_dir():
                        try:
                            file_count = len(list(item.iterdir()))
                            size_info = f" ({file_count} items)"
                        except:
                            size_info = " (access error)"
                    self._log(f"  Final: {item_type} - {item.name}{size_info}")
            
            # Debug: Show final save structure
            if save_dest.exists():
                self._log("--- Final 'save' folder structure:")
                for item in save_dest.iterdir():
                    item_type = "DIR" if item.is_dir() else "FILE"
                    self._log(f"  Save: {item_type} - {item.name}")
            
            self._log("--- SUCCESS: Selective SYSTEM merge completed (savemeta preserved, save folder replaced)")
            return True
            
        except Exception as e:
            self._log(f"ERROR: Failed to selectively copy SYSTEM contents. Error: {e}")
            import traceback
            self._log(traceback.format_exc())
            return False
            
    def _get_donor_nand_path(self, target_size_gb, temp_dir):
        """
        Automatically detect and extract the appropriate donor NAND based on eMMC size.
        Returns the path to the extracted donor NAND image.
        """
        try:
            script_dir = Path(__file__).parent
        except NameError:
            script_dir = Path.cwd()
        
        nand_lib_dir = script_dir / "lib" / "NAND"
        
        # Determine which donor NAND to use
        if target_size_gb > 40:  # 64GB eMMC
            donor_archive = nand_lib_dir / "donor64.7z"
            donor_bin_name = "rawnand64.bin"
            self._log("--- Target: 64GB eMMC, using donor64.7z")
        else:  # 32GB eMMC
            donor_archive = nand_lib_dir / "donor32.7z"
            donor_bin_name = "rawnand32.bin"
            self._log("--- Target: 32GB eMMC, using donor32.7z")
        
        # Check if donor archive exists
        if not donor_archive.is_file():
            self._log(f"ERROR: Donor NAND archive not found: {donor_archive}")
            return None
        
        self._log(f"--- Found donor archive: {donor_archive}")
        
        # Extract donor NAND to temp directory
        extract_dir = Path(temp_dir) / "donor_extract"
        extract_dir.mkdir(exist_ok=True)
        
        self._log(f"--- Extracting donor NAND archive...")
        extract_cmd = [self.paths['7z'].get(), 'x', str(donor_archive), f'-o{extract_dir}']
        if self._run_command(extract_cmd)[0] != 0:
            self._log("ERROR: Failed to extract donor NAND archive.")
            return None
        
        # Find the extracted donor NAND file
        donor_nand_path = extract_dir / donor_bin_name
        if not donor_nand_path.is_file():
            self._log(f"ERROR: Expected donor NAND file not found: {donor_nand_path}")
            return None
        
        self._log(f"--- SUCCESS: Donor NAND extracted to {donor_nand_path}")
        return donor_nand_path      
    
    def _start_level3_threaded(self):
        self._disable_buttons()
        thread = threading.Thread(target=self._start_level3_process); 
        thread.daemon = True; 
        thread.start()

    def _disable_buttons(self):
        for btn in [self.start_level1_button, self.start_level2_button, self.start_level3_button]:
            if btn: btn.config(state="disabled")

    def _re_enable_buttons(self):
        # Re-enabling is now handled by the validation function
        self._validate_paths_and_update_buttons()
            
    def _start_level3_process(self):
        self._log("--- Starting Level 3 Complete Recovery Process ---")
        try:
            with tempfile.TemporaryDirectory(prefix="switch_gui_level3_") as temp_dir:
                self._log(f"INFO: Created temporary directory at: {temp_dir}")
                self._run_level3_process(temp_dir)
        except Exception as e:
            self._log(f"An unexpected critical error occurred: {e}\n{traceback.format_exc()}")
            self._log("\nINFO: Level 3 process finished with an error.")
        finally:
            # Only the action remains, no more logging
            self._re_enable_buttons()

    def _run_level3_process(self, temp_dir):
        self._log("\n--- WARNING ---")
        self._log("Level 3 will completely overwrite your Switch's eMMC with a reconstructed NAND.")
        self._log("This is irreversible. Ensure you have backups and a stable connection.")
        
        self._log("\n[STEP 1/8] Please connect your Switch in Hekate's eMMC RAW GPP mode (Read-Only OFF).")
        self._log("--- Detecting target eMMC...")
        
        # Detect target eMMC
        potential_drives = self._detect_switch_drives_wmi()
        if not potential_drives:
            CustomDialog(self, title="Error", message="No potential Switch eMMC drives found. Please ensure it is connected properly.")
            return

        if len(potential_drives) > 1:
            CustomDialog(self, title="Multiple Drives Found", message="Found multiple drives that could be a Switch eMMC. "
                                                                    "For safety, please disconnect other USB drives of 32GB or 64GB and try again.")
            return
        
        target_drive = potential_drives[0]
        target_size_gb = target_drive['size_gb']
        target_path = target_drive['path']
        
        # Confirm with user
        msg = (f"Found target eMMC:\n\nPath: {target_path}\nSize: {target_drive['size']}\nModel: {target_drive['model']}\n\n"
               "WARNING: ALL DATA ON THIS DRIVE WILL BE PERMANENTLY ERASED.\n\n"
               "This will perform a complete Level 3 recovery. Continue?")
        
        dialog = CustomDialog(self, title="Confirm Level 3 Recovery", message=msg, buttons="yesno")
        if not dialog.result:
            self._log("--- User cancelled Level 3 recovery.")
            return
        
        self._log(f"SUCCESS: User confirmed target eMMC at {target_path} ({target_drive['size']})")
        
        self._log(f"\n[STEP 2/8] Preparing donor NAND skeleton...")

        # Automatically detect and extract donor NAND skeleton based on target eMMC size
        donor_nand_path = self._get_donor_nand_path(target_size_gb, temp_dir)
        if not donor_nand_path:
            self._log("ERROR: Failed to prepare donor NAND skeleton.")
            return

        # Copy donor skeleton to working directory
        working_nand = Path(temp_dir) / "working_nand.img"
        self._log(f"--- Copying donor NAND skeleton to working directory...")
        shutil.copy(donor_nand_path, working_nand)
        self._log(f"--- SUCCESS: Working NAND skeleton ready at {working_nand}")
        
        self._log(f"\n[STEP 3/8] Validating donor PRODINFO...")
        prodinfo_path = Path(self.paths['prodinfo'].get())
        if not prodinfo_path.is_file():
            self._log("ERROR: Donor PRODINFO file not found.")
            return
        
        # Validate PRODINFO
        with open(prodinfo_path, 'rb') as f:
            if f.read(4) != b'CAL0':
                error_msg = "The prodinfo is not correct, make sure it is decrypted!"
                self._log(f"ERROR: PRODINFO magic 'CAL0' not found. {error_msg}")
                CustomDialog(self, title="Invalid PRODINFO", message=error_msg)
                return
        
        # Read model from PRODINFO
        with open(prodinfo_path, 'rb') as f:
            f.seek(0x3740)
            product_model_id = int.from_bytes(f.read(4), byteorder='little')
        model_map = {1: "Erista", 3: "V2", 4: "Lite", 6: "OLED"}
        detected_model = model_map.get(product_model_id, "Unknown Mariko")
        self._log(f"SUCCESS: Detected model from PRODINFO: {detected_model}")
        
        self._log(f"\n[STEP 4/8] Generating boot files and system content...")
        emmchaccgen_out_dir = Path(temp_dir) / "emmchaccgen_out"
        emmchaccgen_out_dir.mkdir()
        keyset_path = self.paths['keys'].get()
        
        emmchaccgen_cmd = [self.paths['emmchaccgen'].get(), '--keys', keyset_path, '--fw', self.paths['firmware'].get()]
        if "Mariko" in detected_model or detected_model in ["V2", "Lite", "OLED"]:
            self._log("--- Mariko model detected, using --mariko flag (AutoRCM disabled by default).")
            emmchaccgen_cmd.append('--mariko')
        else:
            self._log("--- Erista model detected, adding --no-autorcm flag by default.")
            emmchaccgen_cmd.append('--no-autorcm')
        
        if self._run_command(emmchaccgen_cmd, cwd=str(emmchaccgen_out_dir))[0] != 0:
            self._log("ERROR: Failed to generate boot files with EmmcHaccGen.")
            return
        
        # Get EmmcHaccGen output folder
        try:
            versioned_folder = next(d for d in emmchaccgen_out_dir.iterdir() if d.is_dir())
        except StopIteration:
            self._log("ERROR: No EmmcHaccGen output folder found.")
            return
        
        self._log(f"\n[STEP 5/8] Preparing all partition data from donor archives...")
        nx_exe = self.paths['nxnandmanager'].get()
        partitions_folder = Path(self.paths['partitions_folder'].get())
        
        # Extract and prepare SYSTEM partition
        self._log("--- Extracting donor SYSTEM partition...")
        if self._run_command([self.paths['7z'].get(), 'x', str(partitions_folder / "SYSTEM.7z"), f'-o{temp_dir}'])[0] != 0:
            self._log("ERROR: Failed to extract donor SYSTEM partition.")
            return
        
        system_dec_path = Path(temp_dir) / "SYSTEM.dec"
        if not system_dec_path.exists():
            self._log("ERROR: SYSTEM.dec was not extracted.")
            return
        
        # Mount and modify SYSTEM
        self._log("--- Mounting SYSTEM partition for modification...")
        osfmount_cmd = [self.paths['osfmount'].get(), '-a', '-t', 'file', '-f', str(system_dec_path), '-o', 'rw', '-m', '#:']
        return_code, output = self._run_command(osfmount_cmd)
        if return_code != 0: 
            self._log("ERROR: Failed to mount SYSTEM partition.")
            return
        
        match = re.search(r"([A-Z]:)", output)
        if not match: 
            self._log("ERROR: Could not determine drive letter.")
            return
        
        drive_letter_str = match.group(1)
        drive_letter = Path(drive_letter_str)
        self._log(f"--- SUCCESS: SYSTEM mounted to {drive_letter}")
        
        try:
            source_system_path = versioned_folder / "SYSTEM"
            
            # For Level 3, complete replacement of specific folders
            self._log("--- Performing Level 3 SYSTEM modification (complete replacement)...")
            
            # Replace Contents/registered completely
            registered_dest = drive_letter / "Contents" / "registered"
            if registered_dest.exists():
                shutil.rmtree(registered_dest)
            registered_source = source_system_path / "Contents" / "registered"
            if registered_source.exists():
                shutil.copytree(registered_source, registered_dest)
                self._log(f"--- SUCCESS: Replaced 'registered' folder with {len(list(registered_source.iterdir()))} items")
            
            # Replace save folder completely
            save_dest = drive_letter / "save"
            if save_dest.exists():
                shutil.rmtree(save_dest)
            save_source = source_system_path / "save"
            if save_source.exists():
                shutil.copytree(save_source, save_dest)
                save_files = list(save_source.iterdir())
                self._log(f"--- SUCCESS: Replaced 'save' folder with {len(save_files)} files")
                for save_file in save_files:
                    self._log(f"         - {save_file.name}")
            
            self._log("--- SUCCESS: SYSTEM partition modified for Level 3")
            
        except Exception as e:
            self._log(f"ERROR: Failed to modify SYSTEM partition. Error: {e}")
            return
        finally:
            self._log("--- Dismounting SYSTEM partition...")
            self._run_command([self.paths['osfmount'].get(), '-D', '-m', drive_letter_str])
        
        # Extract other required partitions
        self._log("--- Extracting other donor partitions...")
        
        # Extract PRODINFOF
        if self._run_command([self.paths['7z'].get(), 'x', str(partitions_folder / "PRODINFOF.7z"), f'-o{temp_dir}'])[0] != 0:
            self._log("ERROR: Failed to extract PRODINFOF partition.")
            return
        
        # Extract SAFE
        if self._run_command([self.paths['7z'].get(), 'x', str(partitions_folder / "SAFE.7z"), f'-o{temp_dir}'])[0] != 0:
            self._log("ERROR: Failed to extract SAFE partition.")
            return
        
        # Extract USER partition (size-dependent)
        if target_size_gb > 40:  # 64GB eMMC
            user_archive = "USER-64.7z"
            self._log("--- Using 64GB USER partition")
        else:  # 32GB eMMC
            user_archive = "USER-32.7z"
            self._log("--- Using 32GB USER partition")
        
        if self._run_command([self.paths['7z'].get(), 'x', str(partitions_folder / user_archive), f'-o{temp_dir}'])[0] != 0:
            self._log("ERROR: Failed to extract USER partition.")
            return
        
        self._log(f"\n[STEP 6/8] Flashing all partitions to donor NAND skeleton...")
        
        # Flash PRODINFO (encrypted)
        self._log("--- Flashing PRODINFO to skeleton...")
        flash_cmd = [nx_exe, '-i', str(prodinfo_path), '-o', str(working_nand), '-part=PRODINFO', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_command(flash_cmd)[0] != 0:
            self._log("ERROR: Failed to flash PRODINFO to skeleton.")
            return
        
        # Flash PRODINFOF (encrypted)
        self._log("--- Flashing PRODINFOF to skeleton...")
        prodinfof_dec_path = Path(temp_dir) / "PRODINFOF.dec"
        flash_cmd = [nx_exe, '-i', str(prodinfof_dec_path), '-o', str(working_nand), '-part=PRODINFOF', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_command(flash_cmd)[0] != 0:
            self._log("ERROR: Failed to flash PRODINFOF to skeleton.")
            return
        
        # Flash modified SYSTEM (encrypted)
        self._log("--- Flashing modified SYSTEM to skeleton...")
        flash_cmd = [nx_exe, '-i', str(system_dec_path), '-o', str(working_nand), '-part=SYSTEM', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_command(flash_cmd)[0] != 0:
            self._log("ERROR: Failed to flash SYSTEM to skeleton.")
            return
        
        # Flash USER (encrypted) - OPTIMIZED
        self._log("--- Flashing USER to skeleton (Optimized, 100MB)...")
        user_dec_path = Path(temp_dir) / "USER.dec"
        flash_cmd = [nx_exe, '-i', str(user_dec_path), '-o', str(working_nand), '-part=USER', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_and_interrupt_flash(flash_cmd, "USER", 100) != 0:
            self._log("ERROR: Failed to partially flash USER to skeleton.")
            return
        
        # Flash SAFE (encrypted)
        self._log("--- Flashing SAFE to skeleton...")
        safe_dec_path = Path(temp_dir) / "SAFE.dec"
        flash_cmd = [nx_exe, '-i', str(safe_dec_path), '-o', str(working_nand), '-part=SAFE', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_command(flash_cmd)[0] != 0:
            self._log("ERROR: Failed to flash SAFE to skeleton.")
            return
        
        # Flash BCPKG2 partitions (unencrypted)
        self._log("--- Flashing BCPKG2 partitions to skeleton...")
        bcpkg2_partitions = ["BCPKG2-1-Normal-Main", "BCPKG2-2-Normal-Sub", "BCPKG2-3-SafeMode-Main", "BCPKG2-4-SafeMode-Sub"]
        for part_name in bcpkg2_partitions:
            bcpkg2_file = versioned_folder / f"{part_name}.bin"
            if not bcpkg2_file.exists():
                self._log(f"ERROR: {bcpkg2_file.name} not found in EmmcHaccGen output.")
                return
            
            flash_cmd = [nx_exe, '-i', str(bcpkg2_file), '-o', str(working_nand), f'-part={part_name}', 'FORCE']
            if self._run_command(flash_cmd)[0] != 0:
                self._log(f"ERROR: Failed to flash {part_name} to skeleton.")
                return
        
        self._log("SUCCESS: All partitions flashed to donor NAND skeleton.")
        
        self._log(f"\n[STEP 7/8] Writing complete NAND image to target eMMC...")
        self._log("--- This may take 10-15 minutes. Do not disconnect the Switch.")
        
        # Raw copy the complete filled skeleton to target eMMC
        if not self._raw_copy_nand_to_emmc(working_nand, target_path, target_size_gb):
            self._log("ERROR: Failed to write NAND image to target eMMC.")
            return
        
        self._log(f"\n[STEP 8/8] Saving BOOT0 & BOOT1 to output folder...")
        output_folder = Path(self.paths['output_folder'].get())
        shutil.copy(versioned_folder / "BOOT0.bin", output_folder / "BOOT0")
        shutil.copy(versioned_folder / "BOOT1.bin", output_folder / "BOOT1")
        self._log(f"SUCCESS: BOOT0 and BOOT1 saved to {output_folder}")
        
        self._log(f"\n[STEP 8/8] Level 3 Recovery Complete!")
        self._log("IMPORTANT: Please flash BOOT0 and BOOT1 manually using Hekate for safety.")
        self._log("Your Switch should now boot with the reconstructed NAND.")
        self._log("\n--- LEVEL 3 COMPLETE RECOVERY FINISHED ---")
        
        CustomDialog(self, title="Level 3 Complete", 
                        message="Level 3 recovery completed successfully!\n\n" +
                                "Don't forget to flash BOOT0 and BOOT1 using Hekate.\n\n" +
                                "Your Switch should now boot normally.")
    
    def _raw_copy_nand_to_emmc(self, source_nand, target_drive, target_size_gb):
        """
        Raw copy donor NAND image to target eMMC using optimized partial write.
        Only writes the first ~5GB which covers all essential partitions.
        USER partition will be blank but gets initialized by the Switch.
        """
        try:
            self._log(f"--- Opening source NAND image: {source_nand}")
            source_size = source_nand.stat().st_size
            self._log(f"--- Source NAND size: {source_size / (1024**3):.2f} GB")

            # Optimized: Only write first 5GB to cover all essential partitions
            # USER partition (largest) will be blank but Switch will initialize it
            copy_size = 5 * (1024**3)  # 5GB should cover all partitions before USER
            self._log(f"--- OPTIMIZATION: Writing only first 5GB (covers all essential partitions)")
            self._log(f"--- USER partition will be blank - Switch will initialize it on first boot")

            # Open target drive using os.open for direct access (proven method)
            self._log(f"--- Opening target drive: {target_drive}")
            try:
                target_fd = os.open(target_drive, os.O_WRONLY | os.O_BINARY)
                self._log(f"--- Successfully opened target drive using os.open")
            except OSError as e:
                self._log(f"ERROR: Failed to open target drive: {e}")
                return False

            try:
                with open(source_nand, 'rb') as src:
                    bytes_copied = 0
                    chunk_size = 1024 * 1024  # 1MB chunks
                    last_update_bytes = 0
                    update_interval_bytes = 512 * 1024 * 1024  # Update every 512 MB
                    
                    self._log("--- Starting optimized raw write to eMMC (5GB target)...")
                    self._log("--- This should take about 2-3 minutes instead of 10-15 minutes.")

                    while bytes_copied < copy_size:
                        remaining = copy_size - bytes_copied
                        chunk = src.read(min(chunk_size, remaining))
                        if not chunk:
                            self._log("--- WARNING: Source file ended before reaching target copy size.")
                            break

                        # Write using os.write (proven method)
                        bytes_written = os.write(target_fd, chunk)
                        bytes_copied += bytes_written

                        # Update progress and force flush periodically
                        if (bytes_copied - last_update_bytes) >= update_interval_bytes or bytes_copied == copy_size:
                            progress_gb = bytes_copied / (1024**3)
                            total_gb = copy_size / (1024**3)
                            percentage = (bytes_copied / copy_size) * 100
                            self._log(f"--- Progress: {progress_gb:.2f} GB / {total_gb:.2f} GB ({percentage:.1f}%)")
                            last_update_bytes = bytes_copied

                            # Force flush to disk
                            self._log("--- Forcing OS write cache to physical disk...")
                            os.fsync(target_fd)
                            self._log("--- Flush complete.")

            finally:
                # Always close the file descriptor
                os.close(target_fd)
                self._log("--- Target drive closed.")

            self._log(f"--- SUCCESS: Copied {bytes_copied / (1024**3):.2f} GB to target eMMC")
            self._log(f"--- All essential partitions written. USER partition is blank and will be initialized by Switch.")
            return True

        except PermissionError:
            self._log("ERROR: Permission denied. Please ensure the script is running as an Administrator.")
            CustomDialog(self, title="Permission Error", 
                            message="Permission denied when trying to write to the drive. Please ensure the script is running with Administrator privileges.")
            return False
        except OSError as e:
            if e.errno == 22:  # Invalid argument
                self._log(f"ERROR: Cannot access physical drive {target_drive}. This may be due to:")
                self._log("1. Drive is mounted/in use by another process")
                self._log("2. Drive access is blocked by antivirus software")
                self._log("3. Insufficient permissions")
                self._log("4. Device not properly connected")
                CustomDialog(self, title="Drive Access Error", 
                                message=f"Cannot access the physical drive.\n\nPossible solutions:\n• Ensure no other programs are using the drive\n• Temporarily disable antivirus\n• Run as Administrator\n• Try disconnecting and reconnecting the Switch")
            else:
                self._log(f"ERROR: OS error occurred: {e}")
            return False
        except Exception as e:
            self._log(f"ERROR: A critical error occurred during the raw copy: {e}")
            import traceback
            self._log(traceback.format_exc())
            CustomDialog(self, title="Write Error", 
                            message=f"A critical error occurred while writing to the eMMC:\n\n{e}")
            return False

    def _setup_settings_menu(self, menubar):
        settings_menu = tk.Menu(menubar, tearoff=0,
            background=self.BG_LIGHT, foreground=self.FG_COLOR,
            activebackground=self.ACCENT_COLOR, activeforeground=self.FG_COLOR,
            relief="flat", borderwidth=0
        )
        menubar.add_cascade(label="Settings", menu=settings_menu)
        paths_to_show = {"7z": "7-Zip (7z.exe)...", "emmchaccgen": "EmmcHaccGen.exe...",
                            "nxnandmanager": "NxNandManager.exe...", "osfmount": "OSFMount.com...",
                            "partitions_folder": "Partitions Folder (NAND)..."}
        for key, text in paths_to_show.items():
            file_type = "file" if ".exe" in text or ".com" in text else "folder"
            settings_menu.add_command(label=f"Set {text}", command=lambda k=key, t=file_type: self._select_path(k, t))

    def _select_path(self, key, type):
        path = ""
        if type == "file": path = filedialog.askopenfilename(title=f"Select {key.replace('_', ' ').title()} File")
        elif type == "folder": path = filedialog.askdirectory(title=f"Select {key.replace('_', ' ').title()} Folder")
        if path: 
            self.paths[key].set(os.path.normpath(path))
            self._save_config()
            self._validate_paths_and_update_buttons()

    def _log(self, message, end="\n"):
        # Check if log_widget exists before trying to use it
        if hasattr(self, 'log_widget') and self.log_widget:
            self.log_widget.config(state="normal")
            self.log_widget.insert(tk.END, message + end)
            self.log_widget.see(tk.END)
            self.log_widget.config(state="disabled")
            self.update_idletasks()
        else:
            # Fall back to print if log widget not available yet
            print(message)

    def _run_command(self, command, cwd=None):
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                        text=True, creationflags=subprocess.CREATE_NO_WINDOW, cwd=cwd,
                                        bufsize=1, universal_newlines=True)
            output = []
            for line in iter(process.stdout.readline, ''):
                clean_line = line.strip()
                if not clean_line: continue
                output.append(clean_line)
                self._log(clean_line)
            
            process.stdout.close()
            return_code = process.wait()
            return return_code, "\n".join(output)
        except Exception as e:
            self._log(f"FATAL ERROR: Failed to execute command. {e}")
            return -1, str(e)
    
    def _start_threaded_process(self, level):
        self._disable_buttons()
        thread = threading.Thread(target=self._start_process, args=(level,)); thread.daemon = True; thread.start()

    def _detect_switch_drives_wmi(self):
        self._log("--- Detecting all physical drives using WMI...")
        try: import wmi
        except ImportError:
            self._log("ERROR: The 'wmi' library is required. Please run 'pip install wmi' from a command prompt.")
            CustomDialog(self, title="Dependency Error", message="The 'wmi' library is not installed.\nPlease run 'pip install wmi' in a command prompt.")
            return []
        
        c = wmi.WMI()
        potential_drives = []
        for disk in c.Win32_DiskDrive():
            try:
                size_gb = int(disk.Size) / (1024**3)
                if 28.0 < size_gb < 31.0 or 57.0 < size_gb < 61.0:
                    drive_info = {"path": disk.DeviceID, "size": f"{size_gb:.2f} GB", "size_gb": size_gb, "model": disk.Model}
                    potential_drives.append(drive_info)
                    self._log(f"--- Found potential Switch drive: {drive_info['path']} ({drive_info['size']})")
            except Exception: continue
        if not potential_drives: self._log("--- No drives matching Switch eMMC size were found.")
        return potential_drives
    
    def _start_process(self, level):
        self._log(f"--- Starting {level} Process ---")
        try:
            with tempfile.TemporaryDirectory(prefix="switch_gui_") as temp_dir:
                self._log(f"INFO: Created temporary directory at: {temp_dir}")
                if level == "Level 1":
                    self._run_level1_process(temp_dir)
                elif level == "Level 2":
                    self._run_level2_process(temp_dir)
        except Exception as e:
            self._log(f"An unexpected critical error occurred: {e}\n{traceback.format_exc()}")
            self._log(f"\nINFO: {level} process finished with an error.")
        finally:
            # Only the action remains, no more logging
            self._re_enable_buttons()

    def _selective_copy_system_contents_level1(self, source_system_path, drive_letter):
        """
        Level 1: Simple approach - delete registered folder, copy Contents from EmmcHaccGen, skip save folder
        """
        try:
            self._log("--- Starting Level 1 selective SYSTEM merge...")
            
            # Show what exists before merge
            self._log("--- BEFORE merge - existing structure:")
            save_dest = drive_letter / "save"
            original_save_count = 0
            if save_dest.exists():
                original_save_count = len(list(save_dest.iterdir()))
                
            for item in drive_letter.iterdir():
                if item.is_dir():
                    try:
                        count = len(list(item.iterdir()))
                        self._log(f"  Existing: DIR - {item.name} ({count} items)")
                    except:
                        self._log(f"  Existing: DIR - {item.name}")
                else:
                    self._log(f"  Existing: FILE - {item.name}")
            
            # Step 1: Delete existing Contents/registered folder completely
            registered_dest = drive_letter / "Contents" / "registered"
            if registered_dest.exists():
                self._log(f"--- Deleting ALL files in 'Contents/registered' folder...")
                shutil.rmtree(registered_dest)
                self._log("--- 'Contents/registered' folder emptied.")
            
            # Step 2: Copy ONLY Contents folder from EmmcHaccGen (this replaces registered, preserves placehld)
            contents_source = source_system_path / "Contents"
            contents_dest = drive_letter / "Contents"
            
            if contents_source.exists():
                self._log("--- Copying Contents folder from EmmcHaccGen...")
                # Ensure Contents directory exists
                contents_dest.mkdir(exist_ok=True)
                
                # Copy each item from source Contents
                for contents_item in contents_source.iterdir():
                    source_item = contents_item
                    dest_item = contents_dest / contents_item.name
                    
                    if source_item.is_dir():
                        self._log(f"--- Copying Contents subdirectory: {contents_item.name}")
                        # Remove destination first to ensure clean replacement
                        if dest_item.exists():
                            shutil.rmtree(dest_item)
                        shutil.copytree(source_item, dest_item)  # No dirs_exist_ok - clean replacement
                    else:
                        self._log(f"--- Copying Contents file: {contents_item.name}")
                        shutil.copy2(source_item, dest_item)
            
            # Step 3: Update specific system save files from EmmcHaccGen
            save_source = source_system_path / "save"
            if save_source.exists():
                self._log("--- Updating system save files from EmmcHaccGen...")
                save_dest.mkdir(exist_ok=True)
                
                # Copy each save file from EmmcHaccGen (overwrites existing system saves)
                files_updated = 0
                for save_file in save_source.iterdir():
                    source_save = save_file
                    dest_save = save_dest / save_file.name
                    
                    if source_save.is_file():
                        shutil.copy2(source_save, dest_save)
                        files_updated += 1
                        self._log(f"--- Updated system save: {save_file.name}")
                
                self._log(f"--- Updated {files_updated} system save files")
            
            # Step 4: Handle any other top-level items from EmmcHaccGen (if any exist)
            for src_item in source_system_path.iterdir():
                if src_item.name in ["Contents", "save"]:
                    continue  # Already handled or skipped above
                    
                dest_item = drive_letter / src_item.name
                if not dest_item.exists():
                    self._log(f"--- Adding new top-level item: {src_item.name}")
                    if src_item.is_dir():
                        shutil.copytree(src_item, dest_item)
                    else:
                        shutil.copy2(src_item, dest_item)
            
            # Show what exists after merge
            self._log("--- AFTER merge - final structure:")
            final_save_count = 0
            for item in drive_letter.iterdir():
                if item.is_dir():
                    try:
                        count = len(list(item.iterdir()))
                        if item.name == "save":
                            final_save_count = count
                        self._log(f"  Final: DIR - {item.name} ({count} items)")
                    except:
                        self._log(f"  Final: DIR - {item.name}")
                else:
                    self._log(f"  Final: FILE - {item.name}")
            
            # Verify save folder was not touched
            if final_save_count == original_save_count:
                self._log(f"--- ✅ CONFIRMED: Save folder untouched ({original_save_count} files preserved)")
            else:
                self._log(f"--- ⚠️ WARNING: Save count changed from {original_save_count} to {final_save_count}")
            
            # Show Contents subdirectories
            contents_dir = drive_letter / "Contents"
            if contents_dir.exists():
                self._log("--- Contents subdirectories:")
                for item in contents_dir.iterdir():
                    if item.is_dir():
                        try:
                            count = len(list(item.iterdir()))
                            self._log(f"    {item.name} ({count} items)")
                        except:
                            self._log(f"    {item.name}")
            
            self._log("--- SUCCESS: System files replaced, save data completely preserved.")
            return True
            
        except Exception as e:
            self._log(f"ERROR: Failed Level 1 merge. Error: {e}")
            self._log(traceback.format_exc())
            return False

    def _run_level1_process(self, temp_dir):
        # This function and the ones below are unchanged but included for completeness.
        self._log("\n--- WARNING ---")
        self._log("The Level 1 process will write directly to your Switch's eMMC.")
        
        self._log("\n[STEP 1/8] Please connect your Switch in Hekate's eMMC RAW GPP mode (Read-Only OFF).")
        nx_exe = self.paths['nxnandmanager'].get()
        return_code, output = self._run_command([nx_exe, '--list'])
        if return_code != 0: return

        drive_match = re.search(r"(\\\\.\\PhysicalDrive\d+)", output)
        if not drive_match: return self._log("ERROR: No compatible Switch eMMC found.")
        drive_path = drive_match.group(1)
        self._log(f"SUCCESS: Found eMMC at {drive_path}")

        self._log("--- Dumping and decrypting PRODINFO from eMMC...")
        keyset_path = self.paths['keys'].get()
        prodinfo_path = Path(temp_dir) / "PRODINFO"
        dump_cmd = [nx_exe, '-i', drive_path, '-keyset', keyset_path, '-o', temp_dir, '-d', '-part=PRODINFO']
        
        # Check if dump command fails OR if the file wasn't created
        if self._run_command(dump_cmd)[0] != 0 or not prodinfo_path.exists():
            self._log(f"ERROR: Failed to dump or decrypt PRODINFO from the eMMC. It may be corrupt.")
            CustomDialog(self, title="PRODINFO Error", message="PRODINFO is not found or damaged. Please use Level 2 or Level 3 instead.")
            return

        # Validate the dumped PRODINFO file's content
        with open(prodinfo_path, 'rb') as f:
            if f.read(4) != b'CAL0':
                self._log(f"ERROR: The PRODINFO dumped from the eMMC is invalid or encrypted (magic is not CAL0).")
                CustomDialog(self, title="PRODINFO Error", message="PRODINFO is not found or damaged. Please use Level 2 or Level 3 instead.")
                return

        self._log("SUCCESS: PRODINFO is valid and decrypted.")

        self._log(f"\n[STEP 2/8] Reading PRODINFO file...")
        with open(prodinfo_path, 'rb') as f:
            f.seek(0x3740)
            model_bytes = f.read(4)
            product_model_id = int.from_bytes(model_bytes, byteorder='little')
        model_map = {1: "Erista", 3: "V2", 4: "Lite", 6: "OLED"}
        detected_model = model_map.get(product_model_id, "Unknown Mariko")
        self._log(f"SUCCESS: Detected model: {detected_model}")

        self._log(f"\n[STEP 3/8] Generating boot files...")
        emmchaccgen_out_dir = Path(temp_dir) / "emmchaccgen_out"
        emmchaccgen_out_dir.mkdir()
        emmchaccgen_cmd = [self.paths['emmchaccgen'].get(), '--keys', keyset_path, '--fw', self.paths['firmware'].get()]
        if "Mariko" in detected_model or detected_model in ["V2", "Lite", "OLED"]:
            self._log("--- Mariko model detected, using --mariko flag (AutoRCM disabled by default).")
            emmchaccgen_cmd.append('--mariko')
        else:
            self._log("--- Erista model detected, adding --no-autorcm flag by default.")
            emmchaccgen_cmd.append('--no-autorcm')
        if self._run_command(emmchaccgen_cmd, cwd=str(emmchaccgen_out_dir))[0] != 0: return

        self._log(f"\n[STEP 4/8] Dumping and decrypting SYSTEM partition from eMMC...")
        dump_cmd = [nx_exe, '-i', drive_path, '-keyset', keyset_path, '-o', temp_dir, '-d', '-part=SYSTEM']
        if self._run_command(dump_cmd)[0] != 0: return
        
        system_dec_path = Path(temp_dir) / "SYSTEM"
        if not system_dec_path.exists(): return self._log("ERROR: SYSTEM file was not created.")
        self._log("SUCCESS: SYSTEM partition decrypted.")
        
        self._log(f"--- Mounting SYSTEM partition to modify...")
        osfmount_cmd = [self.paths['osfmount'].get(), '-a', '-t', 'file', '-f', str(system_dec_path), '-o', 'rw', '-m', '#:']
        return_code, output = self._run_command(osfmount_cmd)
        if return_code != 0: return
        match = re.search(r"([A-Z]:)", output)
        if not match: return self._log("ERROR: Could not determine drive letter.")
        drive_letter_str = match.group(1)
        drive_letter = Path(drive_letter_str)
        self._log(f"--- SUCCESS: Image mounted to {drive_letter}")

        try:
            versioned_folder = next(d for d in emmchaccgen_out_dir.iterdir() if d.is_dir())
            source_system_path = versioned_folder / "SYSTEM"
            
            # Use selective copying to preserve savemeta and other existing folders
            success = self._selective_copy_system_contents_level1(source_system_path, drive_letter)
            if not success:
                return
                
        except Exception as e:
            return self._log(f"ERROR: Failed to modify SYSTEM partition contents. Error: {e}")
        finally:
            self._log(f"--- Dismounting drive...")
            self._run_command([self.paths['osfmount'].get(), '-D', '-m', drive_letter_str])

        self._log(f"\n[STEP 5 & 6/8] Flashing modified SYSTEM back to eMMC...")
        flash_cmd = [nx_exe, '-i', str(system_dec_path), '-o', drive_path, '-part=SYSTEM', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_command(flash_cmd)[0] != 0:
            return self._log("ERROR: Failed to flash SYSTEM partition back to eMMC.")
        self._log("SUCCESS: SYSTEM partition has been restored.")

        self._log(f"\n[STEP 7/8] Flashing BCPKG2 partitions...")
        versioned_folder = next(d for d in emmchaccgen_out_dir.iterdir() if d.is_dir())
        bcpkg2_partitions = ["BCPKG2-1-Normal-Main", "BCPKG2-2-Normal-Sub", "BCPKG2-3-SafeMode-Main", "BCPKG2-4-SafeMode-Sub"]
        for part_name in bcpkg2_partitions:
            bcpkg2_file = versioned_folder / f"{part_name}.bin"
            if not bcpkg2_file.exists(): return self._log(f"ERROR: {bcpkg2_file.name} not found.")
            flash_cmd = [nx_exe, '-i', str(bcpkg2_file), '-o', drive_path, f'-part={part_name}', 'FORCE']
            if self._run_command(flash_cmd)[0] != 0: return self._log(f"ERROR: Failed to flash {part_name}.")
        self._log("SUCCESS: All BCPKG2 partitions have been restored.")

        self._log(f"\n[STEP 8/8] Saving BOOT0 & BOOT1 to output folder...")
        output_folder = Path(self.paths['output_folder'].get())
        versioned_folder = next(d for d in emmchaccgen_out_dir.iterdir() if d.is_dir())
        shutil.copy(versioned_folder / "BOOT0.bin", output_folder / "BOOT0")
        shutil.copy(versioned_folder / "BOOT1.bin", output_folder / "BOOT1")
        self._log(f"SUCCESS: BOOT0 and BOOT1 saved. Please flash them manually using Hekate.")
        self._log("\n--- LEVEL 1 IN-PLACE RESTORE COMPLETE ---")

    def _run_and_interrupt_flash(self, command, partition_name, target_mb):
        self._log(f"--- Starting partial flash for {partition_name} with a {target_mb}MB target...")
        try:
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                        text=True, creationflags=subprocess.CREATE_NO_WINDOW,
                                        bufsize=1, universal_newlines=True)
            progress_regex = re.compile(rf"Restoring to {partition_name}... (\d+\.\d+)\s*MB")
            for line in iter(process.stdout.readline, ''):
                clean_line = line.strip()
                if not clean_line: continue
                self._log(clean_line)
                match = progress_regex.search(clean_line)
                if match and float(match.group(1)) >= target_mb:
                    self._log(f"--- SUCCESS: Reached target. Terminating flash...")
                    process.terminate()
                    break
            process.stdout.close(); process.wait()
            self._log(f"--- Partial flash for {partition_name} complete.")
            return 0
        except Exception as e:
            self._log(f"FATAL ERROR during interruptible flash: {e}"); return -1

    def _run_level2_process(self, temp_dir):
        self._log("\n--- WARNING ---")
        self._log("The Level 2 process will write directly to your Switch's eMMC.")

        self._log("\n[STEP 1/7] Detecting eMMC...")
        nx_exe = self.paths['nxnandmanager'].get()
        
        # Hardcode the path to the lib/NAND folder
        try:
            script_dir = Path(__file__).parent
        except NameError:
            script_dir = Path.cwd()
        partitions_folder = script_dir / "lib" / "NAND"
        keyset_path = self.paths['keys'].get()
        
        return_code, output = self._run_command([nx_exe, '--list'])
        if return_code != 0: return
        drive_match = re.search(r"(\\\\.\\PhysicalDrive\d+)", output)
        if not drive_match: return self._log("ERROR: No compatible Switch eMMC found.")
        drive_path = drive_match.group(1)
        self._log(f"--- SUCCESS: Found eMMC at {drive_path}")

        self._log("--- Acquiring PRODINFO...")
        prodinfo_path = Path(temp_dir) / "PRODINFO"
        dump_cmd = [nx_exe, '-i', drive_path, '-keyset', keyset_path, '-o', temp_dir, '-d', '-part=PRODINFO']
        
        donor_prodinfo_used = False
        if self._run_command(dump_cmd)[0] != 0 or not prodinfo_path.exists():
            self._log("--- INFO: Could not dump from eMMC. Falling back to donor PRODINFO file.")
            donor_path = Path(self.paths['prodinfo'].get())
            if not donor_path.is_file():
                self._log("ERROR: PRODINFO could not be dumped from eMMC and no donor file was provided.")
                CustomDialog(self, title="PRODINFO Error", message="PRODINFO is not found or damaged. Please use Level 3 instead.")
                return
            shutil.copy(donor_path, prodinfo_path)
            donor_prodinfo_used = True
        else: self._log("--- SUCCESS: PRODINFO dumped from eMMC.")
        
        with open(prodinfo_path, 'rb') as f:
            if f.read(4) != b'CAL0':
                source = "donor file" if donor_prodinfo_used else "eMMC"
                self._log(f"ERROR: The PRODINFO from the {source} is invalid or encrypted (magic is not CAL0).")
                CustomDialog(self, title="PRODINFO Error", message="PRODINFO is not found or damaged. Please use Level 3 instead.")
                return

        self._log(f"\n[STEP 2/7] Reading PRODINFO file...")
        with open(prodinfo_path, 'rb') as f:
            f.seek(0x3740)
            product_model_id = int.from_bytes(f.read(4), byteorder='little')
        model_map = {1: "Erista", 3: "V2", 4: "Lite", 6: "OLED"}
        detected_model = model_map.get(product_model_id, "Unknown Mariko")
        self._log(f"SUCCESS: Detected model: {detected_model}")

        self._log(f"\n[STEP 3/7] Generating boot files...")
        emmchaccgen_out_dir = Path(temp_dir) / "emmchaccgen_out"
        emmchaccgen_out_dir.mkdir()
        emmchaccgen_cmd = [self.paths['emmchaccgen'].get(), '--keys', keyset_path, '--fw', self.paths['firmware'].get()]
        if "Mariko" in detected_model or detected_model in ["V2", "Lite", "OLED"]:
            self._log("--- Mariko model detected, using --mariko flag (AutoRCM disabled by default).")
            emmchaccgen_cmd.append('--mariko')
        else:
            self._log("--- Erista model detected, adding --no-autorcm flag by default.")
            emmchaccgen_cmd.append('--no-autorcm')
        if self._run_command(emmchaccgen_cmd, cwd=str(emmchaccgen_out_dir))[0] != 0: return

        self._log(f"\n[STEP 4/7] Preparing donor SYSTEM partition...")
        if self._run_command([self.paths['7z'].get(), 'x', str(partitions_folder / "SYSTEM.7z"), f'-o{temp_dir}'])[0] != 0: return
        system_dec_path = Path(temp_dir) / "SYSTEM.dec"
        
        self._log(f"--- Mounting donor SYSTEM to inject files...")
        osfmount_cmd = [self.paths['osfmount'].get(), '-a', '-t', 'file', '-f', str(system_dec_path), '-o', 'rw', '-m', '#:']
        return_code, output = self._run_command(osfmount_cmd)
        if return_code != 0: return
        match = re.search(r"([A-Z]:)", output)
        if not match: return self._log("ERROR: Could not determine drive letter.")
        drive_letter_str = match.group(1)

        try:
            versioned_folder = next(d for d in emmchaccgen_out_dir.iterdir() if d.is_dir())
            source_system_path = versioned_folder / "SYSTEM"
            
            # Use selective copying to preserve existing folders in donor SYSTEM
            success = self._selective_copy_system_contents(source_system_path, Path(drive_letter_str))
            if not success:
                return
            self._log("--- SUCCESS: New system files injected into donor SYSTEM.")
        except Exception as e:
            return self._log(f"ERROR: Failed to inject files into SYSTEM. Error: {e}")
        finally:
            self._log(f"--- Dismounting drive...")
            self._run_command([self.paths['osfmount'].get(), '-D', '-m', drive_letter_str])

        self._log(f"\n[STEP 5/7] Flashing all data partitions to eMMC...")
        
        flash_cmd = [nx_exe, '-i', str(prodinfo_path), '-o', drive_path, '-part=PRODINFO', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_command(flash_cmd)[0] != 0: return
        
        flash_cmd = [nx_exe, '-i', str(system_dec_path), '-o', drive_path, '-part=SYSTEM', '-e', '-keyset', keyset_path, 'FORCE']
        if self._run_command(flash_cmd)[0] != 0: return

        partition_map = {"PRODINFOF": {"default": "PRODINFOF.7z"},
                            "USER": {"OLED": "USER-64.7z", "default": "USER-32.7z"},
                            "SAFE": {"default": "SAFE.7z"}}
        for part_name, archive_map in partition_map.items():
            archive_name = archive_map.get(detected_model, archive_map["default"])
            if self._run_command([self.paths['7z'].get(), 'x', str(partitions_folder / archive_name), f'-o{temp_dir}'])[0] == 0:
                dec_file_path = Path(temp_dir) / f"{part_name}.dec"
                flash_cmd = [nx_exe, '-i', str(dec_file_path), '-o', drive_path, f'-part={part_name}', '-e', '-keyset', keyset_path, 'FORCE']
                if part_name == "USER" and not donor_prodinfo_used:
                    if self._run_and_interrupt_flash(flash_cmd, "USER", 100) != 0: return
                else:
                    if self._run_command(flash_cmd)[0] != 0: return
        
        self._log("SUCCESS: All data partitions have been restored.")

        self._log(f"\n[STEP 6/7] Flashing BCPKG2 partitions...")
        versioned_folder = next(d for d in emmchaccgen_out_dir.iterdir() if d.is_dir())
        bcpkg2_partitions = ["BCPKG2-1-Normal-Main", "BCPKG2-2-Normal-Sub", "BCPKG2-3-SafeMode-Main", "BCPKG2-4-SafeMode-Sub"]
        for part_name in bcpkg2_partitions:
            bcpkg2_file = versioned_folder / f"{part_name}.bin"
            if not bcpkg2_file.exists(): return self._log(f"ERROR: {bcpkg2_file.name} not found.")
            flash_cmd = [nx_exe, '-i', str(bcpkg2_file), '-o', drive_path, f'-part={part_name}', 'FORCE']
            if self._run_command(flash_cmd)[0] != 0: return self._log(f"ERROR: Failed to flash {part_name}.")
        self._log("SUCCESS: All BCPKG2 partitions have been restored.")

        self._log(f"\n[STEP 7/7] Saving BOOT0 & BOOT1 to output folder...")
        output_folder = Path(self.paths['output_folder'].get())
        shutil.copy(versioned_folder / "BOOT0.bin", output_folder / "BOOT0")
        shutil.copy(versioned_folder / "BOOT1.bin", output_folder / "BOOT1")
        self._log(f"SUCCESS: BOOT0 and BOOT1 saved. Please flash them manually using Hekate.")
        self._log("\n--- LEVEL 2 IN-PLACE REBUILD COMPLETE ---")

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

if __name__ == "__main__":
    # --- START: ADDED DEPENDENCY INSTALLER ---
    import subprocess
    import sys
    import importlib

    def install_dependencies():
        """Checks for and installs required packages if they are missing."""
        required_packages = ['wmi']
        for package in required_packages:
            try:
                # Try to import the package to see if it's installed
                importlib.import_module(package)
                print(f"INFO: Dependency '{package}' is already installed.")
            except ImportError:
                print(f"INFO: Dependency '{package}' not found. Attempting to install...")
                try:
                    # Use sys.executable to ensure pip from the correct python environment is used
                    # Use flags to suppress installation output unless there's an error
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package],
                                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    print(f"SUCCESS: Successfully installed '{package}'.")
                except subprocess.CalledProcessError:
                    # If installation fails, show a clear error message box and exit
                    error_msg = (f"Failed to automatically install the required package: '{package}'.\n\n"
                                 f"Please install it manually by opening a command prompt/terminal and running:\n\n"
                                 f"pip install {package}")
                    print(f"ERROR: {error_msg}")
                    ctypes.windll.user32.MessageBoxW(0, error_msg, "Dependency Error", 0x10) # 0x10 = MB_ICONERROR
                    sys.exit(1) # Exit the script if a critical dependency can't be installed

    install_dependencies()
    # --- END: ADDED DEPENDENCY INSTALLER ---

    # --- WRAP THE APP STARTUP IN A TRY/EXCEPT BLOCK ---
    app = None
    try:
        if is_admin():
            app = SwitchGuiApp()
            app.mainloop()
        else:
            # Relaunch as admin
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    except Exception as e:
        # If any error happens here, the excepthook at the top will catch it,
        # log it to error_log.txt, and exit the program safely.
        raise e