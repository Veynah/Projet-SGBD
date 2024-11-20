import os
import tkinter as tk
from tkinter import ttk, filedialog
from functools import partial

from utilities.utils import center_window, create_menu, show_message, clean_file_path, create_styled_button, \
    set_window_icon


class ConfigUI(tk.Toplevel):
    def __init__(self, master, tabs):
        super().__init__(master)
        self.title("Configuration")
        self.tabs = tabs
        self.entries = {}
        self.config_manager = master.config_manager
        self.init_ui()

    def init_ui(self):
        """Initializes the user interface components and lays out the window."""
        center_window(self, self.master)
        set_window_icon(self)
        self.resizable(False, False)
        self.create_menus()
        self.create_frames()
        self.transient(self.master)
        self.grab_set()
        self.wait_window(self)

    def create_menus(self):
        """Creates menu items for the window."""
        menu_items = [
            {"label": "Sources", "command": partial(self.toggle_frame, "sources")},
            {"label": "Config", "command": partial(self.toggle_frame, ".config")},
        ]
        create_menu(self, menu_items)

    def create_frames(self):
        """Initializes frames for configuration settings and source file settings."""
        self.sources_destinations_frame = ttk.Frame(self)
        self.config_frame = ttk.Frame(self)
        self.build_sources_frame()
        self.build_config_frame()
        self.sources_destinations_frame.grid(row=0, column=0, sticky="nsew")
        self.config_frame.grid(row=0, column=0, sticky="nsew")
        self.config_frame.grid_remove()

        self.sources_destinations_frame.columnconfigure(0, weight=1)
        self.sources_destinations_frame.columnconfigure(1, weight=1)
        self.sources_destinations_frame.columnconfigure(2, weight=1)

    def build_sources_frame(self):
        """Creates user interface for source file settings."""
        for idx, tab_name in enumerate(self.tabs):
            self.create_source_widgets(tab_name, idx)
            self.set_initial_path(tab_name)

    def create_source_widgets(self, tab_name, row):
        """Helper function to create label, entry, and button for a source file."""
        if tab_name.lower() == "cost":
            key = "cost_src"
            label_text = "Cost Reference file:"
        else:
            key = f"{tab_name.lower()}_src"
            label_text = f"{tab_name} Reference file:"

        ttk.Label(self.sources_destinations_frame, text=label_text).grid(row=row * 2, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")

        container_frame = ttk.Frame(self.sources_destinations_frame)
        container_frame.grid(row=row * 2 + 1, column=0, columnspan=3, padx=10, pady=5, sticky="ew")

        entry = ttk.Entry(container_frame)
        entry.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        ttk.Button(container_frame, text="Browse", command=lambda e=entry: self.select_file(e)).grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        container_frame.columnconfigure(0, weight=1)
        container_frame.rowconfigure(0, weight=1)
        container_frame.rowconfigure(1, weight=1)

        self.entries[key] = entry

    def set_initial_path(self, tab_name):
        """Sets the initial path in the entry field from .config."""
        key = f"{tab_name.lower()}_src"
        if tab_name.lower() == "cost":
            key = "cost_src"
        if key in self.config_manager.config_data:
            self.entries[key].delete(0, tk.END)
            self.entries[key].insert(0, self.config_manager.config_data[key])

    def build_config_frame(self):
        """Creates user interface for configuration management."""
        create_styled_button(self.config_frame, "Load Configuration", command=self.load_configuration, width=20).pack(pady=10, expand=True, fill="x")
        create_styled_button(self.config_frame, "View Configuration", command=self.view_configuration, width=20).pack(pady=10, expand=True, fill="x")
        create_styled_button(self.config_frame, "Save Configuration", command=self.save_configuration, width=20).pack(pady=10, expand=True, fill="x")

    def toggle_frame(self, frame_name):
        """Toggles between different frames in the UI based on the user's choice."""
        if frame_name == ".config":
            self.sources_destinations_frame.grid_remove()
            self.config_frame.grid()
        else:
            self.config_frame.grid_remove()
            self.sources_destinations_frame.grid()

    def select_file(self, entry_field):
        """Handles selecting a file and updating the corresponding entry field."""
        file_selected = filedialog.askopenfilename(filetypes=[("All Files", "*.*"), ("Excel Files", "*.xls;*.xlsx"), ("Text Files", "*.txt"), ("CSV Files", "*.csv")])
        if file_selected:
            entry_field.delete(0, tk.END)
            entry_field.insert(0, file_selected)

    def load_configuration(self):
        file_selected = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if file_selected:
            try:
                self.config_manager.load_config(file_selected)
                show_message("Load Configuration", "Configuration loaded successfully!", type="info", master=self, custom=True)
            except Exception as e:
                show_message("Error Loading Configuration", f"An error occurred: {str(e)}", type="info", master=self, custom=True)

    def save_configuration(self):
        """Saves the current configuration settings."""
        for key, entry in self.entries.items():
            cleaned_path = clean_file_path(entry.get())
            self.config_manager.update_config(key, cleaned_path)
        self.config_manager.save_config()
        show_message("Save Configuration", "Configuration saved successfully!", type="info", master=self, custom=True)

    def view_configuration(self):
        config_path = self.config_manager.config_file

        if os.path.exists(config_path):
            os.startfile(config_path)
        else:
            show_message("View Configuration", "Configuration file does not exist.", type="info", master=self, custom=True)

    def center_window(self):
        """Centers the window on the screen based on the master window's size."""
        window_width, window_height = 420, 650
        x = self.master.winfo_x() + (self.master.winfo_width() - window_width) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

