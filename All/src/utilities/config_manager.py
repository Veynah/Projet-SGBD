import json
import os
import sys
from tkinter import ttk, Toplevel, font
import win32com.client as win32

import pandas as pd

from utilities.utils import center_window, show_message, set_window_icon


class ViewDealsLoaderPopup(Toplevel):
    def __init__(self, master, config_manager):
        super().__init__(master)
        self.config_manager = config_manager
        self.config_data = self.config_manager.get_config()
        self.cost_dest = self.config_data.get('cost_dest', '')

        if not self.cost_dest:
            print("Debug: 'cost_dest' is not set or is empty.")
            show_message("Error", "The 'cost_dest' directory is not set or is invalid.", type='error', master=self)
            self.destroy()
            return

        print(f"Debug: 'cost_dest' is set to: {self.cost_dest}")

        self.working_contracts_file = os.path.join(self.cost_dest, 'working_contracts.xlsx')

        if not os.path.exists(self.working_contracts_file):
            print(f"Debug: 'working_contracts.xlsx' file does not exist at: {self.working_contracts_file}")
            show_message("Error", f"'working_contracts.xlsx' file not found in {self.cost_dest}", type='error', master=self)
            self.destroy()
            return

        self.init_ui()
        center_window(self, master, 500, 400)
        self.transient(master)
        self.grab_set()
        master.wait_window(self)

    def init_ui(self):
        self.title("Contract Templates")
        set_window_icon(self)
        self.configure(bg='#f0f0f0')

        sheets = self.get_sheets()

        self.main_frame = ttk.Frame(self, style='Main.TFrame')
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(self.main_frame, columns=('Business Model', 'Path'), show='headings')
        self.tree.heading('Business Model', text='Business Model')
        self.tree.heading('Path', text='Path')

        self.tree.column('Business Model', width=250, stretch=True)
        self.tree.column('Path', width=200, stretch=True)

        bold_font = font.Font(weight="bold")

        for sheet_name in sheets:
            self.tree.insert('', 'end', values=(sheet_name, self.working_contracts_file))

            self.tree.tag_configure('bold', font=bold_font)
            self.tree.item(self.tree.get_children()[-1], tags=('bold',))
        self.tree.pack(fill='both', expand=True)

        self.tree.bind("<Double-1>", self.open_template)

        scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='right', fill='y')

    def get_sheets(self):
        """Retrieve all sheet names from the centralized working_contracts.xlsx file."""
        print(f"Debug: Retrieving sheets from {self.working_contracts_file}")
        xl = pd.ExcelFile(self.working_contracts_file)
        return xl.sheet_names

    def open_template(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            sheet_name = self.tree.item(selected_item, "values")[0]
            excel = win32.Dispatch('Excel.Application')
            workbook = excel.Workbooks.Open(self.working_contracts_file)
            excel.Visible = True

            try:
                sheet = workbook.Sheets(sheet_name)
                sheet.Activate()
            except Exception as e:
                show_message("Error", f"Could not find sheet '{sheet_name}'. Error: {e}", master=self.master,
                             custom=True)

            excel.WindowState = win32.constants.xlMaximized
            excel.Application.ActiveWindow.Activate()


def default_config():
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    output_dir = os.path.join(project_root, 'outputs')

    print(f"Debug: Setting default 'cost_dest' to: {output_dir}")
    return {
        'audience_src': '',
        'audience_dest': '',
        'cost_src': '',
        'cost_dest': output_dir,
        'product_grouping_src': '',
        'channel_grouping_src': '',
    }




class ConfigManager:
    def __init__(self, config_file=None):
        if config_file is None:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            self.config_file = os.path.join(base_dir, '..', '..', '.config', 'config.json')
        else:
            self.config_file = config_file
        self.config_data = {}

    def load_config(self, file_path=None):
        if file_path is None:
            file_path = self.config_file

        try:
            with open(file_path, 'r') as file:
                content = file.read()
                if content.strip():
                    self.config_data = json.loads(content)
                    print(f"Debug: Loaded config from {file_path}.")
                else:
                    raise ValueError("Empty configuration file")
        except (json.JSONDecodeError, ValueError, FileNotFoundError) as e:
            print(f"Debug: Failed to load config from {file_path}. Using default config.")
            self.config_data = default_config()
            self.save_config()

        print(f"Debug: Current 'cost_dest' in config: {self.config_data.get('cost_dest', '')}")
        return self.config_data

    def save_config(self):
        """Write the configuration data to disk."""
        os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
        with open(self.config_file, 'w') as file:
            json.dump(self.config_data, file, indent=4)
        print(f"Debug: Configuration saved to {self.config_file}")

    def update_config(self, key, value):
        if isinstance(value, str) and '\\' in value:
            value = value.replace('\\', '\\\\')
        self.config_data[key] = value
        self.save_config()

    def get_config(self):
        return self.config_data


class ConfigLoaderPopup(Toplevel):
    def __init__(self, master, config_manager, callback, audience_tab):
        super().__init__(master)
        self.config_manager = config_manager
        self.callback = callback
        self.audience_tab = audience_tab
        self.config_data = self.config_manager.get_config()
        self.loaded_files = set()
        self.selected_files = set()
        self.init_ui()
        center_window(self, master, 800, 400)
        self.transient(master)
        self.grab_set()
        master.wait_window(self)

    def init_ui(self):
        set_window_icon(self)
        self.configure(bg='#f0f0f0')

        files_to_load = {k: v for k, v in self.config_data.items() if os.path.isfile(v)}

        data = []
        for key, path in files_to_load.items():
            file_name = os.path.basename(path)
            data.append([key, file_name, path])

        df = pd.DataFrame(data, columns=['Config Variable', 'File Name', 'Path'])

        self.main_frame = ttk.Frame(self, style='Main.TFrame')
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.table_frame = ttk.Frame(self.main_frame, style='Main.TFrame')
        self.table_frame.grid(row=0, column=0, sticky="nsew")

        self.tree = ttk.Treeview(self.table_frame, columns=('Config Variable', 'File Name', 'Path'), show='headings', selectmode='extended')
        self.tree.heading('Config Variable', text='Config Variable')
        self.tree.heading('File Name', text='File Name')
        self.tree.heading('Path', text='Path')

        self.tree.column('Config Variable', width=150, stretch=True)
        self.tree.column('File Name', width=150, stretch=True)
        self.tree.column('Path', width=300, stretch=True)

        for idx, row in df.iterrows():
            self.tree.insert('', 'end', values=(row['Config Variable'], row['File Name'], row['Path']))

        self.tree.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='left', fill='y')

        self.button_frame = ttk.Frame(self.main_frame, style='Main.TFrame')
        self.button_frame.grid(row=1, column=0, pady=10, sticky="ew")

        self.button_frame.columnconfigure(0, weight=1)
        self.button_frame.columnconfigure(1, weight=1)

        load_all_button = ttk.Button(self.button_frame, text="Load All", command=self.load_all_files)
        load_all_button.grid(row=0, column=0, padx=2, sticky="ew")

        next_button = ttk.Button(self.button_frame, text="Load Selection", command=self.load_selected_files)
        next_button.grid(row=0, column=1, padx=2, sticky="ew")

        help_label = ttk.Label(self.main_frame, text="*Sélectionnez plusieurs fichiers à l'aide de la touche Ctrl ou Alt", font=('Helvetica', 10, 'italic'))
        help_label.grid(row=2, column=0, pady=0, padx=2, sticky="w")

        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

    def load_selected_files(self):
        selected_items = self.tree.selection()
        files_to_load = []
        audience_loaded = False
        grouping_loaded = False

        for item in selected_items:
            values = self.tree.item(item, "values")
            key, path = values[0], values[2]
            if path not in self.selected_files:
                self.selected_files.add(path)
                if path not in self.loaded_files:
                    self.loaded_files.add(path)
                    files_to_load.append((key, path))

                    # Check if audience_src or groupings are selected
                    if key == 'audience_src':
                        audience_loaded = True
                    elif key in ['product_grouping_src', 'channel_grouping_src']:
                        grouping_loaded = True

        if files_to_load:
            for key, path in files_to_load:
                try:
                    if os.path.exists(path):
                        self.callback(key, path)
                        show_message("File Loading...", f"Successfully loaded {key} from {path}", master=self.master,
                                     custom=True)
                    else:
                        show_message("Error", f"File not found: {path}", type='error', master=self.master, custom=True)
                except Exception as e:
                    show_message("Error", f"Failed to load {key}: {e}", type='error', master=self.master, custom=True)

        # Check if audience and a grouping file are both loaded, and enable specifics if so
        if audience_loaded and grouping_loaded:
            self.audience_tab.enable_specifics()

        self.selected_files.clear()
        self.destroy()

    def load_all_files(self):
        for item in self.tree.get_children():
            self.tree.selection_add(item)
        self.load_selected_files()
