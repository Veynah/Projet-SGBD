import json
import os
import subprocess
import sys
import time
import traceback
from datetime import datetime
from tkinter import filedialog, Listbox, MULTIPLE, BooleanVar, Toplevel, StringVar
from tkinter import ttk

import pandas as pd

from utilities import utils
from utilities.utils import show_message


class AudienceTab(ttk.Frame):
    def __init__(self, parent, config_manager, base_dir, config_ui_callback=None):
        super().__init__(parent)
        self.lookup_key_to_prod_num = None
        self.prod_num_label = None
        self.bus_chanl_num_map = None
        self.prod_num_map = None
        self.bus_chanl_label = None
        self.output_dir = None
        self.tooltip = None
        self.df = None
        self.config_ui_callback = config_ui_callback
        self.config_manager = config_manager
        self.config_data = config_manager.get_config()
        self.base_dir = base_dir
        self.file_path = None
        self.current_date = datetime.now()
        self.current_year = self.current_date.year
        self.current_month = self.current_date.month
        self.audience_tab_setup()

    def section_filter_bar_setup(self):
        filter_container = ttk.Frame(self.specifics_frame)
        filter_container.pack(side='top', fill='x', padx=10, pady=(5, 10))

        ttk.Label(filter_container, text="Filter:").pack(side='left', padx=(0, 5))

        self.filter_var = StringVar()
        self.filter_bar = ttk.Entry(filter_container, textvariable=self.filter_var, width=40)
        self.filter_bar.pack(side='left', fill='x', expand=True)

        self.filter_bar.bind('<KeyRelease>', self.filter_listboxes)

    def filter_listboxes(self, event=None):
        search_term = self.filter_var.get().lower()

        if not hasattr(self, 'global_selected_channels'):
            self.global_selected_channels = set()

        currently_selected = set(self.bus_chanl_num_listbox.get(i) for i in self.bus_chanl_num_listbox.curselection())
        self.global_selected_channels.update(currently_selected)

        self.bus_chanl_num_listbox.delete(0, 'end')
        for display_value, bus_chanl_num in self.bus_chanl_num_map.items():
            display_value_str = str(display_value).lower()
            bus_chanl_num_str = str(bus_chanl_num).lower()

            if search_term in display_value_str or search_term in bus_chanl_num_str:
                self.bus_chanl_num_listbox.insert('end', display_value)

        self.restore_selection_after_filter()

    def restore_selection_after_filter(self):
        current_items = set(self.bus_chanl_num_listbox.get(0, 'end'))

        selected_items_to_restore = self.global_selected_channels & current_items

        for i in range(self.bus_chanl_num_listbox.size()):
            if self.bus_chanl_num_listbox.get(i) in selected_items_to_restore:
                self.bus_chanl_num_listbox.selection_set(i)

    def section_specifics_listbox_highlight_top(self, listbox):
        selected_indices = listbox.curselection()
        if not selected_indices:
            return

        first_visible_index = listbox.nearest(0)

        selected_items = [listbox.get(i) for i in selected_indices]
        remaining_items = [listbox.get(i) for i in range(listbox.size()) if i not in selected_indices]

        listbox.delete(0, 'end')
        for item in selected_items:
            listbox.insert('end', item)
            listbox.selection_set('end')
        for item in remaining_items:
            listbox.insert('end', item)

        listbox.yview_scroll(first_visible_index - listbox.nearest(0), 'units')

    def audience_tab_setup(self):
        """Sets up user interface components."""
        self.pack(fill='both', expand=True)
        self.tab_style()
        self.section_references_setup()
        self.section_target_setup()
        self.section_specifics_setup()
        self.load_output_directory()


    def section_target_output_location(self):
        """Prompt the user to select an output directory and update the configuration."""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_path.delete(0, 'end')
            normalized_path = os.path.normpath(folder_selected)
            self.output_path.insert(0, normalized_path)
            self.output_dir = normalized_path
            self.config_manager.update_config('audience_dest', normalized_path)

    def load_output_directory(self):
        """Load the output directory from the configuration."""
        stored_output_dir = self.config_data.get('audience_dest')
        if stored_output_dir:
            normalized_path = os.path.normpath(stored_output_dir)
            self.output_path.delete(0, 'end')
            self.output_path.insert(0, normalized_path)
            self.output_dir = normalized_path
        else:
            print("No output directory set in configuration.")


    def show_tooltip(self, event, text):
        self.tooltip = utils.tooltip_show(event, text, self)

    def hide_tooltip(self):
        utils.tooltip_hide(self.tooltip)

    def section_specifics_setup(self):
        container = ttk.Frame(self)
        container.pack(side='top', fill='x', expand=False, padx=20, pady=10)

        label_checkbox_container = ttk.Frame(container)
        label_checkbox_container.pack(side='top', pady=(10, 5), anchor='center')

        specifics_label = ttk.Label(label_checkbox_container, text="CUSTOM SELECTION", style='Title.TLabel')
        specifics_label.pack(side='left', padx=(0, 5))

        self.specifics_var = BooleanVar()
        self.specifics_checkbox = ttk.Checkbutton(label_checkbox_container, text="enable",
                                                  variable=self.specifics_var,
                                                  command=self.section_specifics_checkbox_enable)
        self.specifics_checkbox.pack(side='left')


        button_container = ttk.Frame(container)
        button_container.pack(side='top', expand=False, pady=(5, 5), anchor='center')

        self.channel_grouping_button = ttk.Button(button_container, text="Channels", command=self.grouping_channel_load,
                                                  width=10, style='AudienceTab.TButton')
        self.channel_grouping_button.pack(side='left', padx=(1, 5), pady=(0, 0))
        self.tooltip = None
        self.channel_grouping_button.bind("<Enter>", lambda e: self.show_tooltip(e,
                                                                                 "Channel Grouping\nsheet: Content_Channel_Grouping\ncolumn: CHANNEL_NAME"))
        self.channel_grouping_button.bind("<Leave>", lambda e: self.hide_tooltip())

        # Products button
        self.product_grouping_button = ttk.Button(button_container, text="Products", command=self.load_product_grouping,
                                                  width=10, style='AudienceTab.TButton')
        self.product_grouping_button.pack(side='left', padx=(5, 1), pady=(0, 0))

        # Center the button container within its parent
        button_container.pack(anchor='center')
        self.product_grouping_button.bind("<Enter>", lambda e: self.show_tooltip(e,
                                                                                 "Products Grouping\nsheet: Content_Product_Grouping WS 241\ncolumn: PROD_NUM"))
        self.product_grouping_button.bind("<Leave>", lambda e: self.hide_tooltip())


        filter_container = ttk.Frame(container)
        filter_container.pack(side='top', fill='x', padx=10, pady=(5, 5))

        filter_label = ttk.Label(filter_container, text="Filter:")
        filter_label.pack(side='left', padx=(0, 5))

        self.filter_var = StringVar()
        self.filter_bar = ttk.Entry(filter_container, textvariable=self.filter_var, width=40)
        self.filter_bar.pack(side='left', fill='x', expand=True)
        self.filter_bar.bind('<KeyRelease>', self.filter_listboxes)

        self.specifics_frame = ttk.Frame(container)
        self.specifics_frame.pack(side='top', fill='both', expand=True, padx=10, pady=(5, 15))

        bus_chanl_frame = ttk.Frame(self.specifics_frame)
        bus_chanl_frame.pack(side='left', fill='both', expand=True, padx=0, pady=0)
        self.bus_chanl_label = ttk.Label(bus_chanl_frame, text="BUS_CHANL_NUM")
        self.bus_chanl_label.pack(side='top', padx=5)

        self.bus_chanl_num_listbox = Listbox(bus_chanl_frame, selectmode=MULTIPLE, exportselection=False)
        self.bus_chanl_num_listbox.pack(side='left', fill='both', expand=True)
        bus_chanl_scrollbar = ttk.Scrollbar(bus_chanl_frame, orient="vertical")
        bus_chanl_scrollbar.config(command=self.bus_chanl_num_listbox.yview)
        self.bus_chanl_num_listbox.config(yscrollcommand=bus_chanl_scrollbar.set)
        bus_chanl_scrollbar.pack(side="right", fill="y")

        # Product listbox setup
        prod_num_frame = ttk.Frame(self.specifics_frame)
        prod_num_frame.pack(side='left', fill='both', expand=True, padx=0, pady=0)
        ttk.Label(prod_num_frame, text="PROD_NUM:").pack(side='top', padx=0)

        self.prod_num_listbox = Listbox(prod_num_frame, selectmode=MULTIPLE, exportselection=False)
        self.prod_num_listbox.pack(side='left', fill='both', expand=True)
        prod_num_scrollbar = ttk.Scrollbar(prod_num_frame, orient="vertical")
        prod_num_scrollbar.config(command=self.prod_num_listbox.yview)
        self.prod_num_listbox.config(yscrollcommand=prod_num_scrollbar.set)
        prod_num_scrollbar.pack(side="right", fill="y")

        # Reset row frame
        self.reset_row_frame = ttk.Frame(container)

        style = ttk.Style()
        style.configure("Small.TLabel", font=("Arial", 12))

        self.row_count_label = ttk.Label(self.reset_row_frame, text="Selected Rows: 0", style="Small.TLabel")
        self.row_count_label.pack(side='left', padx=5, pady=0)

        self.prod_count_label = ttk.Label(self.reset_row_frame, text="Selected Products: 0", style="Small.TLabel")
        self.prod_count_label.pack(side='right', padx=5, pady=0)
        self.reset_label = ttk.Label(self.reset_row_frame, text="⟳", style="Small.TLabel", cursor="hand2")
        self.reset_label.pack(padx=5, pady=0)
        self.reset_label.bind("<Button-1>", self.section_specifics_listbox_reset)

        self.reset_row_frame.pack(side='top', fill='x', expand=False, padx=5, pady=(0, 0))
        self.reset_row_frame.pack_forget()

        self.prod_num_listbox.bind('<<ListboxSelect>>', self.section_specifics_counters_update)
        self.bus_chanl_num_listbox.bind('<<ListboxSelect>>', self.section_specifics_counters_update)

        # Enable specifics functionality
        self.section_specifics_checkbox_enable()

    def section_specifics_listboxes_values(self):
        """Loads unique values from the reference file into listboxes for selection."""
        if self.specifics_var.get() and self.df is not None:
            # Only clear listbox if not initialized
            if not hasattr(self, 'prod_num_map') or self.prod_num_map is None:
                self.prod_num_listbox.delete(0, 'end')
                unique_prod_num = sorted(set(str(item) for item in self.df['PROD_NUM'].unique()))
                for value in unique_prod_num:
                    self.prod_num_listbox.insert('end', value)
                self.prod_num_map = {str(value): str(value) for value in unique_prod_num}

            if not hasattr(self, 'bus_chanl_num_map') or self.bus_chanl_num_map is None:
                self.bus_chanl_num_listbox.delete(0, 'end')
                unique_bus_chanl_num = sorted(set(str(item) for item in self.df['BUS_CHANL_NUM'].unique()))
                for value in unique_bus_chanl_num:
                    self.bus_chanl_num_listbox.insert('end', value)
                self.bus_chanl_num_map = {str(value): str(value) for value in unique_bus_chanl_num}

        self.section_specifics_counters_update()
        self.specifics_frame.update_idletasks()

    def grouping_channel_load(self):
        """Loads the channel grouping file and updates the BUS_CHANL_NUM listbox in alphabetical order."""
        self.specifics_var.set(True)
        self.section_specifics_checkbox_enable()

        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            if not file_path:
                return

            channel_grouping_df = pd.read_excel(file_path, sheet_name='Content_Channel_Grouping')

            bus_chanl_nums = self.bus_chanl_num_listbox.get(0, 'end')
            updated_bus_chanl_nums = []

            for bus_chanl_num in bus_chanl_nums:
                try:
                    bus_chanl_num_float = float(bus_chanl_num)
                    match = channel_grouping_df[channel_grouping_df['BUS_CHANNEL_ID'] == bus_chanl_num_float]
                    if not match.empty:
                        updated_bus_chanl_nums.append((bus_chanl_num, match['CHANNEL_NAME'].values[0]))
                    else:
                        updated_bus_chanl_nums.append((bus_chanl_num, bus_chanl_num))
                except ValueError:
                    updated_bus_chanl_nums.append((bus_chanl_num, bus_chanl_num))

            # Sort the list alphabetically by display_value
            updated_bus_chanl_nums.sort(key=lambda x: str(x[1]).lower() if isinstance(x[1], str) else str(x[1]))

            # Update the BUS_CHANL_NUM listbox and map without affecting the product listbox or map
            self.bus_chanl_num_listbox.delete(0, 'end')
            if not hasattr(self, 'bus_chanl_num_map') or self.bus_chanl_num_map is None:
                self.bus_chanl_num_map = {}

            # Retain the existing mappings and just update with new data
            new_bus_chanl_map = {}
            for bus_chanl_num, display_value in updated_bus_chanl_nums:
                self.bus_chanl_num_listbox.insert('end', display_value)
                new_bus_chanl_map[display_value] = bus_chanl_num

            self.bus_chanl_num_map.update(new_bus_chanl_map)
            self.bus_chanl_label.config(text="CHANNEL_NAME")

        except PermissionError:
            show_message("Error", "Permission denied: unable to open the file.", type="error", master=self)
        except ValueError:
            show_message("Error", "Error in reading the Excel file. Ensure the sheet name and columns are correct.",
                         type="error", master=self)
        except Exception as e:
            show_message("Error", f"An unexpected error occurred: {e}", type="error", master=self)

    def load_product_grouping(self):
        """Loads the product grouping file and updates the PROD_NUM listbox in alphabetical order."""
        self.specifics_var.set(True)
        self.section_specifics_checkbox_enable()

        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            if not file_path:
                return

            product_grouping_df = pd.read_excel(file_path, sheet_name='Content_Product_Grouping WS 241')

            prod_nums = self.prod_num_listbox.get(0, 'end')
            updated_prod_nums = []

            for prod_num in prod_nums:
                try:
                    match = product_grouping_df[product_grouping_df['PROD_NUM'].astype(str) == str(prod_num)]
                    if not match.empty:
                        updated_prod_nums.append((prod_num, match['LOOKUP_KEY'].values[0]))
                    else:
                        updated_prod_nums.append((prod_num, prod_num))
                except ValueError:
                    updated_prod_nums.append((prod_num, prod_num))

            # Sort the list alphabetically by display_value (product name)
            updated_prod_nums.sort(key=lambda x: str(x[1]).lower() if isinstance(x[1], str) else str(x[1]))

            # Update only the PROD_NUM listbox without resetting the bus_chanl_num_map
            self.prod_num_listbox.delete(0, 'end')

            # Ensure the product maps are dictionaries before clearing them
            if not hasattr(self, 'prod_num_map') or self.prod_num_map is None:
                self.prod_num_map = {}
            if not hasattr(self, 'lookup_key_to_prod_num') or self.lookup_key_to_prod_num is None:
                self.lookup_key_to_prod_num = {}

            self.prod_num_map.clear()
            self.lookup_key_to_prod_num.clear()

            for prod_num, display_value in updated_prod_nums:
                self.prod_num_listbox.insert('end', display_value)
                self.prod_num_map[display_value] = prod_num
                self.lookup_key_to_prod_num[display_value] = prod_num

        except PermissionError:
            show_message("Error", "Permission denied: unable to open the file.", type="error", master=self)
        except ValueError:
            show_message("Error", "Error in reading the Excel file. Ensure the sheet name and columns are correct.",
                         type="error", master=self)
        except Exception as e:
            show_message("Error", f"An unexpected error occurred: {e}", type="error", master=self)

    def section_specifics_counters_update(self, event=None):
        """Updates the label to show the number of rows and products selected based on listbox selections."""
        if self.df is None:
            self.row_count_label.config(text="Selected Rows: 0")
            self.prod_count_label.config(text="Selected Products: 0")
            return

        selected_prod_nums = [self.prod_num_listbox.get(i) for i in self.prod_num_listbox.curselection()]
        selected_bus_chanl_display_values = [self.bus_chanl_num_listbox.get(i) for i in
                                             self.bus_chanl_num_listbox.curselection()]

        # Mapping to bus_chanl_num values
        selected_bus_chanl_nums = [self.bus_chanl_num_map.get(display_value, display_value) for display_value in
                                   selected_bus_chanl_display_values]

        # select matching
        if selected_bus_chanl_nums:
            related_prod_nums = set(
                self.df[self.df['BUS_CHANL_NUM'].astype(str).isin(selected_bus_chanl_nums)]['PROD_NUM'].dropna().astype(
                    str))

            # Mapping to LOOKUP_KEY values
            if hasattr(self, 'lookup_key_to_prod_num') and self.lookup_key_to_prod_num:
                related_lookup_keys = {key for key, value in self.lookup_key_to_prod_num.items() if
                                       value in related_prod_nums}
            else:
                related_lookup_keys = related_prod_nums

            self.prod_num_listbox.selection_clear(0, 'end')
            for i in range(self.prod_num_listbox.size()):
                if self.prod_num_listbox.get(i) in related_lookup_keys:
                    self.prod_num_listbox.selection_set(i)

        selected_prod_nums = [self.prod_num_listbox.get(i) for i in self.prod_num_listbox.curselection()]

        if not selected_prod_nums:
            selected_prod_nums = self.prod_num_listbox.get(0, 'end')
        if not selected_bus_chanl_nums:
            selected_bus_chanl_nums = list(self.bus_chanl_num_map.values()) if hasattr(self,
                                                                                       'bus_chanl_num_map') else []

        # Mapping LOOKUP_KEYto the original PROD_NUM
        selected_prod_nums_mapped = [self.prod_num_map.get(lookup_key, lookup_key) for lookup_key in selected_prod_nums]

        filtered_df = self.df[
            (self.df['PROD_NUM'].astype(str).isin(selected_prod_nums_mapped)) &
            (self.df['BUS_CHANL_NUM'].astype(str).isin(selected_bus_chanl_nums))
            ]

        self.row_count_label.config(text=f"Selected Rows: {len(filtered_df)}")
        self.prod_count_label.config(text=f"Selected Products: {len(set(selected_prod_nums))}")

        self.section_specifics_listbox_highlight_top(self.bus_chanl_num_listbox)
        self.section_specifics_listbox_highlight_top(self.prod_num_listbox)

    def section_specifics_listbox_reset(self, event=None):
        """Resets the selections in both listboxes, clears the filter, and resets the highlights."""
        self.filter_var.set("")

        if hasattr(self, 'global_selected_channels'):
            self.global_selected_channels.clear()

        self.prod_num_listbox.selection_clear(0, 'end')
        self.bus_chanl_num_listbox.selection_clear(0, 'end')

        self.section_specifics_listboxes_values()

        self.section_specifics_counters_update()

    def enable_specifics(self):
        """Enables the specifics checkbox and loads the necessary DataFrame values."""
        self.specifics_var.set(True)
        self.section_specifics_checkbox_enable()

        # Automatically load the product and channel grouping if the files are preloaded
        if self.config_manager.get_config().get('product_grouping_src'):
            self.load_product_grouping_from_config(self.config_manager.get_config().get('product_grouping_src'))

        if self.config_manager.get_config().get('channel_grouping_src'):
            self.grouping_channel_load_from_config(self.config_manager.get_config().get('channel_grouping_src'))

    def load_product_grouping_from_config(self, file_path):
        """Loads the product grouping file from the preloaded config and updates the PROD_NUM listbox."""
        self.specifics_var.set(True)
        self.section_specifics_checkbox_enable()

        try:
            product_grouping_df = pd.read_excel(file_path, sheet_name='Content_Product_Grouping WS 241')

            prod_nums = self.prod_num_listbox.get(0, 'end')
            updated_prod_nums = []

            for prod_num in prod_nums:
                try:
                    match = product_grouping_df[product_grouping_df['PROD_NUM'].astype(str) == str(prod_num)]
                    if not match.empty:
                        updated_prod_nums.append((prod_num, match['LOOKUP_KEY'].values[0]))
                    else:
                        updated_prod_nums.append((prod_num, prod_num))
                except ValueError:
                    updated_prod_nums.append((prod_num, prod_num))

            updated_prod_nums.sort(key=lambda x: str(x[1]).lower() if isinstance(x[1], str) else str(x[1]))

            self.prod_num_listbox.delete(0, 'end')
            self.prod_num_map = {}
            self.lookup_key_to_prod_num = {}

            for prod_num, display_value in updated_prod_nums:
                self.prod_num_listbox.insert('end', display_value)
                self.prod_num_map[display_value] = prod_num
                self.lookup_key_to_prod_num[display_value] = prod_num

        except Exception as e:
            show_message("Error", f"An error occurred while loading the product grouping: {e}", type="error",
                         master=self)

    def grouping_channel_load_from_config(self, file_path):
        """Loads the channel grouping file from the preloaded config and updates the BUS_CHANL_NUM listbox."""
        self.specifics_var.set(True)
        self.section_specifics_checkbox_enable()

        try:
            channel_grouping_df = pd.read_excel(file_path, sheet_name='Content_Channel_Grouping')

            bus_chanl_nums = self.bus_chanl_num_listbox.get(0, 'end')
            updated_bus_chanl_nums = []

            for bus_chanl_num in bus_chanl_nums:
                try:
                    bus_chanl_num_float = float(bus_chanl_num)
                    match = channel_grouping_df[channel_grouping_df['BUS_CHANNEL_ID'] == bus_chanl_num_float]
                    if not match.empty:
                        updated_bus_chanl_nums.append((bus_chanl_num, match['CHANNEL_NAME'].values[0]))
                    else:
                        updated_bus_chanl_nums.append((bus_chanl_num, bus_chanl_num))
                except ValueError:
                    updated_bus_chanl_nums.append((bus_chanl_num, bus_chanl_num))

            updated_bus_chanl_nums.sort(key=lambda x: str(x[1]).lower() if isinstance(x[1], str) else str(x[1]))

            self.bus_chanl_num_listbox.delete(0, 'end')
            self.bus_chanl_num_map = {}

            for bus_chanl_num, display_value in updated_bus_chanl_nums:
                self.bus_chanl_num_listbox.insert('end', display_value)
                self.bus_chanl_num_map[display_value] = bus_chanl_num

            self.bus_chanl_label.config(text="CHANNEL_NAME")

        except Exception as e:
            show_message("Error", f"An error occurred while loading the channel grouping: {e}", type="error",
                         master=self)

    def section_specifics_checkbox_enable(self):
        """Toggles the visibility and content of the specifics listboxes based on the checkbox state."""
        if self.specifics_var.get():
            # Check if a file is loaded
            if not self.file_path or not os.path.isfile(self.file_path):
                # If no file is loaded, prompt the user to load a file
                self.prompt_excel_load()
                # Check if a file was successfully loaded after prompting
                if self.file_path and os.path.isfile(self.file_path):
                    self.specifics_var.set(True)  # Set the checkbox to checked
                else:
                    self.specifics_var.set(False)  # Uncheck the checkbox if no file was loaded
                    return  # Exit the function if no file is loaded

            # Instead of clearing maps, just load the values into the listboxes
            self.section_specifics_listboxes_values()
            self.reset_row_frame.pack(side='top', fill='x', expand=False, padx=5, pady=(0, 0))

            # Disable the checkbox to prevent further modification
            self.specifics_checkbox.config(state='disabled')
        else:
            # Clear the listboxes and reset labels
            self.prod_num_listbox.delete(0, 'end')
            self.bus_chanl_num_listbox.delete(0, 'end')
            self.bus_chanl_label.config(text="BUS_CHANL_NUM")
            self.row_count_label.config(text="Selected Rows: 0")
            self.prod_count_label.config(text="Selected Products: 0")
            self.reset_row_frame.pack_forget()

            # No need to reset the maps here either; just clear the UI elements

    def start_processing(self):
        """Handles the start of the processing based on the selections."""
        if self.validate_all():
            references_month = self.references_month.get()
            references_year = self.references_year.get()
            target_start_year = self.target_start_year.get()
            target_end_year = self.target_end_year.get()
            file_path = self.file_path

            if self.file_path is None:
                show_message("Error", "No file selected.", type='error', master=self, custom=True)
                return

            specifics_enabled = self.specifics_var.get()

            current_selected_bus_chanl_display_values = [self.bus_chanl_num_listbox.get(i) for i in
                                                         self.bus_chanl_num_listbox.curselection()]
            if hasattr(self, 'global_selected_channels'):
                global_selected_bus_chanl_display_values = list(self.global_selected_channels)
            else:
                global_selected_bus_chanl_display_values = []

            combined_selected_bus_chanl_display_values = set(
                current_selected_bus_chanl_display_values + global_selected_bus_chanl_display_values)

            combined_selected_bus_chanl_nums_values = [self.bus_chanl_num_map.get(display_value, display_value) for
                                                       display_value in combined_selected_bus_chanl_display_values]

            print(f"References Month: {references_month}, Year: {references_year}")
            print(f"Target Start Year: {target_start_year}, End Year: {target_end_year}")
            print(f"File Path: {file_path}")
            print(f"Specifics Enabled: {specifics_enabled}")
            if specifics_enabled:
                print(f"Selected BUS_CHANL_NUMs: {combined_selected_bus_chanl_nums_values}")

            if self.output_dir is None:
                show_message("Error", "No output directory selected.", type='error', master=self, custom=True)
                return

            start_time = time.time()

            self.call_script(
                references_month,
                references_year,
                target_start_year,
                target_end_year,
                file_path,
                specifics_enabled,
                None,
                combined_selected_bus_chanl_nums_values
            )

            end_time = time.time()
            duration = end_time - start_time
            show_message("Info", f"Parsing completed in {duration:.2f} seconds.", type='info', master=self, custom=True)
        else:
            show_message("Error", "Validation failed. Please correct the errors and try again.", type='error',
                         master=self, custom=True)

    def call_script(self, references_month, references_year, target_start_year, target_end_year,
                    file_path, specifics_enabled, prod_nums, bus_chanl_nums):

        if not self.base_dir:
            raise ValueError("Base directory is not set.")
        script_path = os.path.join(self.base_dir, 'parser', 'parser_audience.py')
        script_path = os.path.abspath(script_path)
        print(f"Script Path: {script_path}")

        output_dir = self.output_dir

        args = {
            "references_month": references_month,
            "references_year": references_year,
            "target_start_year": target_start_year,
            "target_end_year": target_end_year,
            "file_path": file_path,
            "output_dir": output_dir,
            "specifics_enabled": specifics_enabled,
            "prod_nums": prod_nums,
            "bus_chanl_nums": bus_chanl_nums
        }

        subprocess.run(["python", script_path, json.dumps(args)])

    def sections_reference_target_datefields(self, parent, context):
        if context == 'REFERENCE':
            ttk.Label(parent, text="Date (MM - YYYY):").pack(side='left')
            self.references_month = ttk.Entry(parent, width=3, validate='key',
                                              validatecommand=(self.register(self.validate_month), '%P'))
            self.references_month.pack(side='left', padx=(0, 2))
            self.references_year = ttk.Entry(parent, width=5, validate='key',
                                             validatecommand=(self.register(self.validate_year), '%P'))
            self.references_year.pack(side='left', padx=(2, 10))

            help_label = ttk.Label(parent, text="?", cursor="hand2", font=("Helvetica", 11))
            help_label.pack(side='left', padx=(2, 10))
            help_label.bind("<Enter>", self.tooltip_reference_update)
            help_label.bind("<Leave>", lambda e: self.hide_tooltip())


            if self.current_month == 1:
                self.references_month.insert(0, str(12))
                self.references_year.insert(0, str(self.current_year - 1))
            else:
                self.references_month.insert(0, str(self.current_month - 1))
                self.references_year.insert(0, str(self.current_year))
            ttk.Button(parent, text="✓", command=self.validate_references, style='AudienceTab.TButton').pack(
                side='right', padx=(0, 10), pady=(0, 0))
        elif context == 'TARGET':
            ttk.Label(parent, text="From:").pack(side='left')
            self.target_start_year = ttk.Entry(parent, width=5, validate='key',
                                               validatecommand=(self.register(self.validate_year), '%P'))
            self.target_start_year.pack(side='left', padx=(0, 2))
            self.target_start_year.insert(0, str(self.current_year + 1))
            ttk.Label(parent, text="To:").pack(side='left')
            self.target_end_year = ttk.Entry(parent, width=5, validate='key',
                                             validatecommand=(self.register(self.validate_year), '%P'))
            self.target_end_year.pack(side='left', padx=(2, 10))
            self.target_end_year.insert(0, str(self.current_year + 1))
            ttk.Button(parent, text="✓", command=self.validate_target, style='AudienceTab.TButton').pack(side='right', padx=(0, 10), pady=(0, 0))

    def tooltip_reference_update(self, event):
        """Dynamically update and show the tooltip based on the current input values."""
        month_names_fr = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre",
                          "Octobre", "Novembre", "Décembre"]
        references_month = self.references_month.get()
        references_year = self.references_year.get()
        target_start_year = self.target_start_year.get()
        target_end_year = self.target_end_year.get()

        if not references_month or not references_year or not target_start_year or not target_end_year:
            help_text = "Remplir toutes les dates pour obtenir de l'aide"
            self.show_tooltip(event, help_text)
            return

        try:
            references_month_int = int(references_month)
            if 1 <= references_month_int <= 12:
                month_str = month_names_fr[references_month_int - 1]
            else:
                month_str = month_names_fr[self.current_month - 1]
        except ValueError:
            month_str = month_names_fr[self.current_month - 1]

        if target_start_year == target_end_year:
            if references_month_int == 12:
                help_text = f'En utilisant toute l\'année {references_year}, calculer {target_start_year}'
            else:
                help_text = f'Sans aller au delà de {month_str} {references_year}, calculer {target_start_year}'
        else:
            if references_month_int == 12:
                help_text = f'En utilisant toute l\'année {references_year}, calculer {target_start_year} à {target_end_year} inclus'
            else:
                help_text = f'Sans aller au delà de {month_str} {references_year}, calculer {target_start_year} à {target_end_year} inclus'

        self.show_tooltip(event, help_text)

    def load_file(self, path):
        src_audience_path = self.config_data.get('audience_src')
        if path:
            self.section_reference_details_update(path)

    def button_select_sources(self, parent, context):
        if context == 'REFERENCE':
            ttk.Button(parent, text="Source File", command=self.prompt_excel_load, style='AudienceTab.TButton').pack(side='left', padx=10)
        if context == 'TARGET':
            ttk.Button(parent, text="Forecast Folder", command=self.section_target_output_location, style='AudienceTab.TButton').pack(side='left', padx=10)

    def setup_buttons_and_entries(self, parent, context):
        """Setup buttons and entry fields for user interaction."""
        self.button_select_sources(parent, context)
        self.sections_reference_target_datefields(parent, context)
        self.setup_show_columns_button(parent, context)



    def references_file_details(self, parent):
        """Configure and place the file details label within the given container."""
        self.file_details_label = ttk.Label(parent, text="Il faut charger le fichier d'audience de référence, Source File.", anchor='w')
        self.file_details_label.pack(side='top', fill='x', expand=False, padx=10, pady=(10, 5))

    def prompt_excel_load(self):
        filetypes = [("Excel files", "*.xlsx *.xls")]
        filepath = filedialog.askopenfilename(filetypes=filetypes)
        if filepath:
            self.section_reference_details_update(filepath)
            self.df = pd.read_excel(filepath)

    def section_reference_details_update(self, file_path):
        self.file_path = file_path
        try:
            df = pd.read_excel(file_path)
            self.config_manager.update_config('audience_src', file_path)
            print("File loaded, checking content...")
            self.df = pd.read_excel(file_path)
            if df.empty:
                print("DataFrame is empty after loading.")
            else:
                rows, cols = df.shape
                relative_path = '/'.join(file_path.split('/')[-3:])
                self.file_details_label.config(text=f".../{relative_path} \t rows: {rows} ~ columns: {cols}")
        except PermissionError as e:
            show_message("Error", f"Exception REFERENCE FILE ALREADY OPEN, CLOSE IT:\n {str(e)}", type='error', master=self, custom=True)
        except Exception as e:
            print(f"Exception occurred: {str(e)}")
            print(traceback.format_exc())
            self.file_details_label.config(text="Failed to load file or file is empty")
            show_message("Error", f"Exception occurred: {str(e)}", type='error', master=self, custom=True)

    def setup_show_columns_button(self, parent, context):
        """Sets up a button to show column names from the loaded DataFrame."""
        if context == 'REFERENCE':
            self.show_columns_button = ttk.Button(parent, text="Metadata Exploration", command=self.section_reference_button_columns_show, style='AudienceTab.TButton')
            self.show_columns_button.pack(side='right', padx=10)

    def section_reference_button_columns_show(self):
        """Displays the column names from the loaded DataFrame."""
        if self.file_path:
            try:
                df = pd.read_excel(self.file_path)
                columns = '\n'.join(df.columns)
                show_message("Columns", f"Columns in the file:\n{columns}", type='info', master=self, custom=True)
            except Exception as e:
                show_message("Error", f"Failed to load file:\n{str(e)}", type='error', master=self, custom=True)
        else:
            show_message("Error", "Load an Excel file first.", type='info', master=self, custom=True)

    def tab_style(self):
        """Configure styles used within the tab."""
        style = ttk.Style(self)
        style.configure('TFrame')
        style.configure('Title.TLabel', font=('Arial', 12, 'underline'))
        style.configure('AudienceTab.TButton', padding=[5, 2], font=('Arial', 10))

    def section_references_setup(self):
        container = ttk.Frame(self)
        container.pack(side='top', fill='x', expand=False, padx=20, pady=10)
        ttk.Label(container, text="REFERENCE YEAR OF CALCULATION", style='Title.TLabel').pack(side='top', padx=10, pady=(10, 5))
        self.references_file_details(container)
        self.setup_buttons_and_entries(container, 'REFERENCE')

    def section_target_setup(self):
        container = ttk.Frame(self)
        container.pack(side='top', fill='x', expand=False, padx=20, pady=10)
        ttk.Label(container, text="TARGET", style='Title.TLabel').pack(side='top', padx=10, pady=(10, 5))

        self.output_path = ttk.Entry(container)
        self.output_path.pack(side='top', fill='x', padx=10, pady=(5, 5))
        self.setup_buttons_and_entries(container, context='TARGET')

        audience_dest = self.config_data.get('audience_dest')
        if audience_dest:
            self.output_path.insert(0, audience_dest)

    def view_result(self):
        output_path = self.output_path.get()
        print(f"Output path: {output_path}")
        result_file = os.path.join(output_path, "forecast_audience.xlsx")

        print(f"Result file path: {result_file}")
        if os.path.isfile(result_file):
            os.startfile(result_file)
        else:
            show_message("Error",
                         "The result file does not exist. Please make sure the processing is completed successfully.",
                         type='error', master=self, custom=True)


    def validate_all(self):
        valid_references = self.validate_references()

        valid_target = self.validate_target()
        return valid_references and valid_target

    def validate_year(self, P):
        """Validate the year entry to ensure it meets specified conditions."""
        if P == "" or (P.isdigit() and P.startswith("2") and len(P) <= 4 and int(P) <= 2064):
            return True
        return False

    def validate_references(self):
        if self.file_path:
            try:
                df = pd.read_excel(self.file_path)
                month = int(self.references_month.get())
                year = int(self.references_year.get())

                current_date = datetime.now()
                reference_date = datetime(year, month, 1)

                if reference_date >= datetime(current_date.year, current_date.month, 1):
                    show_message("Error", "The reference date cannot be in the current month or the future.", type='error', master=self, custom=True)
                    return False
                else:
                    return self.validation_references_dates(df, year, month)
            except ValueError:
                show_message("Error", "Invalid date. Please enter a valid month and year.", type='error', master=self, custom=True)
                return False
            except Exception as e:
                show_message("Error", f"Failed to load file:\n{str(e)}", type='error', master=self, custom=True)
                return False
        else:
            show_message("Error", "Load an Excel file first.", type='error', master=self, custom=True)
            return False

    def validate_target(self):
        if self.file_path:
            try:
                if not self.output_path.get():
                    show_message("Error", "Select an output folder first.", type='error', master=self, custom=True)
                    return False

                if not self.references_year.get() or not self.references_month.get():
                    show_message("Error", "A reference date must be set.", type='error', master=self, custom=True)
                    return False

                reference_year = int(self.references_year.get())
                reference_month = int(self.references_month.get())
                start_year = int(self.target_start_year.get())
                end_year = int(self.target_end_year.get())

                df = pd.read_excel(self.file_path)

                current_year = datetime.now().year

                if start_year == current_year:
                    show_message("Error", "Target start year cannot be the current year.", type='error', master=self,
                                 custom=True)
                    return False

                if start_year > end_year:
                    show_message("Error", "Target 'From' year cannot be after the target 'To' year.", type='error',
                                 master=self, custom=True)
                    return False

                if reference_month != 12:
                    if start_year < reference_year or end_year < reference_year:
                        show_message("Error",
                                     "Target years must be after or equal to the reference year when the reference month is not December.",
                                     type='error', master=self, custom=True)
                        return False
                else:
                    if start_year <= reference_year or end_year <= reference_year:
                        show_message("Error",
                                     "Target years must be strictly after the reference year when the reference month is December.",
                                     type='error', master=self, custom=True)
                        return False

                if abs(start_year - end_year) > 10:
                    show_message("Error", "The difference between start and end year cannot exceed 10 years.",
                                 type='error', master=self, custom=True)
                    return False
                else:
                    show_message("Validation", "Target years are valid.", type='info', master=self, custom=True)
                    return True
            except ValueError:
                show_message("Error", "Invalid target year. Please enter a valid year.", type='error', master=self,
                             custom=True)
                return False
            except Exception as e:
                show_message("Error", f"Failed to load file:\n{str(e)}", type='error', master=self, custom=True)
                return False
        else:
            show_message("Error", "Load an Excel file first.", type='error', master=self, custom=True)
            return False

    def validation_references_dates(self, df, year, month):
        """Checks if the date exists in the loaded data and updates the user."""
        mask = (df['PERIOD_YEAR'] == year) & (df['PERIOD_MONTH'] == month)
        if mask.any():
            show_message("Validation", "Reference file: Date is valid and found in the file.", type='info', master=self, custom=True)
        else:
            specific_data = df[(df['PERIOD_YEAR'] == year)]
            show_message("Validation",
                         f"Date not found in the file. Debug: Year({year}), Month({month})\nSample rows where year matches:\n{specific_data.head()}",
                         type='error', master=self, custom=True)
            return False
        return True

    def validate_month(self, P):
        """Validate the month entry to ensure it's empty or a valid month number."""
        return P == "" or (P.isdigit() and 1 <= int(P) <= 12)

