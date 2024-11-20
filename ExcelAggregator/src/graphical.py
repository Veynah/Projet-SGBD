import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys
import os
import glob
import threading
import shutil


def install_package(package):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])


def check_and_install_packages():
    required_packages = ['openpyxl']
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"{package} not found. Installing...")
            install_package(package)
            print(f"{package} installed successfully.")


class SimpleGUI:
    def __init__(self, root):
        self.root = root
        self.python_executable = os.getenv('PYTHON_EXEC', 'python')
        self.root.title("Simple Excel Parser and Centralizer")
        self.root.geometry('700x350')
        self.apply_dark_mode()

        self.select_button = tk.Button(root, text="Choose Directory", command=self.select_directory, bg="#333333",
                                       fg="white")
        self.select_button.pack(pady=10, anchor='center')

        self.directory_label = tk.Entry(root, bg="#222222", fg="white", bd=0, state='readonly',
                                        readonlybackground="#222222", relief=tk.FLAT, highlightthickness=0)
        self.directory_label.pack(pady=5, fill=tk.X, padx=10, anchor='center')

        self.list_frame = tk.Frame(root, bg="#222222")
        self.list_frame.pack(pady=10, expand=True, fill=tk.BOTH)

        self.xlsx_frame = tk.Frame(self.list_frame, bg="#222222")
        self.xlsx_frame.pack(side=tk.LEFT, padx=10, fill=tk.BOTH, expand=True)

        self.tsv_frame = tk.Frame(self.list_frame, bg="#222222")
        self.tsv_frame.pack(side=tk.RIGHT, padx=10, fill=tk.BOTH, expand=True)

        self.xlsx_label = tk.Label(self.xlsx_frame, text="Your Contracts", bg="#222222", fg="white",
                                   font=('Arial', 10, 'underline'))
        self.xlsx_label.pack(side=tk.TOP, anchor='center')

        self.tsv_label = tk.Label(self.tsv_frame, text="Parsed Contracts", bg="#222222", fg="white",
                                  font=('Arial', 10, 'underline'))
        self.tsv_label.pack(side=tk.TOP, anchor='center')

        self.xlsx_listbox = tk.Listbox(self.xlsx_frame, bg="#333333", fg="white")
        self.xlsx_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.xlsx_listbox.bind('<Double-1>', self.open_xlsx_file)

        self.tsv_listbox = tk.Listbox(self.tsv_frame, bg="#333333", fg="white")
        self.tsv_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.tsv_listbox.bind('<Double-1>', self.open_tsv_file)

        self.button_frame = tk.Frame(root, bg="#222222")
        self.button_frame.pack(side=tk.BOTTOM, pady=10, anchor='center')

        self.status_label = tk.Label(self.button_frame, text="", bg="#222222", fg="white", font=('Arial', 10))
        self.status_label.pack(side=tk.BOTTOM, pady=5, anchor='center')

        self.process_button = tk.Button(self.button_frame, text="Process Contracts", command=self.start_processing,
                                        bg="#555555", fg="white")
        self.process_button.pack(side=tk.LEFT, padx=5, anchor='center')

        self.view_result_button = tk.Button(self.button_frame, text="View Result", command=self.view_result,
                                            bg="#555555", fg="white")
        self.view_result_button.pack(side=tk.RIGHT, padx=5, anchor='center')

        self.extract_button = tk.Button(self.button_frame, text="Extract Results", command=self.extract_results,
                                        bg="#555555", fg="white")
        self.extract_button.pack(side=tk.RIGHT, padx=5, anchor='center')

        self.directory = ""

        self.update_tsv_list()

        self.check_result_file()

        self.create_tooltips()

    def apply_dark_mode(self):
        self.root.configure(bg="#222222")


    def create_tooltips(self):
        self.create_tooltip(self.tsv_label,
                            "outputs/tsv/*.tsv: parsed Contracts from which the centralized spreadsheet is made")
        self.create_tooltip(self.xlsx_label,
                            "Each single spreadsheet before parsing, won't be modified in the process")

    def create_tooltip(self, widget, text):
        tooltip = tk.Toplevel(widget)
        tooltip.withdraw()
        tooltip.wm_overrideredirect(True)
        tooltip.config(bg="#333333", padx=5, pady=5)
        label = tk.Label(tooltip, text=text, bg="#333333", fg="white", wraplength=200, justify=tk.LEFT)
        label.pack()

        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            tooltip.wm_geometry(f"+{x}+{y}")
            tooltip.deiconify()

        def leave(event):
            tooltip.withdraw()

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def select_directory(self):
        initial_dir = os.path.join(os.path.dirname(__file__), 'contracts')
        self.directory = filedialog.askdirectory(initialdir=initial_dir)
        if self.directory:
            self.directory_label.config(state='normal')
            self.directory_label.delete(0, tk.END)
            self.directory_label.insert(0, self.directory)
            self.directory_label.config(state='readonly')
            self.xlsx_listbox.config(selectmode=tk.EXTENDED)
            self.update_xlsx_list()
            self.process_button.pack(side=tk.LEFT, padx=5)

    def update_xlsx_list(self):
        self.xlsx_listbox.delete(0, tk.END)

        xlsx_files = glob.glob(os.path.join(self.directory, "*.xlsx"))

        for file in xlsx_files:
            self.xlsx_listbox.insert(tk.END, os.path.basename(file))

    def update_tsv_list(self):
        self.tsv_listbox.delete(0, tk.END)

        tsv_files = glob.glob(os.path.join(os.path.dirname(__file__), "outputs/tsv/*.tsv"))

        for file in tsv_files:
            self.tsv_listbox.insert(tk.END, os.path.basename(file))

    def get_python_executable(self):
        return sys.executable

    def run_script(self, script_name, file_path):
        try:
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_name)
            command = [self.python_executable, script_path, file_path]  # Now it takes a file path

            print(f"Running command: {' '.join(command)}")

            env = os.environ.copy()
            env['RUNNING_AS_SUBPROCESS'] = '1'

            result = subprocess.run(command, check=True, capture_output=True, text=True, env=env, creationflags=subprocess.CREATE_NO_WINDOW)

            print(result.stdout)

            return True
        except subprocess.CalledProcessError as e:
            print(f"Error occurred while executing {script_name}.")
            print(e.stderr)
            return False

    def process_contracts(self):
        selected_directory = self.directory

        if not selected_directory:
            messagebox.showerror("Error", "Please select a contracts directory first.", parent=self.root)
            return

        selected_indices = self.xlsx_listbox.curselection()
        xlsx_files = [self.xlsx_listbox.get(i) for i in selected_indices] if selected_indices else glob.glob(
            os.path.join(selected_directory, "*.xlsx"))

        for file in xlsx_files:
            file_path = os.path.join(selected_directory, file)
            self.update_status_label(f"Processing: {os.path.basename(file)}")
            success = self.run_script('contract_parser.py', file_path)
            if not success:
                messagebox.showerror("Error", f"Failed to parse {os.path.basename(file_path)}", parent=self.root)
                self.update_status_label("")
                return

        self.update_status_label("Centralizing parsed files...")
        success_centralize = self.run_script('contract_exporter.py', selected_directory)
        self.update_status_label("")
        if success_centralize:
            messagebox.showinfo("Success", "Processed contracts successfully!", parent=self.root)
            self.update_xlsx_list()
            self.update_tsv_list()
            self.check_result_file()
        else:
            messagebox.showerror("Error", "Failed to centralize .tsv files to .xlsx.", parent=self.root)

    def start_processing(self):
        threading.Thread(target=self.process_contracts).start()

    def update_status_label(self, text):
        self.status_label.config(text=text)
        self.root.update_idletasks()

    def open_xlsx_file(self, event):
        selection = self.xlsx_listbox.curselection()
        if selection:
            file_name = self.xlsx_listbox.get(selection[0])
            file_path = os.path.join(self.directory, file_name)
            os.startfile(file_path)

    def open_tsv_file(self, event):
        selection = self.tsv_listbox.curselection()
        if selection:
            file_name = self.tsv_listbox.get(selection[0])
            file_path = os.path.join(os.path.dirname(__file__), "outputs/tsv", file_name)
            os.startfile(file_path)

    def view_result(self):
        result_files = glob.glob(os.path.join(os.path.dirname(__file__), "outputs/xlsx/centralized_data.xlsx"))
        if result_files:
            for file in result_files:
                os.startfile(file)
        else:
            messagebox.showerror("Error", "No result file found.", parent=self.root)

    def check_result_file(self):
        result_file = os.path.join(os.path.dirname(__file__), "outputs/xlsx/centralized_data.xlsx")
        if os.path.exists(result_file):
            self.view_result_button.pack(side=tk.RIGHT, padx=5, anchor='center')
        else:
            self.view_result_button.pack_forget()

    def extract_results(self):
        output_directory = filedialog.askdirectory(title="Select Directory to Save Results")
        if not output_directory:
            return

        internal_results_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
        if not os.path.exists(internal_results_dir):
            messagebox.showerror("Error", "No outputs found to extract.", parent=self.root)
            return

        try:
            shutil.copytree(internal_results_dir, os.path.join(output_directory, 'outputs'))
            messagebox.showinfo("Success", f"Results have been saved to {output_directory}", parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract outputs: {e}", parent=self.root)


if __name__ == "__main__":
    if os.getenv('RUNNING_AS_SUBPROCESS'):
        print("Called as a subprocess, not launching GUI.")
    else:
        check_and_install_packages()
        root = tk.Tk()
        gui = SimpleGUI(root)
        root.mainloop()
