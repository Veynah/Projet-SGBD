import ctypes
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


def set_window_icon(window, icon_name='favicon.ico'):
    """Sets the icon for the given tkinter window.

    Args:
        window: The tkinter window (Tk or Toplevel) where the icon will be set.
        icon_name: The name of the icon file (default is 'favicon.ico').
    """
    icon_path = os.path.join(os.path.dirname(__file__), '..', 'static', icon_name)

    if os.path.exists(icon_path):
        window.iconbitmap(icon_path)
    else:
        print(f"Error: Icon file not found: {icon_path}")

def get_base_dir(main_file_path):
    """Determine the base directory for the application."""
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
        print(f"Application is frozen. _MEIPASS directory is {base_dir}")
    else:
        base_dir = os.path.dirname(os.path.abspath(main_file_path))
        print(f"Application is not frozen. Base directory is {base_dir}")

    return base_dir

def open_file_and_update_config(config_manager, config_key, title="Select a file", filetypes=None):
    """
    Opens a file dialog to select a file and updates the configuration with the selected file path.

    Args:
        config_manager (ConfigManager): The configuration manager to update.
        config_key (str): The key in the configuration to update with the selected file path.
        title (str): The title of the file dialog window.
        filetypes (list): A list of tuples specifying the file types (e.g., [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]).

    Returns:
        str: The selected file path, or None if no file was selected.
    """
    filetypes = [("Excel files", "*.xlsx *.xls")]
    filepath = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if filepath:
        config_manager.update_config(config_key, filepath)
        try:
            df = pd.read_excel(filepath)
            return df
        except Exception as e:
            show_message("Error", f"Failed to load Excel file: {e}", type="error")
    return None

def center_window(window, master=None, width=420, height=650):
    """
    Centers the window on the screen based on the master's size.

    Args:
        window (tk.Toplevel or tk.Tk): The window to be centered.
        master (tk.Tk, optional): The master window.
        width (int): The width of the window to be centered.
        height (int): The height of the window to be centered.
    """
    if master is None:
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
    else:
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()

    position_top = int(screen_height / 2 - height / 2)
    position_right = int(screen_width / 2 - width / 2)
    window.geometry(f"{width}x{height}+{position_right}+{position_top}")


def create_styled_button(parent, text, command=None, width=10):
    """
    Creates a styled button with a specific text and command.

    Args:
        parent (tk.Widget): The parent widget for the button.
        text (str): The text to display on the button.
        command (callable, optional): The function to call when the button is clicked.
        width (int, optional): The width of the button.

    Returns:
        ttk.Button: The created styled button.
    """
    return ttk.Button(parent, text=text, command=command, style='AudienceTab.TButton', width=width)


def create_menu(window, menu_items):
    """
    Creates a menu in the window with given menu items.

    Args:
        window (tk.Tk): The window to attach the menu to.
        menu_items (list): A list of dictionaries, each containing 'label' and 'command' keys.
    """
    menubar = tk.Menu(window)
    window.config(menu=menubar)
    for item in menu_items:
        submenu = item.get("submenu")
        if submenu:
            cascade_menu = tk.Menu(menubar, tearoff=0, background='SystemButtonFace', fg='black')
            for subitem in submenu:
                cascade_menu.add_command(label=subitem["label"], command=subitem["command"])
            menubar.add_cascade(label=item["label"], menu=cascade_menu)
        else:
            menubar.add_command(label=item["label"], command=item["command"])


def select_file(callback, filetypes):
    """
    Handles selecting a file and passes the selected file path to the callback function.

    Args:
        callback (callable): The function to call with the selected file path.
        filetypes (list): A list of tuples specifying the allowed file types.
    """
    file_selected = filedialog.askopenfilename(filetypes=filetypes)
    if file_selected:
        callback(file_selected)

popup_windows = []

def show_message(title, message, type='info', master=None, custom=False):
    """
    Displays a message box of specified type. Optionally uses a custom dialog.

    Args:
        title (str): The title of the message box.
        message (str): The message to display.
        type (str): The type of message box ('info' or 'error').
        master (tk.Widget, optional): The parent widget if using a custom dialog.
        custom (bool): Whether to use the custom dialog.
    """
    if custom and master is not None:
        show_custom_message(master, title, message, type)
    elif type == 'info':
        tk.messagebox.showinfo(title, message)
    elif type == 'error':
        tk.messagebox.showerror(title, message)

def show_custom_message(master, title, message, type='info'):
    """
    Displays a custom message box with selectable text that adapts its height based on content length.
    Scrolls appear after 20 lines.

    Args:
        master (tk.Widget): The parent widget.
        title (str): The title of the message box.
        message (str): The message to display.
        type (str): The type of message box ('info' or 'error').
    """
    global popup_windows

    top = tk.Toplevel(master)
    top.title(title)
    set_window_icon(top)

    screen_width = master.winfo_screenwidth()
    screen_height = master.winfo_screenheight()
    window_width = min(350, max(300, len(message) * 7))

    characters_per_line = window_width // 7
    estimated_lines = len(message) // characters_per_line + (1 if len(message) % characters_per_line else 0)
    window_height = min(150 + 18 * min(20, estimated_lines), screen_height - 100)

    x_coordinate = (screen_width - window_width) // 2
    y_coordinate = (screen_height - window_height) // 2
    top.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    top.transient(master)

    bg_color = '#dddddd' if type == 'info' else '#ffdddd'

    text_frame = tk.Frame(top)
    text_frame.pack(fill='both', expand=True)
    text_scroll = tk.Scrollbar(text_frame)
    text_scroll.pack(side='right', fill='y')
    text_widget = tk.Text(text_frame, wrap='word', yscrollcommand=text_scroll.set,
                          background=bg_color, borderwidth=0, highlightthickness=0)
    text_widget.insert('end', message)
    text_widget.config(state='disabled', height=min(25, estimated_lines))
    text_widget.pack(pady=20, padx=20, fill='both', expand=True)
    text_scroll.config(command=text_widget.yview)

    def on_click(event):
        if event.widget is not text_widget:
            top.destroy()

    def on_close(event=None):
        top.destroy()
        if top in popup_windows:
            popup_windows.remove(top)
        if popup_windows:
            popup_windows[-1].focus_force()

    top.bind("<FocusOut>", on_close)
    top.bind("<Button-1>", on_click)
    top.protocol("WM_DELETE_WINDOW", on_close)
    top.focus_force()

    popup_windows.append(top)
def select_directory(entry_field):
    """
    Opens a directory dialog to select a folder and updates the entry field.

    Args:
        entry_field (ttk.Entry): The entry field to update with the selected directory path.
    """
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder_selected)

def clean_file_path(file_path):
    """
    Cleans up the file path by removing leading and trailing quotes.

    Args:
        file_path (str): The file path to clean.

    Returns:
        str: The cleaned file path.
    """
    return file_path.strip().strip('"')


current_tooltip = None  # Global variable to keep track of the current tooltip

def tooltip_show(event, text, master):
    global current_tooltip

    # If a tooltip is already being shown, destroy it before showing a new one
    if current_tooltip:
        tooltip_hide(current_tooltip)

    x, y = master.winfo_pointerxy()
    tooltip = tk.Toplevel(master)
    tooltip.wm_overrideredirect(True)
    tooltip.wm_geometry(f"+{x + 10}+{y + 10}")
    label = ttk.Label(tooltip, text=text, background="grey", relief="solid", borderwidth=1, padding=5)
    label.pack()

    current_tooltip = tooltip  # Store the current tooltip reference
    return tooltip

def tooltip_hide(tooltip):
    global current_tooltip

    if tooltip and tooltip.winfo_exists():
        tooltip.destroy()

    current_tooltip = None

def prevent_multiple_instances(mutex_name="my_unique_application_mutex"):
    """Prevents multiple instances of the application by using a named mutex."""
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    last_error = ctypes.windll.kernel32.GetLastError()

    if last_error == 183:  # ERROR_ALREADY_EXISTS
        print("Another instance of this application is already running.")
        sys.exit()


