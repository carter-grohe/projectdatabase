### Carter Grohe UVA Statistics/Data Science '27

### This code creates a GUI that allowed Hall MileOne Automotive to archive their entire database, grouped
### by dealership, through *XXX*, a very specific car-dealership software. It required PyAutoGUI and
### old terminal code due to its age, and uses mss for screen capture, as well as numpy and pandas.

### It also requires an Excel sheet for default settings which were in the accompanying folder. (ArchiveAppSettings)

import subprocess
import sys

required_packages = ["pyautogui", "tkinter", "customtkinter", "pandas",
                     "openpyxl", "keyboard", "screeninfo", "mss", "PIL", "numpy"]

def check_and_install_packages():
    for package in required_packages:
        try:
            __import__(package)
            print(f"{package} is already installed.")
        except ImportError:
            print(f"{package} not found. Installing...")
            subprocess.run([sys.executable, "-m", "pip", "install", package], check=True)

check_and_install_packages() # The function installs any packages you need installed for this.

import pyautogui as gui
import ctypes
import time
from time import sleep
import pandas as pd
import keyboard
import threading
from datetime import datetime
from pathlib import Path
import mss
from PIL import ImageChops, Image
import numpy as np
import calendar

current_date = datetime.now().strftime("%d%b%y").upper() ## DDMMMYY Format

today = datetime.today()
last_day = calendar.monthrange(today.year, today.month)[1]
last_dotm = datetime(today.year, today.month, last_day).strftime("%d%b%y").upper()

###########################################################################################################
## Functions for General Use and QoL in *XXX*

def turn_off_capslock():
    if ctypes.WinDLL("User32.dll").GetKeyState(0x14) == 1: # IF CAPSLOCK is on: (CAPSLOCK = 0x14)
        gui.press('capslock') #Turn off.
        sleep(0.4)

turn_off_capslock()

def up_arrow():
    gui.press('capslock')
    gui.keyDown('esc')
    gui.press('[')
    gui.press('a')
    gui.keyUp('esc')
    gui.press('capslock')
    sleep(0.1)

def down_arrow():
    gui.press('capslock')
    gui.keyDown('esc')
    gui.press('[')
    gui.press('b')
    gui.keyUp('esc')
    gui.press('capslock')
    sleep(0.2)

def quit_to_menu():
    sleep(2)
    gui.hotkey('shift' , 'f4')
    down_arrow()
    down_arrow()
    gui.press('enter')

def show_frame(frame):
    frame.tkraise()

###########################################################################################################
#### More Imports and Customization

import customtkinter as ctk
import tkinter as tk
from screeninfo import get_monitors

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)

    def flush(self):
        pass

############################################################################################################
## Bringing the Right files to the .exe and Loading in the Excel spreadsheet
def resource_path(relative_path: str) -> Path:
    try:
        base_path = Path(sys.executable).parent.resolve()
    except Exception:
        base_path = Path(".").resolve()
    return base_path / relative_path
excel_path = resource_path(r"ArchiveAppSettings.xlsx")
print(f"Looking for Excel file at: {excel_path}")
df = pd.read_excel(excel_path, sheet_name='myinfo', dtype=str)

############################################################################################################
### Sets up instructions for intro frame and console
instructions = ("*Instructions*")

###########################################################################################################
#### Static Functions
# Sets up the table inside the frame the right way -> takes Excel and changes it to a readable format in GUI.
def display_table(df, frame):
    start_col = 11
    span = 4

    for col_index, header in enumerate(df.columns):
        if start_col <= col_index < start_col + span:
            if col_index == start_col:
                merged_header = ctk.CTkLabel(
                    frame,
                    text="Doc Numbers",
                    font=("Verdana", 11, "bold"),
                    text_color="#E0BDEC",
                    justify="center",
                    fg_color="transparent"
                )
                merged_header.grid(row=0, column=start_col, columnspan=span, sticky="nsew", padx=10)
            continue
        else:
            header_label = ctk.CTkLabel(
                frame,
                text=header,
                font=("Verdana", 11, "bold"),
                text_color="#E0BDEC",
                justify="center",
                fg_color="transparent",
            )
            header_label.grid(row=0, column=col_index, sticky="nsew")

        frame.grid_columnconfigure(col_index, weight=1)

    for row_index, (_, row) in enumerate(df.iterrows(), start=1):
        for col_index, cell in enumerate(row):
            text = str(cell)
            if text in ("Y", "Yes"):
                display_txt = "✔"
                text_color = "green"
            elif text in ("N", "No", "nan"):
                display_txt = "✗"
                text_color = "red"
            else:
                display_txt = text
                text_color = "#F7F5ED"

            cell_label = ctk.CTkLabel(
                frame,
                text=display_txt,
                font=("Verdana", 11, "bold"),
                text_color=text_color,
                fg_color="transparent",
            )
            cell_label.grid(row=row_index, column=col_index, padx=0, pady=0, sticky="nsew")

# Grabs image from screen to detect if MReport is available for use.
def grab_region(region):
    with mss.mss() as sct:
        left, top, width, height = region
        monitor = {"left": left, "top": top, "width": width, "height": height}
        screenshot = sct.grab(monitor)
        return Image.frombytes("RGB", screenshot.size, screenshot.rgb)

###########################################################################################################

#### Our Class to build the GUI, and all major functions inside the GUI.
def to_gray(img): # Built with Copilot
    arr = np.asarray(img).astype("float32")
    return 0.2989 * arr[:, :, 0] + 0.5870 * arr[:, :, 1] + 0.1140 * arr[:, :, 2]

def region_changed_from_base(base_image, region, threshold=5, interval=5, timeout=50):
    start_time = time.time()
    g1 = to_gray(base_image)

    while True:
        time.sleep(interval)

        if time.time() - start_time > timeout:
            print("No data screen? Continuing.")
            return False

        img2 = grab_region(region)
        g2 = to_gray(img2)

        if g1.shape != g2.shape:
            return True

        mse = np.mean((g1 - g2) ** 2)

        if mse > threshold:
            print(f"Report done with estimation {mse:.4f}")
            return True


class ArchiveApp(ctk.CTk):
    hotkey_registered = False
    hotkey_lock = threading.Lock()
    abort_flag = threading.Event()

    ## Creates all frames, variables, and sets up intro frame.
    def __init__(self):
        super().__init__()

        self.start_up_lock = threading.Lock()
        self.title("Archive Automation Application")
        self.geometry("1300x750")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.all_stores = {}
        self.selected_stores = []

        self.intro_frame = ctk.CTkFrame(self)
        self.intro_frame.grid(row=0, column=0, sticky="nsew")
        self.setup_intro_frame()

        self.setup_frame = ctk.CTkFrame(self)
        self.setup_frame.grid(row=0, column=0, sticky="nsew")
        self.store_ids = set(map(str, self.selected_stores))
        self.setup_setup_frame()

        self.checker_frame = ctk.CTkFrame(self)
        self.checker_frame.grid(row=0, column=0, sticky="nsew")
        store_ids = set(map(str, self.selected_stores))
        self.df_filtered = df[df[df.columns[0]].isin(store_ids)]
        self.setup_checker_frame(self.df_filtered)

        self.fail_frame = ctk.CTkFrame(self)
        self.fail_frame.grid(row=0, column=0, sticky="nsew")
        self.setup_fail_frame()

        self.verify_frame = ctk.CTkFrame(self)
        self.verify_frame.grid(row=0, column=0, sticky="nsew")
        self.scroll_var = tk.StringVar(value="Page 1 of 2 >>")
        self.cell_widgets = {}
        self.widget_dict = {}
        self.checkbox_vars = {}
        self.setup_verify_frame(self.df_filtered)
        self.master_lock = threading.Lock()
        self.master_id = 0
        self.active_master_id = 0

        self.console_frame = ctk.CTkFrame(self)
        self.console_frame.grid(row=0, column=0, sticky="nsew")
        self.final_cell_labels = {}

        show_frame(self.intro_frame)

    # Creates intro frame the right way, with button.
    def setup_intro_frame(self):
        self.intro_frame.grid_rowconfigure(0, weight=0)
        self.intro_frame.grid_rowconfigure(2, weight=1)
        self.intro_frame.grid_rowconfigure(3, weight=1)
        self.intro_frame.grid_rowconfigure(4, weight=0)
        self.intro_frame.grid_rowconfigure(5, weight=0)
        self.intro_frame.grid_columnconfigure(0, weight=1)
        self.red_indicator = False

        welcome_label = ctk.CTkLabel(
            self.intro_frame,
            text="Welcome to the Hall MileOne Archival Automation App!",
            font=("Verdana", 26, "bold"),
            justify="center",
            wraplength=900,
            text_color="#F7F5ED"
        )
        welcome_label.grid(row=0, column=0, padx=40, pady=(100, 10), sticky="nsew")

        container = ctk.CTkFrame(self.intro_frame, fg_color="transparent")
        container.grid(row=1, column=0, rowspan=3, sticky="ew", padx=(0, 0), pady=10)

        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=0)

        intro_text = f"{instructions}"
        intro_label = ctk.CTkLabel(
            container,
            text=intro_text,
            font=("Verdana", 22),
            justify="center",
            wraplength=900,
            text_color="#F7F5ED"
        )
        intro_label.grid(row=0, column=0, columnspan=2, sticky="ew")

        highlight_btn = ctk.CTkButton(
            container,
            text="Where is my Second Monitor?",
            font=("Verdana", 14, "bold"),
            height=40,
            hover=True,
            fg_color="#4A3C3F",
            hover_color="#705E6B",
            text_color="#F7F5ED",
            command=self.highlight_second_monitor
        )
        highlight_btn.grid(row=0, column=1, sticky="ne", pady=(20, 30))

        ok_button = ctk.CTkButton(
            self.intro_frame,
            text="OK",
            font=("Verdana", 26, "bold"),
            height=60,
            width=200,
            hover=True,
            fg_color="#4A3C3F",
            hover_color="#705E6B",
            text_color="#F7F5ED",
            command=lambda: show_frame(self.setup_frame)
        )
        ok_button.grid(row=4, column=0, pady=(20, 40), sticky="s")

        trademark = ctk.CTkLabel(
            self.intro_frame,
            text="Made by Carter Grohe",
            font=("Verdana", 11.8, "bold"),
            text_color="#F7F5ED"
        )
        trademark.grid(row=5, column=0, sticky="se", padx=10, pady=10)

    # This function accompanies the intro frame. Makes second monitor red.
    def highlight_second_monitor(self):
        if getattr(self, "red_indicator", False):
            return
        monitors = get_monitors()

        if len(monitors) < 2:
            print("Second monitor not detected.")
            return

        second_monitor = monitors[1]
        x = second_monitor.x
        y = second_monitor.y
        width = second_monitor.width
        height = second_monitor.height

        red_window = tk.Toplevel(self)
        red_window.overrideredirect(True)
        red_window.geometry(f"{width}x{height}+{x}+{y}")
        red_window.configure(bg="red")
        red_window.attributes("-topmost", True)
        red_window.attributes("-alpha", 0.7)

        label = ctk.CTkLabel(
            red_window,
            text="Here it is! :D\n"
                 "Put *XXX* fullscreen here.\n\n\n"
                 "Click here to close the red frame.",
            font=("Verdana", 36, "bold"),
            text_color="white",
            bg_color="transparent"
        )
        label.place(relx=0.5, rely=0.5, anchor="center")
        self.red_indicator = True

        def close_window(event=None):
            red_window.destroy()
            self.red_indicator = False

        red_window.bind("<Button-1>", close_window)

    # Sets up Store Frame. Checkboxes inside.
    def setup_setup_frame(self):
        self.setup_frame.grid_rowconfigure((0, 1, 3, 4, 5, 6, 8, 10), weight=0)
        self.setup_frame.grid_rowconfigure((2, 7, 9, 11), weight=1)
        self.setup_frame.grid_columnconfigure(0, weight=1)

        self.title_text = "Welcome to the Hall MileOne Archival App!"
        self.subtitle_text = tk.StringVar(value="Selected Stores = None")
        self.all_stores_var = ctk.BooleanVar()
        self.all_stores = {}
        self.df_final = None
        self.df_simple = None
        self.df_filtered = None

        if hasattr(self, "final_frame"):
            for widget in self.final_frame.winfo_children():
                widget.destroy()

        checkbox_frame = ctk.CTkFrame(self.setup_frame, fg_color="transparent")
        checkbox_frame.grid(row=2, column=0, padx=30, pady=25, sticky="nsew")
        for i in range(5):
            checkbox_frame.grid_columnconfigure(i, weight=1)

        for idx in range(1, 25, 1):
            var = tk.BooleanVar()
            self.all_stores[idx] = var

            cb = ctk.CTkCheckBox(
                checkbox_frame,
                text=f"Store {idx}",
                variable=var,
                font=("Verdana", 15, "bold"),
                text_color="#E0BDEC",
                command=self.update_subtitle_text
            )
            cb.grid(row=(idx - 1) // 5, column=(idx - 1) % 5, sticky="nsew", padx=12, pady=10)

        self.toggle_all_stores = self._toggle_all_stores
        select_all = ctk.CTkCheckBox(
            self.setup_frame,
            text="Select All Stores",
            variable=self.all_stores_var,
            font=("Verdana", 16, "bold", "underline"),
            command=self.toggle_all_stores,
            text_color="#E0BDEC"
        )
        select_all.grid(row=8, column=0)

        title_label = ctk.CTkLabel(
            self.setup_frame, text=self.title_text,
            font=("Verdana", 20, "bold"), justify="center", wraplength=900,
            text_color="#F7F5ED"
        )
        title_label.grid(row=0, column=0, sticky="n", pady=(15, 5))

        subtitle_label = ctk.CTkLabel(
            self.setup_frame, textvariable=self.subtitle_text,
            font=("Verdana", 16, "bold", "italic"), justify="center", wraplength=900,
            text_color="#F7F5ED"
        )
        subtitle_label.grid(row=1, column=0, sticky="n", padx=10, pady=(15, 5))

        submit_btn = ctk.CTkButton(
            self.setup_frame, text="Done", command=self.show_selected,
            fg_color="#4A3C3F", hover_color="#705E6B",
            text_color="#F7F5ED", font=("Verdana", 18, "bold"), height=60, width=200
        )
        submit_btn.grid(row=10, column=0, pady=(0,40))

    # This function makes the subtitle 'Selected Stores' update dynamically.
    def update_subtitle_text(self):
        selected = [str(idx) for idx, var in self.all_stores.items() if var.get()]
        if len(selected) == len(self.all_stores):
            self.all_stores_var.set(True)
            display_text = " Selected Stores: All "
        elif selected:
            self.all_stores_var.set(False)
            display_text = "Selected Stores: Stores " + ", ".join(selected)
        else:
            self.all_stores_var.set(False)
            display_text = "Selected Stores: None"

        self.subtitle_text.set(display_text)

    # Function for Select All Stores button.
    def _toggle_all_stores(self):
        select_all = self.all_stores_var.get()
        for var in self.all_stores.values():
            var.set(select_all)
        self.update_subtitle_text()

    # Initializes the viewable excelsheet on the top of the next frame, as well as the right date/doc number string for MReport.
    def show_selected(self):
        self.selected_stores = [idx for idx, var in self.all_stores.items() if var.get()]
        if not self.selected_stores:
            show_frame(self.fail_frame)
            return
        store_ids = set(map(str, self.selected_stores))
        self.df_filtered = df[df[df.columns[0]].isin(store_ids)]

        self.setup_checker_frame(self.df_filtered)
        show_frame(self.checker_frame)

    # Sets up the next frame with the Excel sheet on it.
    def setup_checker_frame(self, df):
        if hasattr(self, "checker_frame"):
            for widget in self.checker_frame.winfo_children():
                widget.destroy()
        if hasattr(self, "final_frame"):
            for widget in self.final_frame.winfo_children():
                widget.destroy()

        self.checker_frame.grid_rowconfigure((0, 2), weight=1)
        self.checker_frame.grid_rowconfigure((1, 3, 4), weight=0)
        self.checker_frame.grid_columnconfigure(0, weight=1)

        table_container = ctk.CTkScrollableFrame(self.checker_frame, fg_color="transparent")
        table_container.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        display_table(df, table_container)

        legend_text = (
            "*NOTE*"
        )
        legend = ctk.CTkLabel(
            self.checker_frame,
            text=legend_text,
            font=("Verdana", 16, "bold"),
            text_color="#E0BDEC",
            justify="left",
            anchor="w"
        )
        legend.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        verifier = ctk.CTkLabel(
            self.checker_frame,
            text="Please make sure all of the above information is correct. If it isn't, click on 'Change Information' button. \n"
                 "Or, return to select different Stores. Then, click Start.",
            font=("Verdana", 14.5, "bold"),
            text_color="#E0BDEC",
            wraplength=900,
            justify="left",
            anchor="w"
        )
        verifier.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)

        button_frame = ctk.CTkFrame(self.checker_frame, fg_color="transparent")
        button_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=10)
        button_frame.grid_columnconfigure((0, 2), weight=0)
        button_frame.grid_columnconfigure(1, weight=1)

        start_button = ctk.CTkButton(
            button_frame, text="Start",
            fg_color="#4A3C3F", hover_color="#705E6B", text_color="#E0BDEC",
            height=60, width=130, font=("Verdana", 18, "bold", "underline"), command=self.start_up
        )
        start_button.grid(row=0, column=2, padx=10, pady=10)

        change_button = ctk.CTkButton(
            button_frame, text="Change Information",
            fg_color="#4A3C3F", hover_color="#705E6B", text_color="#E0BDEC",
            font=("Verdana", 14, "bold"), height=60,
            command=self.commandultra
        )
        change_button.grid(row=0, column=0, padx=(10, 10), pady=10)

        return_button = ctk.CTkButton(
            button_frame, text="Return to Stores List",
            fg_color="#4A3C3F", hover_color="#705E6B", text_color="#E0BDEC",
            font=("Verdana", 14, "bold"), height=60,
            command=lambda: show_frame(self.setup_frame)
        )
        return_button.grid(row=0, column=1, padx=10, pady=10)

    # Double function that ensures a smooth transition to 'Change Information' frame.
    def commandultra(self):
        self.setup_verify_frame(self.df_filtered)
        show_frame(self.verify_frame)

    # Backup redundancy for if no stores are selected.
    def setup_fail_frame(self):
        self.fail_frame.grid_rowconfigure(0, weight=1)  # top spacer
        self.fail_frame.grid_rowconfigure(1, weight=0)  # label
        self.fail_frame.grid_rowconfigure(2, weight=0)  # button
        self.fail_frame.grid_rowconfigure(3, weight=1)  # bottom spacer
        self.fail_frame.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(
            self.fail_frame,
            text="Sorry, you must select at least one store.",
            font=("Verdana", 18, "bold"),
            text_color="#E0BDEC",
            wraplength=400,
            justify="center"
        )
        label.grid(row=1, column=0, pady=(0, 20), sticky="n")

        back_button = ctk.CTkButton(
            self.fail_frame,
            text="Return",
            font=("Verdana", 14, "bold"),
            command=lambda: show_frame(self.setup_frame),
            width=150,
            height=40,
            corner_radius=8,
            text_color="#F7F5ED",
            fg_color="#363030",
            hover_color="#705E6B"
        )
        back_button.grid(row=2, column=0, sticky="n")

    # Setup for dual frame with changeable text.
    def setup_verify_frame(self, df):
        midpoint = len(df.columns) // 2
        self.df_left = df.iloc[:, :midpoint]
        self.df_right = df.iloc[:, midpoint:]

        for widget in self.verify_frame.winfo_children():
            widget.destroy()

        self.verify_frame.grid_rowconfigure(0, weight=1)
        self.verify_frame.grid_rowconfigure(1, weight=0)
        self.verify_frame.grid_columnconfigure(0, weight=1)

        self.current_side = "left"
        self.scroll_var.set("Page 1 of 2 >>")

        self.left_frame = ctk.CTkScrollableFrame(self.verify_frame, fg_color="transparent")
        self.left_frame.grid(row=0, column=0, sticky="nsew")
        self.display_table2(self.df_filtered, self.left_frame, columns=range(0, len(self.df_filtered.columns) // 2))

        self.right_frame = ctk.CTkScrollableFrame(self.verify_frame, fg_color="transparent")
        self.display_table2(self.df_filtered, self.right_frame,
                            columns=range(len(self.df_filtered.columns) // 2, len(self.df_filtered.columns)))
        self.right_frame.grid_forget()

        button_frame = ctk.CTkFrame(self.verify_frame, fg_color="transparent")
        button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)

        def toggle_table():
            if self.current_side == "left":
                self.left_frame.grid_forget()
                self.right_frame.grid(row=0, column=0, sticky="nsew")
                self.current_side = "right"
                self.scroll_var.set("<< Page 2 of 2")
            else:
                self.right_frame.grid_forget()
                self.left_frame.grid(row=0, column=0, sticky="nsew")
                self.current_side = "left"
                self.scroll_var.set("Page 1 of 2 >>")

        toggle_button = ctk.CTkButton(button_frame, textvariable=self.scroll_var, command=toggle_table, fg_color="#9395D3", font=("Verdana", 14, "bold"),
                                      height=40, hover_color="#7a7cc2")
        toggle_button.grid(row=0, column=1, padx=10)

        save_button = ctk.CTkButton(button_frame, text="Save Changes", command=self.save_changes, fg_color="#3A6351", font=("Verdana", 14, "bold"),
                                    height=40, hover_color="#2f5142")
        save_button.grid(row=0, column=2, sticky="e", padx=10)

        cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=self.cancel_changes, fg_color="#9B4444", font=("Verdana", 14, "bold"),
                                      height=40, hover_color="#823838")
        cancel_button.grid(row=0, column=0, sticky="w", padx=10)

    # Function for Save Changes Button.
    def save_changes(self):
        for (row_idx, col_idx), widget in self.widget_dict.items():
            if isinstance(widget, ctk.CTkEntry):
                val = widget.get()
                if val == '' or val.lower() in ['none', 'nan']:
                    val = '-'
                self.df_filtered.iat[row_idx, col_idx] = val
            elif isinstance(widget, ctk.CTkCheckBox):
                var = self.checkbox_vars.get((row_idx, col_idx))
                if var is not None:
                    self.df_filtered.iat[row_idx, col_idx] = "Yes" if var.get() else "No"

        self.df_final = self.df_filtered.copy()
        self.MReport_already_inserted = False
        for widget in self.checker_frame.winfo_children():
            widget.destroy()
        self.setup_checker_frame(self.df_final)
        show_frame(self.checker_frame)

    #Function for Cancel button.
    def cancel_changes(self):
        if not hasattr(self, 'df_final'):
            self.df_final = self.df_filtered.copy()
        self.df_filtered = self.df_final.copy()

        self.widget_dict.clear()
        self.checkbox_vars.clear()

        for widget in self.left_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        midpoint = len(self.df_filtered.columns) // 2
        self.display_table2(self.df_filtered, self.left_frame, columns=range(0, midpoint))
        self.display_table2(self.df_filtered, self.right_frame, columns=range(midpoint, len(self.df_filtered.columns)))

        if self.current_side == "left":
            self.right_frame.grid_forget()
            self.left_frame.grid(row=0, column=0, sticky="nsew")
        else:
            self.left_frame.grid_forget()
            self.right_frame.grid(row=0, column=0, sticky="nsew")
        show_frame(self.checker_frame)

    # The function which splits the dataset in the Verify frame correctly.
    def display_table2(self, df, frame, columns):
        if not hasattr(self, "widget_dict"):
            self.widget_dict = {}
        if not hasattr(self, "checkbox_vars"):
            self.checkbox_vars = {}

        keys_to_remove = []
        for (r, c) in self.widget_dict:
            if c in columns:
                keys_to_remove.append((r, c))
        for key in keys_to_remove:
            del self.widget_dict[key]

        for widget in frame.winfo_children():
            widget.destroy()

        for grid_col_idx, col_idx in enumerate(columns, start=1):
            header = df.columns[col_idx]
            header_label = ctk.CTkLabel(
                frame,
                text=header,
                font=("Verdana", 11, "bold"),
                text_color="#E0BDEC",
                justify="center",
                fg_color="transparent",
            )
            header_label.grid(row=0, column=grid_col_idx, sticky="nsew")
            frame.grid_columnconfigure(grid_col_idx, weight=1)

        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
            frame.grid_rowconfigure(row_idx, weight=1)
            for grid_col_idx, col_idx in enumerate(columns, start=1):
                cell = row.iloc[col_idx]
                text = str(cell).strip()
                text_lower = text.lower()

                if col_idx < 2:
                    widget = ctk.CTkLabel(
                        frame,
                        text=text,
                        font=("Verdana", 11),
                        text_color="#F7F5ED",
                        fg_color="transparent",
                    )
                elif col_idx == 2 or 11 <= col_idx <= 14:
                    widget = ctk.CTkEntry(
                        frame,
                        font=("Verdana", 11),
                        fg_color="#2B2B2B",
                        text_color="#F7F5ED",
                    )
                    widget.insert(0, text)
                elif 3 <= col_idx <= 10 or col_idx == 15:
                    var = tk.BooleanVar(value=text_lower in ("y", "yes", "true", "1"))
                    widget = ctk.CTkCheckBox(
                        frame,
                        variable=var,
                        fg_color="#4A3C3F",
                        text="",
                    )
                    self.checkbox_vars[(row_idx - 1, col_idx)] = var
                else:
                    widget = ctk.CTkLabel(
                        frame,
                        text=text,
                        font=("Verdana", 11),
                        text_color="#F7F5ED",
                        fg_color="transparent",
                    )

                widget.grid(row=row_idx, column=grid_col_idx, padx=1, pady=1, sticky="nsew")
                self.widget_dict[(row_idx - 1, col_idx)] = widget

    # Updates progress dynamically in top half of frame.
    def update_final_cell(self, index_label, col_name, status):
        try:
            row_position = self.df_final.index.get_loc(index_label) + 1
            col_index = self.df_final.columns.get_loc(col_name)
        except KeyError:
            print(f"Column '{col_name}' not found in df_final")
            return
        emoji_map = {
            "done": "🟩",
            "in_progress": "🟡",
            "not_started": "⚪",
            "failed": "🟥"
        }
        color_map = {
            "done": "green",
            "in_progress": "yellow",
            "not_started": "white",
            "failed": "red"
        }
        new_text = emoji_map.get(status, "")
        new_color = color_map.get(status, "#F7F5ED")

        # This function actually updates the method, above is the setup.
        def _update():
            label = self.final_cell_labels.get((row_position, col_index))
            if label:
                label.configure(text=new_text, text_color=new_color)
            else:
                print(f"No label found at row {row_position}, column '{col_name}'")

        self.final_frame.after(0, _update)

    # Replaces 'Doc Numbers' columns with True/False MReport in the final frame. Also stores Doc Codes for Use.
    def replace_columns_with_MReport(self, df):
        if hasattr(self, 'MReport_already_inserted') and self.MReport_already_inserted:
            return df
        cols_to_check = df.columns[11:15]
        MReport_values = {}

        def check_MReport(row):  # This will do the saving of Doc Codes and figure out which stores ot mark Yes/No.
            row_id = row.name
            values = row[cols_to_check].tolist()
            MReport_values[row_id] = values
            for val in values:
                if pd.notna(val) and val not in ['-', '']:
                    return "Yes"
            return "No"

        MReport_series = df.apply(check_MReport, axis=1)
        df = df.drop(cols_to_check, axis=1)
        df.insert(11, 'MReport', MReport_series)
        self.MReport_values_by_index = MReport_values
        self.MReport_already_inserted = True
        return df

    # Sets up final frame inside console frame inside this. Places in dynamic table and console, and buttons.
    def start_up(self):
        with self.start_up_lock:
            if hasattr(self, "cons_button_frame"):
                self.cons_button_frame.grid_forget()
            if not hasattr(self, 'df_final') or self.df_final is None:
                for (row_idx, col_idx), widget in self.widget_dict.items():
                    if isinstance(widget, ctk.CTkEntry):
                        val = widget.get()
                        self.df_filtered.iat[row_idx, col_idx] = val
                    elif isinstance(widget, ctk.CTkCheckBox):
                        var = self.checkbox_vars.get((row_idx, col_idx))
                        if var is not None:
                            self.df_filtered.iat[row_idx, col_idx] = "Yes" if var.get() else "No"

                self.df_final = self.df_filtered.copy()
                self.MReport_already_inserted = False
            if not hasattr(self, 'df_simple') or self.df_simple is None:
                self.df_simple = False

            if not getattr(self, 'df_simple', False):
                self.df_final = self.replace_columns_with_MReport(self.df_final)

            if self.all_stores_var.get():
                msg = "You are printing all stores."
            else:
                msg = f"You are printing Stores {self.selected_stores}."

            self.instructions = ("*Instructions*")
            show_frame(self.console_frame)

            self.console_frame.grid_rowconfigure((0, 3), weight=1)
            self.console_frame.grid_rowconfigure((1, 2, 4), weight=0)
            self.console_frame.grid_columnconfigure(0, weight=1)

            self.final_frame = ctk.CTkScrollableFrame(self.console_frame)
            self.final_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))
            self.final_frame.grid_columnconfigure(0, weight=1)

            if not hasattr(self, 'df_final'):
                print("ERROR: df_final not defined yet!")
                return

            for widget in self.final_frame.winfo_children():
                widget.destroy()

            for col_index, header in enumerate(self.df_final.columns):
                header_label = ctk.CTkLabel(
                    self.final_frame,
                    text=header,
                    font=("Verdana", 11, "bold"),
                    text_color="#E0BDEC",
                    fg_color="transparent"
                )
                header_label.grid(row=0, column=col_index, sticky="nsew")
                self.final_frame.grid_columnconfigure(col_index, weight=1)

            for row_index, (_, row) in enumerate(self.df_final.iterrows(), start=1):
                for col_index, cell in enumerate(row):
                    text = str(cell).strip().lower()

                    if text in ("✔", "yes", "y", "true"):
                        cons_display = "⚫"
                        text_color = "white"
                    elif text in ("✗", "no", "n", "false", "nan"):
                        cons_display = "❌"
                        text_color = "red"
                    else:
                        cons_display = str(cell)
                        text_color = "#F7F5ED"

                    cell_label = ctk.CTkLabel(
                        self.final_frame,
                        text=cons_display,
                        font=("Verdana", 11, "bold"),
                        text_color=text_color,
                        fg_color="transparent"
                    )
                    cell_label.grid(row=row_index, column=col_index, padx=0, pady=0, sticky="nsew")
                    self.final_cell_labels[(row_index, col_index)] = cell_label

            legend_frame = ctk.CTkFrame(self.console_frame)
            legend_frame.grid(row=1, column=0, pady=(0, 10))

            labels = [
                ("🟩 Done", "green"),
                ("🟡 In Progress", "yellow"),
                ("⚫ Not Started", "white"),
                ("❌ Not Running", "red"),
            ]

            for i, (text, color) in enumerate(labels):
                lbl = ctk.CTkLabel(legend_frame, text=text, font=("Verdana", 13, "bold"), text_color=color)
                lbl.grid(row=0, column=i, padx=10)

            self.cons_button_frame = ctk.CTkFrame(self.console_frame, fg_color="transparent")
            self.cons_button_frame.grid(row=4, column=0, padx=10, pady=10)
            self.cons_button_frame.grid_rowconfigure(0, weight=1)
            self.cons_button_frame.grid_columnconfigure((0), weight=1)

            self.stop_button = ctk.CTkButton(
                self.cons_button_frame,
                text="Exit",
                command=sys.exit,
                height=50,
                width=200,
                hover=True,
                fg_color="#4A3C3F",
                hover_color="#705E6B",
                text_color="#F7F5ED",
                font=("Verdana", 16, "bold", "underline"),
            )
            self.stop_button.grid(row=0, column=0, pady=10, padx=20)

            self.edit_button = ctk.CTkButton(
                self.cons_button_frame,
                text="Back to Edit",
                command=lambda: show_frame(self.checker_frame),
                height=50,
                width=200,
                hover=True,
                fg_color="#4A3C3F",
                hover_color="#705E6B",
                text_color="#F7F5ED",
                font=("Verdana", 14, "bold"),
            )
            self.edit_button.grid(row=0, column=1, pady=10, padx=(0, 20))

            if not hasattr(self, "console"):
                self.console = ctk.CTkTextbox(
                    self.console_frame,
                    width=800,
                    font=("Verdana", 14)
                )
                self.console.grid(row=3, column=0, sticky="nsew", padx=20, pady=(5, 10))

                sys.stdout = TextRedirector(self.console)
                sys.stderr = TextRedirector(self.console)

            self.console.delete('1.0', tk.END)
            print(self.instructions)
            sleep(2)
            print("")
            print("Please press *ENTER* when you are ready to begin.")
            ArchiveApp.abort_flag.clear()
            with self.master_lock:
                self.master_id += 1
                my_id = self.master_id
                self.active_master_id = my_id
            if ArchiveApp.hotkey_registered and ArchiveApp.abort_hotkey_id is not None:
                keyboard.remove_hotkey(ArchiveApp.abort_hotkey_id)
                ArchiveApp.hotkey_registered = False
                ArchiveApp.abort_hotkey_id = None

            threading.Thread(target=lambda: self.master_code(my_id), daemon=True).start()

    # This is the manual abort method.
    def abort(self):
        if ArchiveApp.abort_flag.is_set():
            print("\nProcess aborted.\n")
            self.edit_button.grid(row=0, column=1, pady=10, padx=(0, 20))
            raise SystemExit

    # Checks for when to continue after Expense Report
    def check_region_color_change(self, region):
        if not hasattr(self, "img1"):
            print("Initial image not found — aborting check.")
            return False

        current = grab_region(region).convert("RGB")
        diff = ImageChops.difference(self.img1, current)

        return diff.getbbox() is None

    # Checks continuously for Report to be done.
    def wait_for_region_change(self, region, timeout=30):
        start_time = time.time()

        while time.time() - start_time < timeout:
            if self.check_region_color_change(region):
                return True
            else:
                time.sleep(1)

        return False

    def wait_for_region_change_MReport(self, region, timeout=40):
        start_time = time.time()

        while time.time() - start_time < timeout:
            if self.check_region_color_change(region):
                return True
            else:
                time.sleep(1)

        print("Timeout reached without detecting matching screen.")
        return False

    # Checks for when report is done.
    def wait_for_return(self, region, timeout=60):
        start_time = time.time()

        while time.time() - start_time < timeout:
            if self.check_region_color_change(region):
                print(f"Return screen detected - XXX should be done.")
                return True
            else:
                time.sleep(1)

        print("Timeout reached without detecting matching screen.")
        return False

    # Actual code for *XXX* to use.
    def master_code(self, my_id):
            ### *Deleted for privacy*

###########################################################################################################
## This makes everything run in the right order and correctly when needed. Final touch.
if __name__ == "__main__":
    app = ArchiveApp()
    app.mainloop()