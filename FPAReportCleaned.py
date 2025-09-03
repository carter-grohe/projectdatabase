### Carter Grohe UVA Statistics/Data Science '27
### This code compiled and built daily 'FPA' Reports for my internship at Hall MileOne Automotive
### using *XXX*, Python, and PyInstaller. *XXX* required PyAutoGUI, so my code was
### hampered by old software.

import subprocess
import sys

## This is my go-to function for installing Python packages upon running the code.
def ensure_packages_installed(packages):
    for package in required_packages:
        try:
            __import__(package)
            print(f"{package} is already installed.")
        except ImportError:
            print(f"{package} not found. Installing...")
            subprocess.run([sys.executable, "-m", "pip", "install", package], check=True)

required_packages = ["tkinter", "pyautogui", "PIL", "mss", "customtkinter"]

import customtkinter as ctk
import tkinter as tk
import threading
import pyautogui
from time import sleep
from datetime import datetime
from tkinter import filedialog
import mss
from PIL import ImageChops, Image
from pathlib import Path
import ctypes

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def turn_off_capslock():
    if ctypes.WinDLL("User32.dll").GetKeyState(0x14) == 1: # IF CAPSLOCK is on: (CAPSLOCK = 0x14)
        pyautogui.press('capslock') #Turn off.
        sleep(0.4)

turn_off_capslock()

## Go-to function for capturing images. I prefer mss to pyautogui screenshots.
def grab_region(region):
    with mss.mss() as sct:
        left, top, width, height = region
        monitor = {"left": left, "top": top, "width": width, "height": height}
        screenshot = sct.grab(monitor)
        return Image.frombytes("RGB", screenshot.size, screenshot.rgb)

def check_region_color_change(region):
    current_image = grab_region(region)
    if not hasattr(check_region_color_change, "initial_images"):
        check_region_color_change.initial_images = {}

    if region not in check_region_color_change.initial_images:
        check_region_color_change.initial_images[region] = current_image
        return False

    first_image = check_region_color_change.initial_images[region]
    diff = ImageChops.difference(first_image, current_image)
    return diff.getbbox() is not None

class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.after(0, lambda: self._write(text))

    def _write(self, text):
        self.widget.insert("end", text)
        self.widget.see("end")

    def flush(self):
        pass

store = list(range(1, x, 1))
stores_to_skip = [xxx]
current_date = datetime.now().strftime("%y-%m-%d")

## A bit chunky here, but gets the job done to initialize my class and variables.
class ReportGeneratorApp:
    def __init__(self, root):

        self.root = root
        self.root.title("FPA Report Generator")
        self.root.geometry("1000x700")

        self.intro_frame = ctk.CTkFrame(root)
        self.main_frame = ctk.CTkFrame(root)
        self.report_frame = ctk.CTkFrame(root)
        self.fail_frame = ctk.CTkFrame(root)
        self.test_frame = ctk.CTkFrame(root)
        self.stores_to_skip = stores_to_skip
        self.store_count = len(store)
        self.columns = 5

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        for frame in (self.intro_frame, self.main_frame, self.report_frame, self.fail_frame, self.test_frame):
            frame.grid(row=0, column=0, sticky="nsew")

        self.selected_stores = []
        self.store_vars = {}
        self.process = None
        self.all_stores_var = ctk.BooleanVar()
        self.finished_stores = 0

        self.checked_label_var = tk.StringVar(value="Checked Stores: None")
        self.file_path_var = ctk.StringVar(value="//*Network File Path*")

        self.setup_intro_frame()
        self.setup_main_frame()
        self.setup_report_frame()
        self.setup_fail_frame()
        self.setup_test_frame()

        self.show_frame(self.intro_frame)

    def show_frame(self, frame):
        frame.tkraise()

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.file_path_var.set(directory)

    def setup_intro_frame(self):
        self.intro_frame.grid_rowconfigure(0, weight=1)
        self.intro_frame.grid_columnconfigure(0, weight=1)
        intro_text = (
            "*App Instructions:*"
        )
        label = ctk.CTkLabel(
            self.intro_frame,
            text=intro_text,
            font=("Verdana", 24),
            anchor="center",
            justify="center",
            wraplength=800
        )
        label.grid(row=0, column=0, sticky="nsew", padx=40, pady=20)

        button = ctk.CTkButton(
            self.intro_frame,
            text="OK",
            height=60,
            width=200,
            fg_color="#2553A2",
            hover_color="#2571A3",
            font=("Verdana", 26, "bold"),
            command=lambda: self.show_frame(self.main_frame)
        )
        button.grid(row=1, column=0, pady=(0, 60))

        trademark = ctk.CTkLabel(
            self.intro_frame,
            text="Made by Carter Grohe",
            font=("Verdana", 11),
            text_color="lightgray"
        )
        trademark.grid(row=2, column=0, sticky="se", padx=10, pady=10)

    ## below functions created a guide for older and less computer-savvy people to work with.
    ## Highlighted the box where the function would detect for change.
    def draw_box(self, x1, y1, x2, y2, duration=500000):
        self.top = ctk.CTkToplevel(self.root)
        self.top.attributes("-topmost", True)
        self.top.attributes("-transparentcolor", "white")
        self.top.overrideredirect(True)
        self.top.geometry(f"{x2 - x1}x{y2 - y1}+{x1}+{y1}")
        canvas = tk.Canvas(self.top, width=x2 - x1, height=y2 - y1, bg="white", highlightthickness=0)
        canvas.pack()
        canvas.create_rectangle(1, 1, x2 - x1 - 1, y2 - y1 - 1, outline="red", width=3)
        self.top.after(duration, self.top.destroy)

    def commando(self):
        self.show_frame(self.test_frame)
        region = (3300, 760, 3400, 957)
        self.draw_box(region[0], region[1], region[2], region[3], duration=500000)

    def back_to_main(self):
        if hasattr(self, 'top') and self.top.winfo_exists():
            self.top.destroy()
        self.show_frame(self.main_frame)

    def setup_main_frame(self):
        self.main_frame.grid_rowconfigure((0 ,1, 2, 4, 5, 7, 8, 9, 10, 11, 13), weight=0)
        self.main_frame.grid_rowconfigure((3, 6, 12, 14, 15), weight=1)
        title = ctk.CTkLabel(
            self.main_frame,
            text="Welcome to FPA Report Generator 3000",
            font=("Verdana", 28, "bold")
        )
        title.grid(row=0, column=0, columnspan=5, pady=(15, 5), sticky="ew")

        instruction = ctk.CTkLabel(
            self.main_frame,
            text="Select the stores you want to generate reports for.",
            font=("Verdana", 22)
        )
        instruction.grid(row=1, column=0, columnspan=5, pady=(0, 10), sticky="ew")

        checked_label = ctk.CTkLabel(
            self.main_frame,
            textvariable=self.checked_label_var,
            font=("Verdana", 17)
        )
        checked_label.grid(row=2, column=0, columnspan=5, pady=(0, 10), sticky="ew")

        label = ctk.CTkLabel(
            self.main_frame,
            text="Click 'Browse' to choose the folder to save reports.",
            font=("Verdana", 16, "bold")
        )
        label.grid(row=4, column=0, columnspan=5, pady=(0, 10), padx=10, sticky="w")

        path_entry = ctk.CTkEntry(
            self.main_frame,
            textvariable=self.file_path_var,
            width=500,
            height=35,
            font=("Verdana", 13)
        )
        path_entry.grid(row=5, column=0, columnspan=3, padx=(10, 0), pady=(0, 15), sticky="ew")

        browse_button = ctk.CTkButton(
            self.main_frame,
            text="Browse",
            fg_color="#2553A2",
            hover_color="#2571A3",
            command=self.browse_directory,
            height=35
        )
        browse_button.grid(row=5, column=3, padx=(5, 10), pady=(0, 15), sticky="w")

        for i in range(5):
            self.main_frame.grid_columnconfigure(i, weight=1)

        columns = 5
        for i in range(1, 26):
            var = tk.BooleanVar()
            store_num = str(i)
            self.store_vars[store_num] = var

            row = (i - 1) // columns + 7
            col = (i - 1) % columns
            if i == 25:
                # Make checkbox 25 into Select All, since we had 24 stores.
                cb = ctk.CTkCheckBox(
                    self.main_frame,
                    text="Select All Stores",
                    variable=self.all_stores_var,
                    font=("Verdana", 14, "bold"),
                    command=self.toggle_all_stores
                )
            else:
                cb = ctk.CTkCheckBox(
                    self.main_frame,
                    text=f"Store {i}",
                    variable=var,
                    font=("Verdana", 12, "bold"),
                    command=self.update_checked_stores,
                    text_color="#E14849" if int(store_num) in self.stores_to_skip else None
                )
            cb.grid(row=row, column=col, padx=(15, 0), pady=7, sticky="ew")

        # Action buttons
        test_button = ctk.CTkButton(
            self.main_frame,
            text="Test Range",
            font=("Verdana", 14, "bold"),
            fg_color="#2553A2",
            hover_color="#2571A3",
            command=self.commando,
            width=150,
            height=45,
            anchor="center"
        )
        test_button.grid(row=13, column=1, pady=(30, 20))

        generate_button = ctk.CTkButton(
            self.main_frame,
            text="Generate Reports",
            font=("Verdana", 14, "bold"),
            fg_color="#2553A2",
            hover_color="#2571A3",
            command=self.generate_report,
            width=180,
            height=45,
            anchor="center"
        )
        generate_button.grid(row=13, column=3, pady=(30, 20))

    def select_all_stores(self):
        for var in self.store_vars.values():
            var.set(True)
        self.update_checked_stores()

    def toggle_all_stores(self):
        select_all = self.all_stores_var.get()
        for num, var in self.store_vars.items():
            if num != str(self.store_count + 1) and int(num) not in self.stores_to_skip:
                var.set(select_all)
        self.update_checked_stores()

    def update_checked_stores(self):
        self.selected_stores = [store for store, var in self.store_vars.items() if
                                var.get() and int(store) not in self.stores_to_skip]
        all_selected = all(var.get() for num, var in self.store_vars.items() if
                           num != str(self.store_count + 1) and int(num) not in self.stores_to_skip)
        self.all_stores_var.set(all_selected)
        if self.selected_stores and not self.all_stores_var.get():
            display_text = "Selected Stores: Stores " + ", ".join(self.selected_stores)
        elif self.selected_stores and self.all_stores_var.get():
            display_text = "Selected Stores: All Active Stores"
        else:
            display_text = "Selected Stores: None"
        self.checked_label_var.set(display_text)

    def setup_report_frame(self):
        self.report_frame.grid_rowconfigure((0, 2), weight=1)
        self.report_frame.grid_rowconfigure((1, 3, 4), weight=0)
        self.report_frame.grid_columnconfigure(0, weight=1)
        self.report_frame.grid_columnconfigure(1, weight=0)

        self.console = ctk.CTkTextbox(
            self.report_frame,
            width=800,
            height=450,
            font=("Verdana", 14),
            corner_radius=10,
            wrap="word"
        )
        self.console.grid(row=0, column=0, padx=(20, 10), pady=20, sticky="nsew")

        # Scrollbar for Console
        scrollbar = ctk.CTkScrollbar(self.report_frame, command=self.console.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=20, padx=(0, 20))
        self.console.configure(yscrollcommand=scrollbar.set)

        self.progress_bar = ctk.CTkProgressBar(self.report_frame, orientation="horizontal", progress_color="#C788D9", fg_color="#2c2c2c",
                                               width=600, height=24, corner_radius=8)
        self.progress_bar.grid(row=2, column=0, columnspan=2, padx=(30,30), pady=20, sticky="ew")
        self.progress_bar.set(0)

        self.progress_var = ctk.StringVar(value=f"0% -- {self.finished_stores} of {len(self.selected_stores)} stores completed")
        self.progress_label = ctk.CTkLabel(self.report_frame, textvariable=self.progress_var, font=("Verdana", 14), anchor="center")
        self.progress_label.grid(row=3, column=0, columnspan=2, pady=(0, 20), sticky="ew")

        self.button_frame = ctk.CTkFrame(self.report_frame)
        self.button_frame.grid(row=4, column=0, columnspan=2, padx=20, pady=(0, 20))
        self.button_frame.grid_columnconfigure(0, weight=1)
        self.button_frame.grid_columnconfigure(1, weight=1)

        self.stop_button = ctk.CTkButton(
            self.button_frame,
            text="Stop",
            height=45,
            width=140,
            fg_color="#e74c3c",
            hover_color="#c0392b",
            font=("Verdana", 14),
            command=self.stop_process,
        )
        self.stop_button.grid(row=0, column=0, sticky="ew")

        self.back_to_main_button = ctk.CTkButton(
            self.button_frame,
            text="Back to Main Menu",
            height=45,
            width=180,
            fg_color="#3498db",
            hover_color="#2980b9",
            font=("Verdana", 14),
            command=lambda: self.show_frame(self.main_frame)
        )
        self.back_to_main_button.grid(row=0, column=1)
        self.back_to_main_button.grid_remove()

        sys.stdout = TextRedirector(self.console)

    def toggle_back_button(self, show: bool):
        if show:
            self.back_to_main_button.grid()
            self.stop_button.grid_configure(column=0, columnspan=1)
            self.back_to_main_button.grid_configure(column=1)
        else:
            self.back_to_main_button.grid_remove()
            self.stop_button.grid_configure(column=0, columnspan=2)

    def update_progress(self, value: float):
        value = max(0.0, min(1.0, value))
        self.progress_bar.set(value)
        percent = int(value * 100)
        self.progress_var.set(f"{percent}% -- {self.finished_stores} of {len(self.selected_stores)} stores completed")

    def setup_fail_frame(self):
        label = ctk.CTkLabel(self.fail_frame, text="Sorry, you must select at least one store.", font=("Verdana", 16))
        label.pack(pady=100)

        back_button = ctk.CTkButton(self.fail_frame, text="Return", fg_color="#2553A2",
            hover_color="#2571A3", command=lambda: self.show_frame(self.main_frame))
        back_button.pack(pady=20)

    def setup_test_frame(self):
        label = ctk.CTkLabel(self.test_frame, text="Test your ranges here.\nIf the red box does not contain above the bottom yellow bar --\nthis will not run.", font=("Verdana", 16))
        label.pack(pady=100)

        back_button = ctk.CTkButton(self.test_frame, text="Back to Main Menu", fg_color="#2553A2",
            hover_color="#2571A3", command=self.back_to_main)
        back_button.pack(pady=20)

    def generate_report(self):
        turn_off_capslock()
        self.selected_stores = [int(store) for store, var in self.store_vars.items() if var.get()]
        self.back_to_main_button.grid_forget()
        if not self.selected_stores:
            self.show_frame(self.fail_frame)
            return
        self.show_frame(self.report_frame)
        self.console.delete('1.0', ctk.END)
        sys.stdout = TextRedirector(self.console)
        sys.stderr = TextRedirector(self.console)
        path = self.file_path_var.get()
        threading.Thread(target=self.run_mastercode, args=(self.selected_stores, path), daemon=True).start()

    def run_mastercode(self, store_list, save_path):
        try:
            mastercodeFPAreportsSecondMonitor(self, store_list, save_path)
        except Exception as e:
            print(f"Error during report generation: {e}")
        finally:
            print("Report generation finished.")
            self.back_to_main_button.grid()
            self.stop_button.grid_forget()
            sys.stdout = sys.__stdout__

    def stop_process(self):
        sys.exit()

def mastercodeFPAreportsSecondMonitor(self, specificStores, save_path):
    self.finished_stores = 0
    try:
        if len(specificStores) != 0:
            correctStores = sorted(list(set(specificStores) - set(stores_to_skip)))
            storelist = correctStores
        else:
            storelist = [num for num in store if num not in stores_to_skip]
        print("------------------------------------")
        print("")
        print(f"Running stores: " + str(storelist))
        print("")
        print("*Instructions*")
        print("------------------------------------")
        sleep(4)
        click(2353, 294)
        pyautogui.hotkey('f3')
        pyautogui.press('enter')
        pyautogui.typewrite('*code*'.upper())
        pyautogui.press('enter')
        pyautogui.typewrite('*code*'.upper())
        pyautogui.press('enter')
        for num in storelist:
            num_str = f"{num:02}"
            folder_name = Path(save_path)
            folder_name.mkdir(parents=True, exist_ok=True)
            sleep(1)
            ### INSTRUCTION
            sleep(6)
            print(" ")
            print(f"Scanning Store " + str(num) + " now...")
            print(" ")
            region = (3300, 760, 100, 197)
            key = False
            while True:
                if check_region_color_change(region):
                    print(f"Store {num} report finished. Saving...")
                    print(" ")
                    sleep(7)
                    pyautogui.press('tab')
                    pyautogui.press('tab')
                    pyautogui.press('tab')
                    pyautogui.press('space')
                    pyautogui.press('tab')
                    pyautogui.press('tab')
                    pyautogui.press('enter')
                    sleep(5)
                    base_filename = f"{num_str} - Floor Plan"
                    version = 1
                    while True:
                        if version == 1:
                            full_path_excel = folder_name / f"{base_filename}.xls"
                        else:
                            full_path_excel = folder_name / f"{base_filename}(v{version}).xls"
                        if not full_path_excel.exists():
                            break
                        version += 1
                    pyautogui.typewrite(str(full_path_excel))
                    sleep(1)
                    pyautogui.press('enter')
                    sleep(12)
                    pyautogui.press('right')
                    sleep(3)
                    pyautogui.hotkey('alt', 'f4')
                    sleep(1)
                    print("Store " + str(num) + " complete.")
                    print(" ")
                    self.finished_stores += 1
                    self.update_progress(self.finished_stores / len(storelist))
                    key = True
                    break
                else:
                    sleep(1)
        print("All reports completed successfully.")
    except KeyboardInterrupt:
        print("\nKeyboard interrupt received! Cleaning up...")
    except Exception as e:
        print(f"Error during report generation: {e}")

def click(x, y):
    pyautogui.click(x, y)

if __name__ == "__main__":
    ensure_packages_installed(required_packages)
    root = ctk.CTk()
    app = ReportGeneratorApp(root)
    root.mainloop()