### Carter Grohe UVA Statistics/Data Science '27

### This code compiled and built daily 'schedules', which were documents in all the departments
### at each dealership for my internship at Hall MileOne Automotive using *XXX*, Python, and PyInstaller.
### *XXX* required PyAutoGUI, so my code was hampered by old software.

### The script worked with a JSON file with people's unique coordinates to pinpoint where the mouse
### could click with PyAutoGUI. Each person had a key they inputted before running.

### Although the choices were manual (no API compatibility), the code allows you to select

import subprocess
import sys

required_packages = ["pyautogui", "pynput", "keyboard", "setuptools", "customtkinter"]

def check_and_install_packages():
    for package in required_packages:
        try:
            __import__(package)
            print(f"{package} is already installed.")
        except ImportError:
            print(f"{package} not found. Installing...")
            subprocess.run([sys.executable, "-m", "pip", "install", package], check=True)
check_and_install_packages()

import tkinter as tk
import customtkinter
import pyautogui
from time import sleep
from datetime import datetime
from pynput import mouse
import keyboard
from tkinter import filedialog
import threading
from pathlib import Path
import json
import re
import os
import ctypes

## This function worked with JSON files.
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def turn_off_capslock():
    if ctypes.WinDLL("User32.dll").GetKeyState(0x14) == 1: # IF CAPSLOCK is on: (CAPSLOCK = 0x14)
        pyautogui.press('capslock') #Turn off.
        sleep(0.4)

turn_off_capslock()


'''Everything above here is installing and importing the necessary packages for the app.'''
specificStores = []
specificSchedules = []   ### Defining variables here. Do not change please.
global PostAhead
global ZeroBalanced
global Detail

### THINGS THAT CAN BE CHANGED #######

store = list(range(1, x, 1)) ### Currently from 1 to x. To add, change the second number to Store Count plus one. 24 Stores -> Change to 25.
stores_to_skip = [xxx] ### Inactive Stores. If a store becomes inactive, add here. Vice versa for new stores.
schedules_by_store = {
    1: (5, 7, 11, 14, 15, 19, 21), ## example. Each Store had a manual input here (No API possible).
}
'''Add schedules for new stores here.'''

### NO MORE CHANGES NEEDED BELOW THIS LINE ####

unique_values = set()
for tup in schedules_by_store.values():
    for t in tup:
        if t >= 100:
            unique_values.update(tup)
unique_values = sorted(unique_values)

current_date = datetime.now().strftime("%Y-%m-%d")

class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert("end", text)
        self.widget.see("end")

    def flush(self):
        pass

max_iterations = 5

instructions = [
    "*Instructions*"]

if len(specificStores) != 0:
    cutStores = [num for num in specificStores if num not in stores_to_skip]
    correctStores = [num for num in cutStores if num in store]
    storelist = correctStores
else:
    pass

def capture_positions(self):
    if self.position_check:
        for i in range(max_iterations):
            keyboard.wait('enter')
            if keyboard.is_pressed('enter'):
                x, y = pyautogui.position()
                self.positions.append((x, y))
            print(self.messages[i + 1])
            print(" ")
            if i == 0:
                print(f"Next: {instructions[4]}")
            elif i < max_iterations - 1:
                print(f"Next: {instructions[i + 4]}")
            else:
                print("Thank you! There is one last step required. Please wait briefly.")
                sleep(3)
    else:
        self.positions = self.saved_positions

def masterCode(self, storelist, specificSchedules, PostAhead, ZeroBalanced, Detail, Dump, path):
    self.toggle_back_button(False)
    storecount = 0
    print("\n------------------------------------\n")
    if self.all_stores_var.get():
        print("Running all stores.")
    else:
        print(f"Running stores: " + str(storelist))
    if self.all_scheds_var.get() or len(specificSchedules) == 0:
        print(f"Running all schedules.")
    else:
        print(f"Running schedules: " + str(specificSchedules))
    print("Include Post-Ahead:", "Yes" if PostAhead else "No")
    print("Include Zero-Balanced:", "No" if ZeroBalanced else "Yes")
    print("Dump to One Folder? ", "Yes" if Dump else "No")
    print("Detailed" if Detail else "Summary")
    print("\n------------------------------------")
    for code in storelist:
        if code in stores_to_skip:
            continue
        elif code not in store:
            print("Store not found. Breaking...")
            continue
        else:
            storenum = code
            count = 0
            storecount += 1
            code_str = f"{code:02}"
            if Dump:
                folder_name = Path(path) / current_date
                folder_name.mkdir(parents=True, exist_ok=True)
            else:
                folder_name = Path(path) / code_str / f"{current_date} Schedules"
                folder_name.mkdir(parents=True, exist_ok=True)
            sleep(5)
            pyautogui.click(self.positions[1])
            pyautogui.typewrite(str(storenum))
            pyautogui.press('enter')
            sleep(4)
            if Detail:
                pyautogui.click(self.positions[3])
            if PostAhead:
                pyautogui.click(self.positions[4])
            else:
                pass
            if ZeroBalanced:
                pyautogui.click(self.positions[5])
            else:
                pass
            print("\nBeginning Store " + str(code) + "...\n")
            for num in schedules_by_store:
                if code == num:
                    if len(specificSchedules) != 0:
                        scheds = specificSchedules
                    else:
                        scheds = schedules_by_store[num]
                    for sch in scheds:
                        pyautogui.click(self.positions[2])
                        pyautogui.click(self.positions[2])
                        sleep(0.25)
                        if sch not in schedules_by_store[num]:
                            print("Schedule " + str(sch) + " not found for Store " + str(storenum) + "!\n")
                            count += 1
                            continue
                        else:
                            pyautogui.typewrite(str(sch))
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('tab')
                        pyautogui.press('enter')
                        sleep(20)
                        if not self.action_done:
                            print(f"Please hover over and press ENTER on the Export button. DO NOT CLICK.\n")
                            keyboard.wait('enter')
                            if keyboard.is_pressed('enter'):
                                x, y = pyautogui.position()
                                self.positions.append((x, y))
                                if self.user_key not in self.all_positions:
                                    self.all_positions[self.user_key] = self.positions
                                    try:
                                        with open(self.shared_json, "w") as f:
                                            json.dump(self.all_positions, f, indent=2)
                                        print(f"Saved new positions for '{self.user_key}' to shared JSON.")
                                        sleep(3)
                                    except Exception as e:
                                        print(f"Failed to save: {e}")
                                print(f"Thank you. The program is now running. Please do not touch anything until program breaks or is complete.")
                                sleep(2)
                            self.action_done = True
                        sleep(2)
                        pyautogui.click(self.positions[6])
                        sleep(3)
                        suffix_parts = []
                        if PostAhead:
                            suffix_parts.append("PA")
                        if not ZeroBalanced:
                            suffix_parts.append("ZB")
                        if Detail:
                            suffix_parts.append("DETAIL")
                        suffix = "-".join(suffix_parts) if suffix_parts else ""
                        if Dump:
                            base_filename = f"{code_str}"
                        else:
                            base_filename = f"{code_str}-{sch}"
                        ## if suffix:
                            ## base_filename += f"-{suffix}"  ## TO INCLUDE THE SUFFIX, REMOVE THESE HASHTAGS
                        version = 1
                        while True:
                            if version == 1:
                                full_path = folder_name / f"{base_filename}.xlsx"
                            else:
                                full_path = folder_name / f"{base_filename}(v{version}).xlsx"
                            if not full_path.exists():
                                break
                            version += 1
                        pyautogui.typewrite(str(full_path))
                        pyautogui.press('enter')
                        sleep(10)
                        pyautogui.hotkey('alt', 'f4')
                        sleep(10)
                        pyautogui.hotkey('ctrl', 'w')
                        sleep(2)
                        count += 1
                        print("Store " + str(code) + " schedule " + str(count) + " of " + str(len(scheds)) + " complete.")
                        print(f"This is store {storecount} of {len(storelist)}\n")
            print("-----\nAll schedules for Store " + str(code) + " complete.\n-----")
    print("\n\n\n\nAutomation complete. Thank you.")
    self.toggle_back_button(True)

customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("blue")

class AutomationApp(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Accounting Schedule Automation")
        self.geometry("1000x700")
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("dark-blue")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.intro_frame = customtkinter.CTkFrame(self)
        self.intro_frame.grid(row=0, column=0, sticky="nsew")

        self.main_frame = customtkinter.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.report_frame = customtkinter.CTkFrame(self)
        self.fail_frame = customtkinter.CTkFrame(self)
        self.check_frame = customtkinter.CTkFrame(self)
        self.check_frame.grid(row=0, column=0, sticky="nsew")
        self.stores_to_skip = stores_to_skip

        self.store_count = len(store)
        self.unique_values = unique_values
        self.columns = 5
        self.title_text = "Accounting Schedule Automation Store Selection"
        self.subtitle_text = "Select the stores to run here (Invalid Stores in Red). Then, select where to save the files, and click Generate to begin."

        self.selected_stores = []
        self.selected_scheds = []
        self.store_vars = {}
        self.sched_vars = {}
        self.var_post_ahead = customtkinter.BooleanVar()
        self.var_zero_balanced = customtkinter.BooleanVar()
        self.var_detail = customtkinter.BooleanVar()
        self.one_sched = customtkinter.BooleanVar()
        self.shared_json = Path(resource_path(os.path.join("Do not touch please!", "positions.json")))
        self.positions = []
        self.saved_positions = []
        self.position_check = True
        for frame in (self.intro_frame, self.main_frame, self.report_frame, self.fail_frame, self.check_frame):
            frame.grid(row=0, column=0, sticky='nsew')

        self.checked_label_var = tk.StringVar(value="Selected Stores: None")
        self.checked_sched = tk.StringVar(value="Selected Schedules: All by Default")
        self.file_path_var = customtkinter.StringVar(value="*Network File Path*")

        self.setup_intro_frame()
        self.setup_main_frame()
        self.setup_report_frame()
        self.setup_fail_frame()
        self.setup_check_frame()

        self.show_frame(self.intro_frame)

    def load_access_codes(self):
        json_relative_path = os.path.join("Do not touch please!", "positions.json")
        json_path = resource_path(json_relative_path)
        with open(json_path, "r") as f:
            return json.load(f)

    def on_click(self, x, y, button, pressed):
        if pressed:
            self.positions.append((x, y))
            print(f"{self.messages[0]}")
            print(" ")
            print(f"Next: {instructions[3]}")
            return False
        return None

    def show_frame(self, frame):
        frame.tkraise()

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.file_path_var.set(directory)

    def setup_intro_frame(self):
        intro_text = (
            "*Instructions*"
        )

        # Configure grid rows and columns to expand properly
        self.intro_frame.grid_rowconfigure(0, weight=1)  # spacer above label
        self.intro_frame.grid_rowconfigure(1, weight=0)  # label
        self.intro_frame.grid_rowconfigure(2, weight=1)  # spacer below label
        self.intro_frame.grid_rowconfigure(3, weight=0)  # button
        self.intro_frame.grid_rowconfigure(4, weight=1)  # spacer below button
        self.intro_frame.grid_columnconfigure(0, weight=1)

        intro_label = customtkinter.CTkLabel(
            self.intro_frame,
            text=intro_text,
            font=("Verdana", 19),
            justify="center",
            wraplength=900  # limits line width for better readability
        )
        intro_label.grid(row=1, column=0, padx=40, pady=10, sticky="nsew")

        warning_label = customtkinter.CTkLabel(
            self.intro_frame,
            text="PLEASE READ ALL INSTRUCTIONS IN TXT FILE AND IN APP.",
            text_color="#E63946",
            font=("Verdana", 18, "bold"))

        warning_label.grid(row=0, column=0, padx=40, pady=10, sticky="nsew")

        ok_button = customtkinter.CTkButton(
            self.intro_frame,
            text="OK",
            font=("Verdana", 22, "bold"),
            height=60,
            width=200,
            command=lambda: self.show_frame(self.check_frame)
        )
        ok_button.grid(row=3, column=0, pady=(20, 40), sticky="s")

    def setup_check_frame(self):
        self.check_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)
        self.check_frame.grid_columnconfigure(0, weight=1)

        # Instruction label with Verdana font and bigger size
        label = customtkinter.CTkLabel(
            self.check_frame,
            text="*Instructions*",
            font=("Verdana", 18, "bold"),
            wraplength=900
        )
        label.grid(row=0, column=0, pady=(40, 5), padx=20)

        # Sample label below instructions, smaller font and italic for differentiation
        sample_label = customtkinter.CTkLabel(
            self.check_frame,
            text="Please only use numbers or letters. Access codes are case-sensitive.",
            font=("Verdana", 16, "underline", "bold"),
        )
        sample_label.grid(row=1, column=0, pady=(0, 15))

        self.suggestions = self.load_access_codes()

        self.user_entry = customtkinter.CTkEntry(self.check_frame, width=300, font=("Verdana", 14))
        self.user_entry.grid(row=2, column=0, pady=10, padx=20)

        self.listbox = tk.Listbox(
            self.check_frame,
            height=5,
            width=33,
            bg="#2b2b2b",  # dark background
            fg="white",  # white text
            font=("Verdana", 10, "italic"),
            highlightthickness=1,
            highlightbackground="#3a3a3a",  # border color
            selectbackground="#3a7ebf",  # selection background (light blue)
            selectforeground="white",  # selection text color
            relief="flat",  # modern flat look
            bd=1
        )
        self.listbox.grid(row=3, column=0, padx=20, sticky="ew")
        self.listbox.grid_remove()

        self.user_entry.bind("<KeyRelease>", self.on_keyrelease)
        self.listbox.bind("<<ListboxSelect>>", self.on_listbox_select)

        self.json_path_label = customtkinter.CTkLabel(self.check_frame, text="", font=("Verdana", 10),
                                                      text_color="#AAAAAA")
        self.json_path_label.grid(row=6, column=0, pady=(0, 10))

        submit_button = customtkinter.CTkButton(
            self.check_frame,
            text="Submit",
            font=("Verdana", 14),
            command=self.handle_user_position_check
        )
        submit_button.grid(row=4, column=0, pady=10)

        self.check_frame_feedback = customtkinter.CTkLabel(self.check_frame, text="", font=("Verdana", 12))
        self.check_frame_feedback.grid(row=5, column=0, pady=10)

    def on_keyrelease(self, event):
        typed = self.user_entry.get()
        print(f"Typed: '{typed}'")  # Debug
        filtered = [s for s in self.suggestions if s.lower().startswith(typed.lower())]
        print(f"Filtered: {filtered}")  # Debug

        if not typed or not filtered:
            self.listbox.grid_remove()
            return

        self.listbox.delete(0, tk.END)
        for item in filtered:
            self.listbox.insert(tk.END, item)
        self.listbox.grid()
        self.listbox.update()  # force UI update

    def on_listbox_select(self, event):
        if not self.listbox.curselection():
            return
        index = self.listbox.curselection()[0]
        selected = self.listbox.get(index)
        self.user_entry.delete(0, "end")
        self.user_entry.insert(0, selected)
        self.listbox.grid_remove()

    def handle_user_position_check(self):
        user_key = self.user_entry.get().strip()

        if not user_key:
            self.check_frame_feedback.configure(text="Access code cannot be empty.", text_color="red", font=("Verdana", 14))
            return
        if not re.fullmatch(r"[A-Za-z0-9]+", user_key):
            self.check_frame_feedback.configure(text="Access code must contain only letters and numbers (no spaces).",
                                                text_color="red", font=("Verdana", 14))
            return

        self.shared_json = Path(resource_path(os.path.join("Do not touch please!", "positions.json")))

        if self.shared_json.exists():
            try:
                with open(self.shared_json, "r") as f:
                    all_positions = json.load(f)
            except json.JSONDecodeError:
                self.check_frame_feedback.configure(text="JSON is corrupted. Starting fresh.", text_color="orange", font=("Verdana", 14))
                all_positions = {}
        else:
            all_positions = {}

        self.user_key = user_key
        self.all_positions = all_positions

        json_path_str = str(self.shared_json)

        if user_key in all_positions:
            self.positions = all_positions[user_key]
            self.saved_positions = self.positions
            self.position_check = False
            self.check_frame_feedback.configure(
                text=f"Found saved positions for '{user_key}'. Welcome back!",
                text_color="green",
                font=("Verdana", 14)
            )
            self.json_path_label.configure(text=f"JSON file location: {json_path_str}")
            self.action_done = True
        else:
            self.positions = []
            self.action_done = False
            self.position_check = True
            self.check_frame_feedback.configure(
                text=f"No saved positions for '{user_key}', please continue to capture positions.",
                text_color="white",
                font=("Verdana", 14)
            )
            self.json_path_label.configure(text=f"JSON file location: {json_path_str}")

        self.after(1750, lambda: self.show_frame(self.main_frame))

        self.messages = [
            f"*XXX* screen saved to user {user_key}.",
            f"Company button saved to user {user_key}.",
            f"Schedule number button saved to user {user_key}.",
            f"Detail button saved to user {user_key}.",
            f"Post-Ahead Button saved to user {user_key}.",
            f"Zero-Balanced Button saved to user {user_key}.",
            f"Export Button saved to user {user_key}."
        ]

    def setup_main_frame(self):
        # Configure main_frame grid
        for i in range(9):
            self.main_frame.grid_rowconfigure(i, weight=0)
        self.main_frame.grid_rowconfigure(3, weight=1)
        self.main_frame.grid_rowconfigure(4, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Title
        title_label = customtkinter.CTkLabel(
            self.main_frame, text=self.title_text,
            font=("Verdana", 18, "bold"), justify="center", wraplength=800
        )
        title_label.grid(row=0, column=0, sticky="n", pady=(15, 5))

        # Subtitle
        subtitle_label = customtkinter.CTkLabel(
            self.main_frame, text=self.subtitle_text,
            font=("Verdana", 12), justify="center", wraplength=800
        )
        subtitle_label.grid(row=1, column=0, sticky="n", pady=(0, 15))

        # Title row for checkboxes
        checkbox_title_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        checkbox_title_frame.grid(row=2, column=0, sticky="ew", padx=(250, 150))
        checkbox_title_frame.grid_columnconfigure((0, 1), weight=1)

        # Left (dynamic) and right (static) titles
        left_title = customtkinter.CTkLabel(
            checkbox_title_frame, textvariable=self.checked_label_var,
            font=("Verdana", 13, "bold"), anchor="w"
        )
        left_title.grid(row=0, column=0, sticky="w", padx=(0, 20))

        right_title = customtkinter.CTkLabel(
            checkbox_title_frame, textvariable=self.checked_sched,
            font=("Verdana", 13, "bold"), anchor="e", wraplength=200, justify="right"
        )
        right_title.grid(row=0, column=1, sticky="e")

        # Main checkbox section
        checkbox_container = customtkinter.CTkFrame(self.main_frame)
        checkbox_container.grid(row=3, column=0, rowspan=2, sticky="nsew", padx=20)
        checkbox_container.grid_columnconfigure((0, 1), weight=1)
        checkbox_container.grid_rowconfigure(0, weight=1)

        # Left: store checkboxes
        store_frame = customtkinter.CTkFrame(checkbox_container)
        store_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 10))
        store_frame.grid_columnconfigure(tuple(range(self.columns)), weight=1)

        for i in range(self.store_count):
            var = customtkinter.BooleanVar()
            store_num = str(i + 1)
            self.store_vars[store_num] = var

            cb = customtkinter.CTkCheckBox(
                store_frame,
                text=f"Store {store_num}",
                variable=var,
                font=("Verdana", 12, "bold"),
                command=self.update_checked_stores,
                text_color="#FF4444" if int(store_num) in self.stores_to_skip else "white"
            )
            cb.grid(row=i // self.columns, column=i % self.columns, sticky="w", padx=10, pady=6)

        self.all_stores_var = customtkinter.BooleanVar()
        cb_all = customtkinter.CTkCheckBox(
            store_frame,
            text="All Active Stores",
            variable=self.all_stores_var,
            font=("Verdana", 12, "bold"),
            command=self.toggle_all_stores,
            text_color="white"
        )

        self.all_scheds_var = customtkinter.BooleanVar()
        cb_all2 = customtkinter.CTkCheckBox(
            store_frame,
            text="All Schedules",
            variable=self.all_scheds_var,
            font=("Verdana", 12, "bold"),
            command=self.toggle_all_scheds,
            text_color="white"
        )

        cb_all.grid(
            row=(self.store_count // self.columns) + 1,
            column=self.columns - 1,
            sticky="w", padx=10, pady=6)

        cb_all2.grid(
            row=(self.store_count // self.columns) + 2,
            column=self.columns - 1,
            sticky="w", padx=10, pady=6)

        # Right: scrollable schedule options
        schedule_frame = customtkinter.CTkScrollableFrame(checkbox_container, width=200)
        schedule_frame.grid(row=0, column=1,sticky="nsew", padx=(10, 0))

        for sched_value in self.unique_values:
            var2 = customtkinter.BooleanVar()
            self.sched_vars[sched_value] = var2

            cb = customtkinter.CTkCheckBox(
                schedule_frame,
                text=sched_value,
                variable=var2,
                command=self.update_checked_scheds,
                font=("Verdana", 12),
                text_color="white"
            )
            cb.pack(anchor="w", pady=4, padx=10)

        # File path entry
        file_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        file_frame.grid(row=7, column=0, sticky="ew", padx=20, pady=(10, 5))
        file_frame.grid_columnconfigure(1, weight=1)
        file_frame.grid_columnconfigure((0,2), weight=0)

        dump = customtkinter.CTkCheckBox(file_frame, text="Dump to One Folder", height=36, font=("Verdana", 12, "bold"), variable=self.one_sched)
        dump.grid(row=0, column=0, sticky="ew", padx=10)

        path_entry = customtkinter.CTkEntry(
            file_frame,
            textvariable=self.file_path_var,
            placeholder_text="Select a directory...",
            height=36,
            font=("Verdana", 14)
        )
        path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        browse_button = customtkinter.CTkButton(
            file_frame,
            text="Browse",
            command=self.browse_directory,
            width=100,
            height=36,
            font=("Verdana", 14),
            corner_radius=8
        )
        browse_button.grid(row=0, column=2)

        # Bottom frame
        bottom_frame = customtkinter.CTkFrame(self.main_frame)
        bottom_frame.grid(row=8, column=0, sticky="ew", padx=20, pady=(10, 20))
        bottom_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        post_cb = customtkinter.CTkCheckBox(
            bottom_frame,
            text="Click to include Post-Ahead reports",
            variable=self.var_post_ahead,
            font=("Verdana", 12, "bold")
        )
        post_cb.grid(row=0, column=0, sticky="w", padx=(10, 0))

        zero_cb = customtkinter.CTkCheckBox(
            bottom_frame,
            text="Click to include Zero-Balanced reports",
            variable=self.var_zero_balanced,
            font=("Verdana", 12, "bold")
        )
        zero_cb.grid(row=0, column=1, sticky="w", padx=(20, 0))

        detail_cb = customtkinter.CTkCheckBox(
            bottom_frame,
            text="Click to set to Detail",
            variable=self.var_detail,
            font=("Verdana", 12, "bold")
        )
        detail_cb.grid(row=0, column=2, sticky="w", padx=(20, 0))

        self.generate_button = customtkinter.CTkButton(
            bottom_frame,
            text="Generate",
            command=self.on_generate,
            width=140,
            height=36,
            font=("Verdana", 14, "bold")
        )
        self.generate_button.grid(row=0, column=4, sticky="e", padx=(0, 10))

    def toggle_all_stores(self):
        select_all = self.all_stores_var.get()
        for num, var in self.store_vars.items():
            if num != str(self.store_count + 1) and int(num) not in self.stores_to_skip:
                var.set(select_all)
        self.update_checked_stores()

    def toggle_all_scheds(self):
        select_all = self.all_scheds_var.get()
        for num, var in self.sched_vars.items():
            if num != str(len(self.unique_values) + 1):
                var.set(select_all)
        self.update_checked_scheds()

    def setup_report_frame(self):
        # Configure grid rows and columns for resizing
        self.report_frame.grid_rowconfigure((0,1,2,3), weight=1)
        self.report_frame.grid_rowconfigure(4, weight=0)
        self.report_frame.grid_columnconfigure(0, weight=1)

        # Text widget for live console output, bigger font, placed at top
        self.console = customtkinter.CTkTextbox(
            self.report_frame,
            width=800,
            height=200,
            font=("Verdana", 14)  # bigger font size
        )
        self.console.grid(row=0, column=0, rowspan=3, padx=20, pady=10, sticky="nsew")

        # Scrollbar on the right side of console
        scrollbar = customtkinter.CTkScrollbar(self.report_frame, command=self.console.yview)
        scrollbar.grid(row=0, column=1, rowspan=3, sticky="ns", pady=10, padx=(0, 20))
        self.console.configure(yscrollcommand=scrollbar.set)

        # Redirect stdout and stderr to console
        sys.stdout = TextRedirector(self.console)
        sys.stderr = TextRedirector(self.console)

        # Frame for Stop button centered
        self.button_frame = customtkinter.CTkFrame(self.report_frame)
        self.button_frame.grid(row=4, column=0, columnspan=2, padx=20, pady=(0, 20))
        self.button_frame.grid_columnconfigure(0, weight=1)
        self.button_frame.grid_columnconfigure(1, weight=1)

        self.stop_button = customtkinter.CTkButton(self.button_frame, text="Stop", width=160,
            height=50, command=self.stop_process)
        self.stop_button.grid(row=0, column=0, columnspan=2)

        self.back_to_main_button = customtkinter.CTkButton(
            self.button_frame,
            text="Back to Main Menu",
            font=("Verdana", 12),
            width=160,
            height=50,
            command=lambda: self.show_frame(self.main_frame)
        )
        self.back_to_main_button.grid(row=0, column=1)
        self.back_to_main_button.grid_remove()

    def toggle_back_button(self, show: bool):
        if show:
            self.back_to_main_button.grid()
            self.stop_button.grid_configure(column=0, columnspan=1)
            self.back_to_main_button.grid_configure(column=1)
        else:
            self.back_to_main_button.grid_remove()
            self.stop_button.grid_configure(column=0, columnspan=2)

    def setup_fail_frame(self):
        self.fail_frame.grid_rowconfigure(0, weight=1)  # top spacer
        self.fail_frame.grid_rowconfigure(1, weight=0)  # label
        self.fail_frame.grid_rowconfigure(2, weight=0)  # button
        self.fail_frame.grid_rowconfigure(3, weight=1)  # bottom spacer
        self.fail_frame.grid_columnconfigure(0, weight=1)

        label = customtkinter.CTkLabel(
            self.fail_frame,
            text="Sorry, you must select at least one store.",
            font=("Verdana", 18, "bold"),
            text_color="#FF4444",
            wraplength=400,
            justify="center"
        )
        label.grid(row=1, column=0, pady=(0, 20), sticky="n")

        # Return button
        back_button = customtkinter.CTkButton(
            self.fail_frame,
            text="Return",
            font=("Verdana", 14),
            command=lambda: self.show_frame(self.main_frame),
            width=150,
            height=40,
            corner_radius=8,
            fg_color="#007ACC",
            hover_color="#005F99"
        )
        back_button.grid(row=2, column=0, sticky="n")


    # ======================================================================

    def stop_process(self):
        sys.exit()

    def on_generate(self):
        self.selected_stores = [int(store) for store, var in self.store_vars.items() if var.get()]
        self.selected_scheds = [int(unique_values) for unique_values, var in self.sched_vars.items() if var.get()]
        post_ahead = self.var_post_ahead.get()
        zero_balanced =  not self.var_zero_balanced.get()
        detail = self.var_detail.get()
        one_sched = self.one_sched.get()
        if not self.selected_stores:
            self.show_frame(self.fail_frame)
            return
        self.show_frame(self.report_frame)
        self.console.delete('1.0', tk.END)
        sys.stdout = TextRedirector(self.console)
        path = self.file_path_var.get()
        threading.Thread(target=self.run_master_code, args=(self.selected_stores, self.selected_scheds, post_ahead, zero_balanced, detail, one_sched, path), daemon=True).start()

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
        print("Selected stores:", self.selected_stores)

    def update_checked_scheds(self):
        self.selected_scheds = [sched for sched, var in self.sched_vars.items() if var.get()]
        if self.selected_scheds and not self.all_scheds_var.get():
            display_text = "Selected Schedules: Schedules " + ", ".join(str(s) for s in self.selected_scheds)
        elif self.selected_scheds and self.all_scheds_var.get():
            display_text = "Selected Schedules: All Schedules"
        else:
            display_text = "Selected Schedules: All by Default"
        self.checked_sched.set(display_text)
        print("Selected schedules:", self.selected_scheds)

    # ======================================================================

    def run_master_code(self, selected_stores, selected_schedules, PostAhead, ZeroBalanced, Detail, Dump, path):
        global storelist
        if self.position_check:
            print("\n".join(instructions[0:3]))
            with mouse.Listener(on_click=self.on_click) as listener:
                listener.join()
            capture_positions(self)
        else:
            pass
        masterCode(self, selected_stores, selected_schedules, PostAhead, ZeroBalanced, Detail, Dump, path)

if __name__ == "__main__":
    check_and_install_packages()
    app = AutomationApp()
    app.mainloop()