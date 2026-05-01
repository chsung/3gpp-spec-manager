import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
from tkinter import StringVar
from tkinter import font
import os
import ftplib
import json
import requests
from bs4 import BeautifulSoup
import webbrowser
import datetime
import tempfile
import zipfile
import threading
import queue
import configparser
import shutil
import sys
import re
import time
import ctypes
try:
    import pystray
    from PIL import Image
except ImportError:
    pystray = None
    Image = None

PROGRAM_NAME = "3GPP Spec Manager"
PROGRAM_VERSION = "1"
PROGRAM_DIR = "3gppSpecManager"
ERROR_ALREADY_EXISTS = 183
SINGLE_INSTANCE_MUTEX_NAME = f"Local\\{PROGRAM_DIR}.SingleInstance"
SINGLE_INSTANCE_ACTIVATE_EVENT_NAME = f"Local\\{PROGRAM_DIR}.Activate"

FTP_HOST_SERVER = "ftp.3gpp.org"
FTP_SERIES_PATH = "/Specs/archive/{}_series/"
FTP_DOWNLOAD_BLOCKSIZE = 65536
FTP_TIMEOUT = 30
FTP_MAX_RETRY = 3
FTP_RETRY_SLEEP = 5
TOOLTIP_DURATION_MS = 1500
HTTP_DYNA_REPORT = "https://www.3gpp.org/DynaReport/{}.htm"

SAVE_FILE_SETTINGS = "settings.ini"
SAVE_FILE_LIST = "list.json"

CONFIG_SECTION = "Settings"
CONFIG_KEY_ROOT_FOLDER = "root_folder"
CONFIG_KEY_PREV_VER = "prev_ver"
CONFIG_KEY_WIN_GEOMETRY = "geometry"
CONFIG_KEY_THEME = "theme"
CONFIG_KEY_COMPACT_MODE = "compact_mode"
CONFIG_KEY_SHOW_TOOLTIPS = "show_tooltips"
CONFIG_KEY_MINIMIZE_TO_TRAY = "minimize_to_tray"

FILE_SPEC_NAMES = "spec_names.json"
FILE_APP_ICON = "app_icon.ico"

ICON_CHECK_ON = "☑"
ICON_CHECK_OFF = "☐"
ICON_FOLDER = "📂"
ICON_WEB_LINK = "🔗"
ICON_FAVORITE = "★"
ICON_UNFAVORITE = "☆"
ICON_BLANK = ""
ICON_SEARCH = "🔍"
ICON_SETTINGS = "⚙️"
ICON_REMOVE = "✗"

TEXT_BROWSE_SPECS = "Browse specifications"
TEXT_TOOLTIP_SELECTION_REQUIRED = "Please select an item first"
TEXT_TOOLTIP_ROOT_FOLDER = "Double-click to open folder"
TEXT_TOOLTIP_SETTINGS = "Double-click to toggle Compact mode"
TEXT_TOOLTIP_ADD_SPEC = "Add specification to list"
TEXT_TOOLTIP_DOWNLOAD = "Download latest document"
TEXT_TOOLTIP_REMOVE = "Remove selected from list"
TEXT_TOOLTIP_FAVORITE = "Add selected to favorites"
TEXT_TOOLTIP_UNFAVORITE = "Remove selected from favorites"
TEXT_TOOLTIP_RENAME_FOLDER = "Double-click to open file"
TEXT_TOOLTIP_OPEN_FOLDER = "Open folder"
TEXT_TOOLTIP_GO_TO_WEB = "Go to 3GPP website"
DEFAULT_FAVORITE = ""
MAX_LIST_ITEMS = 100
MAX_FOLDER_NAME = 180
INVALID_FOLDER_CHARS_REGEX = r'[<>:"/\\|?*\x00-\x1F]'

PLACEHOLDER_TEXT_INPUT_SPEC_NUM = "ex) 38.300"
PLACEHOLDER_TEXT_ROOT_FOLDER = "Select a root folder for saving"
TEXT_COLOR_DEFAULT = 'black'
TEXT_COLOR_PLACEHOLDER = 'grey'

# Messagebox Titles
MSG_TITLE_ADD_LIMIT = "Add limit"
MSG_TITLE_INVALID_SPEC = "Invalid specification"
MSG_TITLE_DUPLICATE_SPEC = "Duplicate specification"
MSG_TITLE_ERROR = "Error"
MSG_TITLE_NOT_FOUND = "Not found"
MSG_TITLE_CANCELED = "Canceled"
MSG_TITLE_COMPLETED = "Completed"
MSG_TITLE_OPERATION_STATUS = "Operation Status"
MSG_TITLE_NO_ROOT = "No root"
MSG_TITLE_CONFIRM_REMOVE = "Confirm remove"

# Messagebox Messages
MSG_ADD_LIMIT = f"Cannot add more than {MAX_LIST_ITEMS} items."
MSG_INVALID_SPEC = "Please check specification format.\n(e.g., 38.300 or 38.101-1)."
MSG_DUPLICATE_SPEC = "'{}' already exists in the list."
MSG_FTP_CONNECTION_FAILED = "FTP connection failed.\nPlease try again later."
MSG_SPEC_NOT_FOUND = "Specification '{}' was not found."
MSG_FTP_CONNECTION_FAILED_WEB = "FTP connection failed.\nPlease try again later or download it directly using the web link."
MSG_CANCELED_DOWNLOAD = "The process was canceled.\n\n{}"
MSG_COMPLETED_DOWNLOAD = "{} of {} items processed.\n\n{}"
MSG_NO_ITEMS_PROCESSED = "No items were processed."
MSG_SELECT_ROOT_FOLDER = "Please select a root folder."
MSG_FAILED_OPEN_FOLDER = "Failed to open folder:\n{}"
MSG_FAILED_OPEN_BROWSER = "Failed to open browser:\n{}"
MSG_FAILED_OPEN_FILE = "Failed to open file:\n{}"
MSG_NO_LOCAL_FILE = "No local file found in folder:\n{}"
MSG_CONFIRM_REMOVE = "Remove selected items?\nThis will remove the items from the list only. The files will not be deleted."
MSG_ABOUT_CONTENT = f"Version: {PROGRAM_VERSION}\nFeedback: ch22.sung@gmail.com\n"

KEY_SPEC_NUM = "spec_num"
KEY_FOLDER_NAME = "folder_name"
KEY_LATEST_FTP_FILE_NAME = "latest_ftp_file_name"
KEY_LAST_FTP_CHECK_TIME = "last_ftp_check_time"
KEY_FAVORITE = "favorite"

LIST_IDX_FAVORITE = 0
LIST_IDX_CHECKBOX = 1
LIST_IDX_STATUS = 2
LIST_IDX_FOLDER_NAME = 3
LIST_IDX_FOLDER_ACTION = 4
LIST_IDX_LINK_ACTION = 5
LIST_IDX_SPEC_NUM = 6
LIST_IDX_LATEST_FTP_FILE_BASE_NAME = 7
LIST_IDX_LAST_FTP_CHECK_TIME = 8

LIST_COLUMN_FOLDER_NAME = "#4"
LIST_COLUMN_OPEN_FOLDER = "#5"
LIST_COLUMN_OPEN_WEB_LINK = "#6"

LOCAL_FOLDER_STATUS_LATEST = "Latest"
LOCAL_FOLDER_STATUS_OLD = "Old"
LOCAL_FOLDER_STATUS_EMPTY = "Empty"

SERIES_LIST = {
    "21 series": "Requirements",
    "22 series": "Service aspects",
    "23 series": "Technical realization",
    "24 series": "UE to network protocols",
    "25 series": "UTRA",
    "26 series": "CODECs",
    "27 series": "Data",
    "28 series": "RSS-CN protocols, OAM&P and Charging",
    "29 series": "CN protocols",
    "30 series": "Programme management",
    "31 series": "(U)SIM, IC Cards test",
    "32 series": "OAM&P and Charging",
    "33 series": "Security aspects",
    "34 series": "UE and (U)SIM test",
    "35 series": "Security algorithms",
    "36 series": "LTE",
    "37 series": "Multi-RAT",
    "38 series": "NR"
}

root = None
root_folder_textbox = None
input_spec_number_textbox = None
list_view = None
download_button = None
add_button = None
remove_button = None
checkbox_header = None
PRELOADED_SPEC_NAMES = {}
update_queue = queue.Queue()
download_progress_window = None
download_cancel_event = None
check_ftp_file_thread = None
check_ftp_cancel_event = None
favorite_button = None
unfavorite_button = None
browse_specs_button = None
last_update_ftp_fail_time = None
settings_button = None
settings_menu = None
compact_mode_var = None
minimize_to_tray_var = None
tray_icon = None
show_tooltips_var = None
frame_row1 = None
frame_row2 = None
frame_row3 = None
single_instance_mutex_handle = None
activate_event_handle = None
activate_event_thread = None
activate_event_stop_event = threading.Event()


def acquire_single_instance_mutex():
    global single_instance_mutex_handle
    try:
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateMutexW(None, False, SINGLE_INSTANCE_MUTEX_NAME)
        if not handle:
            return True
        if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
            kernel32.CloseHandle(handle)
            return False
        single_instance_mutex_handle = handle
        return True
    except Exception:
        return True


def release_single_instance_mutex():
    global single_instance_mutex_handle
    handle = single_instance_mutex_handle
    single_instance_mutex_handle = None
    if not handle:
        return
    try:
        ctypes.windll.kernel32.CloseHandle(handle)
    except Exception:
        pass


def notify_existing_instance_to_activate():
    try:
        kernel32 = ctypes.windll.kernel32
        EVENT_MODIFY_STATE = 0x0002
        event_handle = kernel32.OpenEventW(EVENT_MODIFY_STATE, False, SINGLE_INSTANCE_ACTIVATE_EVENT_NAME)
        if event_handle:
            try:
                kernel32.SetEvent(event_handle)
                return True
            finally:
                kernel32.CloseHandle(event_handle)
    except Exception:
        pass
    return False

def scale_pixels(pixels):
    if root is None: return pixels
    try:
        dpi = root.winfo_fpixels('1i')
        return int(pixels * (dpi / 96.0))
    except Exception:
        return pixels

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.widget.tooltip_obj = self
        self.text = text
        self.tooltip_window = None
        self.after_id = None
        self.widget.bind("<Enter>", lambda e: self.show(self.text))
        self.widget.bind("<Leave>", self.hide)

    def show(self, text, duration=TOOLTIP_DURATION_MS):
        if show_tooltips_var is not None and not show_tooltips_var.get():
            return

        self.hide()

        if not text: return

        try:
            tooltip_font = font.Font(family="tahoma", size=8, weight="normal")
            text_width = tooltip_font.measure(text)
            x_offset = self.widget.winfo_rootx() + self.widget.winfo_width() // 2 - text_width // 2
            y_offset = self.widget.winfo_rooty() + self.widget.winfo_height() + scale_pixels(5)

            self.tooltip_window = tk.Toplevel(self.widget)
            self.tooltip_window.wm_overrideredirect(True)
            self.tooltip_window.wm_geometry(f"+{int(x_offset)}+{int(y_offset)}")
            self.tooltip_window.attributes('-topmost', True)

            label = tk.Label(self.tooltip_window, text=text, justify='left',
                             background="#ffffe0", relief='solid', borderwidth=scale_pixels(1),
                             font=("tahoma", "8", "normal"))
            label.pack(ipadx=scale_pixels(1))

            if duration:
                self.after_id = self.tooltip_window.after(duration, self.hide)
                self.tooltip_window.bind("<Button-1>", lambda e: self.hide())

        except Exception as e:
            print(f"Tooltip error: {e}")
            self.hide()

    def hide(self, event=None):
        if self.after_id:
            if self.tooltip_window:
                self.tooltip_window.after_cancel(self.after_id)
            self.after_id = None

        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

def show_tooltip(widget, text, duration=TOOLTIP_DURATION_MS):
    if not widget: return
    
    if hasattr(widget, 'tooltip_obj') and widget.tooltip_obj:
        widget.tooltip_obj.show(text, duration)


class SeriesComboBox(ttk.Combobox):
    def __init__(self, master, series_dict, on_series_change=None, **kwargs):
        self.series_dict = series_dict
        self.all_series_key = "All series"
        self.keys = [self.all_series_key] + list(series_dict.keys())
        self.display_map = [self.all_series_key] + [f"{key} - {series_dict[key]}" for key in series_dict.keys()]
        self.var = StringVar()
        self.on_series_change = on_series_change
        super().__init__(master, textvariable=self.var, values=self.keys, state="readonly", **kwargs)
        self.bind("<<ComboboxSelected>>", self._on_select)
        self.configure(values=self.display_map)
        self.set(self.all_series_key)

    def _on_select(self, event=None):
        selected_series = self.get()
        if self.on_series_change:
            self.on_series_change(selected_series)

class SpecComboBox(ttk.Combobox):
    def __init__(self, master, preloaded_spec_names, **kwargs):
        self.preloaded_spec_names = preloaded_spec_names
        self.filtered_keys = []
        self.display_map = []
        self.selected_index = None
        self.var = StringVar()
        super().__init__(master, textvariable=self.var, values=[], state="readonly", **kwargs)
        self.bind("<<ComboboxSelected>>", self._on_select)
        self.configure(values=self.display_map)

    def update_by_series(self, series_key):
        if series_key == "All series":
            self.filtered_keys = list(self.preloaded_spec_names.keys())
        else:
            prefix = series_key.split()[0]
            self.filtered_keys = [key for key in self.preloaded_spec_names if key.startswith(prefix)]

        self.filtered_keys.sort()
        self.display_map = [
            f"{key} {self.preloaded_spec_names[key]}" for key in self.filtered_keys
        ]
        self.configure(values=self.display_map)
        if self.filtered_keys:
            self.set(self.display_map[0])
            self.selected_index = 0
        else:
            self.set("")
        self.var.set(self.get())

    def _on_select(self, event=None):
        selected = self.get()
        self.selected_index = None
        for i, display in enumerate(self.display_map):
            if selected == display:
                self.selected_index = i
                break

def get_config_folder_path():
    appdata_path = os.getenv('LOCALAPPDATA')
    if not appdata_path or not os.path.isdir(appdata_path):
        appdata_path = os.path.expanduser('~')
    config_folder_path = os.path.join(appdata_path, PROGRAM_DIR)
    return config_folder_path

def get_config_file_path(file_name):
    config_folder_path = get_config_folder_path()
    file_path = os.path.join(config_folder_path, file_name)
    return file_path

def create_config_folder():
    config_folder_path = get_config_folder_path()
    try:
        os.makedirs(config_folder_path, exist_ok=True)
        return True
    except OSError as e:
        print(f"ERROR: create_config_folder failed: {e}")
        return False

def save_config_key_value(key, value):
    if not create_config_folder():
        return

    config = configparser.ConfigParser()
    config_settings_path = get_config_file_path(SAVE_FILE_SETTINGS)
    try:
        if os.path.exists(config_settings_path):
            config.read(config_settings_path, encoding='utf-8')
        if not config.has_section(CONFIG_SECTION):
            config.add_section(CONFIG_SECTION)
        config.set(CONFIG_SECTION, key, value)
        with open(config_settings_path, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
        print(f"save config '{key}':{value}")
    except Exception as e:
        print(f"ERROR: save_config_key_value failed for key '{key}': {e}")

def load_config_key_value(key):
    config = configparser.ConfigParser()
    config_settings_path = get_config_file_path(SAVE_FILE_SETTINGS)
    try:
        if os.path.exists(config_settings_path):
            config.read(config_settings_path, encoding='utf-8')
            if config.has_section(CONFIG_SECTION) and config.has_option(CONFIG_SECTION, key):
                value = config.get(CONFIG_SECTION, key)
                print(f"load config '{key}':{value}")
                return value
    except Exception as e:
        print(f"ERROR: load_config_key_value failed for key '{key}': {e}")
    return None

def delete_file_or_folder(path):
    if not path or not os.path.exists(path):
        return

    try:
        if os.path.isfile(path):
            os.remove(path)
        elif os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        print(f"'{path}' deleted.")
    except Exception as e:
        print(f"ERROR: delete_file_or_folder failed '{path}': {e}")

def copy_file_or_folder(src_path, dst_path):
    try:
        if os.path.exists(src_path):
            if os.path.isfile(src_path):
                shutil.copy2(src_path, dst_path)
            elif os.path.isdir(src_path):
                shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
    except Exception as e:
        print(f"ERROR: copy_file_or_folder failed '{src_path}' to '{dst_path}': {e}")

def bundled_data_path(path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")

    path = os.path.join(base_path, path)
    if os.path.exists(path):
        return path
    return None

def has_invalid_folder_char(name):
    return re.search(INVALID_FOLDER_CHARS_REGEX, name) is not None

def sanitize_folder_name(original_name):
    if not isinstance(original_name, str): return None
    sanitized_name = re.sub(INVALID_FOLDER_CHARS_REGEX, '', original_name)
    sanitized_name = sanitized_name.strip().rstrip('. ')
    return sanitized_name

def shorten_folder_name(name):
    if len(name) > MAX_FOLDER_NAME:
        return name[:150] + "....." + name[-25:]
    return name

def get_numeric_version_from_alpha(alpha_version):
  if 'a' <= alpha_version <= 'z':
    return str(ord(alpha_version) - ord('a') + 10)
  return None

def get_version_postfix_from_code(version_code):
    if not version_code:
        return None
    if len(version_code) != 3:
        return None
    parts = []
    for char_code in version_code:
        if char_code.isalpha():
            numeric_version = get_numeric_version_from_alpha(char_code.lower())
            if numeric_version is None:
                return None
            parts.append(numeric_version)
        elif char_code.isdigit():
            parts.append(char_code)
        else:
            return None
    return f"v{parts[0]}.{parts[1]}.{parts[2]}"

def format_ftp_time_to_yymmdd(time_str):
    if not time_str or time_str == "":
        return ""
    if len(time_str) == 6 and time_str.isdigit():
        return time_str
    if len(time_str) >= 8 and time_str.startswith('20') and time_str[:8].isdigit():
        return f"{time_str[2:8]}"
    return ""

def get_ftp():
    for attempt in range(FTP_MAX_RETRY):
        try:
            ftp = ftplib.FTP(timeout=FTP_TIMEOUT)
            ftp.connect(FTP_HOST_SERVER)
            ftp.login()
            ftp.set_pasv(True)
            print(f"connected to {FTP_HOST_SERVER}")
            return ftp
        except ftplib.all_errors as e:
            print(f"ERROR: attempt:{attempt+1} ftp connect failed: {e}")
            if attempt < FTP_MAX_RETRY - 1:
                time.sleep(FTP_RETRY_SLEEP)
    return None

def open_folder(folder_path):
    try:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path, exist_ok=True)
            print(f"not exist folder, created: {folder_path}")
        os.startfile(os.path.normpath(folder_path))
        print(f"open folder: {folder_path}")
    except Exception as e:
        show_custom_error(MSG_TITLE_ERROR, MSG_FAILED_OPEN_FOLDER.format(e))

def find_latest_file_in_folder(folder_path):
    if not folder_path or not os.path.isdir(folder_path):
        return None

    latest_file_path = None
    latest_mtime = -1.0

    for current_root, _, file_names in os.walk(folder_path):
        for file_name in file_names:
            if file_name.startswith("~$"):
                continue
            candidate_path = os.path.join(current_root, file_name)
            try:
                candidate_mtime = os.path.getmtime(candidate_path)
            except OSError:
                continue
            if candidate_mtime >= latest_mtime:
                latest_mtime = candidate_mtime
                latest_file_path = candidate_path

    return latest_file_path

def open_latest_file_in_folder(folder_path):
    latest_file_path = find_latest_file_in_folder(folder_path)
    if latest_file_path is None:
        show_custom_warning(MSG_TITLE_NOT_FOUND, MSG_NO_LOCAL_FILE.format(folder_path))
        return

    try:
        os.startfile(os.path.normpath(latest_file_path))
        print(f"open latest file: {latest_file_path}")
    except Exception as e:
        show_custom_error(MSG_TITLE_ERROR, MSG_FAILED_OPEN_FILE.format(e))

def open_root_folder():
    root_path = get_root_folder_path_or_prompt()
    if root_path is None:
        return
    open_folder(root_path)

def get_root_folder_path():
    if root_folder_textbox:
        path = root_folder_textbox.get()
        if path and path != PLACEHOLDER_TEXT_ROOT_FOLDER and os.path.isdir(path):
            return path
    return None

def get_root_folder_path_or_prompt():
    root_path = get_root_folder_path()
    if root_path is None:
        show_custom_warning(MSG_TITLE_NO_ROOT, MSG_SELECT_ROOT_FOLDER)
        select_root_folder()
        root_path = get_root_folder_path()
    return root_path

def get_local_folder_path(folder_name):
    root_path = get_root_folder_path()
    if root_path is not None:
        return os.path.join(root_path, folder_name)
    return None

def cancel_check_ftp_thread():
    global check_ftp_cancel_event
    if check_ftp_cancel_event:
        check_ftp_cancel_event.set()
        print("Signal sent to cancel check_ftp_file_thread")

def has_file_starting_with(local_folder_path, prefix):
    if os.path.exists(local_folder_path):
        for entry in os.listdir(local_folder_path):
            if entry.startswith(prefix):
                return True
    return False

def valid_spec_format(spec_number_str):
    return re.match(r'^\d{2}\.\d{3}(-\d{1,2}(-\d{1,2})?)?$', spec_number_str)

def set_icon(obj):
    try:
        if obj:
            app_image_path = bundled_data_path(FILE_APP_ICON)
            if app_image_path is not None:
                obj.iconbitmap(app_image_path)
    except Exception as e:
        print(f"ERROR: set_icon failed {e}")

def get_local_folder_status(spec_number_str, folder_name, latest_file_base_name):
    local_folder_path = get_local_folder_path(folder_name)
    if local_folder_path is None:
        return LOCAL_FOLDER_STATUS_EMPTY

    latest_file_exists = has_file_starting_with(local_folder_path, latest_file_base_name)
    if latest_file_exists:
        return LOCAL_FOLDER_STATUS_LATEST

    spec_number_without_dot = spec_number_str.replace('.', '').split('-')[0]
    file_exists = has_file_starting_with(local_folder_path, spec_number_without_dot)
    if file_exists:
        return LOCAL_FOLDER_STATUS_OLD

    return LOCAL_FOLDER_STATUS_EMPTY

def get_local_folder_status_from_list(list_data):
    return get_local_folder_status(list_data[LIST_IDX_SPEC_NUM], list_data[LIST_IDX_FOLDER_NAME], list_data[LIST_IDX_LATEST_FTP_FILE_BASE_NAME])

def load_spec_names():
    global PRELOADED_SPEC_NAMES
    spec_names_file_path = bundled_data_path(FILE_SPEC_NAMES)
    try:
        if spec_names_file_path is not None:
            with open(spec_names_file_path, 'r', encoding='utf-8') as f:
                PRELOADED_SPEC_NAMES = json.load(f)
            print(f"load {len(PRELOADED_SPEC_NAMES)} spec names.")
        else:
            print(f"ERROR: {FILE_SPEC_NAMES} not found.")
            PRELOADED_SPEC_NAMES = {}
    except Exception as e:
        print(f"ERROR: load_spec_names failed to load {FILE_SPEC_NAMES}: {e}")
        PRELOADED_SPEC_NAMES = {}

def fetch_spec_name_from_web(spec_number):
    spec_number_without_dot = spec_number.replace('.', '')
    url = HTTP_DYNA_REPORT.format(spec_number_without_dot)
    spec_name_from_web = None
    print(f"connect: {url}")
    try:
        headers_web = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers_web, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        title_val_element = soup.find(id="titleVal")
        if title_val_element:
            temp_title = title_val_element.get('value')
            if temp_title is None or not temp_title.strip():
                temp_title = title_val_element.text
            if temp_title is not None:
                fetched_title = temp_title.strip()
                if fetched_title:
                    spec_name_from_web = fetched_title
                    print(f"fetch_spec_name_from_web 'titleVal' found for '{spec_number}': '{spec_name_from_web}'")
    except requests.exceptions.HTTPError as e_http:
        print(f"ERROR: DynaReport HTTP error while fetching title for {spec_number}: {e_http}")
    except requests.exceptions.RequestException as e_req:
        print(f"ERROR: DynaReport request failed while fetching title for {spec_number}: {e_req}")
    except Exception as e_parse:
        print(f"ERROR: Error processing DynaReport content for 'titleVal' for {spec_number}: {e_parse}")

    if not spec_name_from_web:
        print(f"fetch_spec_name_from_web failed for {spec_number}.")
    return spec_name_from_web

def _save_list_data_internal(save_list_of_dicts):
    config_folder_path = get_config_folder_path()
    config_list_path = get_config_file_path(SAVE_FILE_LIST)

    try:
        os.makedirs(config_folder_path, exist_ok=True)
        with open(config_list_path, 'w', encoding='utf-8') as f:
            json.dump(save_list_of_dicts, f, indent=4)
        print(f"saved {len(save_list_of_dicts)} items.")
        return True
    except Exception as e:
        print(f"ERROR: _save_list_data_internal failed: {e}")
    return False

def save_config_list():
    if list_view is None:
        return

    save_list = []
    for item_id in list_view.get_children(''):
        tupld_item = list_view.item(item_id, 'values')
        dict_item = {
            KEY_SPEC_NUM: str(tupld_item[LIST_IDX_SPEC_NUM]),
            KEY_FOLDER_NAME: tupld_item[LIST_IDX_FOLDER_NAME],
            KEY_LATEST_FTP_FILE_NAME: tupld_item[LIST_IDX_LATEST_FTP_FILE_BASE_NAME],
            KEY_LAST_FTP_CHECK_TIME: tupld_item[LIST_IDX_LAST_FTP_CHECK_TIME],
            KEY_FAVORITE: True if tupld_item[LIST_IDX_FAVORITE] == ICON_FAVORITE else False
        }
        save_list.append(dict_item)

    _save_list_data_internal(save_list)

def load_list_from_save_config():
    config_list_path = get_config_file_path(SAVE_FILE_LIST)

    if not os.path.exists(config_list_path):
        print(f"SAVE_FILE_LIST not found.")
        return []

    try:
        with open(config_list_path, 'r', encoding='utf-8') as f:
            data_from_file = json.load(f)
    except Exception as e:
        print(f"ERROR: load SAVE_FILE_LIST failed: {e}.")
        return []

    if not data_from_file:  
        print("SAVE_FILE_LIST is empty.")
        return []
    
    treeview_list = []

    for item_dict in data_from_file:
        treeview_list.append([
            ICON_FAVORITE if item_dict.get(KEY_FAVORITE) else ICON_BLANK,
            ICON_CHECK_OFF,
            get_local_folder_status(item_dict.get(KEY_SPEC_NUM), item_dict.get(KEY_FOLDER_NAME), item_dict.get(KEY_LATEST_FTP_FILE_NAME)),
            item_dict.get(KEY_FOLDER_NAME),
            ICON_FOLDER,
            ICON_WEB_LINK,
            item_dict.get(KEY_SPEC_NUM),
            item_dict.get(KEY_LATEST_FTP_FILE_NAME),
            item_dict.get(KEY_LAST_FTP_CHECK_TIME)
        ])

    print(f"load SAVE_FILE_LIST ({len(treeview_list)} items).")
    return treeview_list

def refresh_list_view(initial_list=None):
    all_list = []
    if initial_list is not None:
        all_list = [list(item) for item in initial_list]
    else:
        for item_id in list_view.get_children(''):
            all_list.append(list(list_view.item(item_id, 'values')))

    def get_sort_key(list_item):
        try:
            spec_number = str(list_item[LIST_IDX_SPEC_NUM])
            is_favorite = 0 if list_item[LIST_IDX_FAVORITE] == ICON_FAVORITE else 1

            str_list = spec_number.split('-', 2)
            main_str_float = float(str_list[0])

            if len(str_list) == 1:
                return (is_favorite, main_str_float, 0, 0)
            elif len(str_list) == 2:
                return (is_favorite, main_str_float, int(str_list[1]), 0)
            elif len(str_list) == 3:
                return (is_favorite, main_str_float, int(str_list[1]), int(str_list[2]))
            else:
                print(f"get_sort_key invalid {spec_number}")
                return (1, float('inf'), float('inf'), float('inf'))
        except (ValueError, IndexError, TypeError) as e:
            print(f"ERROR: get_sort_key for {list_item}: {e}")
            return (1, float('inf'), float('inf'), float('inf'))

    all_list.sort(key=get_sort_key)

    for item_id in list_view.get_children(''):
        list_view.delete(item_id)

    for sorted_list_item in all_list:
        list_view.insert('', tk.END, values=sorted_list_item)

    print(f"refresh_list_view ({len(all_list)} items).")

def exit_app():
    global download_cancel_event
    print("exit_app")
    save_config_key_value(CONFIG_KEY_PREV_VER, PROGRAM_VERSION)
    save_config_key_value(CONFIG_KEY_WIN_GEOMETRY, root.winfo_geometry())
    if download_cancel_event and not download_cancel_event.is_set():
        download_cancel_event.set()

    stop_tray_icon()
    stop_activate_event_listener()

    if root: root.destroy()

def on_closing():
    if minimize_to_tray_var and minimize_to_tray_var.get() and pystray:
        start_tray_icon_thread()
        start_activate_event_listener()
        root.withdraw()
    else:
        exit_app()


def stop_tray_icon():
    global tray_icon
    if tray_icon:
        try:
            tray_icon.stop()
        except Exception:
            pass
        tray_icon = None


def activate_main_window():
    stop_activate_event_listener()
    stop_tray_icon()
    if not root:
        return
    try:
        root.deiconify()
        root.state('normal')
    except Exception:
        pass
    try:
        root.lift()
        root.attributes('-topmost', True)
        root.after(120, lambda: root.attributes('-topmost', False))
        root.focus_force()
    except Exception:
        pass


def run_activate_event_listener():
    global activate_event_handle
    event_handle = activate_event_handle
    if not event_handle:
        return

    try:
        kernel32 = ctypes.windll.kernel32
        WAIT_OBJECT_0 = 0x00000000
        WAIT_TIMEOUT = 0x00000102
        while not activate_event_stop_event.is_set():
            wait_result = kernel32.WaitForSingleObject(event_handle, 250)
            if wait_result == WAIT_OBJECT_0:
                try:
                    if root and root.winfo_exists():
                        root.after(0, activate_main_window)
                except Exception:
                    break
                continue
            if wait_result == WAIT_TIMEOUT:
                continue
            break
    except Exception:
        pass


def start_activate_event_listener():
    global activate_event_handle, activate_event_thread
    if activate_event_handle:
        return True
    try:
        kernel32 = ctypes.windll.kernel32
        event_handle = kernel32.CreateEventW(None, False, False, SINGLE_INSTANCE_ACTIVATE_EVENT_NAME)
    except Exception:
        event_handle = None

    if not event_handle:
        return False

    activate_event_handle = event_handle
    activate_event_stop_event.clear()
    activate_event_thread = threading.Thread(target=run_activate_event_listener, daemon=True)
    activate_event_thread.start()
    return True


def stop_activate_event_listener():
    global activate_event_handle, activate_event_thread
    activate_event_stop_event.set()
    event_handle = activate_event_handle
    activate_event_handle = None
    worker = activate_event_thread
    activate_event_thread = None

    if event_handle:
        try:
            ctypes.windll.kernel32.SetEvent(event_handle)
        except Exception:
            pass

    if worker and worker.is_alive():
        worker.join(timeout=0.4)

    if event_handle:
        try:
            ctypes.windll.kernel32.CloseHandle(event_handle)
        except Exception:
            pass

def init_input_spec_number_textbox():
    if input_spec_number_textbox:
        input_spec_number_textbox.delete(0, tk.END)
        input_spec_number_textbox.insert(0, PLACEHOLDER_TEXT_INPUT_SPEC_NUM)
        input_spec_number_textbox.config(foreground=TEXT_COLOR_PLACEHOLDER)

def focus_in_input_spec_number(event):
    if input_spec_number_textbox and input_spec_number_textbox.get() == PLACEHOLDER_TEXT_INPUT_SPEC_NUM:
        input_spec_number_textbox.delete(0, tk.END)
        input_spec_number_textbox.config(foreground=TEXT_COLOR_DEFAULT)

def focus_out_input_spec_number(event):
    if input_spec_number_textbox and not input_spec_number_textbox.get():
        init_input_spec_number_textbox()

def set_root_folder_textbox(path):
    if root_folder_textbox:
        root_folder_textbox.config(state='normal')
        root_folder_textbox.delete(0, tk.END)
        if path == None:
            root_folder_textbox.insert(0, PLACEHOLDER_TEXT_ROOT_FOLDER)
            root_folder_textbox.config(foreground=TEXT_COLOR_PLACEHOLDER)
        else:
            root_folder_textbox.insert(0, path)
            root_folder_textbox.config(foreground=TEXT_COLOR_DEFAULT)
        root_folder_textbox.config(state='readonly')

def sync_checkbox_header():
    total_items, num_checked = 0, 0
    for item_id in list_view.get_children(''):
        tuple_item = list_view.item(item_id, 'values')
        total_items += 1
        if tuple_item[LIST_IDX_CHECKBOX] == ICON_CHECK_ON:
            num_checked += 1

    checked = total_items > 0 and num_checked == total_items
    set_checkbox_header(checked)

def set_checkbox_header(checked):
    if checkbox_header is not None:
        checkbox_header.set(checked)
    if list_view is not None:
        list_view.heading('select', text=ICON_CHECK_ON if checked else ICON_CHECK_OFF)

def set_all_list_checkbox(checked):
    for item_id in list_view.get_children(''):
        list_item = list(list_view.item(item_id, 'values'))
        list_item[LIST_IDX_CHECKBOX] = ICON_CHECK_ON if checked else ICON_CHECK_OFF
        list_view.item(item_id, values=list_item)

def set_all_checkbox(checked):
    set_checkbox_header(checked)
    set_all_list_checkbox(checked)

def on_checkbox_header_click():
    if checkbox_header is None: return

    checked = not checkbox_header.get()
    set_all_checkbox(checked)

def select_root_folder():
    current_root_folder = get_root_folder_path()
    if current_root_folder is None:
        current_root_folder = os.path.expanduser("~")

    selected_root_folder = filedialog.askdirectory(initialdir=current_root_folder)
    if selected_root_folder:
        print(f"select_root_folder: {selected_root_folder}")
        set_root_folder_textbox(selected_root_folder)
        save_config_key_value(CONFIG_KEY_ROOT_FOLDER, selected_root_folder)
        check_and_update_list()

def fetch_latest_ftp_file_name(ftp, spec_number_str):
    series_number = spec_number_str[:2]
    series_path = FTP_SERIES_PATH.format(series_number)
    ftp.cwd(series_path)

    items = ftp.nlst()
    if spec_number_str in items:
        ftp.cwd(spec_number_str)
    else:
        print(f"no {spec_number_str} folder in {series_path}.")
        return None

    zip_files = sorted([f for f in ftp.nlst() if f.lower().endswith('.zip')])
    if zip_files:
        latest_ftp_file_name = zip_files[-1]
        print(f"latest_ftp_file_name: {latest_ftp_file_name}")
        return latest_ftp_file_name
    else:
        print(f"no {spec_number_str} files in ftp folder.")
        return None

def get_spec_number_from_textbox():
    text = input_spec_number_textbox.get().strip()
    if text == PLACEHOLDER_TEXT_INPUT_SPEC_NUM or text == "":
        return None
    return text

def add_item():
    if len(list_view.get_children('')) >= MAX_LIST_ITEMS:
        show_custom_warning(MSG_TITLE_ADD_LIMIT, MSG_ADD_LIMIT)
        return

    spec_number = get_spec_number_from_textbox()
    if spec_number is None:
        print("add_item but no spec_number")
        return

    if not valid_spec_format(spec_number):
        show_custom_warning(MSG_TITLE_INVALID_SPEC, MSG_INVALID_SPEC)
        return

    for item_id in list_view.get_children(''):
        tuple_item = list_view.item(item_id, 'values')
        if tuple_item[LIST_IDX_SPEC_NUM] == spec_number:
            show_custom_warning(MSG_TITLE_DUPLICATE_SPEC, MSG_DUPLICATE_SPEC.format(spec_number))
            return

    cancel_check_ftp_thread()

    latest_ftp_file_name = None

    ftp = get_ftp()
    if ftp is None:
        show_custom_error(MSG_TITLE_ERROR, MSG_FTP_CONNECTION_FAILED)
        return

    with ftp:
        latest_ftp_file_name = fetch_latest_ftp_file_name(ftp, spec_number)

    if latest_ftp_file_name is None:
        show_custom_warning(MSG_TITLE_NOT_FOUND, MSG_SPEC_NOT_FOUND.format(spec_number))
        return

    spec_name = None
    if spec_number in PRELOADED_SPEC_NAMES:
        spec_name = PRELOADED_SPEC_NAMES[spec_number]
        print(f"found '{spec_name}' for '{spec_number}' in PRELOADED_SPEC_NAMES")

    if spec_name is None:
        print(f"not found '{spec_number}' in PRELOADED_SPEC_NAMES")
        spec_name = fetch_spec_name_from_web(spec_number)

    folder_name = f"{spec_number} {spec_name}" if spec_name else spec_number
    folder_name = sanitize_folder_name(folder_name)
    folder_name = shorten_folder_name(folder_name)

    latest_ftp_file_base_name = os.path.splitext(latest_ftp_file_name)[0]
    last_ftp_file_check_time = datetime.datetime.now().isoformat()

    new_item_values = [
        DEFAULT_FAVORITE,
        ICON_CHECK_OFF,
        get_local_folder_status(spec_number, folder_name, latest_ftp_file_base_name),
        folder_name,
        ICON_FOLDER,
        ICON_WEB_LINK,
        spec_number,
        latest_ftp_file_base_name,
        last_ftp_file_check_time
    ]

    list_view.insert('', tk.END, values=new_item_values)
    print(f"'{spec_number}' added to list.")

    refresh_list_view()
    save_config_list()

    init_input_spec_number_textbox()
    if add_button:
        add_button.focus_set()

    check_and_update_list()

def set_button_enabled(enabled):
    buttons = [add_button, download_button, remove_button, favorite_button, unfavorite_button, browse_specs_button, settings_button]
    for button in buttons:
        if button: 
            button.config(state='normal' if enabled else 'disabled')

def _process_gui_updates():
    global update_queue, root, download_progress_window, list_view, last_update_ftp_fail_time
    try:
        while True:
            try:
                message = update_queue.get_nowait()
            except queue.Empty:
                break

            msg_type = message.get('type')
            can_update_progress_window = download_progress_window and download_progress_window.winfo_exists()

            if msg_type == 'item_progress':
                if can_update_progress_window and hasattr(download_progress_window, 'item_label'):
                    download_progress_window.item_label.config(text=message.get('text', ''))
            elif msg_type == 'percent_progress':
                if can_update_progress_window and hasattr(download_progress_window, 'percent_label'):
                    download_progress_window.percent_label.config(text=message.get('text', ''))
            
            elif msg_type == 'update_list_item':
                item_id = message.get('item_id')
                list_item = message.get('values')
                print(f"update_list_item item_id: {item_id}, values: {list_item}")
                if list_view.exists(item_id):
                    list_view.item(item_id, values=list_item)

            elif msg_type == 'ftp_log': 
                print(f"{message.get('level', 'INFO')} [{message.get('spec','N/A')} THREAD]: {message.get('message','')}")

            elif msg_type == 'failed_download':
                if download_progress_window and download_progress_window.winfo_exists():
                    download_progress_window.destroy()
                download_progress_window = None

                set_button_enabled(True)
                set_all_checkbox(False)

                show_custom_error(MSG_TITLE_ERROR, MSG_FTP_CONNECTION_FAILED_WEB)

                return

            elif msg_type == 'finished_download' or msg_type == 'canceled_download':
                summary_list = message.get('summary_list', [])

                if download_progress_window and download_progress_window.winfo_exists():
                    download_progress_window.destroy()
                download_progress_window = None

                set_button_enabled(True)
                set_all_checkbox(False)
                check_and_update_list()

                task_results = "\n".join(summary_list)
                if not task_results:
                    task_results = MSG_NO_ITEMS_PROCESSED

                title = MSG_TITLE_OPERATION_STATUS
                result_message = task_results

                if msg_type == 'canceled_download':
                    title = MSG_TITLE_CANCELED
                    result_message = MSG_CANCELED_DOWNLOAD.format(task_results)
                else:
                    success_count = message.get('success_count', 0)
                    title = MSG_TITLE_COMPLETED
                    result_message = MSG_COMPLETED_DOWNLOAD.format(success_count, len(summary_list) if summary_list else 0, task_results)
                
                if root and root.winfo_exists():
                    show_custom_info(title, result_message)

                print(f"msg_type: {msg_type} _process_gui_updates stop")
                return

            elif msg_type == 'finished_check_ftp_latest_file' or msg_type == 'failed_check_ftp_latest_file':
                set_button_enabled(True)
                save_config_list()

                if msg_type == 'failed_check_ftp_latest_file':
                    last_update_ftp_fail_time = datetime.datetime.now().isoformat()

                print(f"msg_type: {msg_type} _process_gui_updates stop")
                return

            if can_update_progress_window and msg_type not in ['finished_download', 'canceled_download', 'failed_download', 'finished_check_ftp_latest_file', 'failed_check_ftp_latest_file']:
                try:
                    download_progress_window.update_idletasks()
                except tk.TclError:
                    pass

    except Exception as e:
        print(f"ERROR: _process_gui_updates stop: {e}")
        if download_progress_window and download_progress_window.winfo_exists():
            download_progress_window.destroy()
            download_progress_window = None
        set_button_enabled(True)
        return

    if root and root.winfo_exists():
        root.after(100, _process_gui_updates)

def _check_ftp_latest_file_in_thread(ftp_check_list, q, cancel_event):
    if not ftp_check_list:
        q.put({'type': 'finished_check_ftp_latest_file'})
        return
    
    failed = False

    if cancel_event.is_set():
        q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': 'StaleCheck', 'message': "Check cancelled."})
        q.put({'type': 'finished_check_ftp_latest_file'})
        return

    ftp = get_ftp()
    if ftp is None:
        q.put({'type': 'failed_check_ftp_latest_file'})
        return

    with ftp:
        q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': 'StaleCheck', 'message': "FTP connected for stale check."})

        for item_info in ftp_check_list:
            if cancel_event.is_set():
                q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': 'StaleCheck', 'message': "Check cancelled by user action."})
                break

            list_item = item_info["values"]
            spec_number = list_item[LIST_IDX_SPEC_NUM]
            try:
                q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': list_item[LIST_IDX_SPEC_NUM], 'message': f"FTP CWD for stale check."})

                latest_ftp_file_name = fetch_latest_ftp_file_name(ftp, spec_number)
                latest_ftp_file_base_name = os.path.splitext(latest_ftp_file_name)[0]

                if latest_ftp_file_base_name != list_item[LIST_IDX_LATEST_FTP_FILE_BASE_NAME]:
                    q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 
                            'message': f"latest_ftp_file_base_name changed: '{list_item[LIST_IDX_LATEST_FTP_FILE_BASE_NAME]}' -> '{latest_ftp_file_base_name}'"})
                    
                list_item[LIST_IDX_STATUS] = get_local_folder_status(spec_number, list_item[LIST_IDX_FOLDER_NAME], latest_ftp_file_base_name)
                list_item[LIST_IDX_LATEST_FTP_FILE_BASE_NAME] = latest_ftp_file_base_name
                list_item[LIST_IDX_LAST_FTP_CHECK_TIME] = datetime.datetime.now().isoformat()

                q.put({
                    'type': 'update_list_item',
                    'item_id': item_info["item_id"],
                    'values': list_item 
                })

            except ftplib.all_errors as e_series_ftp:
                failed = True
                q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': f"FTP error processing '{spec_number}' for stale check: {e_series_ftp}"})
            except Exception as e_series_general:
                failed = True
                q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': f"General error processing '{spec_number}' for stale check: {e_series_general}"})

    if failed:
        q.put({'type': 'failed_check_ftp_latest_file'})
    else:
        q.put({'type': 'finished_check_ftp_latest_file'})

def check_and_update_list():
    global update_queue, list_view, check_ftp_file_thread, check_ftp_cancel_event

    root_path = get_root_folder_path()
    if root_path is None:
        print("check_and_update_list: no root folder")
        return

    print("check_and_update_list")

    ftp_check_list = []
    for item_id in list_view.get_children(''):
        list_item = list(list_view.item(item_id, 'values'))
        status = get_local_folder_status_from_list(list_item)

        last_ftp_check_time = list_item[LIST_IDX_LAST_FTP_CHECK_TIME]
        is_stale_item = (datetime.datetime.now() - datetime.datetime.fromisoformat(last_ftp_check_time)).days > 5

        if status == LOCAL_FOLDER_STATUS_LATEST and is_stale_item:
            ftp_check_list.append({
                "item_id": item_id,
                "values": list_item
            })
        else:
            list_item[LIST_IDX_STATUS] = status
            list_view.item(item_id, values=list_item)

    if not ftp_check_list:
        print("check_and_update_list: No ftp_check_list")
        return

    if check_ftp_file_thread and check_ftp_file_thread.is_alive():
        print("check_and_update_list: check_ftp_file_thread is already in progress.")
        return
    
    if last_update_ftp_fail_time and ((datetime.datetime.now() - datetime.datetime.fromisoformat(last_update_ftp_fail_time)).total_seconds() < 3600):
        print(f"check_and_update_list: skip last_update_ftp_fail_time:{last_update_ftp_fail_time}")
        return

    print(f"_check_ftp_latest_file_in_thread for {len(ftp_check_list)} items.")
    check_ftp_cancel_event = threading.Event()
    check_ftp_file_thread = threading.Thread(
        target=_check_ftp_latest_file_in_thread,
        args=(ftp_check_list, update_queue, check_ftp_cancel_event),
        daemon=True
    )
    check_ftp_file_thread.start()
    _process_gui_updates()

def _download_ftp_file_in_thread(selected_items, root_path, q, cancel_event):
    results_summary_list = []
    download_success_count = 0
    cancelled_by_user = False

    if len(selected_items) == 0:
        q.put({'type': 'ftp_log', 'level': 'INFO', 'message': "No items selected for download."})
        return results_summary_list

    ftp = get_ftp()
    if ftp is None:
        q.put({'type': 'failed_download', 'summary_list': results_summary_list, 'success_count': download_success_count})
        return

    try:
        for idx, item_data in enumerate(selected_items):
            item_id = item_data["item_id"]
            list_item = item_data["values"]
            spec_number = list_item[LIST_IDX_SPEC_NUM]
            folder_name = list_item[LIST_IDX_FOLDER_NAME]

            if cancelled_by_user or cancel_event.is_set():
                cancelled_by_user = True
                results_summary_list.append(f"{spec_number}: Canceled")
                continue

            q.put({'type': 'item_progress', 'text': f"Processing: {spec_number} ({idx+1}/{len(selected_items)})"})
            q.put({'type': 'percent_progress', 'text': "Connecting..."})

            local_folder_path = os.path.join(root_path, folder_name)
            operational_outcome_english = "Error (Unknown)"

            temp_dir_for_item = None

            try:
                q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': "Download process for item started..."})
                os.makedirs(local_folder_path, exist_ok=True)

                if cancel_event.is_set():
                    raise ConnectionAbortedError("User canceled before listing spec folder.")

                latest_ftp_file_name = fetch_latest_ftp_file_name(ftp, spec_number)
                latest_ftp_file_base_name = os.path.splitext(latest_ftp_file_name)[0]
                list_item[LIST_IDX_LATEST_FTP_FILE_BASE_NAME] = latest_ftp_file_base_name
                list_item[LIST_IDX_LAST_FTP_CHECK_TIME] = datetime.datetime.now().isoformat()
                ftp_mod_time = ""
                try:
                    mdtm_response = ftp.sendcmd(f'MDTM {latest_ftp_file_name}')
                    if mdtm_response.startswith('213 '): ftp_mod_time = mdtm_response[4:].strip()
                except Exception as e_mdtm:
                    q.put({'type': 'ftp_log', 'level': 'WARN', 'spec': spec_number, 'message': f"Failed to get FTP file mod time: {e_mdtm}"})

                ftp_mod_time_yymmdd = format_ftp_time_to_yymmdd(ftp_mod_time)
                latest_file_exists = has_file_starting_with(local_folder_path, latest_ftp_file_base_name)

                if latest_file_exists:
                    operational_outcome_english = "Skipped (latest version already present)"
                    q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': operational_outcome_english})
                    download_success_count += 1
                else:
                    temp_dir_for_item = tempfile.mkdtemp(prefix=f"spec_dl_{spec_number.replace('.', '_')}_")
                    local_zip_path = os.path.join(temp_dir_for_item, latest_ftp_file_name)
                    total_file_size = 0
                    try:
                        ftp.voidcmd('TYPE I')
                        size_response = ftp.sendcmd(f'SIZE {latest_ftp_file_name}')
                        if size_response.startswith('213 '):
                            total_file_size = int(size_response[4:])
                    except Exception as e_size:
                        q.put({'type': 'ftp_log', 'level': 'WARN', 'spec': spec_number, 'message': f"Could not get file size: {e_size}"})

                    download_state = {'downloaded_bytes': 0, 'last_reported_percent': -1, 'total_size': total_file_size}
                    q.put({'type': 'percent_progress', 'text': "Downloading: 0%"})

                    with open(local_zip_path, 'wb') as local_file_handle:
                        def download_callback(data_block):
                            if cancel_event.is_set():
                                raise ConnectionAbortedError("User canceled during file transfer.")

                            download_state['downloaded_bytes'] += len(data_block)
                            local_file_handle.write(data_block)

                            current_byte = download_state['downloaded_bytes']
                            total_size = download_state['total_size']
                            last_reported_percent = download_state['last_reported_percent']
                            percent_text_to_send = ""

                            if total_size > 0:
                                percent = int((current_byte / total_size) * 100)
                                if percent > last_reported_percent or (current_byte == total_size and last_reported_percent < 100):
                                    percent_text_to_send = f"Downloading: {percent}%"
                                    download_state['last_reported_percent'] = percent
                            else:
                                kb_downloaded = current_byte // 1024
                                if kb_downloaded > last_reported_percent :
                                    percent_text_to_send = f"Downloading: {kb_downloaded} KB"
                                    download_state['last_reported_percent'] = kb_downloaded
                            
                            if percent_text_to_send:
                                q.put({'type': 'percent_progress', 'text': percent_text_to_send})

                        ftp.retrbinary(f'RETR {latest_ftp_file_name}', download_callback, blocksize=FTP_DOWNLOAD_BLOCKSIZE)

                    q.put({'type': 'percent_progress', 'text': "Download: 100%"})
                    q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': f"Zip file downloaded: {local_zip_path}"})

                    extracted_files_path = os.path.join(temp_dir_for_item, "extracted_content")
                    os.makedirs(extracted_files_path, exist_ok=True)
                    with zipfile.ZipFile(local_zip_path, 'r') as zip_ref:
                        zip_ref.extractall(extracted_files_path)
                    q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': f"Zip file extracted: {extracted_files_path}"})

                    top_level_extracted_items = os.listdir(extracted_files_path)
                    num_top_level_items = len(top_level_extracted_items)

                    if num_top_level_items == 0:
                        operational_outcome_english = f"Failed (Downloaded zip is empty)"
                        q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': f"{operational_outcome_english}: {e}"})
                        
                    else:
                        dst_path = ""
                        final_name = latest_ftp_file_base_name

                        if '-' in latest_ftp_file_base_name:
                            parts = latest_ftp_file_base_name.split('-')
                            version_code = parts[-1] if len(parts) > 1 else None

                        version_postfix = None
                        if version_code is not None:
                            version_postfix = get_version_postfix_from_code(version_code)

                        name_parts = [latest_ftp_file_base_name]
                        if version_postfix:
                            name_parts.append(version_postfix)
                        if ftp_mod_time_yymmdd != "": 
                            name_parts.append(ftp_mod_time_yymmdd)

                        final_name = "_".join(name_parts)

                        if num_top_level_items == 1:
                            item_name = top_level_extracted_items[0]
                            src_path = os.path.join(extracted_files_path, item_name)

                            try:
                                if os.path.isfile(src_path):
                                    file_ext = os.path.splitext(item_name)[1]
                                    dst_path = os.path.join(local_folder_path, f"{final_name}{file_ext}")
                                else:
                                    dst_path = os.path.join(local_folder_path, item_name)

                                copy_file_or_folder(src_path, dst_path)
                                q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': f"Single item: {dst_path}"})
                                operational_outcome_english = "Download complete"
                                download_success_count += 1
                            except Exception as e:
                                operational_outcome_english = f"Failed (Copy error: {type(e).__name__})"
                                q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': f"{operational_outcome_english}: {e}"})

                        elif num_top_level_items >= 2:
                            dst_folder_path = os.path.join(local_folder_path, final_name)
                            q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': f"Multiple items in zip. Creating subfolder: {dst_folder_path}"})
                            try:
                                os.makedirs(dst_folder_path, exist_ok=True)
                                for item_in_extracted in top_level_extracted_items:
                                    src_path = os.path.join(extracted_files_path, item_in_extracted)
                                    dst_path = os.path.join(dst_folder_path, item_in_extracted)
                                    copy_file_or_folder(src_path, dst_path)

                                q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': f"All {num_top_level_items} extracted items copied to subfolder '{dst_folder_path}'."})
                                operational_outcome_english = "Download complete"
                                download_success_count += 1
                            except ConnectionAbortedError as cae_multi:
                                operational_outcome_english = "Canceled"
                                q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': str(cae_multi)})
                            except Exception as e:
                                operational_outcome_english = f"Failed (Error copying to subfolder: {type(e).__name__})"
                                q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': f"{operational_outcome_english}: {e}"})

            except ConnectionAbortedError as e_aborted:
                operational_outcome_english = "Canceled"
                q.put({'type': 'ftp_log', 'level': 'INFO', 'spec': spec_number, 'message': f"{operational_outcome_english}: {e_aborted}"})
                cancelled_by_user = True 
            except ftplib.all_errors as e_ftp:
                operational_outcome_english = f"Failed (FTP error: {type(e_ftp).__name__})"
                q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': f"FTP operation: {e_ftp}"})
            except zipfile.BadZipFile:
                operational_outcome_english = "Failed (Invalid Zip file)"
                q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': "Zip processing: BadZipFile"})
            except Exception as e_download:
                operational_outcome_english = f"Failed (General Error: {type(e_download).__name__})"
                q.put({'type': 'ftp_log', 'level': 'ERROR', 'spec': spec_number, 'message': f"Download processing: {e_download}"})
            finally:
                delete_file_or_folder(temp_dir_for_item)
                results_summary_list.append(f"{spec_number}: {operational_outcome_english}")
                list_item[LIST_IDX_STATUS] = get_local_folder_status(spec_number, folder_name, latest_ftp_file_base_name)

                q.put({'type': 'update_list_item', 'item_id': item_id, 'values': list_item})
    finally:
        if cancelled_by_user:
            try:
                ftp.close()
                q.put({'type': 'ftp_log', 'level': 'INFO', 'message': "ftp.close after download cancel."})
            except Exception as e:
                q.put({'type': 'ftp_log', 'level': 'ERROR', 'message': f"ftp.close error after downloadcancel: {e}"})
                pass
            q.put({'type': 'canceled_download', 'summary_list': results_summary_list})
        else:
            try:
                ftp.quit()
                q.put({'type': 'ftp_log', 'level': 'INFO', 'message': "ftp.quit after download."})
            except ftplib.all_errors as e:
                q.put({'type': 'ftp_log', 'level': 'ERROR', 'message': f"ftp.quit error after download: {e}"})
            q.put({'type': 'finished_download', 'summary_list': results_summary_list, 'success_count': download_success_count})    

def on_download_progress_window_close():
    global download_cancel_event
    if download_cancel_event:
        download_cancel_event.set()
    print("download_cancel_event by user")

def create_download_progress_window():
    global download_progress_window
    
    if download_progress_window and download_progress_window.winfo_exists():
        download_progress_window.destroy()
    download_progress_window = None

    download_progress_window = tk.Toplevel(root)
    download_progress_window.title("Download Progress")
    download_progress_window.transient(root)
    download_progress_window.resizable(False, False)
    win_width, win_height = scale_pixels(350), scale_pixels(120)
    x = root.winfo_x() + (root.winfo_width() - win_width) // 2
    y = root.winfo_y() + (root.winfo_height() - win_height) // 2
    download_progress_window.geometry(f'{win_width}x{win_height}+{x}+{y}')
    set_icon(download_progress_window)

    download_progress_window.info_label = ttk.Label(download_progress_window, text="Download documents...")
    download_progress_window.info_label.pack(pady=(scale_pixels(10),0))
    download_progress_window.item_label = ttk.Label(download_progress_window, text="Initializing...")
    download_progress_window.item_label.pack(pady=scale_pixels(5), padx=scale_pixels(10), fill='x')
    download_progress_window.percent_label = ttk.Label(download_progress_window, text="")
    download_progress_window.percent_label.pack(pady=(0,scale_pixels(10)), padx=scale_pixels(10), fill='x')
    download_progress_window.update()
    download_progress_window.protocol("WM_DELETE_WINDOW", on_download_progress_window_close)

def _center_dialog_to_root(dialog):
    dialog.update_idletasks()
    dialog_width = dialog.winfo_reqwidth()
    dialog_height = dialog.winfo_reqheight()
    x = root.winfo_rootx() + (root.winfo_width() - dialog_width) // 2
    y = root.winfo_rooty() + (root.winfo_height() - dialog_height) // 2
    dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

def show_custom_message_dialog(title, message, level="info", ok_text="OK"):
    if root is None:
        print(f"{title}: {message}")
        return

    style_map = {
        "info": "primary",
        "warning": "warning",
        "error": "warning",
    }
    button_style = style_map.get(level, "secondary")

    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.transient(root)
    dialog.resizable(False, False)
    dialog.grab_set()
    set_icon(dialog)

    frame = ttk.Frame(dialog, padding=scale_pixels(12))
    frame.pack(fill='both', expand=True)

    message_label = ttk.Label(frame, text=message, justify='left', wraplength=scale_pixels(360))
    message_label.pack(fill='x')

    button_frame = ttk.Frame(frame)
    button_frame.pack(pady=(scale_pixels(12), 0), anchor='center')

    def on_close(event=None):
        dialog.destroy()

    ok_button = ttk.Button(button_frame, text=ok_text, command=on_close, width=10, bootstyle=button_style)
    ok_button.pack()

    dialog.protocol("WM_DELETE_WINDOW", on_close)
    dialog.bind("<Return>", on_close)
    dialog.bind("<Escape>", on_close)

    _center_dialog_to_root(dialog)
    ok_button.focus_set()
    dialog.wait_window()

def show_custom_warning(title, message):
    show_custom_message_dialog(title, message, level="warning")

def show_custom_error(title, message):
    show_custom_message_dialog(title, message, level="error")

def show_custom_info(title, message):
    show_custom_message_dialog(title, message, level="info")

def show_custom_confirm_dialog(title, message, yes_text="Yes", no_text="No"):
    if root is None:
        return False

    dialog_result = {'confirmed': False}
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.transient(root)
    dialog.resizable(False, False)
    dialog.grab_set()
    set_icon(dialog)

    frame = ttk.Frame(dialog, padding=scale_pixels(12))
    frame.pack(fill='both', expand=True)

    message_label = ttk.Label(frame, text=message, justify='left', wraplength=scale_pixels(320))
    message_label.pack(fill='x')

    button_frame = ttk.Frame(frame)
    button_frame.pack(pady=(scale_pixels(12), 0), anchor='center')

    def on_yes(event=None):
        dialog_result['confirmed'] = True
        dialog.destroy()

    def on_no(event=None):
        dialog.destroy()

    yes_button = ttk.Button(button_frame, text=yes_text, command=on_yes, width=10)
    yes_button.pack(side='left', padx=(0, scale_pixels(8)))
    no_button = ttk.Button(button_frame, text=no_text, command=on_no, width=10, bootstyle="secondary")
    no_button.pack(side='left')

    dialog.protocol("WM_DELETE_WINDOW", on_no)
    dialog.bind("<Return>", on_yes)
    dialog.bind("<Escape>", on_no)

    _center_dialog_to_root(dialog)
    yes_button.focus_set()

    dialog.wait_window()
    return dialog_result['confirmed']

def download_selected_item():
    global update_queue, download_progress_window, download_cancel_event

    selected_items = []
    if list_view is None: return
    for item_id in list_view.get_children(''):
        tuple_item = list_view.item(item_id, 'values')
        if tuple_item[LIST_IDX_CHECKBOX] == ICON_CHECK_ON:
            selected_items.append({
                "item_id": item_id,
                "values": list(tuple_item)
            })

    if not selected_items:
        show_tooltip(download_button, TEXT_TOOLTIP_SELECTION_REQUIRED)
        return

    root_path = get_root_folder_path_or_prompt()
    if root_path is None:
        return

    cancel_check_ftp_thread()

    create_download_progress_window()
    set_button_enabled(False)

    download_cancel_event = threading.Event()
    download_ftp_file_thread = threading.Thread(target=_download_ftp_file_in_thread,
                                       args=(selected_items, root_path, update_queue, download_cancel_event),
                                       daemon=True)
    download_ftp_file_thread.start()
    _process_gui_updates()

def remove_selected_item():
    selected_items = []
    if list_view is None: return
    for item_id in list_view.get_children(''):
        tuple_item = list_view.item(item_id, 'values')
        if tuple_item[LIST_IDX_CHECKBOX] == ICON_CHECK_ON:
            selected_items.append(item_id)

    if not selected_items:
        show_tooltip(remove_button, TEXT_TOOLTIP_SELECTION_REQUIRED)
        return

    if show_custom_confirm_dialog(MSG_TITLE_CONFIRM_REMOVE, MSG_CONFIRM_REMOVE):
        cancel_check_ftp_thread()
        print(f"remove items: {selected_items}")
        for item_id in selected_items:
            if list_view.exists(item_id):
                list_view.delete(item_id)

        set_checkbox_header(False)
        save_config_list()
    
    check_and_update_list()

def favorite_selected_item(isFavorite, widget=None):
    if list_view is None: return
    selected_item_count = 0
    for item_id in list_view.get_children(''):
        list_item = list(list_view.item(item_id, 'values'))
        if list_item[LIST_IDX_CHECKBOX] == ICON_CHECK_ON:
            list_item[LIST_IDX_FAVORITE] = ICON_FAVORITE if isFavorite else ICON_BLANK
            list_item[LIST_IDX_CHECKBOX] = ICON_CHECK_OFF
            list_view.item(item_id, values=list_item)
            selected_item_count += 1

    print(f"favorite_selected_item {selected_item_count} items set to {isFavorite}")
    if selected_item_count:
        cancel_check_ftp_thread()
        refresh_list_view()
        set_checkbox_header(False)
        save_config_list()
    else:
        show_tooltip(widget, TEXT_TOOLTIP_SELECTION_REQUIRED)

def set_favorite_selected_item():
    print(f"set_favorite_selected_item")
    favorite_selected_item(True, favorite_button)

def unset_favorite_selected_item():
    print(f"unset_favorite_selected_item")
    favorite_selected_item(False, unfavorite_button)

def on_folder_name_edit_inline(item_id, column_id_str):
    bbox = list_view.bbox(item_id, column_id_str)
    if not bbox:
        return
    x, y, width, height = bbox

    list_item = list(list_view.item(item_id, 'values'))
    old_folder_name = list_item[LIST_IDX_FOLDER_NAME]
    spec_number = list_item[LIST_IDX_SPEC_NUM]

    if old_folder_name.startswith(spec_number):
        default_edit_folder_name = old_folder_name[len(spec_number)+1:]
    else:
        default_edit_folder_name = old_folder_name

    def check_folder_name(name):
        no_invalid_char = not has_invalid_folder_char(name)
        is_valid_folder_name_length = len(f"{spec_number} {name}") <= MAX_FOLDER_NAME
        return no_invalid_char and is_valid_folder_name_length

    vcmd = (list_view.register(check_folder_name), '%P')
    entry = tk.Entry(list_view, validate='key', validatecommand=vcmd)
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, default_edit_folder_name)
    entry.focus_set()

    def submit_folder_name_input(event=None):
        input_folder_name = entry.get().strip()
        if input_folder_name != default_edit_folder_name:
            new_folder_name = f"{spec_number} {input_folder_name}"
            new_folder_name = new_folder_name.strip().rstrip('. ')
            list_item[LIST_IDX_FOLDER_NAME] = new_folder_name
            list_item[LIST_IDX_STATUS] = get_local_folder_status(spec_number, new_folder_name, list_item[LIST_IDX_LATEST_FTP_FILE_BASE_NAME])
            list_view.item(item_id, values=list_item)
            save_config_list()
        entry.destroy()
        set_button_enabled(True)

        check_and_update_list()

    def cancel_folder_name_input(event=None):
        entry.destroy()
        set_button_enabled(True)

    entry.bind("<Return>", submit_folder_name_input)
    entry.bind("<FocusOut>", cancel_folder_name_input)
    entry.bind("<Escape>", cancel_folder_name_input)

def on_list_item_click(event):
    if list_view is None: return
    region = list_view.identify_region(event.x, event.y)
    if region != "cell": return

    item_id = list_view.identify_row(event.y)
    column_id_str = list_view.identify_column(event.x)

    if not item_id or not column_id_str: return

    try:
        clicked_column_index = int(column_id_str.replace('#', '')) - 1

        if not list_view.exists(item_id): return
        list_item = list(list_view.item(item_id, 'values'))

        if clicked_column_index == LIST_IDX_CHECKBOX:
            list_item[LIST_IDX_CHECKBOX] = ICON_CHECK_ON if list_item[LIST_IDX_CHECKBOX] == ICON_CHECK_OFF else ICON_CHECK_OFF
            list_view.item(item_id, values=list_item)
            print(f"{list_item[LIST_IDX_SPEC_NUM]}({item_id}) checkbox state: {list_item[LIST_IDX_CHECKBOX]}")
            sync_checkbox_header()

        elif clicked_column_index == LIST_IDX_FOLDER_ACTION:
            print(f"{list_item[LIST_IDX_SPEC_NUM]}({item_id}) click folder")
            root_path = get_root_folder_path_or_prompt()
            if root_path:
                folder_path = os.path.join(root_path, list_item[LIST_IDX_FOLDER_NAME])
                open_folder(folder_path)
            
            check_and_update_list()

        elif clicked_column_index == LIST_IDX_LINK_ACTION:
            print(f"{list_item[LIST_IDX_SPEC_NUM]}({item_id}) click link")
            spec_number = str(list_item[LIST_IDX_SPEC_NUM])
            url_to_open = HTTP_DYNA_REPORT.format(spec_number.replace('.', ''))
            try:
                webbrowser.open_new_tab(url_to_open)
                print(f"open: {url_to_open}")
            except Exception as e:
                show_custom_error(MSG_TITLE_ERROR, MSG_FAILED_OPEN_BROWSER.format(e))

            check_and_update_list()

    except Exception as e:
        print(f"ERROR: on_list_item_click: {e}")

def on_list_item_double_click(event):
    if list_view is None: return
    region = list_view.identify_region(event.x, event.y)
    if region != "cell": return

    item_id = list_view.identify_row(event.y)
    column_id_str = list_view.identify_column(event.x)
    if not item_id or not column_id_str or not list_view.exists(item_id):
        return

    try:
        clicked_column_index = int(column_id_str.replace('#', '')) - 1
        if clicked_column_index != LIST_IDX_FOLDER_NAME:
            return

        list_item = list(list_view.item(item_id, 'values'))
        root_path = get_root_folder_path_or_prompt()
        if root_path is None:
            return
        
        folder_path = os.path.join(root_path, list_item[LIST_IDX_FOLDER_NAME])
        open_latest_file_in_folder(folder_path)
        check_and_update_list()
    except Exception as e:
        print(f"ERROR: on_list_item_double_click: {e}")

def on_list_item_right_click(event):
    if list_view is None: return
    region = list_view.identify_region(event.x, event.y)
    if region != "cell": return

    item_id = list_view.identify_row(event.y)
    column_id_str = list_view.identify_column(event.x)
    if not item_id or not column_id_str or not list_view.exists(item_id):
        return

    try:
        clicked_column_index = int(column_id_str.replace('#', '')) - 1
        if clicked_column_index != LIST_IDX_FOLDER_NAME:
            return

        list_view.selection_set(item_id)
        list_view.focus(item_id)
        on_folder_name_edit_inline(item_id, column_id_str)
    except Exception as e:
        print(f"ERROR: on_list_item_right_click: {e}")

def on_browse_specs():
    dialog = tk.Toplevel(root)
    dialog.withdraw()
    dialog.title(TEXT_BROWSE_SPECS)
    dialog.transient(root)
    dialog.grab_set()
    dialog.resizable(True, False)
    dialog.columnconfigure(0, weight=1)
    set_icon(dialog)

    series_cb = SeriesComboBox(dialog, SERIES_LIST)
    series_cb.grid(row=0, column=0, padx=scale_pixels(15), pady=(scale_pixels(15), scale_pixels(5)), sticky="ew")

    spec_cb = SpecComboBox(dialog, PRELOADED_SPEC_NAMES)
    spec_cb.grid(row=1, column=0, padx=scale_pixels(15), pady=(scale_pixels(5), scale_pixels(15)), sticky="ew")

    def on_series_change(selected_series):
        spec_cb.update_by_series(selected_series)
    series_cb.on_series_change = on_series_change
    spec_cb.update_by_series(series_cb.get())

    btn_frame = ttk.Frame(dialog)
    btn_frame.grid(row=2, column=0, columnspan=2, pady=(0, scale_pixels(15)))

    def on_select():
        selected_index = getattr(spec_cb, 'selected_index', None)
        if selected_index is not None and selected_index < len(spec_cb.filtered_keys):
            selected_spec = spec_cb.filtered_keys[selected_index]
        else:
            selected_spec = None

        if selected_spec:
            input_spec_number_textbox.delete(0, tk.END)
            input_spec_number_textbox.insert(0, selected_spec)
            input_spec_number_textbox.config(foreground=TEXT_COLOR_DEFAULT)
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    select_btn = ttk.Button(btn_frame, text="Select", command=on_select, width=10)
    cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel, width=10)
    select_btn.grid(row=0, column=0, padx=(0, scale_pixels(10)))
    cancel_btn.grid(row=0, column=1)

    dialog.bind("<Return>", lambda e: on_select())
    dialog.bind("<Escape>", lambda e: on_cancel())

    series_cb.focus_set()

    dialog.update_idletasks()
    dialog_width = root.winfo_width() - scale_pixels(50)
    dialog_height = dialog.winfo_reqheight()
    dialog_x = root.winfo_rootx() + scale_pixels(20)
    dialog_y = root.winfo_rooty() + scale_pixels(20)
    dialog.geometry(f"{dialog_width}x{dialog_height}+{dialog_x}+{dialog_y}")
    dialog.minsize(scale_pixels(250), dialog_height)
    dialog.deiconify()

def add_listview_tooltips(treeview):
    tooltip = None
    tooltip_col = None

    def on_motion(event):
        nonlocal tooltip, tooltip_col

        if show_tooltips_var is not None and not show_tooltips_var.get():
            if tooltip:
                try:
                    tooltip.destroy()
                except tk.TclError:
                    pass
                tooltip = None
                tooltip_col = None
            return

        region = treeview.identify_region(event.x, event.y)
        col = treeview.identify_column(event.x)
        
        text = None
        if region == "cell":
            if col == LIST_COLUMN_FOLDER_NAME:
                text = TEXT_TOOLTIP_RENAME_FOLDER
            elif col == LIST_COLUMN_OPEN_FOLDER:
                text = TEXT_TOOLTIP_OPEN_FOLDER
            elif col == LIST_COLUMN_OPEN_WEB_LINK:
                text = TEXT_TOOLTIP_GO_TO_WEB

        if text:
            if tooltip is not None and tooltip_col != col:
                try:
                    tooltip.destroy()
                except tk.TclError:
                    pass
                tooltip = None
                tooltip_col = None

            if tooltip is None:
                x = treeview.winfo_rootx() + event.x
                y = treeview.winfo_rooty() + event.y + scale_pixels(20)
                tooltip = tk.Toplevel(treeview)
                tooltip.wm_overrideredirect(True)
                tooltip.wm_geometry(f"+{x}+{y}")
                label = tk.Label(tooltip, text=text, background="#ffffe0", relief='solid', borderwidth=scale_pixels(1), font=("tahoma", "8", "normal"))
                label.pack(ipadx=scale_pixels(1))
                tooltip_col = col

                def auto_destroy():
                    nonlocal tooltip, tooltip_col
                    if tooltip:
                        try:
                            tooltip.destroy()
                        except tk.TclError:
                            pass
                        tooltip = None
                        tooltip_col = None

                tooltip.after(TOOLTIP_DURATION_MS, auto_destroy)
        else:
            if tooltip:
                try:
                    tooltip.destroy()
                except tk.TclError:
                    pass
                tooltip = None
                tooltip_col = None

    def on_leave(event):
        nonlocal tooltip, tooltip_col
        if tooltip:
            try:
                tooltip.destroy()
            except tk.TclError:
                pass
            tooltip = None
            tooltip_col = None

    treeview.bind("<Motion>", on_motion)
    treeview.bind("<Leave>", on_leave)

def configure_styles():
    style = ttk.Style()
    default_font_family = font.nametofont("TkDefaultFont").actual("family")

    style.configure("Treeview", rowheight=scale_pixels(30), font=(default_font_family, 10))
    style.configure("Treeview.Heading", font=(default_font_family, 10))
    style.configure("ThemeList.TRadiobutton", font=(default_font_family, 10))

    icon_font_size = scale_pixels(12)
    style.configure("Icon.TButton", font=(default_font_family, -icon_font_size))
    style.configure("Icon.TMenubutton", font=(default_font_family, -icon_font_size))

def change_theme(theme_name):
    try:
        root.style.theme_use(theme_name)
    except tk.TclError:
        pass
    save_config_key_value(CONFIG_KEY_THEME, theme_name)
    configure_styles()

def quit_from_tray(icon, item):
    root.after(0, exit_app)

def show_from_tray(icon, item):
    root.after(0, activate_main_window)

def run_tray_icon():
    global tray_icon
    if not pystray: return

    image_path = bundled_data_path(FILE_APP_ICON)
    if image_path and os.path.exists(image_path):
        image = Image.open(image_path)
    else:
        image = Image.new('RGB', (64, 64), color=(73, 109, 137))

    menu = pystray.Menu(
        pystray.MenuItem("Show", show_from_tray, default=True),
        pystray.MenuItem("Quit", quit_from_tray)
    )

    tray_icon = pystray.Icon(PROGRAM_NAME, image, PROGRAM_NAME, menu)
    tray_icon.run()

def start_tray_icon_thread():
    if not pystray: 
        return
    threading.Thread(target=run_tray_icon, daemon=True).start()

def open_theme_selector():
    theme_window = tk.Toplevel(root)
    theme_window.withdraw()
    theme_window.transient(root)
    theme_window.grab_set()
    set_icon(theme_window)
    theme_window.title("Change theme")
    theme_window.resizable(False, False)

    content_frame = ttk.Frame(theme_window)
    content_frame.pack(fill='both', expand=True, padx=scale_pixels(10), pady=(scale_pixels(15), scale_pixels(20)))

    current_theme = root.style.theme.name
    var = StringVar(value=current_theme)

    for theme_name in sorted(root.style.theme_names()):
        rb = ttk.Radiobutton(
            content_frame, 
            text=theme_name, 
            variable=var, 
            value=theme_name,
            command=lambda t=theme_name: change_theme(t),
            style="ThemeList.TRadiobutton"
        )
        rb.pack(anchor='w', padx=scale_pixels(5), pady=scale_pixels(2), fill='x')

    button_frame = ttk.Frame(content_frame)
    button_frame.pack(fill='x', pady=(scale_pixels(10), 0))
    ttk.Button(button_frame, text="Close", command=theme_window.destroy).pack()

    theme_window.update_idletasks()
    win_w = max(scale_pixels(200), theme_window.winfo_reqwidth())
    win_h = theme_window.winfo_reqheight()
    x = root.winfo_x() + (root.winfo_width() - win_w) // 2
    y = root.winfo_y() + (root.winfo_height() - win_h) // 2
    theme_window.geometry(f'{win_w}x{win_h}+{x}+{y}')
    theme_window.deiconify()
    theme_window.focus_set()

def toggle_show_tooltips():
    if show_tooltips_var:
        save_config_key_value(CONFIG_KEY_SHOW_TOOLTIPS, str(show_tooltips_var.get()))

def toggle_layout():
    if settings_button:
        settings_button.configure(menu='')

    if settings_menu:
        settings_menu.unpost()
    
    if root:
        root.focus_set()
        root.after(100, _toggle_layout_delayed)

def _toggle_layout_delayed():
    current_mode = load_config_key_value(CONFIG_KEY_COMPACT_MODE)
    new_mode = "False" if current_mode == "True" else "True"
    save_config_key_value(CONFIG_KEY_COMPACT_MODE, new_mode)
    apply_layout_mode(resize_window=True)

    if settings_button and settings_menu:
        settings_button.configure(menu=settings_menu)

def apply_layout_mode(resize_window=False):
    is_compact = load_config_key_value(CONFIG_KEY_COMPACT_MODE) == "True"
    
    if compact_mode_var:
        compact_mode_var.set(is_compact)

    if settings_button:
        settings_button.grid_forget()

    if is_compact:
        if root: 
            w, h = scale_pixels(400), scale_pixels(200)
            root.minsize(w, h)
            if resize_window:
                try:
                    current_geo = root.geometry()
                    match = re.match(r"(\d+)x(\d+)(.*)", current_geo)
                    if match:
                        root.geometry(f"{w}x{h}{match.group(3)}")
                    else:
                        root.geometry(f"{w}x{h}")
                except:
                    root.geometry(f"{w}x{h}")

        if frame_row1: frame_row1.grid_remove()
        if frame_row2: frame_row2.grid_remove()
        if frame_row3: frame_row3.grid_configure(pady=(scale_pixels(5), 0))
        if settings_button and frame_row3:
            settings_button.grid(in_=frame_row3, row=0, column=11, sticky='e', padx=(scale_pixels(5), 0))
            settings_button.lift()
    else:
        if root: root.minsize(scale_pixels(450), scale_pixels(300))
        if frame_row1: frame_row1.grid()
        if frame_row2: frame_row2.grid()
        if frame_row3: frame_row3.grid_configure(pady=0)
        if settings_button and frame_row1:
            settings_button.grid(in_=frame_row1, row=0, column=1, padx=(scale_pixels(5),0))
            settings_button.lift()

def show_about():
    if not root:
        return

    about_window = tk.Toplevel(root)
    about_window.withdraw()
    about_window.transient(root)
    about_window.grab_set()
    set_icon(about_window)
    about_window.title(PROGRAM_NAME)
    about_window.resizable(False, False)

    content_frame = ttk.Frame(about_window)
    content_frame.pack(fill='both', expand=True, padx=scale_pixels(10), pady=(scale_pixels(12), scale_pixels(14)))

    ttk.Label(
        content_frame,
        text=MSG_ABOUT_CONTENT.strip(),
        justify='left',
        anchor='w',
    ).pack(fill='x')

    button_frame = ttk.Frame(content_frame)
    button_frame.pack(fill='x', pady=(scale_pixels(10), 0))
    ttk.Button(button_frame, text="Close", command=about_window.destroy).pack()

    about_window.update_idletasks()
    win_w = max(scale_pixels(260), about_window.winfo_reqwidth())
    win_h = about_window.winfo_reqheight()
    x = root.winfo_x() + (root.winfo_width() - win_w) // 2
    y = root.winfo_y() + (root.winfo_height() - win_h) // 2
    about_window.geometry(f"{win_w}x{win_h}+{x}+{y}")
    about_window.deiconify()
    about_window.focus_set()

def load_gui():
    global root, root_folder_textbox, input_spec_number_textbox
    global add_button, download_button, remove_button
    global browse_specs_button
    global favorite_button, unfavorite_button
    global list_view, checkbox_header
    global settings_button, settings_menu, compact_mode_var, show_tooltips_var, minimize_to_tray_var
    global frame_row1, frame_row2, frame_row3

    loaded_theme = load_config_key_value(CONFIG_KEY_THEME)
    if not loaded_theme:
        loaded_theme = "superhero"

    root = ttk.Window(themename=loaded_theme)
    root.withdraw()
    root.title(f"{PROGRAM_NAME}")
    set_icon(root)

    current_dpi = root.winfo_fpixels('1i')
    root.tk.call('tk', 'scaling', current_dpi / 72.0)

    root.update_idletasks()
    
    if load_config_key_value(CONFIG_KEY_COMPACT_MODE) == "True":
        root.minsize(scale_pixels(400), scale_pixels(250))
    else:
        root.minsize(scale_pixels(450), scale_pixels(300))

    loaded_geometry = load_config_key_value(CONFIG_KEY_WIN_GEOMETRY)
    if loaded_geometry:
        root.geometry(loaded_geometry)
    else:
        root.geometry(f"{scale_pixels(500)}x{scale_pixels(350)}")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(3, weight=1)
    root.protocol("WM_DELETE_WINDOW", on_closing)

    configure_styles()

    checkbox_header = tk.BooleanVar(value=False)

    frame_row1 = ttk.Frame(root, padding=scale_pixels(5))
    frame_row1.grid(row=0, column=0, sticky="ew", padx=scale_pixels(5), pady=(scale_pixels(5),0))
    frame_row1.columnconfigure(0, weight=1)

    root_folder_textbox = ttk.Entry(frame_row1, state='readonly', width=60)
    root_folder_textbox.grid(row=0, column=0, sticky="ew")
    root_folder_textbox.bind('<Double-1>', lambda event: open_root_folder())
    Tooltip(root_folder_textbox, TEXT_TOOLTIP_ROOT_FOLDER)
    TEXT_COLOR_DEFAULT = root_folder_textbox.cget('foreground')
    set_root_folder_textbox(load_config_key_value(CONFIG_KEY_ROOT_FOLDER))
    
    settings_button = ttk.Menubutton(root, text=ICON_SETTINGS, width=3, style="Icon.TMenubutton")
    settings_button.grid(in_=frame_row1, row=0, column=1, padx=(scale_pixels(5),0))
    settings_button.bind('<Double-1>', lambda event: toggle_layout())

    default_font_family = font.nametofont("TkDefaultFont").actual("family")
    settings_menu = tk.Menu(settings_button, tearoff=0, font=(default_font_family, 10))
    
    is_compact_init = load_config_key_value(CONFIG_KEY_COMPACT_MODE) == "True"
    compact_mode_var = tk.BooleanVar(value=is_compact_init)

    show_tooltips_val = load_config_key_value(CONFIG_KEY_SHOW_TOOLTIPS)
    is_show_tooltips = True
    if show_tooltips_val is not None:
        is_show_tooltips = show_tooltips_val == "True"
    show_tooltips_var = tk.BooleanVar(value=is_show_tooltips)

    minimize_to_tray_val = load_config_key_value(CONFIG_KEY_MINIMIZE_TO_TRAY)
    is_minimize_to_tray = True if minimize_to_tray_val is None else (minimize_to_tray_val == "True")
    minimize_to_tray_var = tk.BooleanVar(value=is_minimize_to_tray)

    settings_menu.add_command(label="Select root folder", command=select_root_folder)
    settings_menu.add_command(label="Change theme", command=open_theme_selector)
    settings_menu.add_checkbutton(label="Show tooltips", variable=show_tooltips_var, command=toggle_show_tooltips)
    settings_menu.add_checkbutton(label="Close to tray", variable=minimize_to_tray_var, command=lambda: save_config_key_value(CONFIG_KEY_MINIMIZE_TO_TRAY, str(minimize_to_tray_var.get())))
    settings_menu.add_checkbutton(label="Compact mode", variable=compact_mode_var, command=toggle_layout)
    settings_menu.add_separator()
    settings_menu.add_command(label="About", command=show_about)
    settings_button.configure(menu=settings_menu)

    Tooltip(settings_button, TEXT_TOOLTIP_SETTINGS)

    frame_row2 = ttk.Frame(root, padding=scale_pixels(5))
    frame_row2.grid(row=1, column=0, sticky="ew", padx=scale_pixels(5))
    frame_row2.columnconfigure(0, weight=1)

    browse_specs_button = ttk.Button(frame_row2, text=ICON_SEARCH, command=on_browse_specs, width=3, style="Icon.TButton")
    browse_specs_button.grid(row=0, column=1, padx=(0, scale_pixels(1)))
    Tooltip(browse_specs_button, TEXT_BROWSE_SPECS)

    input_spec_number_textbox = ttk.Entry(frame_row2, justify='right', width=11)
    input_spec_number_textbox.grid(row=0, column=2, sticky="e")
    init_input_spec_number_textbox()
    input_spec_number_textbox.bind('<FocusIn>', focus_in_input_spec_number)
    input_spec_number_textbox.bind('<FocusOut>', focus_out_input_spec_number)
    input_spec_number_textbox.bind('<Return>', lambda event=None: add_button.invoke())

    add_button = ttk.Button(frame_row2, text="Add", command=add_item)
    add_button.grid(row=0, column=3, padx=(scale_pixels(5), 0))
    Tooltip(add_button, TEXT_TOOLTIP_ADD_SPEC)

    frame_row3 = ttk.Frame(root, padding=scale_pixels(5))
    frame_row3.grid(row=2, column=0, sticky="ew", padx=scale_pixels(5), pady=(scale_pixels(5),0))
    frame_row3.columnconfigure(10, weight=1)
    download_button = ttk.Button(frame_row3, text="Download", command=download_selected_item)
    download_button.grid(row=0, column=0)
    Tooltip(download_button, TEXT_TOOLTIP_DOWNLOAD)
    remove_button = ttk.Button(frame_row3, text=ICON_REMOVE, command=remove_selected_item, width=2, style="Icon.TButton")
    remove_button.grid(row=0, column=1, padx=(scale_pixels(5), 0))
    Tooltip(remove_button, TEXT_TOOLTIP_REMOVE)
    favorite_button = ttk.Button(frame_row3, text=ICON_FAVORITE, command=set_favorite_selected_item, width=2, style="Icon.TButton")
    favorite_button.grid(row=0, column=2, padx=(scale_pixels(5), 0))
    Tooltip(favorite_button, TEXT_TOOLTIP_FAVORITE)
    unfavorite_button = ttk.Button(frame_row3, text=ICON_UNFAVORITE, command=unset_favorite_selected_item, width=2, style="Icon.TButton")
    unfavorite_button.grid(row=0, column=3, padx=(scale_pixels(5), 0))
    Tooltip(unfavorite_button, TEXT_TOOLTIP_UNFAVORITE)

    frame_row4 = ttk.Frame(root, padding=scale_pixels(5))
    frame_row4.grid(row=3, column=0, sticky="nsew", padx=scale_pixels(5), pady=(0,scale_pixels(5)))
    frame_row4.columnconfigure(0, weight=1)
    frame_row4.rowconfigure(1, weight=1)

    columns = ('favorite', 'select', 'status', 'folder_name', 'open_folder_action', 'open_web_link_action')
    list_view = ttk.Treeview(frame_row4, columns=columns, show='headings', selectmode='browse')

    col_text_select = ICON_CHECK_OFF
    col_text_status = 'Status'
    col_text_spec = 'Specification'
    col_text_folder = 'Folder'
    col_text_web = 'Web'

    list_view.heading('favorite', text='')
    list_view.heading('select', text=col_text_select, command=on_checkbox_header_click)
    list_view.heading('status', text=col_text_status)
    list_view.heading('folder_name', text=col_text_spec)
    list_view.heading('open_folder_action', text=col_text_folder)
    list_view.heading('open_web_link_action', text=col_text_web)
    add_listview_tooltips(list_view)

    heading_font = font.Font(family=default_font_family, size=10)

    def get_col_width(text, min_w=0):
        return heading_font.measure(text) + scale_pixels(14)

    w = scale_pixels(20)
    list_view.column('favorite', width=w, minwidth=w, anchor='e', stretch=tk.NO) 
    w = get_col_width(col_text_select, 30)
    list_view.column('select', width=w, minwidth=w, anchor='center', stretch=tk.NO)
    w = get_col_width(col_text_status, 60)
    list_view.column('status', width=w, minwidth=w, anchor='center', stretch=tk.NO)
    w = scale_pixels(140)
    list_view.column('folder_name', width=w, minwidth=w, anchor='w', stretch=tk.YES)
    w = get_col_width(col_text_folder, 65)
    list_view.column('open_folder_action', width=w, minwidth=w, anchor='center', stretch=tk.NO)
    w = get_col_width(col_text_web, 50)
    list_view.column('open_web_link_action', width=w, minwidth=w, anchor='center', stretch=tk.NO)

    list_view.grid(row=1, column=0, sticky='nsew')
    vsb = ttk.Scrollbar(frame_row4, orient="vertical", command=list_view.yview)
    vsb.grid(row=1, column=1, sticky='ns')
    hsb = ttk.Scrollbar(frame_row4, orient="horizontal", command=list_view.xview)
    hsb.grid(row=2, column=0, sticky='ew')
    list_view.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    list_view.bind('<ButtonRelease-1>', on_list_item_click)
    list_view.bind('<Double-1>', on_list_item_double_click)
    list_view.bind('<ButtonRelease-3>', on_list_item_right_click)

    refresh_list_view(initial_list=load_list_from_save_config())
    
    apply_layout_mode(resize_window=False)

def main():
    try:
        appid = f'{PROGRAM_DIR}'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)
        print(f"AppUserModelID set to: {appid}")
    except Exception as e:
        print(f"WARNING: Could not set AppUserModelID: {e}")

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    load_spec_names()
    load_gui()
    check_and_update_list()

    if "--minimized" in sys.argv:
        if pystray:
            start_tray_icon_thread()
            start_activate_event_listener()
        else:
            root.deiconify()
    else:
        root.deiconify()

    root.mainloop()

if __name__ == "__main__":
    if not acquire_single_instance_mutex():
        notify_existing_instance_to_activate()
        sys.exit(0)
    try:
        main()
    finally:
        release_single_instance_mutex()
