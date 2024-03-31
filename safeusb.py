import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font as tkFont
import tkinter.messagebox as messagebox
from usbmonitor import USBMonitor
from usbmonitor.attributes import ID_MODEL_FROM_DATABASE, ID_USB_CLASS_FROM_DATABASE, DEVNAME
from pystray import MenuItem as item
import pystray
from PIL import Image, ImageTk
import pyWinhook as pyHook
import pythoncom
import multiprocessing
from notifypy import Notify
import os
import subprocess
import sys
import keyboard
import json
import winreg
import configparser
import time
import win32evtlog
import win32evtlogutil

BUNDLE_DIR = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
APP_ICON = os.path.abspath(os.path.join(BUNDLE_DIR, r"safeusb-data\favicon.ico"))
WARNING_ICON = os.path.abspath(os.path.join(BUNDLE_DIR, r"safeusb-data\warning.png"))
INFO_ICON = os.path.abspath(os.path.join(BUNDLE_DIR, r"safeusb-data\information.png")) 

CWD = os.path.abspath(os.path.dirname(sys.executable))
SAFE_DATABASE = os.path.join(CWD, "safedatabase.txt")
KEYWORDS = os.path.join(CWD, "keywords.txt")
CONFIG_FILE = os.path.join(CWD, "config.ini")

class App:
    def __init__(self, root, usb_enumerator, intrusion_handler, keymon, config_handler, registry_manager):
        self.root = root
        self.usb_enumerator = usb_enumerator
        self.intrusion_handler = intrusion_handler
        self.keymon = keymon
        self.config_handler = config_handler
        self.registry_manager = registry_manager
        self.startup_checkbox = tk.IntVar()
        self.setup_window()
        self.setup_tab_control()
        self.setup_device_table()
        self.setup_registered_device_table()
        self.setup_status_labels()
        self.setup_buttons()
        self.setup_autostartcheckbox()
        self.setup_keymonconfig()
        self.refresh_registered_device()

    def setup_window(self):
        self.root.title("SafeUSB")
        icon = Image.open(APP_ICON)
        icon = ImageTk.PhotoImage(icon)
        self.root.iconphoto(True, icon)
        width, height = 680, 290
        screenwidth, screenheight = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.root.geometry(alignstr)
        self.root.resizable(width=False, height=False)

    def setup_tab_control(self):
        self.tabControl = ttk.Notebook(self.root)
        self.tab1, self.tab2, self.tab3 = ttk.Frame(self.tabControl), ttk.Frame(self.tabControl), ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab1, text='Active Devices')
        self.tabControl.add(self.tab2, text='Safe Devices')
        self.tabControl.add(self.tab3, text='Configuration')
        self.tabControl.pack(expand=1, fill="both")

    def setup_table(self, tab, columns, scrollbar_x, scrollbar_y):
        scrollbar = ttk.Scrollbar(tab)
        scrollbar.place(x=scrollbar_x, y=scrollbar_y, height=207)
        table = ttk.Treeview(tab, selectmode="extended", show="headings", yscrollcommand=scrollbar.set)
        scrollbar.configure(command=table.yview)
        table['columns'] = columns
        for col in table['columns']:
            table.heading(col, text=col)
            table.column(col, width=tkFont.Font().measure(col))
        table.place(x=10, y=10, width=647, height=207)
        return table

    def setup_device_table(self):
        self.deviceTable = self.setup_table(self.tab1, ('Device Name', 'Class', 'Status', 'Device ID'), 657, 10)
        self.deviceTable.tag_configure('Safe', background='green')
        self.deviceTable.tag_configure('Unregistered', background='yellow')

    def setup_registered_device_table(self):
        self.registeredDeviceTable = self.setup_table(self.tab2, ('Device Name', 'Device Class', 'Device ID'), 657, 10)

    def setup_status_labels(self):
        self.keystroke_status_label = tk.Label(self.root, text="Keystroke Monitoring: Stopped", fg="grey")
        self.keystroke_status_label.place(x=320, y=253)

        self.keyboard_block_status_label = tk.Label(self.root, text="Keyboard Unblocked", fg="green")
        self.keyboard_block_status_label.place(x=520, y=253)

    def setup_buttons(self):
        self.authButton = ttk.Button(self.tab1, text="Register Device as Safe", command=self.register_selected_devices)
        self.authButton.place(x=10, y=230)
        self.unauthButton = ttk.Button(self.tab2, text="Unregister Device", command=self.unregister_selected_devices)
        self.unauthButton.place(x=10, y=230)
        
    def setup_autostartcheckbox(self):
        try:
            self.startup_checkbox.set(self.config_handler.load_from_config('Autostart', 'run_at_startup'))
        except (configparser.NoSectionError, configparser.NoOptionError):
            self.startup_checkbox.set(self.registry_manager.check_autostart_registry("SafeUSB"))

        self.chk = ttk.Checkbutton(self.tab3, text="Run at startup", variable=self.startup_checkbox, command=self.toggle_autostart)
        self.chk.place(x=10, y=10)
        
    def setup_keymonconfig(self):
        self.limit_label = ttk.Label(self.tab3, text="Keystroke speed threshold")
        self.limit_label.place(x=10, y=50)
        self.limit_entry = ttk.Entry(self.tab3)
        self.limit_entry.place(x=10, y=70)
        self.limit_entry.insert(0, self.keymon.limit)

        self.size_label = ttk.Label(self.tab3, text="Keystroke size")
        self.size_label.place(x=10, y=110)
        self.size_entry = ttk.Entry(self.tab3)
        self.size_entry.place(x=10, y=130)
        self.size_entry.insert(0, self.keymon.size)

        self.save_limit_button = ttk.Button(self.tab3, text="Save", command=self.save_keymonconfig)
        self.save_limit_button.place(x=10, y=230)
        
    def save_keymonconfig(self):
        limit = int(self.limit_entry.get())
        size = int(self.size_entry.get())
        self.keymon.limit = limit
        self.keymon.size = size
        self.config_handler.save_int_to_config('KeystrokeMonitoring', 'limit', limit)
        self.config_handler.save_int_to_config('KeystrokeMonitoring', 'size', size)
        messagebox.showwarning("Info", "Keystroke Monitoring configuration changed, SafeUSB will restart")
        self.restart_program()
        
    def toggle_autostart(self):
        app_name = "SafeUSB"
        key_data = sys.executable
        autostart = self.startup_checkbox.get()

        if autostart:
            if not self.registry_manager.set_autostart_registry(app_name, key_data):
                messagebox.showerror("Error", "Failed to set autostart.")
        else:
            if not self.registry_manager.set_autostart_registry(app_name, key_data, autostart=False):
                messagebox.showerror("Error", "Failed to remove autostart.")

        self.config_handler.save_to_config('Autostart', 'run_at_startup', str(autostart))

    def register_selected_devices(self):
        selected_items = self.deviceTable.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "No device selected.")
            return

        with open(SAFE_DATABASE, 'r') as f:
            registered_devices = [line.strip().split(',') for line in f]

        for item in selected_items:
            device_name, device_class, device_status, device_id = self.deviceTable.item(item, "values")
            if device_status == 'Safe':
                messagebox.showwarning("Warning", f"Device {device_name} is already registered.")
                continue

            matching_devices = [device for device in self.usb_enumerator.devices.values() if device['DEVNAME'] == device_id]
            if not matching_devices:
                messagebox.showwarning("Warning", f"Device {device_name} not found.")
                continue

            for device in matching_devices:
                if any(rd[0] == device_name and rd[1] == device_class and rd[2] == device_id for rd in registered_devices):
                    messagebox.showwarning("Warning", f"Device {device_name} is already registered.")
                    continue
                self.usb_enumerator.write_to_database(device_name, device_class, device_id)
                device['Status'] = 'Safe'  # Update the 'Status' key in the device dictionary
                self.deviceTable.set(item, 'Status', 'Safe')
                self.deviceTable.item(item, tags=('Safe',))
                self.refresh_registered_device()
        # Check for unregistered devices after a device is registered
        self.usb_enumerator.check_unregistered_devices()
        
    def unregister_selected_devices(self):
        selected_items = self.registeredDeviceTable.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "No device selected.")
            return

        for item in selected_items:
            device_name, device_class, device_id = self.registeredDeviceTable.item(item, "values")
            self.usb_enumerator.remove_from_database(device_name, device_class, device_id)
            self.registeredDeviceTable.delete(item)
        
        messagebox.showwarning("Info", "Device(s) unregistered, SafeUSB will restart")
        self.restart_program()

    def refresh_registered_device(self):    
        for i in self.registeredDeviceTable.get_children():
            self.registeredDeviceTable.delete(i)
        if not os.path.isfile(SAFE_DATABASE):
            open(SAFE_DATABASE, 'w').close()
        with open(SAFE_DATABASE, 'r') as f:
            for line in f:
                device_name, device_class, device_id = line.strip().split(',')
                self.registeredDeviceTable.insert('', 'end', values=(device_name, device_class, device_id))   

    def hide_window(self):
        runNotify = Notify()
        runNotify.title = "SafeUSB is active"
        runNotify.message = "SafeUSB is running in the background"
        runNotify.icon = INFO_ICON
        runNotify.send()
        root.withdraw()
        image=Image.open(APP_ICON)
        menu=(item('Show', self.show_window), item('Quit', self.quit_program))
        icon=pystray.Icon("name", image, "SafeUSB", menu)
        icon.run()

    def show_window(self, icon, item):
        icon.stop()
        root.after(0,lambda: root.deiconify())

    def quit_program(self, icon, item):
        icon.stop()
        if self.usb_enumerator.p is not None and self.usb_enumerator.p.is_alive():
            self.usb_enumerator.p.terminate()
        root.destroy()
        
    def restart_program(self):
        if self.usb_enumerator.p is not None and self.usb_enumerator.p.is_alive():
            self.usb_enumerator.p.terminate()
            self.usb_enumerator.p.join()  # Ensure the process has time to terminate
        root.destroy()
        python = sys.executable
        subprocess.Popen([python] + sys.argv)
        sys.exit()

    def update_gui(self):
        while not q.empty():
            action, *data = q.get()
            if action == 'connect':
                device_name, device_class, device_status, device_id = data
                if device_status == 'Unregistered':
                    self.deviceTable.insert('', 0, values=(device_name, device_class, device_status, device_id), tags=(device_status,))
                else:
                    self.deviceTable.insert('', 'end' if device_status == 'Safe' else 0, values=(device_name, device_class, device_status, device_id), tags=(device_status,))
            elif action == 'disconnect':
                device_name = data[0]
                for item in self.deviceTable.get_children():
                    if self.deviceTable.item(item, "values")[0] == device_name:
                        self.deviceTable.delete(item)
                        break
            elif action == 'keystroke_monitoring_started':
                    self.keystroke_status_label.config(text="Keystroke Monitoring: Active", fg="red")
            elif action == 'keystroke_monitoring_stopped':
                    self.keystroke_status_label.config(text="Keystroke Monitoring: Stopped", fg="green")
            elif action == 'keyboard_blocked':
                    self.keyboard_block_status_label.config(text="Keyboard Blocked", fg="red")
            elif action == 'keyboard_unblocked':
                    self.keyboard_block_status_label.config(text="Keyboard Unblocked", fg="green")
        self.root.after(100, self.update_gui)  # Schedule the next call to this function
     
class ConfigHandler:
    def __init__(self, config_file):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
    
        if not os.path.exists(self.config_file):
            self.create_default_config()
            
    def create_default_config(self):
        self.config['KeystrokeMonitoring'] = {'limit': '30', 'size': '10'}
        with open(self.config_file, 'x+') as configfile:
            self.config.write(configfile)

    def save_to_config(self, section, option, value):
        self.config[section] = {option: value}
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def load_from_config(self, section, option):
        self.config.read(self.config_file)
        return self.config.getboolean(section, option)

    def save_int_to_config(self, section, option, value):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, option, str(value))
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def load_int_from_config(self, section, option):
        self.config.read(self.config_file)
        if self.config.has_section(section):
            return self.config.getint(section, option)
        else:
            return None

class RegistryManager:    
    def set_autostart_registry(self, app_name, key_data, autostart: bool = True) -> bool:
        with winreg.OpenKey(
                key=winreg.HKEY_CURRENT_USER,
                sub_key=r'Software\Microsoft\Windows\CurrentVersion\Run',
                reserved=0,
                access=winreg.KEY_ALL_ACCESS,
        ) as key:
            try:
                if autostart:
                    winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, key_data)
                else:
                    winreg.DeleteValue(key, app_name)
            except OSError:
                return False
        return True

    def check_autostart_registry(self, value_name):
        with winreg.OpenKey(
                key=winreg.HKEY_CURRENT_USER,
                sub_key=r'Software\Microsoft\Windows\CurrentVersion\Run',
                reserved=0,
                access=winreg.KEY_ALL_ACCESS,
        ) as key:
            idx = 0
            while idx < 1_000:     # Max 1.000 values
                try:
                    key_name, _, _ = winreg.EnumValue(key, idx)
                    if key_name == value_name:
                        return True
                    idx += 1
                except OSError:
                    break
        return False

class USBEnumerator:
    def __init__(self, queue, keymon, intrusion_handler, callback=None):
        self.queue = queue
        self.callback = callback
        self.keymon = keymon
        self.intrusion_handler = intrusion_handler
        self.usb_monitor = USBMonitor()
        self.keystroke_monitoring_started = False
        self.p = None
        self.devices = {}  # Store the current devices
        self.usb_enum()
        self.usb_monitor.start_monitoring(on_connect=self.usb_enum, on_disconnect=self.usb_enum)

    def usb_enum(self, *args):        
        new_devices = self.usb_monitor.get_available_devices()
        registered_devices = self.load_registered_devices()

        for key, device in new_devices.items():
            device_name = f"{device['ID_MODEL_FROM_DATABASE']}"
            device_class = f"{device['ID_USB_CLASS_FROM_DATABASE']}"
            device_id = f"{device['DEVNAME']}"

            if any(rd[0] == device_name and rd[1] == device_class and rd[2] == device_id for rd in registered_devices):
                device_status = 'Safe'
            elif device_class == 'HIDClass':
                device_status = 'Unregistered'
                self.start_keystroke_monitoring()
            else:
                device_status = 'Safe'
                self.write_to_database(device_name, device_class, device_id)

            device['Status'] = device_status  # Store the status in the device dictionary

        self.check_new_devices(new_devices, registered_devices)
        self.check_disconnected_devices(new_devices)
        self.devices = new_devices
        self.check_unregistered_devices()

    def load_registered_devices(self):
        if not os.path.isfile(SAFE_DATABASE):
            open(SAFE_DATABASE, 'x+').close()
        with open(SAFE_DATABASE, 'r') as f:
            registered_devices = [line.strip().split(',') for line in f]
        if not registered_devices:
            messagebox.showinfo("Enumerating Device", "SafeUSB is enumerating device for the first time.")
        return registered_devices

    def check_new_devices(self, new_devices, registered_devices):
        for key, device in new_devices.items():
            if key not in self.devices:
                device_name = f"{device['ID_MODEL_FROM_DATABASE']}"
                device_class = f"{device['ID_USB_CLASS_FROM_DATABASE']}"
                device_id = f"{device['DEVNAME']}"

                if any(rd[0] == device_name and rd[1] == device_class and rd[2] == device_id for rd in registered_devices):
                    device_status = 'Safe'
                elif device_class == 'HIDClass':
                    device_status = 'Unregistered'
                    self.start_keystroke_monitoring()
                else:
                    device_status = 'Safe'
                    self.write_to_database(device_name, device_class, device_id)

                device['Status'] = device_status  # Store the status in the device dictionary
                self.queue.put(('connect', device_name, device_class, device_status, device_id))

    def start_keystroke_monitoring(self):
        if not self.keystroke_monitoring_started:
            self.p = multiprocessing.Process(target=self.keymon.start)
            self.p.start()
            self.keystroke_monitoring_started = True  
            self.queue.put(('keystroke_monitoring_started',))  

    def check_disconnected_devices(self, new_devices):
        for key in list(self.devices.keys()):
            if key not in new_devices:
                self.queue.put(('disconnect', self.devices[key]['ID_MODEL_FROM_DATABASE']))

    def check_unregistered_devices(self):
        unregistered_devices_left = any(device['Status'] == 'Unregistered' for device in self.devices.values())
        if not unregistered_devices_left and self.keystroke_monitoring_started:
            self.terminate_keystroke_monitoring()
            self.intrusion_handler.unblock_keyboard()

    def terminate_keystroke_monitoring(self):
        if self.p is not None and self.p.is_alive():
            self.p.terminate()
        self.keystroke_monitoring_started = False
        self.queue.put(('keystroke_monitoring_stopped',))  

    def write_to_database(self, device_name, device_class, device_id):
        devices = self.read_database()
        new_device = f"{device_name},{device_class},{device_id}\n"
        if new_device not in devices:
            self.append_to_database(new_device)
            if self.callback:
                self.callback()
                
    def remove_from_database(self, device_name, device_class, device_id):
        devices = self.read_database()
        device = f"{device_name},{device_class},{device_id}\n"
        if device in devices:
            devices.remove(device)
            with open(SAFE_DATABASE, 'w') as f:
                f.writelines(devices)
            if self.callback:
                self.callback()

    def read_database(self):
        if not os.path.isfile(SAFE_DATABASE):
            open(SAFE_DATABASE, 'w').close()
        with open(SAFE_DATABASE, 'r') as f:
            devices = f.readlines()
        return devices

    def append_to_database(self, new_device):
        if not os.path.isfile(SAFE_DATABASE):
            open(SAFE_DATABASE, 'w').close()
        with open(SAFE_DATABASE, 'a') as f:
            f.write(new_device)

class IntrusionHandler:
    def __init__(self, queue):
        self.queue = queue
        self.notification_sent = False

    def send_intrusion_warning(self):
        intrusionWarning = Notify()
        intrusionWarning.title = "Intrusion Detected"
        intrusionWarning.message = "HID keystroke injection by BadUSB detected"
        intrusionWarning.icon = WARNING_ICON
        intrusionWarning.send()
        messagebox.showwarning("Intrusion Detected by SafeUSB", "Possible HID keystroke injection by BadUSB detected.\n\nAll keyboard input will be blocked.\n\nTo unblock, register any unregistered device (if you believe this warning is a false positive) or immediately check your physical USB port and disconnect any malicious device")

    def write_to_event_log(self):
        # Define the event source detailspowershell
        source = 'SafeUSB'
        event_id = 1337  # The security ID structure is invalid.
        descr = ["HID keystroke injection by BadUSB detected"]
        # Write to the event log
        win32evtlogutil.ReportEvent(source, event_id, eventType=win32evtlog.EVENTLOG_WARNING_TYPE, strings=descr, data=None)
    
    def block_keyboard(self):
        for i in range(150):
            keyboard.block_key(i) 
        self.queue.put(('keyboard_blocked',)) 
    
    def unblock_keyboard(self):
        keyboard.unhook_all()
        self.queue.put(('keyboard_unblocked',)) 

class KeystrokeMonitoring:
    def __init__(self, intrusion_handler, config_handler):
        self.intrusion_handler = intrusion_handler # Create an instance of IntrusionHandler
        self.config_handler = config_handler
        self.limit = self.config_handler.load_int_from_config('KeystrokeMonitoring', 'limit')
        self.size = self.config_handler.load_int_from_config('KeystrokeMonitoring', 'size')
        self.speed = 0
        self.prev = -1
        self.i = 0
        self.speedIntrusion = False
        self.history = [self.limit+1] * self.size
        self.keylogged = "".lower()
        self.keyWords = self.read_keywords()
        self.contentIntrusion = False
    
    def read_keywords(self):
        filename = KEYWORDS
        default_keywords = ["POWERSHELL", "CMD.EXE", "USER", "HOSTNAME", "TASK", "NEWOem_MinusOBJECT", "LwinX", "LwinR", "LcontrolLmenuDelete"]

        # Check if file exists
        if not os.path.exists(filename):
            with open(filename, 'x+') as f:
                json.dump(default_keywords, f)
            return default_keywords

        # Load keywords from file
        with open(filename, 'r') as f:
            try:
                keywords = json.load(f)
                # If file is empty, write the default keywords
                if not keywords:
                    raise ValueError("Keyword list is empty")
            except Exception as e:
                messagebox.showerror("Error", str(e))
                with open(filename, 'w') as f:
                    json.dump(default_keywords, f)
                return default_keywords
        return keywords

    def KeyboardEvent(self, event):
        self.log_key(event.Key)
        self.detect_keywords()
        self.calculate_speed(event.Time)
        self.detect_intrusion()
        return True

    def log_key(self, key):
        print("Keystroke : " + key)
        self.keylogged += key

    def detect_keywords(self):
        for word in self.keyWords:
            if word in self.keylogged:
                print(f"[*] Key Words Detected: [{word}]")
                self.contentIntrusion = True
                self.keylogged = ""

    def calculate_speed(self, time):
        if (self.prev == -1):
            self.prev = time
            return

        if (self.i >= len(self.history)): self.i = 0

        self.history[self.i] = time - self.prev
        print(time, "-", self.prev, "=", self.history[self.i])
        self.prev = time
        self.speed = sum(self.history) / float(len(self.history))
        self.i = self.i + 1

        print("\rAverage Typing Speed:", self.speed)

        if (self.speed < self.limit):
            self.speedIntrusion = True
        else:
            self.speedIntrusion = False

    def detect_intrusion(self):
        if (self.speedIntrusion or self.contentIntrusion) and not self.intrusion_handler.notification_sent:
            self.intrusion_handler.write_to_event_log()
            self.intrusion_handler.block_keyboard()
            self.intrusion_handler.send_intrusion_warning()
            self.intrusion_handler.notification_sent = True
        elif not self.speedIntrusion and not self.contentIntrusion:
            self.intrusion_handler.notification_sent = False

    def start(self):
        keyhook = pyHook.HookManager()
        keyhook.KeyDown = self.KeyboardEvent
        keyhook.HookKeyboard()
        pythoncom.PumpMessages()

if __name__ == "__main__":
    multiprocessing.freeze_support() #freeze_support must be enabled when compiling to exe with pyinstaller with multiprocessing
    root = tk.Tk()
    config_handler = ConfigHandler(CONFIG_FILE)
    registry_manager = RegistryManager()
    q = multiprocessing.Queue()
    handler = IntrusionHandler(q)
    keymon = KeystrokeMonitoring(handler, config_handler)
    usb_enumerator = USBEnumerator(q, keymon, handler)
    app = App(root, usb_enumerator, handler, keymon, config_handler, registry_manager)
    root.protocol('WM_DELETE_WINDOW', app.hide_window)
    
    app.update_gui()  # Start the periodic call to the function
    app.hide_window()  # Add this line to hide the window on startup
    root.mainloop()