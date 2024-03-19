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
import keyboard
import json

class App:
    def __init__(self, root, usb_enumerator, intrusion_handler, keymon):
        self.root = root
        self.usb_enumerator = usb_enumerator
        self.intrusion_handler = intrusion_handler
        self.keymon = keymon
        self.setup_window()
        self.setup_tab_control()
        self.setup_device_table()
        self.setup_registered_device_table()
        self.setup_status_labels()
        self.setup_buttons()
        self.refresh_registered_device()

    def setup_window(self):
        self.root.title("SafeUSB")
        icon = Image.open("favicon.ico")
        icon = ImageTk.PhotoImage(icon)
        self.root.iconphoto(True, icon)
        width, height = 680, 290
        screenwidth, screenheight = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.root.geometry(alignstr)
        self.root.resizable(width=False, height=False)

    def setup_tab_control(self):
        self.tabControl = ttk.Notebook(self.root)
        self.tab1, self.tab2 = ttk.Frame(self.tabControl), ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab1, text='Active Devices')
        self.tabControl.add(self.tab2, text='Safe Devices')
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

    def register_selected_devices(self):
        selected_items = self.deviceTable.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "No device selected.")
            return

        with open('safe.txt', 'r') as f:
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
                self.usb_enumerator.write_to_file(device_name, device_class, device_id)
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
            self.usb_enumerator.remove_from_file(device_name, device_class, device_id)
            self.registeredDeviceTable.delete(item)
        self.usb_enumerator.check_unregistered_devices()

    def refresh_registered_device(self):    
        for i in self.registeredDeviceTable.get_children():
            self.registeredDeviceTable.delete(i)
        if not os.path.isfile('safe.txt'):
            open('safe.txt', 'w').close()
        with open('safe.txt', 'r') as f:
            for line in f:
                device_name, device_class, device_id = line.strip().split(',')
                self.registeredDeviceTable.insert('', 'end', values=(device_name, device_class, device_id))   

    def hide_window(self):
        runNotify = Notify()
        runNotify.title = "SafeUSB is active"
        runNotify.message = "SafeUSB is running in the background"
        runNotify.icon = "information.png"
        runNotify.send()
        root.withdraw()
        image=Image.open("favicon.ico")
        menu=(item('Show', self.show_window), item('Quit', self.quit_window))
        icon=pystray.Icon("name", image, "SafeUSB", menu)
        icon.run()

    def show_window(self, icon, item):
        icon.stop()
        root.after(0,lambda: root.deiconify())

    def quit_window(self, icon, item):
        icon.stop()
        if self.usb_enumerator.p is not None and self.usb_enumerator.p.is_alive():
            self.usb_enumerator.p.terminate()
        root.destroy()
           
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
                self.write_to_file(device_name, device_class, device_id)

            device['Status'] = device_status  # Store the status in the device dictionary

        self.check_new_devices(new_devices, registered_devices)
        self.check_disconnected_devices(new_devices)
        self.devices = new_devices
        self.check_unregistered_devices()

    def load_registered_devices(self):
        if not os.path.isfile('safe.txt'):
            open('safe.txt', 'w').close()
        with open('safe.txt', 'r') as f:
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
                    self.write_to_file(device_name, device_class, device_id)

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

    def write_to_file(self, device_name, device_class, device_id):
        devices = self.read_current_contents()
        new_device = f"{device_name},{device_class},{device_id}\n"
        if new_device not in devices:
            self.append_to_file(new_device)
            if self.callback:
                self.callback()
                
    def remove_from_file(self, device_name, device_class, device_id):
        devices = self.read_current_contents()
        device = f"{device_name},{device_class},{device_id}\n"
        if device in devices:
            devices.remove(device)
            with open('safe.txt', 'w') as f:
                f.writelines(devices)
            if self.callback:
                self.callback()

    def read_current_contents(self):
        if not os.path.isfile('safe.txt'):
            open('safe.txt', 'w').close()
        with open('safe.txt', 'r') as f:
            devices = f.readlines()
        return devices

    def append_to_file(self, new_device):
        if not os.path.isfile('safe.txt'):
            open('safe.txt', 'w').close()
        with open('safe.txt', 'a') as f:
            f.write(new_device)

class IntrusionHandler:
    def __init__(self, queue):
        self.queue = queue
        self.notification_sent = False

    def send_intrusion_warning(self):
        intrusionWarning = Notify()
        intrusionWarning.title = "Intrusion Detected"
        intrusionWarning.message = "HID keystroke injection by BadUSB detected"
        intrusionWarning.icon = "warning.png"
        intrusionWarning.send()
        messagebox.showwarning("Intrusion Detected by SafeUSB", "Possible HID keystroke injection by BadUSB detected.\n\nAll keyboard input will be blocked.\n\nTo unblock, register any unregistered device (if you believe this warning is a false positive) or immediately check your physical USB port and disconnect any malicious device")

    def block_keyboard(self):
        for i in range(150):
            keyboard.block_key(i) 
        self.queue.put(('keyboard_blocked',)) 
    
    def unblock_keyboard(self):
        keyboard.unhook_all()
        self.queue.put(('keyboard_unblocked',)) 

class KeystrokeMonitoring:
    def __init__(self, intrusion_handler):
        self.intrusion_handler = intrusion_handler # Create an instance of IntrusionHandler
        self.limit = 30
        self.size = 25
        self.speed = 0
        self.prev = -1
        self.i = 0
        self.speedIntrusion = False
        self.history = [self.limit+1] * self.size
        self.keylogged = ""
        self.keyWords = self.read_keywords()
        self.contentIntrusion = False
    
    def read_keywords(self):
        filename = "keywords.txt"
        default_keywords = ["POWERSHELL", "CMD.EXE", "USER", "HOSTNAME", "TASK", "NEW-OBJECT"]

        # Check if file exists
        if not os.path.exists(filename):
            with open(filename, 'w') as f:
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
            if word in self.keylogged.upper():
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
    q = multiprocessing.Queue()
    handler = IntrusionHandler(q)
    keymon = KeystrokeMonitoring(handler)
    usb_enumerator = USBEnumerator(q, keymon, handler)
    app = App(root, usb_enumerator, handler, keymon)
    root.protocol('WM_DELETE_WINDOW', app.hide_window)
    
    app.update_gui()  # Start the periodic call to the function
    app.hide_window()  # Add this line to hide the window on startup
    root.mainloop()