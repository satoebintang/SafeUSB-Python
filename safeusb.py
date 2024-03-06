import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font as tkFont
import tkinter.messagebox as messagebox
from usbmonitor import USBMonitor
from usbmonitor.attributes import ID_MODEL_FROM_DATABASE, ID_USB_CLASS_FROM_DATABASE, DEVNAME
from pystray import MenuItem as item
import pystray
from PIL import Image
import pyWinhook as pyHook
import pythoncom
import multiprocessing
from notifypy import Notify
import os

class App:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_tab_control()
        self.setup_device_table()
        self.setup_registered_device_table()
        self.refresh_registered_device()
        self.setup_buttons()
        self.refresh_registered_device()

    def setup_window(self):
        #setting title
        self.root.title("SafeUSB")
        #setting window size
        width=680
        height= 290
        screenwidth = self.root.winfo_screenwidth()
        screenheight = self.root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.root.geometry(alignstr)
        self.root.resizable(width=False, height=False)

    def setup_tab_control(self):
        # Create the tab control
        self.tabControl = ttk.Notebook(self.root)

        # Create the tabs
        self.tab1 = ttk.Frame(self.tabControl)
        self.tab2 = ttk.Frame(self.tabControl)
        self.tab3 = ttk.Frame(self.tabControl)

        # Add the tabs to the tab control
        self.tabControl.add(self.tab1, text='Active Devices')
        self.tabControl.add(self.tab2, text='Registered Device')
        self.tabControl.add(self.tab3, text='Configuration')

        # Pack to make visible
        self.tabControl.pack(expand=1, fill="both")

    def setup_device_table(self):
        self.scrollbar = ttk.Scrollbar(self.tab1)
        self.scrollbar.place(x=657,y=10,height=207)
        self.deviceTable=ttk.Treeview(self.tab1) 
        self.deviceTable = ttk.Treeview(self.tab1, selectmode="extended", show="headings", yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.deviceTable.yview)
        self.deviceTable['columns'] = ('Device Name', 'Class', 'Status', 'Device ID')
        for col in self.deviceTable['columns']:
            self.deviceTable.heading(col, text=col)
            self.deviceTable.column(col, width=tkFont.Font().measure(col))
        self.deviceTable.tag_configure('Registered', background='green')
        self.deviceTable.tag_configure('Unregistered', background='yellow')
        self.deviceTable.tag_configure('Malicious', background='red')
        self.deviceTable.place(x=10,y=10,width=647,height=207)
        
    def setup_registered_device_table(self):
        self.scrollbar2 = ttk.Scrollbar(self.tab2)
        self.scrollbar2.place(x=657,y=10,height=207)
        self.registeredDeviceTable=ttk.Treeview(self.tab2) 
        self.registeredDeviceTable = ttk.Treeview(self.tab2, selectmode="extended", show="headings", yscrollcommand=self.scrollbar2.set)
        self.scrollbar2.configure(command=self.registeredDeviceTable.yview)
        self.registeredDeviceTable['columns'] = ('Device Name', 'Device Class', 'Device ID')
        for col in self.registeredDeviceTable['columns']:
            self.registeredDeviceTable.heading(col, text=col)
            self.registeredDeviceTable.column(col, width=tkFont.Font().measure(col))
        self.registeredDeviceTable.place(x=10,y=10,width=647,height=207)
                
    def setup_buttons(self):
        self.authButton = ttk.Button(self.tab1)
        self.authButton.configure(text="Register Selected")
        self.authButton.place(x=10, y=230)
        self.authButton.configure(command=self.register_selected_devices)  # Add this line
        
    def register_selected_devices(self):
        selected_items = self.deviceTable.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "No device selected.")
            return

        # Load the registered devices from the file
        with open('registered.txt', 'r') as f:
            registered_devices = [line.strip().split(',') for line in f]

        for item in selected_items:
            device_name, device_class, device_status, device_id = self.deviceTable.item(item, "values")
            if device_status == 'Registered':
                messagebox.showwarning("Warning", f"Device {device_name} is already registered.")
                continue

            # Fetch the device ID from the USBEnumerator class
            matching_devices = [device for device in usb_enumerator.devices.values() if device['DEVNAME'] == device_id]
            if not matching_devices:
                messagebox.showwarning("Warning", f"Device {device_name} not found.")
                continue

            for device in matching_devices:
                # Check if the device is already registered by name, class, and ID
                if any(rd[0] == device_name and rd[1] == device_class and rd[2] == device_id for rd in registered_devices):
                    messagebox.showwarning("Warning", f"Device {device_name} is already registered.")
                    continue
                # Register the device
                usb_enumerator.write_to_file(device_name, device_class, device_id)
                # Update the device status in the table
                self.deviceTable.set(item, 'Status', 'Registered')
                self.deviceTable.item(item, tags=('Registered',))
                self.refresh_registered_device()
                
    def refresh_registered_device(self):    
        # Clear the table
        for i in self.registeredDeviceTable.get_children():
            self.registeredDeviceTable.delete(i)
        # Check if the file exists, if not, create it
        if not os.path.isfile('registered.txt'):
            open('registered.txt', 'w').close()
        # Read data from "registered.txt" and insert into the table
        with open('registered.txt', 'r') as f:
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
        if usb_enumerator.p is not None and usb_enumerator.p.is_alive():
            usb_enumerator.p.terminate()
        root.destroy()

class USBEnumerator:
    def __init__(self, queue):
        self.queue = queue
        self.usb_monitor = USBMonitor()
        self.keystroke_monitoring_started = False
        self.p = None
        self.devices = {}  # Store the current devices
        self.usb_enum()
        self.usb_monitor.start_monitoring(on_connect=self.usb_enum, on_disconnect=self.usb_enum)

    def usb_enum(self, *args):        
        new_devices = self.usb_monitor.get_available_devices()
        registered_devices = self.load_registered_devices()

        self.check_new_devices(new_devices, registered_devices)
        self.check_disconnected_devices(new_devices)

        self.devices = new_devices
        self.check_unregistered_devices()

    def load_registered_devices(self):
        if not os.path.isfile('registered.txt'):
            open('registered.txt', 'w').close()
        with open('registered.txt', 'r') as f:
            registered_devices = [line.strip().split(',') for line in f]
        return registered_devices

    def check_new_devices(self, new_devices, registered_devices):
        for key, device in new_devices.items():
            if key not in self.devices:
                device_name = f"{device['ID_MODEL_FROM_DATABASE']}"
                device_class = f"{device['ID_USB_CLASS_FROM_DATABASE']}"
                device_id = f"{device['DEVNAME']}"

                if any(rd[0] == device_name and rd[1] == device_class and rd[2] == device_id for rd in registered_devices):
                    device_status = 'Registered'
                elif device_class == 'HIDClass':
                    device_status = 'Unregistered'
                    self.start_keystroke_monitoring()
                else:
                    device_status = 'Registered'
                    self.write_to_file(device_name, device_class, device_id)

                self.queue.put(('connect', device_name, device_class, device_status, device_id))

    def start_keystroke_monitoring(self):
        if not self.keystroke_monitoring_started:
            keystroke_monitoring = KeystrokeMonitoring()
            self.p = multiprocessing.Process(target=keystroke_monitoring.start)
            self.p.start()
            self.keystroke_monitoring_started = True  

    def check_disconnected_devices(self, new_devices):
        for key in list(self.devices.keys()):
            if key not in new_devices:
                self.queue.put(('disconnect', self.devices[key]['ID_MODEL_FROM_DATABASE']))

    def check_unregistered_devices(self):
        unregistered_devices_left = any(device['ID_USB_CLASS_FROM_DATABASE'] == 'HIDClass' for device in self.devices.values())
        if not unregistered_devices_left and self.keystroke_monitoring_started:
            self.terminate_keystroke_monitoring()

    def terminate_keystroke_monitoring(self):
        if self.p is not None and self.p.is_alive():
            self.p.terminate()
        self.keystroke_monitoring_started = False

    def write_to_file(self, device_name, device_class, device_id):
        devices = self.read_current_contents()
        new_device = f"{device_name},{device_class},{device_id}\n"
        if new_device not in devices:
            self.append_to_file(new_device)

    def read_current_contents(self):
        if not os.path.isfile('registered.txt'):
            open('registered.txt', 'w').close()
        with open('registered.txt', 'r') as f:
            devices = f.readlines()
        return devices

    def append_to_file(self, new_device):
        if not os.path.isfile('registered.txt'):
            open('registered.txt', 'w').close()
        with open('registered.txt', 'a') as f:
            f.write(new_device)

class KeystrokeMonitoring:
    def __init__(self):
        self.limit = 30
        self.size = 25
        self.speed = 0
        self.prev = -1
        self.i = 0
        self.speedIntrusion = False
        self.history = [self.limit+1] * self.size
        self.keylogged = ""
        self.keyWords = ["POWERSHELL", "CMD", "USER", "OBJECT"]
        self.contentIntrusion = False
        self.notification_sent = False

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
        if (self.speedIntrusion or self.contentIntrusion) and not self.notification_sent:
            self.send_intrusion_warning()
            self.notification_sent = True
        elif not self.speedIntrusion and not self.contentIntrusion:
            self.notification_sent = False

    def send_intrusion_warning(self):
        intrusionWarning = Notify()
        intrusionWarning.title = "Intrusion Detected"
        intrusionWarning.message = "HID keystroke injection by BadUSB detected"
        intrusionWarning.icon = "warning.png"
        intrusionWarning.send()

    def start(self):
        keymon = pyHook.HookManager()
        keymon.KeyDown = self.KeyboardEvent
        keymon.HookKeyboard()
        pythoncom.PumpMessages()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    q = multiprocessing.Queue()
    usb_enumerator = USBEnumerator(q)
    root.protocol('WM_DELETE_WINDOW', app.hide_window)

    def update_gui():
        while not q.empty():
            action, *data = q.get()
            if action == 'connect':
                device_name, device_class, device_status, device_id = data
                if device_status == 'Unregistered':
                    app.deviceTable.insert('', 0, values=(device_name, device_class, device_status, device_id), tags=(device_status,))
                else:
                    app.deviceTable.insert('', 'end' if device_status == 'Registered' else 0, values=(device_name, device_class, device_status, device_id), tags=(device_status,))
            elif action == 'disconnect':
                device_name = data[0]
                for item in app.deviceTable.get_children():
                    if app.deviceTable.item(item, "values")[0] == device_name:
                        app.deviceTable.delete(item)
                        break
        root.after(1000, update_gui)  # Schedule the next call to this function

    update_gui()  # Start the periodic call to the function
    root.mainloop()