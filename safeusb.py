import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font as tkFont
from usbmonitor import USBMonitor
from usbmonitor.attributes import ID_MODEL_FROM_DATABASE, ID_USB_CLASS_FROM_DATABASE
from pystray import MenuItem as item
import pystray
from PIL import Image
import pyWinhook as pyHook
import pythoncom
import multiprocessing
from notifypy import Notify

class App:
    def __init__(self, root):
        #setting title
        root.title("SafeUSB")
        #setting window size
        width=680
        height= 290
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        # Create the tab control
        tabControl = ttk.Notebook(root)

        # Create the tabs
        self.tab1 = ttk.Frame(tabControl)
        self.tab2 = ttk.Frame(tabControl)

        # Add the tabs to the tab control
        tabControl.add(self.tab1, text='Active Devices')
        tabControl.add(self.tab2, text='Protection History')

        # Pack to make visible
        tabControl.pack(expand=1, fill="both")
        
        self.scrollbar = ttk.Scrollbar(self.tab1)
        self.scrollbar.place(x=657,y=10,height=207)
        self.deviceTable=ttk.Treeview(self.tab1) 
        self.deviceTable = ttk.Treeview(self.tab1, selectmode="extended", show="headings", yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.deviceTable.yview)
        self.deviceTable['columns'] = ('Device Name', 'Class', 'Status')
        for col in self.deviceTable['columns']:
            self.deviceTable.heading(col, text=col)
            self.deviceTable.column(col, width=tkFont.Font().measure(col))
        self.deviceTable.tag_configure('Safe', background='green')
        self.deviceTable.tag_configure('Suspicious', background='yellow')
        self.deviceTable.tag_configure('Malicious', background='red')
        self.deviceTable.place(x=10,y=10,width=647,height=207)
        
        # Create a button to authorize all device
        self.authButton = ttk.Button(self.tab1)
        self.authButton.configure(text="Authorize Selected")
        self.authButton.place(x=10, y=230)
        
        self.authAllButton = ttk.Button(self.tab1)
        self.authAllButton.configure(text="Authorize All")
        self.authAllButton.place(x=120, y=230)    

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
        root.destroy()

class USBEnumerator:
    def __init__(self, deviceTable):
        self.deviceTable = deviceTable
        self.usb_monitor = USBMonitor()
        self.keystroke_monitoring_started = False
        self.p = None
        self.devices = {}  # Store the current devices
        self.usb_enum()
        self.usb_monitor.start_monitoring(on_connect=self.usb_enum, on_disconnect=self.usb_enum)

    def insert_device_details(self, device_name, device_class, device_status):
        # Insert the device details into the deviceTable
        if device_status == 'Suspicious':
            self.deviceTable.insert('', 0, values=(device_name, device_class, device_status), tags=(device_status,))
        else:
            self.deviceTable.insert('', 'end' if device_status == 'Safe' else 0, values=(device_name, device_class, device_status), tags=(device_status,))
            
    def usb_enum(self, *args):        
        new_devices = self.usb_monitor.get_available_devices()

        # Check for new devices
        for key, device in new_devices.items():
            if key not in self.devices:
                # Get the device details
                device_name = f"{device['ID_MODEL_FROM_DATABASE']}"
                device_class = f"{device['ID_USB_CLASS_FROM_DATABASE']}"
                # Check if the device class is 'HIDClass'
                if device_class == 'HIDClass':
                    device_status = 'Suspicious'
                    if not self.keystroke_monitoring_started:
                        keystroke_monitoring = KeystrokeMonitoring()
                        self.p = multiprocessing.Process(target=keystroke_monitoring.start)
                        self.p.start()
                        self.keystroke_monitoring_started = True  
                else:
                    device_status = 'Safe'
                self.insert_device_details(device_name, device_class, device_status)

        # Check for disconnected devices
        for key in list(self.devices.keys()):
            if key not in new_devices:
                # Find the item in the deviceTable and remove it
                for item in self.deviceTable.get_children():
                    if self.deviceTable.item(item, "values")[0] == self.devices[key]['ID_MODEL_FROM_DATABASE']:
                        self.deviceTable.delete(item)
                        break

        self.devices = new_devices

        # Check if there are any suspicious devices left
        suspicious_devices_left = any(device['ID_USB_CLASS_FROM_DATABASE'] == 'HIDClass' for device in self.devices.values())
        if not suspicious_devices_left and self.keystroke_monitoring_started:
            # Terminate the keystroke monitoring process
            if self.p is not None and self.p.is_alive():
                self.p.terminate()
            self.keystroke_monitoring_started = False
        
class KeystrokeMonitoring:
    def __init__(self):
        self.limit = 30
        self.size = 25
        self.speed = 0
        self.prev = -1
        self.i = 0
        self.speedIintrusion = False
        self.history = [self.limit+1] * self.size
        self.keylogged = ""
        self.keyWords = ["POWERSHELL", "CMD", "USER", "OBJECT"]
        self.contentIntrusion = False
        self.notification_sent = False

    def KeyboardEvent(self, event):
        print("Keystroke : " + event.Key)
        self.keylogged += event.Key

        for word in self.keyWords:
            if word in self.keylogged.upper():
                print(f"[*] Key Words Detected: [{word}]")
                self.contentIntrusion = True
                self.keylogged = ""

        if (self.prev == -1):
            self.prev = event.Time
            return True

        if (self.i >= len(self.history)): self.i = 0

        self.history[self.i] = event.Time - self.prev
        print(event.Time, "-", self.prev, "=", self.history[self.i])
        self.prev = event.Time
        self.speed = sum(self.history) / float(len(self.history))
        self.i = self.i + 1

        print("\rAverage Typing Speed:", self.speed)

        if (self.speed < self.limit):
            self.speedIintrusion = True
            # do something (plan : mark device as malicious, disconnect device and add to blacklist)
        else:
            self.speedIintrusion = False

        if (self.speedIintrusion or self.contentIntrusion) and not self.notification_sent:
            intrusionWarning = Notify()
            intrusionWarning.title = "Intrusion Detected"
            intrusionWarning.message = "HID keystroke injection by BadUSB detected"
            intrusionWarning.icon = "warning.png"
            intrusionWarning.send()
            self.notification_sent = True  # Set the flag to True after sending a notification
        elif not self.speedIintrusion and not self.contentIntrusion:
            self.notification_sent = False  # Reset the flag when the conditions are no longer met
        return True

    def start(self):
        keymon = pyHook.HookManager()
        keymon.KeyDown = self.KeyboardEvent
        keymon.HookKeyboard()
        pythoncom.PumpMessages()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    usb_enumerator = USBEnumerator(app.deviceTable)
    root.protocol('WM_DELETE_WINDOW', app.hide_window)
    root.mainloop()
