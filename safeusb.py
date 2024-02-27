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
        width=670
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
        
        self.deviceTable=ttk.Treeview(self.tab1)  # Change this line
        self.deviceTable.configure(selectmode="extended", show="headings")
        self.deviceTable['columns'] = ('Device Name', 'Class', 'Status')
        for col in self.deviceTable['columns']:
            self.deviceTable.heading(col, text=col)
            self.deviceTable.column(col, width=tkFont.Font().measure(col))  # Add this line
        self.deviceTable.tag_configure('Safe', background='green')
        self.deviceTable.tag_configure('Suspicious', background='yellow')
        self.deviceTable.tag_configure('Malicious', background='red')
        self.deviceTable.place(x=10,y=10,width=647,height=207)
        
        # Create a button to authorize all device
        self.authAllButton = ttk.Button(self.tab1)
        self.authAllButton.configure(text="Authorize Suspicious Device")
        self.authAllButton.place(x=10, y=230)  
          
        self.usb_monitor = USBMonitor()
        self.usb_enumerator = USBEnumerator(self.deviceTable, self.usb_monitor)
        self.usb_enumerator.usb_enum()
        self.usb_monitor.start_monitoring(on_connect=self.usb_enumerator.usb_enum, on_disconnect=self.usb_enumerator.usb_enum)

class USBEnumerator:
    def __init__(self, deviceTable, usb_monitor):
        self.deviceTable = deviceTable
        self.usb_monitor = usb_monitor
        self.keystroke_monitoring_started = False
        self.p = None

    def insert_device_details(self, device_name, device_class, device_status):
        # Insert the device details into the deviceTable
        self.deviceTable.insert('', 'end', values=(device_name, device_class, device_status), tags=(device_status,))

    def usb_enum(self, *args):        
        # Clear the deviceTable
        self.deviceTable.delete(*self.deviceTable.get_children())
        devices = self.usb_monitor.get_available_devices()
        suspicious_device_detected = False
        for key, device in devices.items():
            # Get the device details
            device_name = f"{device['ID_MODEL_FROM_DATABASE']}"
            device_class = f"{device['ID_USB_CLASS_FROM_DATABASE']}"
            # Check if the device class is 'HIDClass'
            if device_class == 'HIDClass':
                device_status = 'Suspicious'
                suspicious_device_detected = True
                if not self.keystroke_monitoring_started:
                    keystroke_monitoring = KeystrokeMonitoring()
                    self.p = multiprocessing.Process(target=keystroke_monitoring.start)
                    self.p.start()
                    self.keystroke_monitoring_started = True  
            else:
                device_status = 'Safe'
            self.insert_device_details(device_name, device_class, device_status)
        if not suspicious_device_detected and self.keystroke_monitoring_started:
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
        self.intrusion = False
        self.history = [self.limit+1] * self.size
        self.keylogged = ""
        self.keyWords = ["POWERSHELL", "CMD", "USER", "OBJECT"]
        self.keystroke_words = False

    def KeyboardEvent(self, event):
        print("Keystroke : " + event.Key)
        self.keylogged += event.Key

        for word in self.keyWords:
            if word in self.keylogged.upper():
                print(f"[*] Key Words Detected: [{word}]")
                self.keystroke_words = True
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
            self.intrusion = True
            # do something (plan : mark device as malicious, disconnect device and add to blacklist)
        else:
            self.intrusion = False

        if self.intrusion:
            intrusionWarning = Notify()
            intrusionWarning.title = "Intrusion Detected"
            intrusionWarning.message = "HID keystroke injection by BadUSB detected"
            intrusionWarning.icon = "warning.png"
            intrusionWarning.send()
        return True

    def start(self):
        keymon = pyHook.HookManager()
        keymon.KeyDown = self.KeyboardEvent
        keymon.HookKeyboard()
        pythoncom.PumpMessages()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    # Define a function for quit the window
    def quit_window(icon, item):
        icon.stop()
        root.destroy()
    
    # Define a function to show the window again
    def show_window(icon, item):
        icon.stop()
        root.after(0,lambda: root.deiconify())
    
    # Hide the window and show on the system taskbar
    def hide_window():
        # ignore pylance issues
        runNotify = Notify()
        runNotify.title = "SafeUSB is active"
        runNotify.message = "SafeUSB is running in the background"
        runNotify.icon = "information.png"
        runNotify.send()
        root.withdraw()
        image=Image.open("favicon.ico")
        menu=(item('Show', show_window), item('Quit', quit_window))
        icon=pystray.Icon("name", image, "SafeUSB", menu)
        icon.run()
    
    root.protocol('WM_DELETE_WINDOW', hide_window)
    root.mainloop()
