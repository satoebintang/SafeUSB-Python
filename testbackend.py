import tkinter as tk
import tkinter.ttk as ttk
import tkinter.font as tkFont
import wmi

# Define a function to get the USB devices from WMI
def usbenum():
    c = wmi.WMI()
    devices = []
    # Use a WQL query to filter the devices by their DeviceID
    for device in c.query("SELECT * FROM Win32_PnPEntity WHERE DeviceID LIKE 'USB%'"):
        devices.append({
            'FriendlyName': device.Name,
            'Class': device.PNPClass,
        })
    return devices

# Define a function to update the treeview with the USB devices
def update_treeview(devicelist, devices):
    # Clear the treeview
    devicelist.delete(*devicelist.get_children())
    # Insert the devices with tags for coloring
    for device in devices:
        status = "Safe" # For now, mark all devices as safe
        devicelist.insert('', 'end', values=(device['FriendlyName'], device['Class'], status), tags=(status,))
    # Adjust the column widths to fit the content
    for col in devicelist['columns']:
        devicelist.column(col, width=tkFont.Font().measure(col.title()))
        for item in devicelist.get_children():
            cell_value = devicelist.item(item, 'values')[devicelist['columns'].index(col)]
            cell_width = tkFont.Font().measure(cell_value)
            if devicelist.column(col, 'width') < cell_width:
                devicelist.column(col, width=cell_width)

class NewprojectApp:
    def __init__(self, master=None):
        # build ui
        self.mainwindow = tk.Tk() if master is None else tk.Toplevel(master)
        self.mainwindow.configure(height=200, width=200)
        self.mainwindow.geometry("900x550")
        self.mainwindow.resizable(False, False)
        self.mainwindow.title("SafeUSB")
        
        # Create a treeview to display the USB devices
        self.devicelist = ttk.Treeview(self.mainwindow)
        self.devicelist.configure(selectmode="extended", show="headings")
        self.devicelist['columns'] = ('Device Name', 'Class', 'Status')
        # Create the headings for the columns
        for col in self.devicelist['columns']:
            self.devicelist.heading(col, text=col)
        # Create the tags for coloring the rows
        self.devicelist.tag_configure('Safe', background='green')
        self.devicelist.tag_configure('Suspicious', background='yellow')
        self.devicelist.tag_configure('Potentially Malicious', background='orange')
        self.devicelist.tag_configure('Malicious', background='red')
        # Place the treeview in the window
        self.devicelist.place(anchor="nw", height=450, width=650, x=30, y=30)
        
        # Create a button to authorize the selected device
        button1 = ttk.Button(self.mainwindow)
        button1.configure(text="Authorize Device")
        button1.place(anchor="nw", x=30, y=500)

        # Create radio buttons to select the mode
        self.mode = tk.IntVar()  # Create a variable to store the mode

        self.radio_normalmode = ttk.Radiobutton(self.mainwindow)
        self.radio_normalmode.configure(text="Normal Mode", value=1, variable=self.mode, command=self.handle_mode_selection)
        self.radio_normalmode.place(anchor="nw", x=700, y=30)

        self.radio_paranoidmode = ttk.Radiobutton(self.mainwindow, command=self.handle_mode_selection)
        self.radio_paranoidmode.configure(
            text="Paranoid Mode", value=2, variable=self.mode
        )
        self.radio_paranoidmode.place(anchor="nw", x=700, y=60)

        self.radio_superparanoidmode = ttk.Radiobutton(self.mainwindow, command=self.handle_mode_selection)
        self.radio_superparanoidmode.configure(
            text="Super Paranoid Mode", value=3, variable=self.mode
        )
        self.radio_superparanoidmode.place(anchor="nw", x=700, y=90)

        self.mode.set(1)  # Set the default mode to 1
        self.handle_mode_selection()

        # Main widget
        self.mainwindow = self.mainwindow

        # Get the initial list of USB devices
        devices = usbenum()
        # Update the treeview with the devices
        update_treeview(self.devicelist, devices)
        # Convert the list to a set of tuples
        self.device_set = set((device['FriendlyName'], device['Class']) for device in devices)
        # Schedule the periodic check for USB events
        self.check_usb_events()
        self.hid_check()

    # Define a method to check for USB events
    def check_usb_events(self):
        # Get the current list of USB devices
        devices = usbenum()
        # Convert the list to a set of tuples
        device_set = set((device['FriendlyName'], device['Class']) for device in devices)
        # Check if the set has changed
        if device_set.symmetric_difference(self.device_set):
            # Update the treeview with the new devices
            update_treeview(self.devicelist, devices)
            # Save the new set
            self.device_set = device_set
        # Schedule the next check
        self.mainwindow.after(1000, self.check_usb_events)
        self.hid_check()

    def hid_check(self):
        hid_devices = [device for device in usbenum() if device['Class'] == 'HIDClass']
        for item in self.devicelist.get_children():
            values = self.devicelist.item(item, 'values')
            if (values[0], values[1]) in [(device['FriendlyName'], device['Class']) for device in hid_devices]:
                self.devicelist.item(item,values=(values[0], values[1], 'Suspicious'), tags=('Suspicious',))
     
    def handle_mode_selection(self):
        selected_mode = self.mode.get()
        print("Selected Mode:", selected_mode)
        
        # Add logic based on the selected mode
        if selected_mode == 1:
            print("Normal Mode Selected: Perform actions for Normal Mode")
            self.hid_check()
        elif selected_mode == 2:
            print("Paranoid Mode Selected: Perform actions for Paranoid Mode")
        elif selected_mode == 3:
            print("Super Paranoid Mode Selected: Perform actions for Super Paranoid Mode")
    
    def run(self):
        self.mainwindow.mainloop()

if __name__ == "__main__":
    app = NewprojectApp()
    app.run()
