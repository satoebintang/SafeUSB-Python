from usbmonitor import USBMonitor
from usbmonitor.attributes import ID_MODEL, ID_MODEL_ID, ID_VENDOR_ID, ID_USB_CLASS_FROM_DATABASE, ID_MODEL_FROM_DATABASE

# Create the USBMonitor instance
monitor = USBMonitor()

# Get the current devices
devices_dict = monitor.get_available_devices()

# Print them
for device_id, device_info in devices_dict.items():
    print(f"{device_id} -- {device_info[ID_MODEL_FROM_DATABASE]} - {device_info[ID_USB_CLASS_FROM_DATABASE]})")