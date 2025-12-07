import usb.core

print("Listing all connected USB devices:")

devices = usb.core.find(find_all=True)
for dev in devices:
    print(f"VID={hex(dev.idVendor)}, PID={hex(dev.idProduct)}")