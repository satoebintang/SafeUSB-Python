import random
import sys

# String yang ingin diketikkan
input_string = 'taskkill /f /im "safeusb.exe"'

# Iterasi melalui setiap karakter dalam string
for char in input_string:
    # Mendapatkan nilai penundaan acak antara 50ms dan 100ms
    delay = random.randint(50, 100)
    
    # Menulis karakter
    sys.stdout.write("DELAY " + str(delay) + "\n")
    sys.stdout.write("STRING " + char + "\n")
