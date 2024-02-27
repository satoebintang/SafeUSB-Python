import win32ui
import pyWinhook as pyHook
import pythoncom
from keyboard import block_key

limit = 30 #Speed Threshold between keystrokes in Milliseconds
size = 25 # Size of array that holds the history of keystroke speeds
speed = 0
prev = -1
i = 0
intrusion = False
history = [limit+1] * size
keylogged = ""
keyWords = ["POWERSHELL", "CMD", "USER", "OBJECT"]
keystroke_words = False

def KeyboardEvent(event):
    global limit
    global speed, prev, i, history, intrusion, keylogged, keystroke_words
    
    print("Keystroke : " + event.Key)
    keylogged += event.Key
    
    # Check for keywords
    for word in keyWords:
        if word in keylogged.upper():
            print(f"[*] Key Words Detected: [{word}]")
            keystroke_words = True
            keylogged = ""  # Reset the keylogged
    
    # Initial Condition
    if (prev == -1):
        prev = event.Time
        return True

    if (i >= len(history)): i = 0
    
    # TypeSpeed = NewKeyTime - OldKeyTime
    history[i] = event.Time - prev
    print(event.Time, "-", prev, "=", history[i])
    prev = event.Time
    speed = sum(history) / float(len(history))
    i = i + 1
    
    print("\rAverage Typing Speed:", speed)
    
    # Intrusion detected
    if (speed < limit):
        intrusion = True
        for p in range(150):
            block_key(p)
    
    else:
        intrusion = False
        
    if intrusion:
        win32ui.MessageBox("HID keystroke injection by BadUSB detected.\nAll keyboard inputs will be blocked for 10 seconds.\nPlease check your USB port for malicious USB device.", "BadUSB Detected", 4096)
    
    # pass execution to next hook registered
    return True

keymon = pyHook.HookManager()
keymon.KeyDown = KeyboardEvent
keymon.HookKeyboard()
pythoncom.PumpMessages()
