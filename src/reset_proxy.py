import pyautogui
import time
import subprocess

def reset_IE_to_defaults():
    command = "RunDll32.exe InetCpl.cpl,ResetIEtoDefaults"
    try:
        # Start the command and don't wait for it to complete
        subprocess.Popen(command, shell=True)
        print("Command executed, waiting for window...")
    except subprocess.CalledProcessError as e:
        print("Failed to execute command:", e)
        return

    # Wait for the window to appear
    time.sleep(2)  # Adjust if necessary

    # Use a loop to wait for the window to appear
    window_found = False
    timeout = time.time() + 10  # 10-second timeout
    while not window_found and time.time() < timeout:
        window_found = pyautogui.getActiveWindow().title.startswith("Reset Internet Explorer Settings")
        time.sleep(1)  # Check every second

    if not window_found:
        print("The 'Reset Internet Explorer Settings' window did not appear.")
        return

    print("Window found, sending keystrokes...")
    # Send the necessary keystrokes to navigate and press "Reset"
    pyautogui.hotkey('shift', 'tab')  # Press Shift+Tab to navigate backwards in the dialog
    pyautogui.press('enter')  # Press Enter to confirm reset

    # Wait for a potential confirmation dialog after the reset
    time.sleep(10)  # Wait for 5 seconds before pressing Enter again
    pyautogui.press('enter')  # Press Enter to close the confirmation dialog

    print("Reset process completed.")


