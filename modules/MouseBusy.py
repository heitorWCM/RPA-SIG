import win32gui
import win32con
import time
import sys

def is_mouse_busy():
    """
    Checks if the mouse cursor is currently in a busy or waiting state on Windows.
    Returns True if busy, False otherwise.
    """
    cursor_info = win32gui.GetCursorInfo()
    current_cursor = cursor_info[1]
    
    # Load the system busy/waiting cursors to get their handles
    wait_cursor = win32gui.LoadCursor(0, win32con.IDC_WAIT)
    appstarting_cursor = win32gui.LoadCursor(0, win32con.IDC_APPSTARTING)
    
    # Compare the current cursor handle with busy cursor handles
    return current_cursor in [wait_cursor, appstarting_cursor]

def MouseBusy(timeout=300):
    
    start = time.time()

    tempoAtual = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(start))
    print(f"Entou em loop - {tempoAtual}")

    while True:
        if is_mouse_busy():
            print("Mouse is busy, waiting...")
        else:
            break

        # Timeout
        if time.time() - start >= timeout:
            print("Timeout reached — quitting script.")
            sys.exit()

        time.sleep(2)

# Example usage:
if __name__ == "__main__":
    print("Checking mouse cursor status...")
    print("Move your mouse over a loading application or website to test.\n")
    
    try:
        while True:
            if is_mouse_busy():
                print("🕐 Mouse is busy/waiting.")
            else:
                print("✓ Mouse is normal.")
            time.sleep(0.5)  # Check twice per second for better responsiveness
    except KeyboardInterrupt:
        print("\nStopped monitoring.")