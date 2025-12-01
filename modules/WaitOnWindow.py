import pygetwindow as gw
import time


def WaitOnWindow(window_name, wait=0.5, timeout=60):
    """
    Waits for a window with the specified name to appear within the given timeout period.

    :param window_name: The title of the window to wait for.
    :param timeout: The maximum time to wait in seconds.
    :return: The window object if found, None otherwise.
    """

    # Convert a single image into a list
    if isinstance(window_name, str):
        window_list = [window_name]
    else:
        window_list = window_name

    start_time = time.time()
    while True:
        for window in window_list:
            print(f"Checking for window: {window}")
            windows_found = gw.getWindowsWithTitle(window)
            if windows_found:
                print(f"Window found: {window}")
                return windows_found[0]
            elif time.time() - start_time > timeout:
                print(f"Timeout reached while waiting for window: {window}")
                return None
            time.sleep(wait)

# Example usage:
if __name__ == "__main__":
    window = WaitOnWindow("PROSYST [\PROSYST\WPROSYST\PR07416.RPT ]", timeout=30)
    if window:
        print("Successfully found the window.")
    else:
        print("Failed to find the window within the timeout period.")