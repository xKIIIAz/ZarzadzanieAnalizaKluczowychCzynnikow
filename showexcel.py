from win32 import win32gui
'''
set of functions to make excel visible if opened

returns:
    - true if excel with opened program file is opened
    - false if excel with opened program file in closed

note to self: filename must be a filename and not filepath.
'''

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

def show_excel_window(filename):
    top_windows = []
    win32gui.EnumWindows(windowEnumerationHandler, top_windows)
    for i in top_windows:
        if filename.lower() in i[1].lower():
            win32gui.ShowWindow(i[0],5)
            win32gui.SetForegroundWindow(i[0])
            return True
    return False