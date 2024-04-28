import pygetwindow as gw
import psutil
import win32process
from ctypes import Array, byref, c_char, memset, sizeof
from ctypes import c_int, c_void_p, POINTER
from ctypes.wintypes import *
from enum import Enum
import ctypes
from PIL import Image, ImageTk
from io import BytesIO
import tkinter as tk
from tkinter import ttk

BI_RGB = 0
DIB_RGB_COLORS = 0


class ICONINFO(ctypes.Structure):
    _fields_ = [
        ("fIcon", BOOL),
        ("xHotspot", DWORD),
        ("yHotspot", DWORD),
        ("hbmMask", HBITMAP),
        ("hbmColor", HBITMAP)
    ]

class RGBQUAD(ctypes.Structure):
    _fields_ = [
        ("rgbBlue", BYTE),
        ("rgbGreen", BYTE),
        ("rgbRed", BYTE),
        ("rgbReserved", BYTE),
    ]

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", DWORD),
        ("biWidth", LONG),
        ("biHeight", LONG),
        ("biPlanes", WORD),
        ("biBitCount", WORD),
        ("biCompression", DWORD),
        ("biSizeImage", DWORD),
        ("biXPelsPerMeter", LONG),
        ("biYPelsPerMeter", LONG),
        ("biClrUsed", DWORD),
        ("biClrImportant", DWORD)
    ]

class BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", RGBQUAD * 1),
    ]


shell32 = ctypes.WinDLL("shell32", use_last_error=True)
user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

gdi32.CreateCompatibleDC.argtypes = [HDC]
gdi32.CreateCompatibleDC.restype = HDC
gdi32.GetDIBits.argtypes = [
    HDC, HBITMAP, UINT, UINT, LPVOID, c_void_p, UINT
]
gdi32.GetDIBits.restype = c_int
gdi32.DeleteObject.argtypes = [HGDIOBJ]
gdi32.DeleteObject.restype = BOOL
shell32.ExtractIconExW.argtypes = [
    LPCWSTR, c_int, POINTER(HICON), POINTER(HICON), UINT
]
shell32.ExtractIconExW.restype = UINT
user32.GetIconInfo.argtypes = [HICON, POINTER(ICONINFO)]
user32.GetIconInfo.restype = BOOL
user32.DestroyIcon.argtypes = [HICON]
user32.DestroyIcon.restype = BOOL


class IconSize(Enum):
    SMALL = 1
    LARGE = 2

    @classmethod
    def to_wh(cls, size: "IconSize") -> tuple[int, int]:
        """
        Return the actual (width, height) values for the specified icon size.
        """
        size_table = {
            cls.SMALL: (16, 16),
            cls.LARGE: (32, 32)
        }
        return size_table[size]


def extract_icon(filename: str, size: IconSize) -> Array[c_char]:
    """
    Extract the icon from the specified `filename`, which might be
    either an executable or an `.ico` file.
    """
    dc: HDC = gdi32.CreateCompatibleDC(0)
    if dc == 0:
        raise ctypes.WinError()

    hicon: HICON = HICON()
    extracted_icons: UINT = shell32.ExtractIconExW(
        filename,
        0,
        byref(hicon) if size == IconSize.LARGE else None,
        byref(hicon) if size == IconSize.SMALL else None,
        1
    )
    if extracted_icons != 1:
        print(f"Error extracting icon from {filename}, returned {extracted_icons} icons instead of one.")
        return

    def cleanup() -> None:
        if icon_info.hbmColor != 0:
            gdi32.DeleteObject(icon_info.hbmColor)
        if icon_info.hbmMask != 0:
            gdi32.DeleteObject(icon_info.hbmMask)
        user32.DestroyIcon(hicon)

    icon_info: ICONINFO = ICONINFO(0, 0, 0, 0, 0)
    if not user32.GetIconInfo(hicon, byref(icon_info)):
        cleanup()
        raise ctypes.WinError()

    w, h = IconSize.to_wh(size)
    bmi: BITMAPINFO = BITMAPINFO()
    memset(byref(bmi), 0, sizeof(bmi))
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = w
    bmi.bmiHeader.biHeight = -h
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = BI_RGB
    bmi.bmiHeader.biSizeImage = w * h * 4
    bits = ctypes.create_string_buffer(bmi.bmiHeader.biSizeImage)
    copied_lines = gdi32.GetDIBits(
        dc, icon_info.hbmColor, 0, h, bits, byref(bmi), DIB_RGB_COLORS
    )
    if copied_lines == 0:
        cleanup()
        raise RuntimeError(f"Error copying bitmap bits from icon in {filename}.")

    cleanup()
    return bits


def get_open_windows():
    windows = gw.getAllWindows()
    return windows

def get_process_id(hwnd):
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return pid

def is_system_process(pid):
    try:
        process = psutil.Process(pid)
        return process.name().lower() in {'system', 'idle'} or 'system32' in process.exe().lower()
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        return True  # Treat as system process if it fails to retrieve process information

def get_executable_path(pid):
    try:
        process = psutil.Process(pid)
        return process.exe()
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        return None

def main():
    open_windows = get_open_windows()
    for window in open_windows:
        hwnd = window._hWnd
        pid = get_process_id(hwnd)
        if not is_system_process(pid):
            executable_path = get_executable_path(pid)
            if executable_path:
                print(f"Title: {window.title}, Executable Path: {executable_path}")
                result = extract_icon(executable_path, IconSize.SMALL)
                if result is not None:
                    print(result)
                else:
                    print("No small icon found in the executable")

def load_icon(executable_path):
    # Extract icon from executable path
    icon_data = extract_icon(executable_path, IconSize.SMALL)
    if icon_data is not None:
        # Convert icon data to an image
        image = Image.frombytes('RGBA', (16, 16), icon_data, 'raw', 'BGRA', 0, 1)
        # Resize image to fit the list
        image = image.resize((15, 15))
        # Convert image to PhotoImage
        return ImageTk.PhotoImage(image=image)
    else:
        return None

def populate_list(tree):
    open_windows = get_open_windows()
    for window in open_windows:
        hwnd = window._hWnd
        pid = get_process_id(hwnd)
        if not is_system_process(pid):
            executable_path = get_executable_path(pid)
            if executable_path and window.title:
                # Load icon
                icon = load_icon(executable_path)
                if icon:
                    # Create a label to display the icon
                    icon_label = ttk.Label(tree, image=icon, compound='left', padding=(5, 0, 0, 0))
                    icon_label.image = icon  # Keep a reference to the image to prevent garbage collection
                    # Insert program with icon into the treeview
                    tree.insert('', 'end', values=(window.title,), image=icon, tags=('custom',))


def main():
    # Create main window
    root = tk.Tk()
    root.title("Program List")

    root.geometry("400x400")

    # Create treeview widget
    tree = ttk.Treeview(root)
    tree['columns'] = ('Icon', 'Window')  # Two columns: Icon and Program
    tree.heading('#0', text='')  # Heading for icon column
    tree.heading('#1', text='Window')  # Heading for program column
    tree.column('#0', width=50, stretch=False)  # Set width of icon column
    tree.column('#1', stretch=True)  # Set stretch=True to make the program column fill the window width
    tree.pack(fill='both', expand=True)

    # Function to update column width based on window width
    def update_column_width(event):
        width = root.winfo_width() - 60
        tree.column('#1', width=width)

    # Bind window resize event to update_column_width function
    root.bind("<Configure>", update_column_width)

    # Populate treeview with program list and icons
    populate_list(tree)

    # Start Tkinter event loop
    root.mainloop()

if __name__ == "__main__":
    main()
