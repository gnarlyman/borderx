import pygetwindow as gw
import psutil
import win32process
from ctypes import Array, byref, c_char, memset, sizeof
from ctypes import c_int, c_void_p, POINTER
from ctypes.wintypes import *
from enum import Enum
import ctypes
from PIL import Image, ImageTk, ImageDraw
import tkinter as tk
from tkinter import ttk
import win32gui
import win32con
from pystray import Icon as icon, MenuItem as item, Menu
import threading

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

def create_image():
    # Create a simple black and white image for the tray icon.
    image = Image.new('RGB', (64, 64), color='black')
    dc = ImageDraw.Draw(image)

    # Draw a black "X" on the white background
    # Drawing two diagonal lines from the corners of a square inside the image
    dc.line((8, 8, 56, 56), fill='white', width=5)  # Diagonal from top-left to bottom-right
    dc.line((8, 56, 56, 8), fill='white', width=5)  # Diagonal from bottom-left to top-right

    return image

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

def populate_list():
    tree.delete(*tree.get_children())
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

def make_borderless_fullscreen(window_title):
    try:
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window = gw.getWindowsWithTitle(window_title)[0]
        hwnd = window._hWnd
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        style &= ~(win32con.WS_CAPTION | win32con.WS_THICKFRAME)
        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, screen_width, screen_height, win32con.SWP_FRAMECHANGED)
        print(f'Made "{window_title}" borderless and fullscreen.')
    except Exception as e:
        print(f"Error: {e}")

def on_select(event):
    selected_item = event.widget.selection()[0]
    global selected_window
    selected_window = event.widget.item(selected_item, 'values')[0]
    print("Selected program:", selected_window)

def on_button_click():
    make_borderless_fullscreen(selected_window)
    print(f'Made "{selected_window}" borderless and fullscreen.')

def toggle_window_visibility(icon, item=None):
    if root.state() == 'withdrawn':
        root.deiconify()  # Show the window
    else:
        root.withdraw()  # Hide the window

def exit_application(icon, item):
    icon.stop()  # Stop the tray icon
    root.after(0, root.quit)  # Schedule the root.quit to be run on the main thread

def run_tray_icon():
    tray = icon('Window Manager', create_image(), menu=Menu(item('Toggle Window', toggle_window_visibility), item('Exit', exit_application)))
    tray.run()

def set_icon(window, image):
    # Convert PIL image to PhotoImage and set as Tkinter window icon
    photo = ImageTk.PhotoImage(image)
    window.iconphoto(False, photo)
    return photo  # Return photo to keep a reference

def main():
    # Create main window
    global root
    root = tk.Tk()
    root.title("Program List")

    root.geometry("400x400")
    # Generate the icon image and set it as the window icon
    icon_image = create_image()
    icon_photo = set_icon(root, icon_image)

    # Create treeview widget
    global tree
    tree = ttk.Treeview(root)
    tree['columns'] = ('Icon', 'Window')  # Two columns: Icon and Program
    tree.heading('#0', text='')  # Heading for icon column
    tree.heading('#1', text='Window')  # Heading for program column
    tree.column('#0', width=50, stretch=False)  # Set width of icon column
    tree.column('#1', stretch=True)  # Set stretch=True to make the program column fill the window width
    tree.pack(fill='both', expand=True)
    tree.bind('<<TreeviewSelect>>', on_select)
    refresh_button = ttk.Button(root, text="Refresh Window List", command=populate_list)
    refresh_button.pack(pady=5)
    borderless_button = ttk.Button(root, text="Make Borderless Fullscreen", command=on_button_click)
    borderless_button.pack(pady=5)

    # Function to update column width based on window width
    def update_column_width(event):
        width = root.winfo_width()
        tree.column('#1', width=width)

    # Bind window resize event to update_column_width function
    root.bind("<Configure>", update_column_width)

    # Populate treeview with program list and icons
    populate_list()

    # Run the tray icon in a separate thread
    thread = threading.Thread(target=run_tray_icon)
    thread.start()

    # Start Tkinter event loop
    root.mainloop()

if __name__ == "__main__":
    main()
