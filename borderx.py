import threading
import tkinter as tk
from tkinter import ttk
import pygetwindow as gw
import win32gui
import win32con
from PIL import Image, ImageDraw
from pystray import Icon as icon, MenuItem as item, Menu

def create_image():
    # Create a simple black and white image for the tray icon.
    image = Image.new('RGB', (64, 64), color='white')
    dc = ImageDraw.Draw(image)
    dc.rectangle((8, 8, 56, 56), fill='black')
    return image

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

def update_window_list():
    listbox.delete(0, tk.END)
    for window in gw.getAllWindows():
        if window.title:
            listbox.insert(tk.END, window.title)

def on_select(evt):
    global selected_window
    w = evt.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    selected_window = value

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

# Setup the main window
root = tk.Tk()
root.title("Window Manager")
root.geometry("300x200")
root.protocol("WM_DELETE_WINDOW", root.withdraw)  # Minimize to tray instead of closing

listbox = tk.Listbox(root, width=60, height=10)
listbox.pack(pady=20)
listbox.bind('<<ListboxSelect>>', on_select)
refresh_button = ttk.Button(root, text="Refresh Window List", command=update_window_list)
refresh_button.pack(pady=5)
borderless_button = ttk.Button(root, text="Make Borderless Fullscreen", command=on_button_click)
borderless_button.pack(pady=5)
update_window_list()

# Run the tray icon in a separate thread
thread = threading.Thread(target=run_tray_icon)
thread.start()

root.mainloop()
