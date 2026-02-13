import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import re
import json
import subprocess
import sys
import ctypes

CONFIG_FILE = "config.json"

# -----------------------------
# Windows App Identity (Taskbar Icon Fix)
# -----------------------------
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
        "zaparoo.custom.launcher.builder"
    )
except Exception:
    pass

# -----------------------------
# Resource Path (PyInstaller Safe)
# -----------------------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# -----------------------------
# Utility
# -----------------------------
def sanitize_id(name):
    return re.sub(r'[^A-Za-z0-9]', '', name).upper()

def escape_windows_path(path):
    path = os.path.normpath(path)
    return path.replace("\\", "\\\\")

def clean_emulator_name(name):
    name = name.lower().replace("-qt", "").replace("_qt", "")
    return sanitize_id(name)

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "custom_root" in data and os.path.exists(data["custom_root"]):
                custom_root_path.set(data["custom_root"])
                location_mode.set("Custom")
        except:
            pass

def save_config():
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"custom_root": custom_root_path.get()}, f, indent=4)

def get_launchers_folder():
    if location_mode.get() == "Default":
        return os.path.join(os.getenv("LOCALAPPDATA"), "zaparoo", "launchers")

    root = custom_root_path.get().strip()
    if not root:
        return ""

    root = os.path.normpath(root)

    if os.path.basename(root).lower() == "user":
        return os.path.join(root, "launchers")

    return os.path.join(root, "user", "launchers")

# -----------------------------
# Browse Functions
# -----------------------------
def browse_emulator():
    path = filedialog.askopenfilename(filetypes=[("Executable", "*.exe")])
    if path:
        emulator_path.set(path)

def browse_rom_dir():
    path = filedialog.askdirectory()
    if path:
        rom_path.set(path)

def browse_core():
    path = filedialog.askopenfilename(filetypes=[("DLL", "*.dll")])
    if path:
        core_path.set(path)

def browse_zaparoo_root():
    path = filedialog.askdirectory()
    if path:
        custom_root_path.set(path)
        location_mode.set("Custom")
        save_config()

# -----------------------------
# Logic
# -----------------------------
def is_retroarch():
    return "retroarch" in emulator_path.get().lower()

def update_ui(*args):
    if launcher_type.get() == "Direct":
        emulator_widgets(False)
        core_widgets(False)
    else:
        emulator_widgets(True)
        if is_retroarch():
            core_widgets(True)
        else:
            core_widgets(False)

def emulator_widgets(show=False):
    for w in emulator_row_widgets:
        if show:
            w.grid()
        else:
            w.grid_remove()

def core_widgets(show=False):
    for w in core_row_widgets:
        if show:
            w.grid()
        else:
            w.grid_remove()

# -----------------------------
# Launcher Generation
# -----------------------------
def generate_launcher():
    system = system_var.get().strip()
    rom = rom_path.get().strip()
    exts = extensions_var.get().strip()

    if not system or not rom or not exts:
        messagebox.showerror("Missing Fields", "Please fill all required fields.")
        return

    escaped_rom = escape_windows_path(rom)

    # ID
    if launcher_type.get() == "Emulator":
        emu = emulator_path.get().strip()
        if not emu:
            messagebox.showerror("Missing Emulator", "Select an emulator.")
            return

        emu_name = os.path.splitext(os.path.basename(emu))[0]
        emu_name = clean_emulator_name(emu_name)
        launcher_id = emu_name + sanitize_id(system)
    else:
        launcher_id = sanitize_id(system) + "DIRECT"

    # Execute
    if launcher_type.get() == "Direct":
        execute = (
            "powershell -WindowStyle Hidden -NoProfile "
            "-ExecutionPolicy Bypass -Command "
            "Start-Process -FilePath '[[media_path]]'"
        )
    else:
        emu = emulator_path.get().strip()
        escaped_emu = escape_windows_path(emu)

        if is_retroarch():
            if not core_path.get():
                messagebox.showerror("Missing Core", "Select a RetroArch core.")
                return

            escaped_core = escape_windows_path(core_path.get())

            execute = (
                "powershell -WindowStyle Hidden -NoProfile "
                "-ExecutionPolicy Bypass -Command "
                f"Start-Process -FilePath '{escaped_emu}' "
                f"-ArgumentList '-L', '{escaped_core}', '\\\"[[media_path]]\\\"'"
            )
        else:
            execute = (
                "powershell -WindowStyle Hidden -NoProfile "
                "-ExecutionPolicy Bypass -Command "
                f"Start-Process -FilePath '{escaped_emu}' "
                f"-ArgumentList '\\\"[[media_path]]\\\"'"
            )

    # Extensions
    clean_exts = []
    for e in exts.split(","):
        e = e.strip()
        if not e.startswith("."):
            e = "." + e
        clean_exts.append(f'"{e}"')
    file_exts = ",".join(clean_exts)

    toml = f"""[[launchers.custom]]
id = "{launcher_id}"
system = "{system}"
media_dirs = ["{escaped_rom}"]
file_exts = [{file_exts}]
execute = "{execute}"
"""

    save_launcher(toml, launcher_id)

# -----------------------------
# Save / Open
# -----------------------------
def save_launcher(toml_content, launcher_id):
    folder = get_launchers_folder()

    if not folder:
        messagebox.showerror("Error", "Invalid Zaparoo folder.")
        return

    os.makedirs(folder, exist_ok=True)

    file_path = os.path.join(folder, f"{launcher_id}.toml")

    if os.path.exists(file_path):
        if not messagebox.askyesno("Launcher Exists", f"{launcher_id}.toml exists.\nOverwrite?"):
            return

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(toml_content)

    messagebox.showinfo("Success", f"Launcher saved to:\n{file_path}")

def open_launchers_folder():
    folder = get_launchers_folder()

    if not folder or not os.path.exists(folder):
        messagebox.showerror("Error", "Launchers folder does not exist.")
        return

    subprocess.Popen(["explorer", folder])

# -----------------------------
# UI Setup
# -----------------------------
root = tk.Tk()
root.title("Zaparoo Custom Launcher Builder v1.0")
root.geometry("720x520")
root.resizable(False, False)

# Set PNG window icon
try:
    icon_path = resource_path("icon.png")
    icon_image = tk.PhotoImage(file=icon_path)
    root.iconphoto(True, icon_image)
except Exception:
    pass

system_var = tk.StringVar()
emulator_path = tk.StringVar()
rom_path = tk.StringVar()
extensions_var = tk.StringVar()
core_path = tk.StringVar()

launcher_type = tk.StringVar(value="Emulator")
location_mode = tk.StringVar(value="Default")
custom_root_path = tk.StringVar()

main = ttk.Frame(root, padding=25)
main.pack(fill="both", expand=True)

main.columnconfigure(0, minsize=170)
main.columnconfigure(1, minsize=380)
main.columnconfigure(2, minsize=90)

row = 0

ttk.Label(main, text="Launcher Type").grid(row=row, column=0, sticky="w", pady=8)
ttk.Radiobutton(main, text="Emulator", variable=launcher_type, value="Emulator").grid(row=row, column=1, sticky="w")
ttk.Radiobutton(main, text="Direct / Shortcut", variable=launcher_type, value="Direct").grid(row=row, column=1, padx=150, sticky="w")
row += 1

ttk.Label(main, text="System").grid(row=row, column=0, sticky="w", pady=8)
ttk.Entry(main, textvariable=system_var).grid(row=row, column=1, sticky="ew")
row += 1

emulator_label = ttk.Label(main, text="Emulator Path")
emulator_entry = ttk.Entry(main, textvariable=emulator_path)
emulator_button = ttk.Button(main, text="Browse", command=browse_emulator)

emulator_label.grid(row=row, column=0, sticky="w", pady=8)
emulator_entry.grid(row=row, column=1, sticky="ew")
emulator_button.grid(row=row, column=2)

emulator_row_widgets = [emulator_label, emulator_entry, emulator_button]
row += 1

ttk.Label(main, text="ROM Directory").grid(row=row, column=0, sticky="w", pady=8)
ttk.Entry(main, textvariable=rom_path).grid(row=row, column=1, sticky="ew")
ttk.Button(main, text="Browse", command=browse_rom_dir).grid(row=row, column=2)
row += 1

core_label = ttk.Label(main, text="RetroArch Core Path")
core_entry = ttk.Entry(main, textvariable=core_path)
core_button = ttk.Button(main, text="Browse", command=browse_core)

core_label.grid(row=row, column=0, sticky="w", pady=8)
core_entry.grid(row=row, column=1, sticky="ew")
core_button.grid(row=row, column=2)

core_row_widgets = [core_label, core_entry, core_button]
row += 1

ttk.Label(main, text="File Extensions (iso,zip,lnk)").grid(row=row, column=0, sticky="w", pady=8)
ttk.Entry(main, textvariable=extensions_var).grid(row=row, column=1, sticky="ew")
row += 1

ttk.Label(main, text="Zaparoo Location").grid(row=row, column=0, sticky="w", pady=12)
location_frame = ttk.Frame(main)
location_frame.grid(row=row, column=1, sticky="w")
ttk.Radiobutton(location_frame, text="Default (AppData Local)", variable=location_mode, value="Default").pack(anchor="w")
ttk.Radiobutton(location_frame, text="Custom Folder", variable=location_mode, value="Custom").pack(anchor="w")
row += 1

ttk.Entry(main, textvariable=custom_root_path).grid(row=row, column=1, sticky="ew")
ttk.Button(main, text="Browse", command=browse_zaparoo_root).grid(row=row, column=2)
row += 1

button_frame = ttk.Frame(main)
button_frame.grid(row=row, column=0, columnspan=3, pady=30)
ttk.Button(button_frame, text="Generate & Save Launcher", width=28, command=generate_launcher).pack(pady=6)
ttk.Button(button_frame, text="Open Launchers Folder", width=28, command=open_launchers_folder).pack()

launcher_type.trace_add("write", update_ui)
emulator_path.trace_add("write", update_ui)

load_config()
update_ui()

root.mainloop()
