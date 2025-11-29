import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import json
import sys
import ctypes
import re


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def relaunch_as_admin():
    try:
        params = " ".join([f'"{arg}"' for arg in sys.argv])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        sys.exit(0)
    except Exception as e:
        messagebox.showerror("Error", f"Restart as administrator failed:\n{e}")


def read_json_tolerant(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = f.read()
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            fixed = re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', raw)
            return json.loads(fixed)
    except Exception:
        return {}


def choose_exe():
    file_path = filedialog.askopenfilename(
        title="Select start_protected_game.exe",
        filetypes=[("Executable", "*.exe"), ("All Files", "*.*")]
    )
    if not file_path:
        return

    file_path = os.path.normpath(file_path)

    if os.path.basename(file_path).lower() != "start_protected_game.exe":
        messagebox.showerror("Error", "Please select the file named 'start_protected_game.exe'.")
        return

    exe_var.set(file_path)

    path_escaped = file_path.replace('"', r'\"')
    cmd_preview = (
        f'netsh advfirewall firewall add rule name="EAC Disabled" '
        f'dir=out program="{path_escaped}" action=block'
    )
    cmd_var.set(cmd_preview)


def run_command():
    exe_path = exe_var.get().strip()
    if not exe_path:
        messagebox.showerror("Error", "No EXE selected. Please select start_protected_game.exe first.")
        return
    if not os.path.isfile(exe_path):
        messagebox.showerror("Error", "The selected EXE path is not valid.")
        return

    exe_path = os.path.normpath(exe_path)
    path_escaped = exe_path.replace('"', r'\"')

    cmd = (
        f'netsh advfirewall firewall add rule name="EAC Disabled" '
        f'dir=out program="{path_escaped}" action=block'
    )

    try:
        result = subprocess.run(cmd, shell=True)
        if result.returncode == 0:
            messagebox.showinfo("OK", "OK.")
        else:
            messagebox.showerror("Error", f"Command failed.\nReturn code: {result.returncode}")
    except Exception as e:
        messagebox.showerror("Error", f"Error while executing command:\n{e}")


def check_rule():
    cmd = 'netsh advfirewall firewall show rule name="EAC Disabled"'
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            encoding="utf-8",
            errors="ignore"
        )
        stdout = result.stdout or ""
        stderr = result.stderr or ""
        if result.returncode == 0 and "EAC Disabled" in stdout:
            messagebox.showinfo("Firewall rule", f"Rule 'EAC Disabled' exists.\n\n{stdout[:4000]}")
        else:
            messagebox.showwarning(
                "Firewall rule",
                "Rule 'EAC Disabled' does not seem to exist or could not be read.\n\n"
                f"{(stdout + stderr)[:4000]}"
            )
    except Exception as e:
        messagebox.showerror("Error", f"Error while checking rule:\n{e}")


def remove_rule():
    cmd = 'netsh advfirewall firewall delete rule name="EAC Disabled"'
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            encoding="utf-8",
            errors="ignore"
        )
        stdout = result.stdout or ""
        stderr = result.stderr or ""
        combined = (stdout + stderr)[:4000]
        if result.returncode == 0:
            if "No rules match the specified criteria" in combined:
                messagebox.showwarning("Firewall rule", "Rule 'EAC Disabled' was not found.")
            else:
                messagebox.showinfo("Firewall rule", f"Rule 'EAC Disabled' removed.\n\n{combined}")
        else:
            messagebox.showerror("Error", f"Failed to remove rule.\n\n{combined}")
    except Exception as e:
        messagebox.showerror("Error", f"Error while removing rule:\n{e}")


def choose_settings():
    file_path = filedialog.askopenfilename(
        title="Select Setting.json",
        filetypes=[("JSON file", "*.json"), ("All Files", "*.*")]
    )
    if not file_path:
        return

    file_path = os.path.normpath(file_path)
    settings_path_var.set(file_path)
    load_settings_into_entries()


def load_settings_into_entries():
    settings_path = settings_path_var.get().strip()
    if not settings_path or not os.path.exists(settings_path):
        return
    try:
        data = read_json_tolerant(settings_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read JSON file:\n{e}")
        return

    entry_productid.delete(0, tk.END)
    entry_productid.insert(0, str(data.get("productid", "")))

    entry_sandboxid.delete(0, tk.END)
    entry_sandboxid.insert(0, str(data.get("sandboxid", "")))

    entry_deploymentid.delete(0, tk.END)
    entry_deploymentid.insert(0, str(data.get("deploymentid", "")))


def save_json():
    settings_path = settings_path_var.get().strip()
    if not settings_path:
        messagebox.showerror("Error", "Please select your Setting.json first.")
        return

    data = read_json_tolerant(settings_path)

    productid = entry_productid.get().strip()
    sandboxid = entry_sandboxid.get().strip()
    deploymentid = entry_deploymentid.get().strip()

    if not productid or not sandboxid or not deploymentid:
        messagebox.showerror("Error", "Please fill in all three fields.")
        return

    data["productid"] = productid
    data["sandboxid"] = sandboxid
    data["deploymentid"] = deploymentid

    try:
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        messagebox.showinfo("OK", f"Setting.json saved:\n{settings_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error while writing file:\n{e}")


root = tk.Tk()
root.title("EAC Disable + Settings Tool")

if not is_admin():
    answer = messagebox.askyesno(
        "Administrator privileges required",
        "This program should be run as administrator.\n\nRestart as administrator now?"
    )
    if answer:
        relaunch_as_admin()

exe_var = tk.StringVar()
cmd_var = tk.StringVar()
settings_path_var = tk.StringVar()

frame_exe = tk.LabelFrame(root, text="Step 1: Disable EAC")
frame_exe.pack(fill="both", padx=10, pady=10)

button_choose_exe = tk.Button(frame_exe, text="Select start_protected_game.exe", command=choose_exe)
button_choose_exe.pack(padx=5, pady=5)

label_exe_path = tk.Label(frame_exe, textvariable=exe_var, wraplength=500, fg="gray")
label_exe_path.pack(anchor="w", padx=5, pady=5)

label_cmd_preview = tk.Label(frame_exe, textvariable=cmd_var, wraplength=500, fg="blue")
label_cmd_preview.pack(anchor="w", padx=5, pady=5)

button_run_cmd = tk.Button(frame_exe, text="Add firewall rule", command=run_command)
button_run_cmd.pack(padx=5, pady=5)

button_check_rule = tk.Button(frame_exe, text="Check firewall rule", command=check_rule)
button_check_rule.pack(padx=5, pady=5)

button_remove_rule = tk.Button(frame_exe, text="Remove firewall rule", command=remove_rule)
button_remove_rule.pack(padx=5, pady=5)

frame_json = tk.LabelFrame(root, text="Step 2: Edit Setting.json")
frame_json.pack(fill="both", padx=10, pady=10)

label_settings_info = tk.Label(frame_json, text="Select your Setting.json file manually:", fg="gray")
label_settings_info.grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=3)

label_settings_path = tk.Label(frame_json, textvariable=settings_path_var, wraplength=500, fg="gray")
label_settings_path.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=3)

button_browse_settings = tk.Button(frame_json, text="Browse Setting.json", command=choose_settings)
button_browse_settings.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

label_productid = tk.Label(frame_json, text="productid:")
label_productid.grid(row=3, column=0, sticky="w", padx=5, pady=5)
entry_productid = tk.Entry(frame_json)
entry_productid.grid(row=3, column=1, sticky="we", padx=5, pady=5)

label_sandboxid = tk.Label(frame_json, text="sandboxid:")
label_sandboxid.grid(row=4, column=0, sticky="w", padx=5, pady=5)
entry_sandboxid = tk.Entry(frame_json)
entry_sandboxid.grid(row=4, column=1, sticky="we", padx=5, pady=5)

label_deploymentid = tk.Label(frame_json, text="deploymentid:")
label_deploymentid.grid(row=5, column=0, sticky="w", padx=5, pady=5)
entry_deploymentid = tk.Entry(frame_json)
entry_deploymentid.grid(row=5, column=1, sticky="we", padx=5, pady=5)

frame_json.columnconfigure(1, weight=1)

button_save = tk.Button(frame_json, text="Save Setting.json", command=save_json)
button_save.grid(row=6, column=0, columnspan=2, padx=5, pady=10)

root.mainloop()
