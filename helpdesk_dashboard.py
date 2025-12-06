import os
import socket
import platform
import subprocess
import shutil
import tkinter as tk
from tkinter import ttk, messagebox

# external lib for live stats
try:
    import psutil
except ImportError:
    psutil = None


# =============================
#   HELPERS: POPUP MESSAGE
# =============================

def show_info(title: str, text: str):
    messagebox.showinfo(title, text)


def show_error(title: str, text: str):
    messagebox.showerror(title, text)


# =============================
#   DIAGNOSTICS FUNCTIONS
# =============================

def get_system_info():
    try:
        ram_gb = None
        if psutil:
            ram_gb = round(psutil.virtual_memory().total / (1024 ** 3), 2)

        info_lines = [
            f"OS: {platform.system()} {platform.release()}",
            f"Version: {platform.version()}",
            f"Computer Name: {platform.node()}",
        ]

        try:
            info_lines.append(f"User: {os.getlogin()}")
        except Exception:
            pass

        info_lines.append(f"CPU: {platform.processor() or 'N/A'}")

        if ram_gb is not None:
            info_lines.append(f"RAM: {ram_gb} GB (total)")

        show_info("System Info", "\n".join(info_lines))
    except Exception as e:
        show_error("System Info Error", str(e))


def get_network_info():
    try:
        hostname = socket.gethostname()
        try:
            ip = socket.gethostbyname(hostname)
        except Exception:
            ip = "N/A"

        lines = [
            f"Hostname: {hostname}",
            f"IP Address: {ip}",
        ]

        # Try to get default gateway (Windows)
        gateway = "N/A"
        try:
            proc = subprocess.run(
                ["ipconfig"],
                capture_output=True,
                text=True,
                shell=True
            )
            out = proc.stdout.lower()
            if "default gateway" in out:
                for line in proc.stdout.splitlines():
                    if "default gateway" in line.lower():
                        parts = line.split(":")
                        if len(parts) > 1:
                            gw_candidate = parts[1].strip()
                            if gw_candidate:
                                gateway = gw_candidate
                                break
        except Exception:
            pass

        lines.append(f"Default Gateway: {gateway}")
        show_info("Network Info", "\n".join(lines))
    except Exception as e:
        show_error("Network Info Error", str(e))


def ping_google():
    try:
        proc = subprocess.run(
            ["ping", "8.8.8.8", "-n", "4"],
            capture_output=True,
            text=True,
            shell=True
        )
        if proc.returncode == 0:
            show_info("Ping Test", proc.stdout)
        else:
            show_error("Ping Failed", proc.stdout or proc.stderr)
    except Exception as e:
        show_error("Ping Error", str(e))


# =============================
#   FIX / REPAIR FUNCTIONS
# =============================

def flush_dns():
    try:
        proc = subprocess.run(
            ["ipconfig", "/flushdns"],
            capture_output=True,
            text=True,
            shell=True
        )
        if proc.returncode == 0:
            show_info("Flush DNS", "DNS cache flushed successfully.")
        else:
            show_error("Flush DNS Error", proc.stdout or proc.stderr)
    except Exception as e:
        show_error("Flush DNS Error", str(e))


def clear_temp():
    try:
        temp_dir = os.environ.get("TEMP") or os.environ.get("TMP")
        if not temp_dir:
            show_error("Temp Cleanup", "Could not locate TEMP directory.")
            return

        deleted = 0
        failed = 0

        for name in os.listdir(temp_dir):
            path = os.path.join(temp_dir, name)
            try:
                if os.path.isfile(path) or os.path.islink(path):
                    os.remove(path)
                    deleted += 1
                elif os.path.isdir(path):
                    shutil.rmtree(path, ignore_errors=True)
                    deleted += 1
            except Exception:
                failed += 1

        msg = f"Temp cleanup finished.\nDeleted items: {deleted}\nFailed: {failed}"
        show_info("Temp Cleanup", msg)
    except Exception as e:
        show_error("Temp Cleanup Error", str(e))


def reset_network():
    try:
        subprocess.run("netsh winsock reset", shell=True)
        subprocess.run("netsh int ip reset", shell=True)
        show_info("Network Reset", "Network stack reset.\nPlease restart your computer.")
    except Exception as e:
        show_error("Network Reset Error", str(e))


def restart_explorer():
    try:
        subprocess.run("taskkill /F /IM explorer.exe", shell=True)
        subprocess.Popen("explorer.exe", shell=True)
        show_info("Explorer Restart", "Explorer restarted.")
    except Exception as e:
        show_error("Explorer Restart Error", str(e))


def restart_spooler():
    try:
        subprocess.run("net stop spooler", shell=True)
        subprocess.run("net start spooler", shell=True)
        show_info("Print Spooler", "Print Spooler restarted.")
    except Exception as e:
        show_error("Print Spooler Error", str(e))


def run_sfc():
        try:
            show_info("SFC", "Running SFC scan... This may take 10-20 minutes.")
            subprocess.run("sfc /scannow", shell=True)
            show_info("SFC", "SFC scan completed.")
        except Exception as e:
            show_error("SFC Error", str(e))


def run_dism():
    try:
        show_info("DISM", "Running DISM... This may take 10-20 minutes.")
        subprocess.run("DISM /Online /Cleanup-Image /RestoreHealth", shell=True)
        show_info("DISM", "DISM completed.")
    except Exception as e:
        show_error("DISM Error", str(e))


# =============================
#   TOOL LAUNCHERS (WINDOWS)
# =============================

def open_task_manager():
    try:
        subprocess.Popen("taskmgr", shell=True)
    except Exception as e:
        show_error("Task Manager Error", str(e))


def open_services():
    try:
        subprocess.Popen("services.msc", shell=True)
    except Exception as e:
        show_error("Services Error", str(e))


def open_event_viewer():
    try:
        subprocess.Popen("eventvwr.msc", shell=True)
    except Exception as e:
        show_error("Event Viewer Error", str(e))


def open_network_connections():
    try:
        subprocess.Popen("ncpa.cpl", shell=True)
    except Exception as e:
        show_error("Network Connections Error", str(e))


def open_control_panel():
    try:
        subprocess.Popen("control", shell=True)
    except Exception as e:
        show_error("Control Panel Error", str(e))


# =============================
#   LIVE DASHBOARD UPDATE
# =============================

def update_stats(cpu_label, ram_label, disk_label, host_label, ip_label, root):
    try:
        if psutil:
            cpu = psutil.cpu_percent(interval=None)
            ram = psutil.virtual_memory().percent
            try:
                disk = psutil.disk_usage("C:\\").percent
            except Exception:
                disk = 0.0
        else:
            cpu = 0.0
            ram = 0.0
            disk = 0.0

        cpu_label.config(text=f"CPU: {cpu:.1f}%")
        ram_label.config(text=f"RAM: {ram:.1f}%")
        disk_label.config(text=f"Disk C:: {disk:.1f}%")

        hostname = socket.gethostname()
        host_label.config(text=f"Host: {hostname}")

        try:
            ip = socket.gethostbyname(hostname)
        except Exception:
            ip = "N/A"
        ip_label.config(text=f"IP: {ip}")

    except Exception:
        pass

    root.after(1000, update_stats, cpu_label, ram_label, disk_label, host_label, ip_label, root)


# =============================
#   GUI
# =============================

def create_gui():
    root = tk.Tk()
    root.title("Helpdesk Technician Dashboard")
    root.geometry("700x500")

    main_frame = ttk.Frame(root, padding=10)
    main_frame.pack(fill=tk.BOTH, expand=True)

    status_frame = ttk.LabelFrame(main_frame, text="Live Status", padding=10)
    status_frame.pack(fill=tk.X, expand=False, pady=(0, 10))

    cpu_label = ttk.Label(status_frame, text="CPU: N/A", width=18)
    ram_label = ttk.Label(status_frame, text="RAM: N/A", width=18)
    disk_label = ttk.Label(status_frame, text="Disk C:: N/A", width=18)
    host_label = ttk.Label(status_frame, text="Host: N/A", width=26)
    ip_label = ttk.Label(status_frame, text="IP: N/A", width=22)

    cpu_label.grid(row=0, column=0, padx=5, pady=2, sticky="w")
    ram_label.grid(row=0, column=1, padx=5, pady=2, sticky="w")
    disk_label.grid(row=0, column=2, padx=5, pady=2, sticky="w")
    host_label.grid(row=1, column=0, padx=5, pady=2, sticky="w")
    ip_label.grid(row=1, column=1, padx=5, pady=2, sticky="w")

    buttons_frame = ttk.Frame(main_frame)
    buttons_frame.pack(fill=tk.BOTH, expand=True)

    diag_frame = ttk.LabelFrame(buttons_frame, text="Diagnostics", padding=10)
    diag_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

    ttk.Button(diag_frame, text="System Info", width=20, command=get_system_info).pack(pady=3)
    ttk.Button(diag_frame, text="Network Info", width=20, command=get_network_info).pack(pady=3)
    ttk.Button(diag_frame, text="Ping Google", width=20, command=ping_google).pack(pady=3)

    fix_frame = ttk.LabelFrame(buttons_frame, text="Fix / Repair", padding=10)
    fix_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

    ttk.Button(fix_frame, text="Flush DNS", width=20, command=flush_dns).pack(pady=3)
    ttk.Button(fix_frame, text="Clear Temp Files", width=20, command=clear_temp).pack(pady=3)
    ttk.Button(fix_frame, text="Reset Network", width=20, command=reset_network).pack(pady=3)
    ttk.Button(fix_frame, text="Restart Explorer", width=20, command=restart_explorer).pack(pady=3)
    ttk.Button(fix_frame, text="Restart Spooler", width=20, command=restart_spooler).pack(pady=3)
    ttk.Button(fix_frame, text="Run SFC", width=20, command=run_sfc).pack(pady=3)
    ttk.Button(fix_frame, text="Run DISM", width=20, command=run_dism).pack(pady=3)

    tools_frame = ttk.LabelFrame(buttons_frame, text="Tools", padding=10)
    tools_frame.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")

    ttk.Button(tools_frame, text="Task Manager", width=20, command=open_task_manager).pack(pady=3)
    ttk.Button(tools_frame, text="Services", width=20, command=open_services).pack(pady=3)
    ttk.Button(tools_frame, text="Event Viewer", width=20, command=open_event_viewer).pack(pady=3)
    ttk.Button(tools_frame, text="Network Connections", width=20, command=open_network_connections).pack(pady=3)
    ttk.Button(tools_frame, text="Control Panel", width=20, command=open_control_panel).pack(pady=3)

    buttons_frame.columnconfigure(0, weight=1)
    buttons_frame.columnconfigure(1, weight=1)
    buttons_frame.columnconfigure(2, weight=1)

    update_stats(cpu_label, ram_label, disk_label, host_label, ip_label, root)

    root.mainloop()


if __name__ == "__main__":
    if psutil is None:
        msg = (
            "psutil is not installed.\n\n"
            "Run:\n\n"
            "    pip install psutil\n\n"
            "and then run this tool again."
        )
        messagebox.showerror("Missing Dependency", msg)
    else:
        create_gui()
