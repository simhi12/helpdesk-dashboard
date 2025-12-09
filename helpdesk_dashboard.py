import os
import socket
import platform
import subprocess
import shutil
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import datetime

# =============================
#   DEPENDENCY
# =============================

try:
    import psutil
except ImportError:
    psutil = None


# =============================
#   POPUP HELPERS
# =============================

def show_info(title: str, text: str):
    messagebox.showinfo(title, text)


def show_error(title: str, text: str):
    messagebox.showerror(title, text)


# =============================
#   TEMPERATURE
# =============================

def get_temperature_str() -> str:
    """Try to read CPU temp via psutil. If not available -> N/A."""
    if not psutil:
        return "N/A"
    try:
        temps = psutil.sensors_temperatures()
        if not temps:
            return "N/A"

        for key in ("cpu-thermal", "cpu_thermal", "coretemp", "acpitz"):
            if key in temps and temps[key]:
                return f"{temps[key][0].current:.1f}°C"

        for entries in temps.values():
            if entries:
                return f"{entries[0].current:.1f}°C"
    except Exception:
        pass
    return "N/A"


# =============================
#   SYSTEM / NETWORK INFO
# =============================

def get_system_info():
    try:
        if psutil:
            ram_gb = round(psutil.virtual_memory().total / (1024 ** 3), 2)
        else:
            ram_gb = "N/A"

        lines = [
            f"OS: {platform.system()} {platform.release()}",
            f"Version: {platform.version()}",
            f"Computer Name: {platform.node()}",
        ]
        try:
            lines.append(f"User: {os.getlogin()}")
        except Exception:
            pass
        lines.append(f"CPU: {platform.processor() or 'N/A'}")
        if isinstance(ram_gb, (int, float)):
            lines.append(f"RAM: {ram_gb} GB (total)")
        else:
            lines.append(f"RAM: {ram_gb}")

        show_info("System Info", "\n".join(lines))
    except Exception as e:
        show_error("System Info Error", str(e))


def get_network_info():
    """Parse ipconfig /all to show IP, mask, gateway, DNS, MAC."""
    try:
        hostname = socket.gethostname()
        try:
            ip = socket.gethostbyname(hostname)
        except Exception:
            ip = "N/A"

        proc = subprocess.run(
            ["ipconfig", "/all"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore"
        )
        text = proc.stdout or ""
        lines = text.splitlines()

        subnet_mask = "N/A"
        gateway = "N/A"
        dns_server = "N/A"
        mac = "N/A"

        for line in lines:
            l = line.strip()
            if "Subnet Mask" in l and ":" in l:
                subnet_mask = l.split(":", 1)[1].strip()
            if "Default Gateway" in l and ":" in l:
                gw = l.split(":", 1)[1].strip()
                if gw:
                    gateway = gw
            if "DNS Servers" in l and ":" in l:
                dns_server = l.split(":", 1)[1].strip()
            if "Physical Address" in l and ":" in l:
                mac = l.split(":", 1)[1].strip()

        out = [
            f"Hostname: {hostname}",
            f"IP Address: {ip}",
            f"Subnet Mask: {subnet_mask}",
            f"Default Gateway: {gateway}",
            f"DNS Server: {dns_server}",
            f"MAC Address: {mac}",
        ]
        show_info("Network Info", "\n".join(out))
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
#   RAM HEALTH (POPUP)
# =============================

def check_ram_health():
    try:
        if not psutil:
            show_error("RAM Health", "psutil is required.")
            return

        vm = psutil.virtual_memory()
        total = round(vm.total / (1024**3), 2)
        used = round(vm.used / (1024**3), 2)
        percent = vm.percent

        if percent < 85:
            status = "OK"
        elif percent < 95:
            status = "WARNING – High Memory Usage"
        else:
            status = "CRITICAL – Memory Pressure"

        out = [
            f"Total RAM: {total} GB",
            f"Used RAM: {used} GB",
            f"Usage: {percent}%",
            f"Status: {status}",
        ]
        show_info("RAM Health", "\n".join(out))
    except Exception as e:
        show_error("RAM Health Error", str(e))


# =============================
#   SMART STATUS (POPUP)
# =============================

def check_smart_status():
    try:
        proc = subprocess.run(
            'wmic diskdrive get model,serialnumber,status',
            capture_output=True,
            text=True,
            shell=True
        )

        output = proc.stdout or ""
        if not output.strip():
            show_error("SMART Status", "No SMART data found.")
            return

        lines = output.strip().splitlines()
        parsed = []
        for line in lines[1:]:
            parts = line.split()
            if len(parts) >= 2:
                model = " ".join(parts[:-2])
                serial = parts[-2]
                status = parts[-1]
                parsed.append(f"Model: {model}\nSerial: {serial}\nStatus: {status}\n")

        if parsed:
            show_info("SMART Status", "\n".join(parsed))
        else:
            show_error("SMART Status", "Unable to parse SMART output.")
    except Exception as e:
        show_error("SMART Status Error", str(e))


# =============================
#   FIX / REPAIR
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


# =============================
#   TOOL LAUNCHERS
# =============================

def _open_with_start(target: str, title: str):
    try:
        subprocess.Popen(f'start "" "{target}"', shell=True)
    except Exception as e:
        show_error(f"{title} Error", str(e))


def open_task_manager():
    _open_with_start("taskmgr", "Task Manager")


def open_services():
    _open_with_start("services.msc", "Services")


def open_event_viewer():
    _open_with_start("eventvwr.msc", "Event Viewer")


def open_network_connections():
    _open_with_start("ncpa.cpl", "Network Connections")


def open_control_panel():
    _open_with_start("control", "Control Panel")


def open_rdp():
    _open_with_start("mstsc", "RDP")


def open_windows_defender():
    _open_with_start("windowsdefender:", "Windows Defender")


def open_windows_update():
    _open_with_start("ms-settings:windowsupdate", "Windows Update")


def open_registry_editor():
    _open_with_start("regedit", "Registry Editor")


def open_troubleshooter():
    _open_with_start("ms-settings:troubleshoot", "Troubleshooter")


# =============================
#   DIAGNOSTIC REPORT
# =============================

def log_line(widget: tk.Text, text: str):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    widget.insert(tk.END, f"[{ts}] {text}\n")
    widget.see(tk.END)


def run_command_to_file(cmd: str, outfile: Path, log_widget: tk.Text):
    """
    Run a shell command and dump output into a text file.
    If return code != 0, we log a clear message (e.g. for Security logs when not Admin).
    """
    try:
        with outfile.open("w", encoding="utf-8", errors="ignore") as f:
            proc = subprocess.run(
                cmd,
                stdout=f,
                stderr=subprocess.STDOUT,
                text=True,
                shell=True
            )
        if proc.returncode == 0:
            log_line(log_widget, f"OK: {outfile.name}")
        else:
            log_line(
                log_widget,
                f"ERROR (code {proc.returncode}) while running command for {outfile.name}. "
                f"If this is a Security log, run as Administrator."
            )
    except Exception as e:
        log_line(log_widget, f"ERROR writing {outfile.name}: {e}")


def collect_full_diagnostic(root: tk.Tk, log_widget: tk.Text, progress_var: tk.DoubleVar):
    """
    Collect extended diagnostic data into a folder on Desktop and zip it.
    Includes:
    - Security log events (4624, 4625, 1102, 4672)
    - System WHEA (18), BugCheck (1001)
    - Windows Defender status+threats
    - Systeminfo, ipconfig /all, tasklist /v
    - DiskDrive info
    - RAM_Health.txt
    - SMART_Status.txt
    """
    try:
        log_widget.delete("1.0", tk.END)
        progress_var.set(0)
        root.update_idletasks()

        desktop = Path.home() / "Desktop"
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"HelpdeskReport_{timestamp}"
        report_dir = desktop / folder_name
        report_dir.mkdir(parents=True, exist_ok=True)

        log_line(log_widget, f"Report folder: {report_dir}")

        steps = [
            (5, "Collecting Security log: Successful logons (4624)...",
             'wevtutil qe Security /q:"*[System[(EventID=4624)]]" /f:text',
             "Security_Logons_4624.txt"),

            (10, "Collecting Security log: Failed logons (4625)...",
             'wevtutil qe Security /q:"*[System[(EventID=4625)]]" /f:text',
             "Security_FailedLogons_4625.txt"),

            (15, "Collecting Security log: Log cleared (1102)...",
             'wevtutil qe Security /q:"*[System[(EventID=1102)]]" /f:text',
             "Security_LogCleared_1102.txt"),

            (20, "Collecting Security log: Admin privilege granted (4672)...",
             'wevtutil qe Security /q:"*[System[(EventID=4672)]]" /f:text',
             "Security_AdminPrivilege_4672.txt"),

            (30, "Collecting System log: WHEA hardware errors (18)...",
             'wevtutil qe System /q:"*[System[(EventID=18)]]" /f:text',
             "System_WHEA_18.txt"),

            (40, "Collecting System log: BugCheck (1001)...",
             'wevtutil qe System /q:"*[System[(EventID=1001)]]" /f:text',
             "System_BugCheck_1001.txt"),

            (50, "Collecting Windows Defender computer status...",
             'powershell -Command "Get-MpComputerStatus | Format-List *"',
             "Defender_ComputerStatus.txt"),

            (60, "Collecting Windows Defender threat detections...",
             'powershell -Command "Get-MpThreatDetection | Format-List *"',
             "Defender_ThreatDetections.txt"),

            (70, "Collecting system info (systeminfo)...",
             "systeminfo",
             "SystemInfo.txt",

             ),

            (80, "Collecting network info (ipconfig /all)...",
             "ipconfig /all",
             "Network_IpconfigAll.txt"),

            (90, "Collecting running processes (tasklist /v)...",
             "tasklist /v",
             "Processes_Tasklist_v.txt"),

            (95, "Collecting disk/drive info (wmic diskdrive)...",
             "wmic diskdrive get model,serialnumber,size,firmwarerevision",
             "DiskDrive_Info.txt"),
        ]

        for pct, desc, cmd, filename in steps:
            progress_var.set(pct)
            root.update_idletasks()
            log_line(log_widget, desc)
            run_command_to_file(cmd, report_dir / filename, log_widget)

        # RAM_Health.txt
        try:
            if psutil:
                vm = psutil.virtual_memory()
                total = round(vm.total / (1024 ** 3), 2)
                used = round(vm.used / (1024 ** 3), 2)
                percent = vm.percent

                if percent < 85:
                    status = "OK"
                elif percent < 95:
                    status = "WARNING"
                else:
                    status = "CRITICAL"

                with (report_dir / "RAM_Health.txt").open("w", encoding="utf-8") as f:
                    f.write(f"Total RAM: {total} GB\n")
                    f.write(f"Used RAM: {used} GB\n")
                    f.write(f"Usage: {percent}%\n")
                    f.write(f"Status: {status}\n")
                log_line(log_widget, "OK: RAM_Health.txt")
            else:
                with (report_dir / "RAM_Health.txt").open("w", encoding="utf-8") as f:
                    f.write("psutil is not installed, RAM metrics unavailable.\n")
                log_line(log_widget, "OK: RAM_Health.txt (no psutil)")
        except Exception as e:
            log_line(log_widget, f"ERROR RAM_Health.txt: {e}")

        # SMART_Status.txt
        try:
            proc = subprocess.run(
                'wmic diskdrive get model,serialnumber,status',
                capture_output=True,
                text=True,
                shell=True
            )
            out = proc.stdout or ""
            with (report_dir / "SMART_Status.txt").open("w", encoding="utf-8") as f:
                if out.strip():
                    f.write(out)
                else:
                    f.write("No SMART data returned from WMIC.\n")
            log_line(log_widget, "OK: SMART_Status.txt")
        except Exception as e:
            log_line(log_widget, f"ERROR SMART_Status.txt: {e}")

        log_line(log_widget, "Creating ZIP archive...")
        shutil.make_archive(str(report_dir), "zip", root_dir=report_dir)
        zip_path = report_dir.with_suffix(".zip")

        progress_var.set(100)
        root.update_idletasks()
        show_info("Full Diagnostic", f"Diagnostic package created:\n{zip_path}")
    except Exception as e:
        show_error("Diagnostic Error", str(e))
        log_line(log_widget, f"ERROR: {e}")


# =============================
#   LIVE STATS
# =============================

def update_stats(cpu_label, ram_label, disk_label, temp_label, host_label, ip_label, root):
    try:
        if psutil:
            cpu = psutil.cpu_percent(interval=None)
            ram = psutil.virtual_memory().percent
            try:
                disk = psutil.disk_usage("C:\\").percent
            except Exception:
                disk = 0.0
        else:
            cpu = ram = disk = 0.0

        cpu_label.config(text=f"CPU: {cpu:.1f}%")
        ram_label.config(text=f"RAM: {ram:.1f}%")
        disk_label.config(text=f"Disk C: {disk:.1f}%")
        temp_label.config(text=f"Temp: {get_temperature_str()}")

        hostname = socket.gethostname()
        host_label.config(text=f"Host: {hostname}")
        try:
            ip = socket.gethostbyname(hostname)
        except Exception:
            ip = "N/A"
        ip_label.config(text=f"IP: {ip}")
    except Exception:
        pass

    root.after(
        1000,
        update_stats,
        cpu_label,
        ram_label,
        disk_label,
        temp_label,
        host_label,
        ip_label,
        root
    )


# =============================
#   GUI
# =============================

def create_gui():
    root = tk.Tk()
    root.title("Helpdesk Technician Dashboard v2.0.0")
    root.geometry("1000x620")
    root.configure(bg="#1e1e1e")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", background="#1e1e1e", foreground="white")
    style.configure("TFrame", background="#1e1e1e")
    style.configure("TLabelframe", background="#1e1e1e", foreground="cyan", borderwidth=2)
    style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
    style.map("TButton", background=[("active", "#333333")])

    main_frame = ttk.Frame(root, padding=10)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Live
    status_frame = ttk.LabelFrame(main_frame, text="Live Status", padding=10)
    status_frame.pack(fill=tk.X, expand=False, pady=(0, 10))

    cpu_label = ttk.Label(status_frame, text="CPU: N/A", width=18)
    ram_label = ttk.Label(status_frame, text="RAM: N/A", width=18)
    disk_label = ttk.Label(status_frame, text="Disk C: N/A", width=18)
    temp_label = ttk.Label(status_frame, text="Temp: N/A", width=18)
    host_label = ttk.Label(status_frame, text="Host: N/A", width=26)
    ip_label = ttk.Label(status_frame, text="IP: N/A", width=26)

    cpu_label.pack(side="left", padx=15)
    ram_label.pack(side="left", padx=15)
    disk_label.pack(side="left", padx=15)
    temp_label.pack(side="left", padx=15)
    host_label.pack(side="left", padx=15)
    ip_label.pack(side="left", padx=15)

    # Panels
    middle_frame = ttk.Frame(main_frame)
    middle_frame.pack(fill=tk.BOTH, expand=True)

    # Diagnostics
    diag_frame = ttk.LabelFrame(middle_frame, text="Diagnostics", padding=10)
    diag_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

    ttk.Button(diag_frame, text="System Info", width=24, command=get_system_info).pack(pady=6)
    ttk.Button(diag_frame, text="Network Info", width=24, command=get_network_info).pack(pady=6)
    ttk.Button(diag_frame, text="Ping Google (8.8.8.8)", width=24, command=ping_google).pack(pady=6)

    ttk.Button(diag_frame, text="Check RAM Health", width=24, command=check_ram_health).pack(pady=6)
    ttk.Button(diag_frame, text="Check SMART Status", width=24, command=check_smart_status).pack(pady=6)

    progress_var = tk.DoubleVar(value=0)

    ttk.Button(
        diag_frame,
        text="Collect Full Diagnostic",
        width=24,
        command=lambda: collect_full_diagnostic(root, log_text, progress_var)
    ).pack(pady=(12, 8))

    ttk.Label(diag_frame, text="Diagnostic Log:").pack(anchor="center")

    progress_bar = ttk.Progressbar(
        diag_frame,
        orient="horizontal",
        mode="determinate",
        variable=progress_var,
        maximum=100
    )
    progress_bar.pack(fill=tk.X, padx=5, pady=(0, 5))

    log_text = tk.Text(diag_frame, height=14, wrap="word", bg="#000000", fg="#00ff00")
    log_scroll = ttk.Scrollbar(diag_frame, orient="vertical", command=log_text.yview)
    log_text.configure(yscrollcommand=log_scroll.set)
    log_text.pack(side="left", fill=tk.BOTH, expand=True)
    log_scroll.pack(side="right", fill="y")

    # Fix / Repair
    fix_frame = ttk.LabelFrame(middle_frame, text="Fix / Repair", padding=10)
    fix_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

    ttk.Button(fix_frame, text="Flush DNS", width=24, command=flush_dns).pack(pady=10)
    ttk.Button(fix_frame, text="Clear Temp Files", width=24, command=clear_temp).pack(pady=10)

    # Tools
    tools_frame = ttk.LabelFrame(middle_frame, text="Tools", padding=10)
    tools_frame.grid(row=0, column=2, padx=5, pady=5, sticky="nsew")

    ttk.Button(tools_frame, text="Task Manager", width=24, command=open_task_manager).pack(pady=4)
    ttk.Button(tools_frame, text="Services", width=24, command=open_services).pack(pady=4)
    ttk.Button(tools_frame, text="Event Viewer", width=24, command=open_event_viewer).pack(pady=4)
    ttk.Button(tools_frame, text="Network Connections", width=24, command=open_network_connections).pack(pady=4)
    ttk.Button(tools_frame, text="Control Panel", width=24, command=open_control_panel).pack(pady=4)
    ttk.Button(tools_frame, text="Open RDP", width=24, command=open_rdp).pack(pady=4)

    ttk.Separator(tools_frame, orient="horizontal").pack(fill=tk.X, pady=6)

    ttk.Button(tools_frame, text="Windows Defender", width=24, command=open_windows_defender).pack(pady=4)
    ttk.Button(tools_frame, text="Windows Update", width=24, command=open_windows_update).pack(pady=4)
    ttk.Button(tools_frame, text="Registry Editor", width=24, command=open_registry_editor).pack(pady=4)
    ttk.Button(tools_frame, text="Troubleshooter", width=24, command=open_troubleshooter).pack(pady=4)

    middle_frame.columnconfigure(0, weight=2)
    middle_frame.columnconfigure(1, weight=1)
    middle_frame.columnconfigure(2, weight=1)

    update_stats(cpu_label, ram_label, disk_label, temp_label, host_label, ip_label, root)
    root.mainloop()


# =============================
#   ENTRY POINT
# =============================

if __name__ == "__main__":
    if psutil is None:
        messagebox.showerror(
            "Missing Dependency",
            "psutil is not installed.\n\nRun:\n\n    pip install psutil\n"
        )
    else:
        create_gui()
