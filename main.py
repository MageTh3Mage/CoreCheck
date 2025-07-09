import os
import platform
import subprocess
import socket
import psutil
import GPUtil
import wmi
import datetime
import time
import threading
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import darkdetect
import sv_ttk

class SystemDiagnosticsTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows PC Diagnostics Toolkit")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)

        self.set_theme()

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        self.create_header()
        self.create_main_frame()
        self.create_tabs()
        self.create_footer()

        self.c = wmi.WMI()

        self.update_system_info()
    
    def set_theme(self):
        try:
            if darkdetect.isDark():
                sv_ttk.set_theme("dark")
            else:
                sv_ttk.set_theme("light")
        except:
            sv_ttk.set_theme("light")
    
    def create_header(self):
        header_frame = ttk.Frame(self.root)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=LEFT, fill=BOTH, expand=True)
        
        self.logo_img = self.load_image("logo.png", (40, 40))
        logo_label = ttk.Label(title_frame, image=self.logo_img)
        logo_label.pack(side=LEFT, padx=(0, 10))
        
        title_label = ttk.Label(
            title_frame, 
            text="Mage's CoreCheck Toolkit", 
            font=("Segoe UI", 16, "bold")
        )
        title_label.pack(side=LEFT)

        button_frame = ttk.Frame(header_frame)
        button_frame.pack(side=RIGHT)
        
        refresh_btn = ttk.Button(
            button_frame, 
            text="Refresh", 
            command=self.update_system_info,
            style="Accent.TButton"
        )
        refresh_btn.pack(side=LEFT, padx=5)
        
        theme_btn = ttk.Button(
            button_frame, 
            text="Toggle Theme", 
            command=self.toggle_theme
        )
        theme_btn.pack(side=LEFT, padx=5)
        
        exit_btn = ttk.Button(
            button_frame, 
            text="Exit", 
            command=self.root.quit
        )
        exit_btn.pack(side=LEFT, padx=5)
    
    def load_image(self, filename, size):
        try:
            img = Image.open(filename)
        except:
            img = Image.new('RGB', size, color='gray')
        
        img = img.resize(size)
        return ImageTk.PhotoImage(img)
    
    def create_main_frame(self):
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
    
    def create_tabs(self):
        self.tab_control = ttk.Notebook(self.main_frame)
        self.tab_control.pack(expand=1, fill="both")
        
        self.create_system_info_tab()
        self.create_performance_tab()
        self.create_network_tab()
        self.create_disk_tab()
        self.create_processes_tab()
        self.create_tools_tab()
    
    def create_system_info_tab(self):
        tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab, text="System Info")
        
        left_frame = ttk.Frame(tab)
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        
        right_frame = ttk.Frame(tab)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=5, pady=5)
        
        overview_frame = ttk.LabelFrame(left_frame, text="System Overview", padding=10)
        overview_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        self.system_overview_text = Text(
            overview_frame, 
            wrap=WORD, 
            font=("Consolas", 10), 
            height=10,
            padx=5,
            pady=5
        )
        self.system_overview_text.pack(fill=BOTH, expand=True)
        
        hardware_frame = ttk.LabelFrame(left_frame, text="Hardware Details", padding=10)
        hardware_frame.pack(fill=BOTH, expand=True)
        
        self.hardware_details_text = Text(
            hardware_frame, 
            wrap=WORD, 
            font=("Consolas", 10), 
            height=15,
            padx=5,
            pady=5
        )
        self.hardware_details_text.pack(fill=BOTH, expand=True)
        
        windows_frame = ttk.LabelFrame(right_frame, text="Windows Information", padding=10)
        windows_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        self.windows_info_text = Text(
            windows_frame, 
            wrap=WORD, 
            font=("Consolas", 10), 
            height=10,
            padx=5,
            pady=5
        )
        self.windows_info_text.pack(fill=BOTH, expand=True)
        
        battery_frame = ttk.LabelFrame(right_frame, text="Battery Information", padding=10)
        battery_frame.pack(fill=BOTH, expand=True)
        
        self.battery_info_text = Text(
            battery_frame, 
            wrap=WORD, 
            font=("Consolas", 10), 
            height=5,
            padx=5,
            pady=5
        )
        self.battery_info_text.pack(fill=BOTH, expand=True)
    
    def create_performance_tab(self):
        tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab, text="Performance")

        cpu_frame = ttk.LabelFrame(tab, text="CPU Usage", padding=10)
        cpu_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        self.cpu_label = ttk.Label(cpu_frame, text="CPU: 0%", font=("Segoe UI", 12))
        self.cpu_label.pack(anchor=W)
        
        self.cpu_canvas = Canvas(cpu_frame, height=150, bg='white')
        self.cpu_canvas.pack(fill=BOTH, expand=True)
        self.cpu_data = [0] * 60

        mem_frame = ttk.LabelFrame(tab, text="Memory Usage", padding=10)
        mem_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        self.mem_label = ttk.Label(mem_frame, text="Memory: 0% (0 GB / 0 GB)", font=("Segoe UI", 12))
        self.mem_label.pack(anchor=W)
        
        self.mem_canvas = Canvas(mem_frame, height=150, bg='white')
        self.mem_canvas.pack(fill=BOTH, expand=True)
        self.mem_data = [0] * 60

        self.gpu_frame = ttk.LabelFrame(tab, text="GPU Usage", padding=10)
        self.gpu_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        self.gpu_label = ttk.Label(self.gpu_frame, text="No GPU detected", font=("Segoe UI", 12))
        self.gpu_label.pack(anchor=W)
        
        self.gpu_canvas = Canvas(self.gpu_frame, height=150, bg='white')
        self.gpu_canvas.pack(fill=BOTH, expand=True)
        self.gpu_data = [0] * 60

        disk_frame = ttk.LabelFrame(tab, text="Disk Activity", padding=10)
        disk_frame.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        self.disk_label = ttk.Label(disk_frame, text="Disk: 0% (Read: 0 MB/s | Write: 0 MB/s)", font=("Segoe UI", 12))
        self.disk_label.pack(anchor=W)
        
        self.disk_canvas = Canvas(disk_frame, height=150, bg='white')
        self.disk_canvas.pack(fill=BOTH, expand=True)
        self.disk_data = [0] * 60

        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        self.update_performance()
    
    def create_network_tab(self):

        tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab, text="Network")

        interfaces_frame = ttk.LabelFrame(tab, text="Network Interfaces", padding=10)
        interfaces_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        self.interfaces_tree = ttk.Treeview(
            interfaces_frame,
            columns=("name", "ip", "mac", "netmask", "gateway", "speed"),
            show="headings"
        )
        
        self.interfaces_tree.heading("name", text="Interface")
        self.interfaces_tree.heading("ip", text="IP Address")
        self.interfaces_tree.heading("mac", text="MAC Address")
        self.interfaces_tree.heading("netmask", text="Subnet Mask")
        self.interfaces_tree.heading("gateway", text="Gateway")
        self.interfaces_tree.heading("speed", text="Speed (Mbps)")
        
        self.interfaces_tree.column("name", width=150)
        self.interfaces_tree.column("ip", width=120)
        self.interfaces_tree.column("mac", width=120)
        self.interfaces_tree.column("netmask", width=120)
        self.interfaces_tree.column("gateway", width=120)
        self.interfaces_tree.column("speed", width=100)
        
        scrollbar = ttk.Scrollbar(
            interfaces_frame,
            orient="vertical",
            command=self.interfaces_tree.yview
        )
        self.interfaces_tree.configure(yscrollcommand=scrollbar.set)
        
        self.interfaces_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        stats_frame = ttk.LabelFrame(tab, text="Network Statistics", padding=10)
        stats_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        self.network_stats_text = Text(
            stats_frame, 
            wrap=WORD, 
            font=("Consolas", 10), 
            height=10,
            padx=5,
            pady=5
        )
        self.network_stats_text.pack(fill=BOTH, expand=True)

        tools_frame = ttk.LabelFrame(tab, text="Network Tools", padding=10)
        tools_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        ping_frame = ttk.Frame(tools_frame)
        ping_frame.pack(fill=X, pady=(0, 5))
        
        ttk.Label(ping_frame, text="Ping:").pack(side=LEFT)
        self.ping_entry = ttk.Entry(ping_frame)
        self.ping_entry.pack(side=LEFT, fill=X, expand=True, padx=5)
        self.ping_entry.insert(0, "google.com")
        
        ping_btn = ttk.Button(ping_frame, text="Ping", command=self.run_ping)
        ping_btn.pack(side=LEFT)
        
        self.ping_result = Text(
            tools_frame, 
            wrap=WORD, 
            font=("Consolas", 9), 
            height=5,
            padx=5,
            pady=5
        )
        self.ping_result.pack(fill=BOTH, expand=True)

        self.update_network_info()
    
    def create_disk_tab(self):
        """Disk Information Tab"""
        tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab, text="Disks")

        partitions_frame = ttk.LabelFrame(tab, text="Disk Partitions", padding=10)
        partitions_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        self.partitions_tree = ttk.Treeview(
            partitions_frame,
            columns=("device", "mountpoint", "fstype", "total", "used", "free", "percent"),
            show="headings"
        )
        
        self.partitions_tree.heading("device", text="Device")
        self.partitions_tree.heading("mountpoint", text="Mount Point")
        self.partitions_tree.heading("fstype", text="File System")
        self.partitions_tree.heading("total", text="Total Size")
        self.partitions_tree.heading("used", text="Used")
        self.partitions_tree.heading("free", text="Free")
        self.partitions_tree.heading("percent", text="Usage %")
        
        self.partitions_tree.column("device", width=100)
        self.partitions_tree.column("mountpoint", width=150)
        self.partitions_tree.column("fstype", width=100)
        self.partitions_tree.column("total", width=100)
        self.partitions_tree.column("used", width=100)
        self.partitions_tree.column("free", width=100)
        self.partitions_tree.column("percent", width=80)
        
        scrollbar = ttk.Scrollbar(
            partitions_frame,
            orient="vertical",
            command=self.partitions_tree.yview
        )
        self.partitions_tree.configure(yscrollcommand=scrollbar.set)
        
        self.partitions_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        details_frame = ttk.LabelFrame(tab, text="Disk Details", padding=10)
        details_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        self.disk_details_text = Text(
            details_frame, 
            wrap=WORD, 
            font=("Consolas", 10), 
            height=10,
            padx=5,
            pady=5
        )
        self.disk_details_text.pack(fill=BOTH, expand=True)

        tools_frame = ttk.LabelFrame(tab, text="Disk Tools", padding=10)
        tools_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        btn_frame = ttk.Frame(tools_frame)
        btn_frame.pack(fill=X, pady=5)
        
        analyze_btn = ttk.Button(
            btn_frame, 
            text="Analyze Disk Usage", 
            command=self.analyze_disk_usage
        )
        analyze_btn.pack(side=LEFT, padx=5)
        
        cleanup_btn = ttk.Button(
            btn_frame, 
            text="Open Disk Cleanup", 
            command=self.open_disk_cleanup
        )
        cleanup_btn.pack(side=LEFT, padx=5)
        
        self.disk_tools_result = Text(
            tools_frame, 
            wrap=WORD, 
            font=("Consolas", 9), 
            height=5,
            padx=5,
            pady=5
        )
        self.disk_tools_result.pack(fill=BOTH, expand=True)

        self.update_disk_info()
    
    def toggle_windows_processes(self):
        self.update_processes()

    def create_processes_tab(self):
        tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab, text="Processes")

        processes_frame = ttk.LabelFrame(tab, text="Running Processes", padding=10)
        processes_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        self.processes_tree = ttk.Treeview(
            processes_frame,
            columns=("pid", "name", "status", "cpu", "memory", "user"),
            show="headings"
        )
        
        self.processes_tree.heading("pid", text="PID")
        self.processes_tree.heading("name", text="Name")
        self.processes_tree.heading("status", text="Status")
        self.processes_tree.heading("cpu", text="CPU %")
        self.processes_tree.heading("memory", text="Memory (MB)")
        self.processes_tree.heading("user", text="User")
        
        self.processes_tree.column("pid", width=80)
        self.processes_tree.column("name", width=200)
        self.processes_tree.column("status", width=100)
        self.processes_tree.column("cpu", width=80)
        self.processes_tree.column("memory", width=100)
        self.processes_tree.column("user", width=150)
        
        scrollbar = ttk.Scrollbar(
            processes_frame,
            orient="vertical",
            command=self.processes_tree.yview
        )
        self.processes_tree.configure(yscrollcommand=scrollbar.set)
        
        self.processes_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        tools_frame = ttk.LabelFrame(tab, text="Process Tools", padding=10)
        tools_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        btn_frame = ttk.Frame(tools_frame)
        btn_frame.pack(fill=X, pady=5)
        
        refresh_btn = ttk.Button(
            btn_frame, 
            text="Refresh", 
            command=self.update_processes
        )
        refresh_btn.pack(side=LEFT, padx=5)
        
        end_btn = ttk.Button(
            btn_frame, 
            text="End Process", 
            command=self.end_process,
            style="Accent.TButton"
        )
        end_btn.pack(side=LEFT, padx=5)

        self.show_windows_processes = BooleanVar(value=True)
        toggle_win_btn = ttk.Checkbutton(
            btn_frame,
            text=f"Show Default Windows Processes",
            variable=self.show_windows_processes,
            command=self.toggle_windows_processes
        )
        toggle_win_btn.pack(side=LEFT, padx=5)
        
        self.process_details_text = Text(
            tools_frame, 
            wrap=WORD, 
            font=("Consolas", 9), 
            height=5,
            padx=5,
            pady=5
        )
        self.process_details_text.pack(fill=BOTH, expand=True)

        self.update_processes()
    
    def create_tools_tab(self):
        tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab, text="Tools")

        system_tools_frame = ttk.LabelFrame(tab, text="System Tools", padding=10)
        system_tools_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        btn_frame1 = ttk.Frame(system_tools_frame)
        btn_frame1.pack(fill=X, pady=5)
        
        ttk.Button(
            btn_frame1, 
            text="Open System Information", 
            command=self.open_system_info
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame1, 
            text="Open Device Manager", 
            command=self.open_device_manager
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame1, 
            text="Open Event Viewer", 
            command=self.open_event_viewer
        ).pack(side=LEFT, padx=5)
        
        btn_frame2 = ttk.Frame(system_tools_frame)
        btn_frame2.pack(fill=X, pady=5)
        
        ttk.Button(
            btn_frame2, 
            text="Open Task Manager", 
            command=self.open_task_manager
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame2, 
            text="Open Registry Editor", 
            command=self.open_registry_editor
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame2, 
            text="Open Command Prompt", 
            command=self.open_command_prompt
        ).pack(side=LEFT, padx=5)

        diag_tools_frame = ttk.LabelFrame(tab, text="Diagnostic Tools", padding=10)
        diag_tools_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        btn_frame3 = ttk.Frame(diag_tools_frame)
        btn_frame3.pack(fill=X, pady=5)
        
        ttk.Button(
            btn_frame3, 
            text="Run System File Checker", 
            command=self.run_sfc
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame3, 
            text="Run DISM Check", 
            command=self.run_dism
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame3, 
            text="Check Disk", 
            command=self.run_chkdsk
        ).pack(side=LEFT, padx=5)
        
        btn_frame4 = ttk.Frame(diag_tools_frame)
        btn_frame4.pack(fill=X, pady=5)
        
        ttk.Button(
            btn_frame4, 
            text="Generate System Report", 
            command=self.generate_system_report
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame4, 
            text="Check Windows Updates", 
            command=self.check_windows_updates
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame4, 
            text="Restart Explorer", 
            command=self.restart_explorer
        ).pack(side=LEFT, padx=5)

        self.tools_result_text = Text(
            diag_tools_frame, 
            wrap=WORD, 
            font=("Consolas", 9), 
            height=10,
            padx=5,
            pady=5
        )
        self.tools_result_text.pack(fill=BOTH, expand=True)
    
    def create_footer(self):
        footer_frame = ttk.Frame(self.root)
        footer_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        self.status_var = StringVar()
        self.status_var.set("Ready")
        
        status_label = ttk.Label(
            footer_frame,
            textvariable=self.status_var,
            relief=SUNKEN,
            anchor=W,
            padding=5
        )
        status_label.pack(fill=X)
        
        copyright_label = ttk.Label(
            footer_frame,
            text="2025 CoreCheck",
            anchor=E,
            padding=5
        )
        copyright_label.pack(fill=X)
    
    def toggle_theme(self):
        current_theme = sv_ttk.get_theme()
        if current_theme == "light":
            sv_ttk.set_theme("dark")
        else:
            sv_ttk.set_theme("light")
    
    def update_system_info(self):
        self.status_var.set("Updating system information...")
        
        self.system_overview_text.config(state=NORMAL)
        self.system_overview_text.delete(1.0, END)
        self.hardware_details_text.delete(1.0, END)
        self.windows_info_text.delete(1.0, END)
        self.battery_info_text.delete(1.0, END)
        
        overview = self.get_system_overview()
        self.system_overview_text.insert(END, overview)
        self.system_overview_text.config(state=DISABLED)
        
        hardware = self.get_hardware_details()
        self.hardware_details_text.insert(END, hardware)
        self.hardware_details_text.config(state=DISABLED)

        windows_info = self.get_windows_info()
        self.windows_info_text.insert(END, windows_info)
        self.windows_info_text.config(state=DISABLED)

        battery_info = self.get_battery_info()
        self.battery_info_text.insert(END, battery_info)
        self.battery_info_text.config(state=DISABLED)
        
        self.status_var.set("System information updated")
    
    def get_system_overview(self):
        """Get system overview information"""
        try:
            uname = platform.uname()
            boot_time = datetime.datetime.fromtimestamp(psutil.boot_time())
            
            info = f"""System Overview:
- System: {uname.system}
- Node Name: {uname.node}
- Release: {uname.release}
- Version: {uname.version}
- Machine: {uname.machine}
- Processor: {uname.processor}
- Boot Time: {boot_time.strftime("%Y-%m-%d %H:%M:%S")}
- Uptime: {str(datetime.timedelta(seconds=int(time.time() - psutil.boot_time())))}

"""
            return info
        except Exception as e:
            return f"Error getting system overview: {str(e)}"
    
    def get_hardware_details(self):
        """Get detailed hardware information"""
        try:
            info = "Hardware Details:\n"

            info += "\nCPU:\n"
            info += f"- Physical cores: {psutil.cpu_count(logical=False)}\n"
            info += f"- Total cores: {psutil.cpu_count(logical=True)}\n"
            
            cpufreq = psutil.cpu_freq()
            info += f"- Max Frequency: {cpufreq.max:.2f}Mhz\n"
            info += f"- Min Frequency: {cpufreq.min:.2f}Mhz\n"
            info += f"- Current Frequency: {cpufreq.current:.2f}Mhz\n"

            svmem = psutil.virtual_memory()
            info += "\nMemory:\n"
            info += f"- Total: {self.get_size(svmem.total)}\n"
            info += f"- Available: {self.get_size(svmem.available)}\n"
            info += f"- Used: {self.get_size(svmem.used)}\n"
            info += f"- Percentage: {svmem.percent}%\n"

            swap = psutil.swap_memory()
            info += "\nSWAP:\n"
            info += f"- Total: {self.get_size(swap.total)}\n"
            info += f"- Free: {self.get_size(swap.free)}\n"
            info += f"- Used: {self.get_size(swap.used)}\n"
            info += f"- Percentage: {swap.percent}%\n"

            try:
                gpus = GPUtil.getGPUs()
                if gpus:
                    info += "\nGPU:\n"
                    for gpu in gpus:
                        info += f"- ID: {gpu.id}, Name: {gpu.name}\n"
                        info += f"- Load: {gpu.load*100:.1f}%\n"
                        info += f"- Free Memory: {gpu.memoryFree}MB\n"
                        info += f"- Used Memory: {gpu.memoryUsed}MB\n"
                        info += f"- Total Memory: {gpu.memoryTotal}MB\n"
                        info += f"- Temperature: {gpu.temperature} °C\n"
                else:
                    info += "\nNo NVIDIA GPU detected\n"
            except:
                info += "\nGPU information not available\n"
            
            return info
        except Exception as e:
            return f"Error getting hardware details: {str(e)}"
    
    def get_windows_info(self):
        try:
            info = "Windows Information:\n"
            
            # Get Windows edition
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:
                    product_name = winreg.QueryValueEx(key, "ProductName")[0]
                    release_id = winreg.QueryValueEx(key, "ReleaseId")[0] if "ReleaseId" in \
                        [winreg.EnumValue(key, i)[0] for i in range(winreg.QueryInfoKey(key)[1])] else "N/A"
                    build_number = winreg.QueryValueEx(key, "CurrentBuildNumber")[0]
                    
                    info += f"- Edition: {product_name}\n"
                    info += f"- Release ID: {release_id}\n"
                    info += f"- Build Number: {build_number}\n"
            except:
                info += "- Edition: Could not determine\n"

            try:
                for os_info in self.c.Win32_OperatingSystem():
                    install_date = os_info.InstallDate.split('.')[0]
                    install_date = datetime.datetime.strptime(install_date, '%Y%m%d%H%M%S')
                    info += f"- Install Date: {install_date.strftime('%Y-%m-%d %H:%M:%S')}\n"
            except:
                info += "- Install Date: Could not determine\n"

            try:
                info += f"- System Directory: {os.environ['SystemRoot']}\n"
            except:
                info += "- System Directory: Could not determine\n"

            try:
                for os_info in self.c.Win32_OperatingSystem():
                    info += f"- Registered Owner: {os_info.RegisteredUser}\n"
                    info += f"- Organization: {os_info.Organization if os_info.Organization else 'N/A'}\n"
            except:
                info += "- Registered Owner/Org: Could not determine\n"

            try:
                update_session = wmi.WMI().Win32_WindowsUpdateAgentSession()
                last_update = max([session.LastModifiedTime for session in update_session])
                last_update = datetime.datetime.strptime(last_update.split('.')[0], '%Y%m%d%H%M%S')
                info += f"- Last Update Check: {last_update.strftime('%Y-%m-%d %H:%M:%S')}\n"
            except:
                info += "- Last Update Check: Could not determine\n"
            
            return info
        except Exception as e:
            return f"Error getting Windows info: {str(e)}"
    
    def get_battery_info(self):
        try:
            if not hasattr(psutil, "sensors_battery"):
                return "Battery information not available on this system"
            
            battery = psutil.sensors_battery()
            if battery is None:
                return "No battery detected"
            
            info = "Battery Information:\n"
            info += f"- Percent: {battery.percent}%\n"
            info += f"- Power Plugged: {'Yes' if battery.power_plugged else 'No'}\n"
            
            if battery.secsleft != psutil.POWER_TIME_UNLIMITED:
                if battery.secsleft == psutil.POWER_TIME_UNKNOWN:
                    info += "- Time Left: Unknown\n"
                else:
                    info += f"- Time Left: {str(datetime.timedelta(seconds=battery.secsleft))}\n"
            
            return info
        except Exception as e:
            return f"Error getting battery info: {str(e)}"
    
    def update_performance(self):
        cpu_percent = psutil.cpu_percent(interval=0.1)
        self.cpu_label.config(text=f"CPU: {cpu_percent}%")
        self.cpu_data.pop(0)
        self.cpu_data.append(cpu_percent)
        self.draw_graph(self.cpu_canvas, self.cpu_data, "CPU Usage %", "red")

        mem = psutil.virtual_memory()
        self.mem_label.config(text=f"Memory: {mem.percent}% ({self.get_size(mem.used)} / {self.get_size(mem.total)})")

        self.mem_data.pop(0)
        self.mem_data.append(mem.percent)
        self.draw_graph(self.mem_canvas, self.mem_data, "Memory Usage %", "blue")

        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                gpu = gpus[0]
                self.gpu_label.config(text=f"GPU: {gpu.name} - Load: {gpu.load*100:.1f}%")

                self.gpu_data.pop(0)
                self.gpu_data.append(gpu.load*100)
                self.draw_graph(self.gpu_canvas, self.gpu_data, "GPU Load %", "green")
        except:
            pass

        disk = psutil.disk_io_counters()
        read_mb = disk.read_bytes / (1024 * 1024)
        write_mb = disk.write_bytes / (1024 * 1024)
        
        disk_percent = psutil.disk_usage('/').percent if os.name == 'posix' else psutil.disk_usage('C:\\').percent
        self.disk_label.config(text=f"Disk: {disk_percent}% (Read: {read_mb:.2f} MB/s | Write: {write_mb:.2f} MB/s)")

        self.disk_data.pop(0)
        self.disk_data.append(disk_percent)
        self.draw_graph(self.disk_canvas, self.disk_data, "Disk Usage %", "purple")

        self.root.after(1000, self.update_performance)
    
    def draw_graph(self, canvas, data, title, color):
        canvas.delete("all")
        width = canvas.winfo_width()
        height = canvas.winfo_height()

        canvas.create_rectangle(0, 0, width, height, outline="gray")

        canvas.create_text(width//2, 10, text=title, anchor=N, font=("Arial", 8))

        max_val = max(data) if max(data) > 0 else 100
        canvas.create_text(5, 10, text=f"{max_val}", anchor=NW, font=("Arial", 7))
        canvas.create_text(5, height//2, text=f"{max_val//2}", anchor=W, font=("Arial", 7))
        canvas.create_text(5, height-10, text="0", anchor=SW, font=("Arial", 7))

        canvas.create_text(width-5, height-10, text="now", anchor=SE, font=("Arial", 7))
        canvas.create_text(width//2, height-10, text="30s ago", anchor=S, font=("Arial", 7))
        canvas.create_text(5, height-10, text="60s ago", anchor=SW, font=("Arial", 7))

        x_scale = width / len(data)
        y_scale = (height - 20) / max_val
        
        points = []
        for i, value in enumerate(data):
            x = i * x_scale
            y = height - (value * y_scale)
            points.extend([x, y])
        
        if len(points) > 2:
            canvas.create_line(points, fill=color, width=2)

        last_val = data[-1]
        last_y = height - (last_val * y_scale)
        canvas.create_text(width-5, last_y, text=f"{last_val:.1f}", anchor=E, font=("Arial", 8), fill=color)
    
    def update_network_info(self):
        try:
            for item in self.interfaces_tree.get_children():
                self.interfaces_tree.delete(item)
            
            interfaces = psutil.net_if_addrs()
            stats = psutil.net_if_stats()
            io_counters = psutil.net_io_counters(pernic=True)
            
            for interface, addrs in interfaces.items():
                ipv4 = ""
                ipv6 = ""
                mac = ""
                netmask = ""
                
                for addr in addrs:
                    if addr.family == socket.AF_INET:
                        ipv4 = addr.address
                        netmask = addr.netmask
                    elif addr.family == socket.AF_INET6:
                        ipv6 = addr.address
                    elif addr.family == psutil.AF_LINK:
                        mac = addr.address
                gateway = "N/A"
                try:
                    routes = subprocess.check_output(["route", "print"], text=True, shell=True)
                    for line in routes.split('\n'):
                        if interface in line and "0.0.0.0" in line:
                            parts = line.split()
                            if len(parts) > 2:
                                gateway = parts[2]
                                break
                except:
                    pass

                speed = "N/A"
                if interface in stats:
                    speed = stats[interface].speed
                    if speed > 0:
                        speed = f"{speed} Mbps"
                    else:
                        speed = "N/A"

                self.interfaces_tree.insert(
                    "", 
                    "end", 
                    values=(interface, ipv4 if ipv4 else ipv6, mac, netmask, gateway, speed)
                )

            self.network_stats_text.config(state=NORMAL)
            self.network_stats_text.delete(1.0, END)
            
            total_io = psutil.net_io_counters()
            stats_text = f"""Network Statistics:
- Bytes Sent: {self.get_size(total_io.bytes_sent)}
- Bytes Received: {self.get_size(total_io.bytes_recv)}
- Packets Sent: {total_io.packets_sent:,}
- Packets Received: {total_io.packets_recv:,}
- Errors In: {total_io.errin:,}
- Errors Out: {total_io.errout:,}
- Drops In: {total_io.dropin:,}
- Drops Out: {total_io.dropout:,}
"""
            self.network_stats_text.insert(END, stats_text)
            self.network_stats_text.config(state=DISABLED)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get network info: {str(e)}")
    
    def update_disk_info(self):
        try:
            for item in self.partitions_tree.get_children():
                self.partitions_tree.delete(item)
            
            partitions = psutil.disk_partitions()
            
            for partition in partitions:
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    
                    self.partitions_tree.insert(
                        "", 
                        "end", 
                        values=(
                            partition.device,
                            partition.mountpoint,
                            partition.fstype,
                            self.get_size(usage.total),
                            self.get_size(usage.used),
                            self.get_size(usage.free),
                            f"{usage.percent}%"
                        )
                    )
                except:
                    continue

            self.disk_details_text.config(state=NORMAL)
            self.disk_details_text.delete(1.0, END)

            io_counters = psutil.disk_io_counters()
            details_text = f"""Disk I/O Statistics:
- Read Count: {io_counters.read_count:,}
- Write Count: {io_counters.write_count:,}
- Read Bytes: {self.get_size(io_counters.read_bytes)}
- Write Bytes: {self.get_size(io_counters.write_bytes)}
- Read Time: {io_counters.read_time} ms
- Write Time: {io_counters.write_time} ms
"""
            self.disk_details_text.insert(END, details_text)
            self.disk_details_text.config(state=DISABLED)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get disk info: {str(e)}")
    
    def update_processes(self):
        try:
            for item in self.processes_tree.get_children():
                self.processes_tree.delete(item)

            windows_processes = [
                # Core System Processes
                "System",
                "System Idle Process",
                "smss.exe",          # Session Manager Subsystem
                "csrss.exe",         # Client Server Runtime Process
                "wininit.exe",       # Windows Initialization Process
                "winlogon.exe",      # Windows Logon Application
                "services.exe",      # Service Control Manager
                "lsass.exe",         # Local Security Authority Subsystem
                "svchost.exe",       # Service Host Process (multiple instances)
                "dwm.exe",           # Desktop Window Manager
                "explorer.exe",      # Windows Explorer
                "taskhost.exe",      # Host Process for Windows Tasks
                "taskhostw.exe",     # Host Process for Windows Tasks (alternate)
                "taskeng.exe",       # Task Scheduler Engine
                "rundll32.exe",      # Runs DLL functions
                "dllhost.exe",       # COM+ DLL Host Process
                "conhost.exe",       # Console Window Host
                
                # System Services
                "spoolsv.exe",       # Print Spooler
                "msmpeng.exe",       # Windows Defender Antivirus
                "NisSrv.exe",        # Windows Defender Network Inspection
                "SecurityHealthService.exe", # Windows Security Health Service
                "WmiPrvSE.exe",      # WMI Provider Host
                "SearchIndexer.exe", # Windows Search Indexer
                "SearchFilterHost.exe",
                "SearchProtocolHost.exe",
                "audiodg.exe",       # Windows Audio Device Graph
                "sihost.exe",        # Shell Infrastructure Host
                "ctfmon.exe",        # Alternative User Input Text Input Processor
                "RuntimeBroker.exe", # Runtime Broker (UWP app helper)
                "CompPkgSrv.exe",    # Component Package Support Server
                
                # Network Related
                "wlanext.exe",       # Windows Wireless LAN Extension
                "dasHost.exe",       # Device Association Framework Provider Host
                "dns.exe",           # DNS Server
                "iphlpsvc.exe",      # IP Helper Service
                
                # Windows Subsystem
                "fontdrvhost.exe",   # Font Driver Host
                "lsm.exe",           # Local Session Manager
                "mmc.exe",           # Microsoft Management Console
                "mstsc.exe",         # Remote Desktop Connection
                "rdpclip.exe",       # RDP Clipboard Monitor
                
                # Additional Common Processes
                "backgroundTaskHost.exe",
                "Registry",          # Windows Registry
                "Memory Compression",
                "MoUsoCoreWorker.exe", # Windows Update Standalone Installer
                "usocoreworker.exe", # Update Session Orchestrator
                "ApplicationFrameHost.exe", # UWP Application Frame Host
                "StartMenuExperienceHost.exe",
                "ShellExperienceHost.exe",
                "LockApp.exe",
                "SettingSyncHost.exe",
                "PhoneExperienceHost.exe",
                "GameBarPresenceWriter.exe",
                "TextInputHost.exe",
                "Calculator.exe",     # Windows Calculator
                "MicrosoftEdgeCP.exe",# Microsoft Edge Content Process
                "browser_broker.exe", # Microsoft Edge Browser Broker
                "Widgets.exe",        # Windows Widgets
                "WidgetService.exe",
            ]

            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'status', 'cpu_percent', 'memory_info', 'username']):
                try:
                    if (not getattr(self, 'show_windows_processes', True).get() and 
                        proc.info['name'] in windows_processes):
                        continue
                        
                    processes.append((
                        proc.info['pid'],
                        proc.info['name'],
                        proc.info['status'],
                        proc.info['cpu_percent'],
                        proc.info['memory_info'].rss / (1024 * 1024),  # Convert to MB
                        proc.info['username']
                    ))
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue

            processes.sort(key=lambda p: p[3], reverse=True)
            
            for p in processes:
                self.processes_tree.insert(
                    "", 
                    "end", 
                    values=(
                        p[0], 
                        p[1], 
                        p[2], 
                        f"{p[3]:.1f}", 
                        f"{p[4]:.1f}", 
                        p[5]
                    )
                )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get processes: {str(e)}")
    
    def run_ping(self):
        host = self.ping_entry.get().strip()
        if not host:
            messagebox.showwarning("Warning", "Please enter a host to ping")
            return
        
        self.ping_result.config(state=NORMAL)
        self.ping_result.delete(1.0, END)
        self.ping_result.insert(END, f"Pinging {host}...\n")
        self.ping_result.see(END)
        self.ping_result.config(state=DISABLED)
        
        def ping_thread():
            try:
                result = subprocess.run(
                    ["ping", "-n", "4", host],
                    capture_output=True,
                    text=True,
                    shell=True
                )
                
                self.ping_result.config(state=NORMAL)
                self.ping_result.insert(END, result.stdout)
                if result.stderr:
                    self.ping_result.insert(END, f"\nError: {result.stderr}")
                self.ping_result.see(END)
                self.ping_result.config(state=DISABLED)
                
            except Exception as e:
                self.ping_result.config(state=NORMAL)
                self.ping_result.insert(END, f"\nError: {str(e)}")
                self.ping_result.see(END)
                self.ping_result.config(state=DISABLED)
        
        threading.Thread(target=ping_thread, daemon=True).start()
    
    def analyze_disk_usage(self):
        directory = filedialog.askdirectory(title="Select directory to analyze")
        if not directory:
            return
        
        self.disk_tools_result.config(state=NORMAL)
        self.disk_tools_result.delete(1.0, END)
        self.disk_tools_result.insert(END, f"Analyzing disk usage for: {directory}\nThis may take a while...\n")
        self.disk_tools_result.see(END)
        self.disk_tools_result.config(state=DISABLED)
        
        def analyze_thread():
            try:
                try:
                    result = subprocess.run(
                        ["du", "-h", "--max-depth=1", directory],
                        capture_output=True,
                        text=True,
                        shell=True
                    )
                    
                    output = result.stdout
                    if not output:
                        output = result.stderr
                    
                    self.disk_tools_result.config(state=NORMAL)
                    self.disk_tools_result.insert(END, output)
                    self.disk_tools_result.see(END)
                    self.disk_tools_result.config(state=DISABLED)
                except:
                    sizes = {}
                    total_size = 0
                    
                    for root, dirs, files in os.walk(directory):
                        for name in files:
                            try:
                                path = os.path.join(root, name)
                                size = os.path.getsize(path)
                                sizes[root] = sizes.get(root, 0) + size
                                total_size += size
                            except:
                                continue

                    sorted_sizes = sorted(sizes.items(), key=lambda x: x[1], reverse=True)
                    
                    self.disk_tools_result.config(state=NORMAL)
                    self.disk_tools_result.insert(END, f"Total size: {self.get_size(total_size)}\n\n")
                    self.disk_tools_result.insert(END, "Largest directories:\n")
                    
                    for path, size in sorted_sizes[:20]: 
                        self.disk_tools_result.insert(END, f"{self.get_size(size)} - {path}\n")
                    
                    self.disk_tools_result.see(END)
                    self.disk_tools_result.config(state=DISABLED)
                
            except Exception as e:
                self.disk_tools_result.config(state=NORMAL)
                self.disk_tools_result.insert(END, f"\nError: {str(e)}")
                self.disk_tools_result.see(END)
                self.disk_tools_result.config(state=DISABLED)
        
        threading.Thread(target=analyze_thread, daemon=True).start()
    
    def open_disk_cleanup(self):
        try:
            subprocess.Popen(["cleanmgr"], shell=True)
            self.disk_tools_result.config(state=NORMAL)
            self.disk_tools_result.insert(END, "Launched Disk Cleanup utility\n")
            self.disk_tools_result.see(END)
            self.disk_tools_result.config(state=DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Disk Cleanup: {str(e)}")
    
    def end_process(self):
        selected = self.processes_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a process to end")
            return
        
        item = self.processes_tree.item(selected[0])
        pid = int(item['values'][0])
        name = item['values'][1]
        
        if not messagebox.askyesno("Confirm", f"Are you sure you want to end process {name} (PID: {pid})?"):
            return
        
        try:
            process = psutil.Process(pid)
            process.terminate()
            
            self.process_details_text.config(state=NORMAL)
            self.process_details_text.insert(END, f"Successfully terminated process {name} (PID: {pid})\n")
            self.process_details_text.see(END)
            self.process_details_text.config(state=DISABLED)
            
            # Update processes list
            self.update_processes()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to end process: {str(e)}")
    
    def open_system_info(self):
        try:
            subprocess.Popen(["msinfo32"], shell=True)
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, "Launched System Information\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open System Information: {str(e)}")
    
    def open_device_manager(self):
        try:
            subprocess.Popen(["devmgmt.msc"], shell=True)
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, "Launched Device Manager\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Device Manager: {str(e)}")
    
    def open_event_viewer(self):
        try:
            subprocess.Popen(["eventvwr.msc"], shell=True)
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, "Launched Event Viewer\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Event Viewer: {str(e)}")
    
    def open_task_manager(self):
        try:
            subprocess.Popen(["taskmgr"], shell=True)
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, "Launched Task Manager\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Task Manager: {str(e)}")
    
    def open_registry_editor(self):
        try:
            subprocess.Popen(["regedit"], shell=True)
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, "Launched Registry Editor\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Registry Editor: {str(e)}")
    
    def open_command_prompt(self):
        try:
            subprocess.Popen(["cmd"], shell=True)
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, "Launched Command Prompt\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Command Prompt: {str(e)}")
    
    def run_sfc(self):
        self.tools_result_text.config(state=NORMAL)
        self.tools_result_text.delete(1.0, END)
        self.tools_result_text.insert(END, "Running System File Checker (sfc /scannow)...\nThis may take several minutes.\n")
        self.tools_result_text.see(END)
        self.tools_result_text.config(state=DISABLED)
        
        def sfc_thread():
            try:
                result = subprocess.run(
                    ["sfc", "/scannow"],
                    capture_output=True,
                    text=True,
                    shell=True
                )
                
                self.tools_result_text.config(state=NORMAL)
                self.tools_result_text.insert(END, result.stdout)
                if result.stderr:
                    self.tools_result_text.insert(END, f"\nError: {result.stderr}")
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
                
            except Exception as e:
                self.tools_result_text.config(state=NORMAL)
                self.tools_result_text.insert(END, f"\nError: {str(e)}")
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
        
        threading.Thread(target=sfc_thread, daemon=True).start()
    
    def run_dism(self):
        self.tools_result_text.config(state=NORMAL)
        self.tools_result_text.delete(1.0, END)
        self.tools_result_text.insert(END, "Running DISM health check (DISM /Online /Cleanup-Image /CheckHealth)...\nThis may take a while.\n")
        self.tools_result_text.see(END)
        self.tools_result_text.config(state=DISABLED)
        
        def dism_thread():
            try:
                result = subprocess.run(
                    ["DISM", "/Online", "/Cleanup-Image", "/CheckHealth"],
                    capture_output=True,
                    text=True,
                    shell=True
                )
                
                self.tools_result_text.config(state=NORMAL)
                self.tools_result_text.insert(END, result.stdout)
                if result.stderr:
                    self.tools_result_text.insert(END, f"\nError: {result.stderr}")
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
                
            except Exception as e:
                self.tools_result_text.config(state=NORMAL)
                self.tools_result_text.insert(END, f"\nError: {str(e)}")
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
        
        threading.Thread(target=dism_thread, daemon=True).start()
    
    def run_chkdsk(self):
        if not messagebox.askyesno("Confirm", "CHKDSK will run at the next system restart. Continue?"):
            return
        
        self.tools_result_text.config(state=NORMAL)
        self.tools_result_text.delete(1.0, END)
        self.tools_result_text.insert(END, "Scheduling CHKDSK to run at next boot...\n")
        self.tools_result_text.see(END)
        self.tools_result_text.config(state=DISABLED)
        
        def chkdsk_thread():
            try:
                result = subprocess.run(
                    ["chkdsk", "/f", "C:"],
                    capture_output=True,
                    text=True,
                    shell=True
                )
                
                self.tools_result_text.config(state=NORMAL)
                self.tools_result_text.insert(END, result.stdout)
                if result.stderr:
                    self.tools_result_text.insert(END, f"\nError: {result.stderr}")
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
                
            except Exception as e:
                self.tools_result_text.config(state=NORMAL)
                self.tools_result_text.insert(END, f"\nError: {str(e)}")
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
        
        threading.Thread(target=chkdsk_thread, daemon=True).start()
    
    def generate_system_report(self):
        try:

            report = "=== SYSTEM DIAGNOSTICS REPORT ===\n\n"
            report += self.get_system_overview() + "\n"
            report += self.get_hardware_details() + "\n"
            report += self.get_windows_info() + "\n"
            report += self.get_battery_info() + "\n"
            

            report += "\n=== NETWORK INFORMATION ===\n"
            for item in self.interfaces_tree.get_children():
                values = self.interfaces_tree.item(item)['values']
                report += f"- Interface: {values[0]}\n"
                report += f"  IP: {values[1]}\n"
                report += f"  MAC: {values[2]}\n"
                report += f"  Netmask: {values[3]}\n"
                report += f"  Gateway: {values[4]}\n"
                report += f"  Speed: {values[5]}\n"

            report += "\n=== DISK INFORMATION ===\n"
            for item in self.partitions_tree.get_children():
                values = self.partitions_tree.item(item)['values']
                report += f"- Device: {values[0]}\n"
                report += f"  Mount: {values[1]}\n"
                report += f"  FS: {values[2]}\n"
                report += f"  Size: {values[3]}\n"
                report += f"  Used: {values[4]} ({values[6]})\n"

            report += "\n=== TOP PROCESSES ===\n"
            for item in self.processes_tree.get_children()[:10]:  # Top 10
                values = self.processes_tree.item(item)['values']
                report += f"- {values[1]} (PID: {values[0]}) - CPU: {values[3]}%, Mem: {values[4]} MB\n"

            filename = f"system_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, 'w') as f:
                f.write(report)
            
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, f"System report saved to: {os.path.abspath(filename)}\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)

            if messagebox.askyesno("Report Generated", "System report generated successfully. Open the file now?"):
                os.startfile(filename)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate system report: {str(e)}")
    
    def check_windows_updates(self):
        self.tools_result_text.config(state=NORMAL)
        self.tools_result_text.delete(1.0, END)
        self.tools_result_text.insert(END, "Checking for Windows updates...\n")
        self.tools_result_text.see(END)
        self.tools_result_text.config(state=DISABLED)
        
        def updates_thread():
            try:
                result = subprocess.run(
                    ["powershell", "Get-WindowsUpdate"],
                    capture_output=True,
                    text=True,
                    shell=True
                )
                
                self.tools_result_text.config(state=NORMAL)
                if result.stdout:
                    self.tools_result_text.insert(END, result.stdout)
                else:
                    self.tools_result_text.insert(END, "No updates found or unable to check updates.\n")
                    self.tools_result_text.insert(END, "Try running Windows Update from Settings for more details.\n")
                
                if result.stderr:
                    self.tools_result_text.insert(END, f"\nError: {result.stderr}")
                
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
                
            except Exception as e:
                self.tools_result_text.config(state=NORMAL)
                self.tools_result_text.insert(END, f"\nError: {str(e)}")
                self.tools_result_text.see(END)
                self.tools_result_text.config(state=DISABLED)
        
        threading.Thread(target=updates_thread, daemon=True).start()
    
    def restart_explorer(self):

        try:

            subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], shell=True)

            subprocess.Popen(["explorer"], shell=True)
            
            self.tools_result_text.config(state=NORMAL)
            self.tools_result_text.insert(END, "Windows Explorer restarted successfully\n")
            self.tools_result_text.see(END)
            self.tools_result_text.config(state=DISABLED)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to restart Explorer: {str(e)}")
    
    def get_size(self, bytes):
        for unit in ['', 'K', 'M', 'G', 'T', 'P']:
            if bytes < 1024:
                return f"{bytes:.2f}{unit}B"
            bytes /= 1024
        return f"{bytes:.2f}EB"

def main():

    if os.name != 'nt':
        print("This tool is designed to run on Windows only.")
        return
    
    try:
        import tkinter
        import psutil
        import wmi
        import GPUtil
    except ImportError as e:
        print(f"Error: Required module not found - {str(e)}")
        print("Please install the required modules with:")
        print("pip install psutil wmi gputil darkdetect sv-ttk")
        return
    
    root = Tk()
    app = SystemDiagnosticsTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
