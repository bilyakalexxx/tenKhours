import sqlite3
import time
import threading
import json
import os
from datetime import datetime
import pandas as pd
import customtkinter as ctk
from PIL import Image

# Windows-specific libraries for active window tracking
import win32gui
import win32process
import psutil

# Configuration & Theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

GROUPS_FILE = "group_containers.json"
TARGET_HOURS = 10000.0


class DatabaseManager:
    """Handles persistent storage using SQLite and Excel exports."""
    def __init__(self, db_name="time_tracker.db"):
        self.db_name = db_name
        self.init_db()

    def init_db(self):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS time_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    group_name TEXT,
                    program_name TEXT,
                    window_title TEXT,
                    start_time TEXT,
                    end_time TEXT,
                    duration_seconds INTEGER,
                    date TEXT
                )
            """)
            conn.commit()

    def log_session(self, group_name, program_name, window_title, start_dt, end_dt, duration_sec):
        if duration_sec < 1:
            return
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO time_logs (group_name, program_name, window_title, start_time, end_time, duration_seconds, date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                group_name,
                program_name,
                window_title,
                start_dt.strftime("%Y-%m-%d %H:%M:%S"),
                end_dt.strftime("%Y-%m-%d %H:%M:%S"),
                duration_sec,
                start_dt.strftime("%Y-%m-%d")
            ))
            conn.commit()

    def get_group_total_seconds(self, group_name):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT SUM(duration_seconds) FROM time_logs WHERE group_name = ?", (group_name,))
            result = cursor.fetchone()[0]
            return result if result else 0

    def get_program_total_seconds(self, group_name, program_name):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT SUM(duration_seconds) FROM time_logs WHERE group_name = ? AND LOWER(program_name) = ?", 
                (group_name, program_name.lower())
            )
            result = cursor.fetchone()[0]
            return result if result else 0

    def export_to_excel(self, filename="time_tracker_report.xlsx"):
        with sqlite3.connect(self.db_name) as conn:
            df = pd.read_sql_query("SELECT * FROM time_logs", conn)
            if df.empty:
                return False, "No data available to export."

            df['Formatted Duration'] = df['duration_seconds'].apply(
                lambda x: f"{x // 3600:02d}:{(x % 3600) // 60:02d}:{x % 60:02d}"
            )
            df.to_excel(filename, index=False)
            return True, f"Successfully exported to {filename}"


def get_active_window_info():
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return "Unknown", "Unknown"

        window_title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        
        try:
            process = psutil.Process(pid)
            program_name = process.name()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            program_name = "System"

        return program_name, window_title if window_title else "No Title"
    except Exception:
        return "Unknown", "Unknown"


class GroupManagerWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Group & Program Manager")
        self.geometry("400x560")
        self.attributes("-topmost", True)

        self.setup_ui()

    def setup_ui(self):
        lbl_add_group = ctk.CTkLabel(self, text="1. Create New Group Container", font=ctk.CTkFont(size=13, weight="bold"))
        lbl_add_group.pack(anchor="w", padx=15, pady=(15, 2))

        self.entry_group_name = ctk.CTkEntry(self, placeholder_text="e.g. 3D Work")
        self.entry_group_name.pack(fill="x", padx=15, pady=5)

        btn_create_group = ctk.CTkButton(self, text="Create Group", fg_color="#2b5c8f", command=self.add_group)
        btn_create_group.pack(padx=15, pady=(0, 10))

        lbl_add_app = ctk.CTkLabel(self, text="2. Manage Selected Group", font=ctk.CTkFont(size=13, weight="bold"))
        lbl_add_app.pack(anchor="w", padx=15, pady=(10, 2))

        group_sel_frame = ctk.CTkFrame(self, fg_color="transparent")
        group_sel_frame.pack(fill="x", padx=15, pady=5)

        self.dropdown_groups = ctk.CTkOptionMenu(
            group_sel_frame, 
            values=list(self.parent.group_containers.keys()),
            command=self.on_group_selected
        )
        self.dropdown_groups.pack(side="left", fill="x", expand=True, padx=(0, 5))

        btn_delete_group = ctk.CTkButton(
            group_sel_frame,
            text="🗑 Delete Group",
            width=100,
            fg_color="#8c2b2b",
            hover_color="#ab3737",
            command=self.delete_group
        )
        btn_delete_group.pack(side="right")

        self.entry_exe_name = ctk.CTkEntry(self, placeholder_text="Paste path or enter program (e.g. maya.exe)")
        self.entry_exe_name.pack(fill="x", padx=15, pady=5)

        btn_add_exe = ctk.CTkButton(self, text="Add Program (.exe)", fg_color="#2b5c8f", command=self.add_exe)
        btn_add_exe.pack(padx=15, pady=(0, 10))

        lbl_assigned = ctk.CTkLabel(self, text="Assigned Programs in Selected Group:", font=ctk.CTkFont(size=11, weight="bold"))
        lbl_assigned.pack(anchor="w", padx=15, pady=(5, 2))

        self.scroll_frame = ctk.CTkScrollableFrame(self, height=130)
        self.scroll_frame.pack(fill="both", expand=True, padx=15, pady=5)

        self.lbl_dialog_status = ctk.CTkLabel(self, text="", text_color="gray", font=ctk.CTkFont(size=11))
        self.lbl_dialog_status.pack(pady=5)

        self.refresh_program_list()

    def on_group_selected(self, selected_group):
        self.refresh_program_list()

    def refresh_program_list(self):
        for child in self.scroll_frame.winfo_children():
            child.destroy()

        target_group = self.dropdown_groups.get()
        if not target_group or target_group not in self.parent.group_containers:
            return

        programs = self.parent.group_containers[target_group]

        if not programs:
            lbl_empty = ctk.CTkLabel(self.scroll_frame, text="No programs added yet.", text_color="gray")
            lbl_empty.pack(pady=10)
            return

        for exe in programs:
            row_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
            row_frame.pack(fill="x", pady=2)

            lbl_exe = ctk.CTkLabel(row_frame, text=exe, anchor="w")
            lbl_exe.pack(side="left", padx=5)

            btn_remove = ctk.CTkButton(
                row_frame, 
                text="❌ Remove", 
                width=75, 
                height=22,
                fg_color="#8c2b2b", 
                hover_color="#ab3737",
                command=lambda e=exe: self.remove_exe(e)
            )
            btn_remove.pack(side="right", padx=5)

    def add_group(self):
        group_name = self.entry_group_name.get().strip()
        if group_name and group_name not in self.parent.group_containers:
            self.parent.group_containers[group_name] = []
            self.parent.save_groups()
            self.parent.refresh_group_dropdown()
            
            groups = list(self.parent.group_containers.keys())
            self.dropdown_groups.configure(values=groups)
            self.dropdown_groups.set(group_name)
            self.entry_group_name.delete(0, 'end')
            
            self.lbl_dialog_status.configure(text=f"Created group: '{group_name}'")
            self.refresh_program_list()

    def delete_group(self):
        target_group = self.dropdown_groups.get()
        if not target_group:
            return

        if len(self.parent.group_containers) <= 1:
            self.lbl_dialog_status.configure(text="Cannot delete the last remaining group.")
            return

        del self.parent.group_containers[target_group]
        self.parent.save_groups()
        self.parent.refresh_group_dropdown()

        groups = list(self.parent.group_containers.keys())
        self.dropdown_groups.configure(values=groups)
        self.dropdown_groups.set(groups[0])
        self.lbl_dialog_status.configure(text=f"Deleted group: '{target_group}'")
        self.refresh_program_list()

    def add_exe(self):
        target_group = self.dropdown_groups.get()
        raw_input = self.entry_exe_name.get().strip()

        if not raw_input:
            return

        exe_name = os.path.basename(raw_input).lower()

        if not exe_name.endswith(".exe"):
            exe_name += ".exe"

        if target_group and exe_name:
            if exe_name not in self.parent.group_containers[target_group]:
                self.parent.group_containers[target_group].append(exe_name)
                self.parent.save_groups()
                self.parent.build_program_progress_bars()
                self.entry_exe_name.delete(0, 'end')
                self.lbl_dialog_status.configure(text=f"Added '{exe_name}' to '{target_group}'")
                self.refresh_program_list()

    def remove_exe(self, exe_name):
        target_group = self.dropdown_groups.get()
        if target_group in self.parent.group_containers:
            if exe_name in self.parent.group_containers[target_group]:
                self.parent.group_containers[target_group].remove(exe_name)
                self.parent.save_groups()
                self.parent.build_program_progress_bars()
                self.lbl_dialog_status.configure(text=f"Removed '{exe_name}' from '{target_group}'")
                self.refresh_program_list()


class TimeTrackerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("tenKhours")
        self.geometry("380x720")
        self.attributes("-topmost", True)

        self.db = DatabaseManager()

        self.group_containers = self.load_groups()
        self.active_target_group = list(self.group_containers.keys())[0]

        self.tracking_enabled = False
        self.is_session_active = False
        self.current_program = ""
        self.current_title = ""
        self.start_timestamp = None
        self.elapsed_seconds = 0

        self.program_progress_widgets = {}

        # Corgi Animation State Variables
        self.corgi_frames = []
        self.corgi_idle = None
        self.corgi_anim_index = 0

        self.load_corgi_sprites()
        self.setup_ui()
        self.update_progress_display()

        # Start animation loop
        self.animate_corgi()

        self.monitor_thread = threading.Thread(target=self.track_loop, daemon=True)
        self.monitor_thread.start()

    def load_corgi_sprites(self):
        """Loads and resizes Corgi PNG frames from the 'corgie' subfolder."""
        size = (48, 48)  # Display size inside UI
        corgie_dir = os.path.join(os.path.dirname(__file__), "corgie")

        # 1. Load punch frames (co_01.png -> co_07.png)
        for i in range(1, 8):
            filepath = os.path.join(corgie_dir, f"co_0{i}.png")
            if os.path.exists(filepath):
                pil_img = Image.open(filepath)
                self.corgi_frames.append(ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=size))

        # 2. Load idle frame (co_idle.png)
        idle_path = os.path.join(corgie_dir, "co_idle.png")
        if os.path.exists(idle_path):
            pil_idle = Image.open(idle_path)
            self.corgi_idle = ctk.CTkImage(light_image=pil_idle, dark_image=pil_idle, size=size)

    def load_groups(self):
        if os.path.exists(GROUPS_FILE):
            try:
                with open(GROUPS_FILE, "r") as f:
                    data = json.load(f)
                    if isinstance(data, dict) and data:
                        return data
            except Exception:
                pass
        return {"Default Group": []}

    def save_groups(self):
        try:
            with open(GROUPS_FILE, "w") as f:
                json.dump(self.group_containers, f, indent=4)
        except Exception as e:
            print(f"Error saving groups: {e}")

    def setup_ui(self):
        self.title_label = ctk.CTkLabel(self, text="tenKhours", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(pady=(15, 5))

        lbl_select_group = ctk.CTkLabel(self, text="Active Group Container:", font=ctk.CTkFont(size=12))
        lbl_select_group.pack(anchor="w", padx=25, pady=(5, 0))

        self.opt_group = ctk.CTkOptionMenu(
            self, 
            values=list(self.group_containers.keys()), 
            command=self.select_active_group
        )
        self.opt_group.pack(fill="x", padx=20, pady=(2, 5))

        self.btn_manage_groups = ctk.CTkButton(
            self, 
            text="⚙ Manage Groups & Programs", 
            fg_color="#3a3d40", 
            hover_color="#4f5256",
            command=self.open_group_manager
        )
        self.btn_manage_groups.pack(fill="x", padx=20, pady=5)

        self.progress_frame = ctk.CTkFrame(self)
        self.progress_frame.pack(fill="x", padx=20, pady=10)

        self.lbl_progress_title = ctk.CTkLabel(
            self.progress_frame, 
            text="Group Total Mastery Progress", 
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_progress_title.pack(anchor="w", padx=10, pady=(8, 2))

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, height=14)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.set(0.0)

        self.lbl_progress_stats = ctk.CTkLabel(
            self.progress_frame, 
            text="0.00 / 10,000 hrs (0.0000%)", 
            font=ctk.CTkFont(size=11), 
            text_color="gray"
        )
        self.lbl_progress_stats.pack(anchor="w", padx=10, pady=(0, 8))

        # Active App Info Card with Corgi Sprite side-by-side
        self.info_frame = ctk.CTkFrame(self)
        self.info_frame.pack(fill="x", padx=20, pady=5)

        self.lbl_corgi = ctk.CTkLabel(self.info_frame, text="", image=self.corgi_idle)
        self.lbl_corgi.pack(side="left", padx=(10, 5), pady=5)

        self.lbl_program = ctk.CTkLabel(self.info_frame, text="App: None", font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_program.pack(side="left", padx=5, pady=8)

        self.lbl_timer = ctk.CTkLabel(self, text="00:00:00", font=ctk.CTkFont(size=32, weight="bold"))
        self.lbl_timer.pack(pady=5)

        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(pady=5)

        self.btn_track = ctk.CTkButton(
            self.btn_frame, 
            text="▶ Track", 
            width=110, 
            fg_color="#2af598", 
            text_color="black", 
            hover_color="#22c479", 
            command=self.start_tracking
        )
        self.btn_track.grid(row=0, column=0, padx=5)

        self.btn_stop = ctk.CTkButton(
            self.btn_frame, 
            text="⏹ Stop", 
            width=110, 
            fg_color="#8c2b2b", 
            hover_color="#ab3737", 
            command=self.stop_tracking
        )
        self.btn_stop.grid(row=0, column=1, padx=5)

        self.btn_export = ctk.CTkButton(self, text="📊 Export to Excel", fg_color="#3c4043", hover_color="#5f6368", command=self.export_excel)
        self.btn_export.pack(pady=5)

        lbl_prog_bars_title = ctk.CTkLabel(self, text="Assigned Programs Progress:", font=ctk.CTkFont(size=11, weight="bold"))
        lbl_prog_bars_title.pack(anchor="w", padx=25, pady=(5, 2))

        self.scroll_programs = ctk.CTkScrollableFrame(self, height=130)
        self.scroll_programs.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        self.lbl_status = ctk.CTkLabel(self, text="Status: Stopped", font=ctk.CTkFont(size=11), text_color="gray")
        self.lbl_status.pack(side="bottom", pady=5)

        self.build_program_progress_bars()

    def animate_corgi(self):
        """Switches frames during active sessions or resets to idle pose."""
        if self.is_session_active and self.corgi_frames:
            current_frame = self.corgi_frames[self.corgi_anim_index]
            self.lbl_corgi.configure(image=current_frame)
            self.corgi_anim_index = (self.corgi_anim_index + 1) % len(self.corgi_frames)
        else:
            if self.corgi_idle:
                self.lbl_corgi.configure(image=self.corgi_idle)
            self.corgi_anim_index = 0

        self.after(120, self.animate_corgi)

    def build_program_progress_bars(self):
        for child in self.scroll_programs.winfo_children():
            child.destroy()
        self.program_progress_widgets.clear()

        assigned_programs = self.group_containers.get(self.active_target_group, [])

        if not assigned_programs:
            lbl_empty = ctk.CTkLabel(self.scroll_programs, text="No programs assigned to group.", text_color="gray", font=ctk.CTkFont(size=11))
            lbl_empty.pack(pady=10)
            return

        for exe in assigned_programs:
            card = ctk.CTkFrame(self.scroll_programs, fg_color="#2b2d30")
            card.pack(fill="x", pady=4, padx=2)

            header_frame = ctk.CTkFrame(card, fg_color="transparent")
            header_frame.pack(fill="x", padx=8, pady=(4, 2))

            lbl_name = ctk.CTkLabel(header_frame, text=exe, font=ctk.CTkFont(size=11, weight="bold"))
            lbl_name.pack(side="left")

            lbl_stats = ctk.CTkLabel(header_frame, text="0.00 hrs (0.0000%)", font=ctk.CTkFont(size=10), text_color="gray")
            lbl_stats.pack(side="right")

            pbar = ctk.CTkProgressBar(card, height=8)
            pbar.pack(fill="x", padx=8, pady=(0, 6))
            pbar.set(0.0)

            self.program_progress_widgets[exe.lower()] = {
                "bar": pbar,
                "stats": lbl_stats
            }

        self.update_progress_display()

    def select_active_group(self, group_name):
        self.active_target_group = group_name
        if self.is_session_active:
            self.end_session()
        self.build_program_progress_bars()

    def refresh_group_dropdown(self):
        groups = list(self.group_containers.keys())
        self.opt_group.configure(values=groups)
        if self.active_target_group not in groups and groups:
            self.active_target_group = groups[0]
            self.opt_group.set(groups[0])
        self.build_program_progress_bars()

    def open_group_manager(self):
        GroupManagerWindow(self)

    def update_progress_display(self):
        past_seconds = self.db.get_group_total_seconds(self.active_target_group)
        current_session_seconds = self.elapsed_seconds if self.is_session_active else 0
        total_seconds = past_seconds + current_session_seconds

        total_hours = total_seconds / 3600.0
        progress_ratio = min(total_hours / TARGET_HOURS, 1.0)
        percentage = (total_hours / TARGET_HOURS) * 100

        self.progress_bar.set(progress_ratio)
        self.lbl_progress_stats.configure(
            text=f"{total_hours:.2f} / 10,000 hrs ({percentage:.4f}%)"
        )

        assigned_programs = self.group_containers.get(self.active_target_group, [])
        for exe in assigned_programs:
            exe_key = exe.lower()
            if exe_key in self.program_progress_widgets:
                prog_past_sec = self.db.get_program_total_seconds(self.active_target_group, exe)
                
                if self.is_session_active and self.current_program.lower() == exe_key:
                    prog_total_sec = prog_past_sec + self.elapsed_seconds
                else:
                    prog_total_sec = prog_past_sec

                prog_hours = prog_total_sec / 3600.0
                prog_ratio = min(prog_hours / TARGET_HOURS, 1.0)
                prog_percentage = (prog_hours / TARGET_HOURS) * 100

                self.program_progress_widgets[exe_key]["bar"].set(prog_ratio)
                self.program_progress_widgets[exe_key]["stats"].configure(
                    text=f"{prog_hours:.2f} hrs ({prog_percentage:.4f}%)"
                )

    def start_tracking(self):
        self.tracking_enabled = True
        self.lbl_status.configure(text="Status: Active Tracking Enabled")

    def stop_tracking(self):
        if self.is_session_active:
            self.end_session()
        self.tracking_enabled = False
        self.elapsed_seconds = 0
        self.update_timer_display()
        self.update_progress_display()
        self.lbl_program.configure(text="App: Stopped")
        self.lbl_status.configure(text="Status: Stopped")

    def track_loop(self):
        while True:
            if self.tracking_enabled:
                prog, title = get_active_window_info()
                prog_lower = prog.lower()

                target_apps = self.group_containers.get(self.active_target_group, [])
                target_apps_lower = [a.lower() for a in target_apps]

                is_matching_app = prog_lower in target_apps_lower

                if is_matching_app:
                    if not self.is_session_active or prog != self.current_program or title != self.current_title:
                        if self.is_session_active:
                            self.end_session()
                        
                        self.current_program = prog
                        self.current_title = title
                        self.begin_session()
                else:
                    if self.is_session_active:
                        self.end_session()
                    
                    self.lbl_program.configure(text=f"App: {prog} [Untracked]")

            if self.is_session_active and self.tracking_enabled:
                self.elapsed_seconds = int(time.time() - self.start_timestamp)
                self.update_timer_display()
                self.update_progress_display()

            time.sleep(1)

    def begin_session(self):
        self.is_session_active = True
        self.start_timestamp = time.time()
        self.lbl_program.configure(text=f"App: {self.current_program}")
        self.lbl_status.configure(text="Status: Tracking focused app")

    def end_session(self):
        if not self.is_session_active or not self.start_timestamp:
            return

        end_time = datetime.now()
        start_time = datetime.fromtimestamp(self.start_timestamp)
        duration = int(time.time() - self.start_timestamp)

        self.db.log_session(
            group_name=self.active_target_group,
            program_name=self.current_program,
            window_title=self.current_title,
            start_dt=start_time,
            end_dt=end_time,
            duration_sec=duration
        )

        self.is_session_active = False
        self.elapsed_seconds = 0
        self.update_timer_display()
        self.update_progress_display()

    def update_timer_display(self):
        hrs, remainder = divmod(self.elapsed_seconds, 3600)
        mins, secs = divmod(remainder, 60)
        self.lbl_timer.configure(text=f"{hrs:02d}:{mins:02d}:{secs:02d}")

    def export_excel(self):
        success, message = self.db.export_to_excel()
        self.lbl_status.configure(text=f"Status: {message}")


if __name__ == "__main__":
    app = TimeTrackerApp()
    app.mainloop()