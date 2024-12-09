import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import datetime
import subprocess
import os
import webbrowser
from pystray import Icon as TrayIcon, MenuItem as Item, Menu
from PIL import Image
import json
import logging

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from comtypes import CLSCTX_ALL
except ImportError:
    AudioUtilities = None

# 配置日志记录
logging.basicConfig(
    filename='log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class Task:
    def __init__(self, start_time, end_time, path, days, volume):
        self.start_time = start_time
        self.end_time = end_time
        self.path = path
        self.days = days
        self.volume = volume
        self.process = None

class SchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("定时播放工具")
        self.tasks = []
        self.tray_icon_initialized = False

        # 创建托盘图标
        self.tray_icon = TrayIcon("定时播放软件", self.load_icon(), menu=Menu(
            Item('显示', self.show_window),
            Item('退出', self.exit_app)
        ))

        self.create_widgets()  # 先创建UI组件
        self.load_tasks()      # 然后加载任务

        self.update_time()
        self.check_tasks()

        # 重写关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

    def create_widgets(self):
        frame = tk.Frame(self.root)
        frame.pack(padx=10, pady=10)

        tk.Label(frame, text="开始时间 (HH:MM):").grid(row=0, column=0, sticky="e")
        self.start_time_entry = tk.Entry(frame)
        self.start_time_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame, text="结束时间 (HH:MM):").grid(row=1, column=0, sticky="e")
        self.end_time_entry = tk.Entry(frame)
        self.end_time_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame, text="音量:").grid(row=0, column=2, sticky="e")
        self.volume_scale = tk.Scale(frame, from_=0, to=100, orient=tk.HORIZONTAL)
        self.volume_scale.set(100)
        self.volume_scale.grid(row=0, column=3, padx=5, pady=5, rowspan=2)

        tk.Label(frame, text="选择重复的天:").grid(row=2, column=0, sticky="e")
        self.days_vars = []
        days_frame = tk.Frame(frame)
        days_frame.grid(row=2, column=1, columnspan=3, padx=5, pady=5)

        days = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        for day in days:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(days_frame, text=day, variable=var)
            chk.pack(side=tk.LEFT)
            self.days_vars.append(var)

        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)

        tk.Button(button_frame, text="选择播放软件", command=self.choose_file).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="添加任务", command=self.add_task).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="删除任务", command=self.delete_task).pack(side=tk.LEFT, padx=5)

        list_frame = tk.Frame(self.root)
        list_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.task_list = tk.Listbox(list_frame, width=50, height=10)
        self.task_list.grid(row=0, column=0, sticky="nsew")

        y_scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.task_list.yview)
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        self.task_list.config(yscrollcommand=y_scrollbar.set)

        x_scrollbar = tk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.task_list.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.task_list.config(xscrollcommand=x_scrollbar.set)

        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.time_label = tk.Label(self.root, text="", font=("Arial", 10))
        self.time_label.pack(pady=5)

        author_label = tk.Label(self.root, text="作者: 掌昆运维部", font=("Arial", 10), fg="blue", cursor="hand2")
        author_label.pack(pady=5)
        author_label.bind("<Button-1>", lambda e: webbrowser.open("https://ewain.top"))

    def update_time(self):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=f"当前时间: {now}")
        self.root.after(1000, self.update_time)

    def choose_file(self):
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            messagebox.showinfo("选择路径", f"已选择: {self.file_path}")

    def add_task(self):
        start_time_str = self.start_time_entry.get()
        end_time_str = self.end_time_entry.get()
        days = [var.get() for var in self.days_vars]
        volume = self.volume_scale.get() / 100
        try:
            start_time = datetime.datetime.strptime(start_time_str, "%H:%M").time()
            end_time = datetime.datetime.strptime(end_time_str, "%H:%M").time()
            if hasattr(self, 'file_path') and self.file_path:
                task = Task(start_time, end_time, self.file_path, days, volume)
                self.tasks.append(task)
                days_str = ','.join([day for day, var in zip(["周一", "周二", "周三", "周四", "周五", "周六", "周日"], days) if var])
                if not days_str:
                    days_str = "一次性"
                self.task_list.insert(tk.END, f"{start_time} - {end_time} - {os.path.basename(self.file_path)} - 音量: {volume*100}% - {days_str}")
                self.start_time_entry.delete(0, tk.END)
                self.end_time_entry.delete(0, tk.END)
                for var in self.days_vars:
                    var.set(False)
                self.volume_scale.set(100)
                self.save_tasks()
            else:
                messagebox.showwarning("错误", "请选择播放软件路径")
        except ValueError:
            messagebox.showwarning("错误", "时间格式错误，请使用 HH:MM 格式")

    def delete_task(self):
        selected = self.task_list.curselection()
        if selected:
            index = selected[0]
            self.tasks.pop(index)
            self.task_list.delete(index)
            self.save_tasks()
        else:
            messagebox.showwarning("错误", "请选择要删除的任务")

    def check_tasks(self):
        now = datetime.datetime.now()
        weekday = now.weekday()
        current_time = now.time()

        for index in range(len(self.tasks) - 1, -1, -1):
            task = self.tasks[index]
            if (task.days[weekday] or not any(task.days)) and task.start_time <= current_time < task.end_time:
                if task.process is None:
                    self.run_task(task)
            elif task.process and current_time >= task.end_time:
                self.end_task(task)
                if not any(task.days):  # 如果是一次性任务
                    self.tasks.pop(index)
                    self.task_list.delete(index)
                    self.save_tasks()  # 更新保存的任务

        self.root.after(1000, self.check_tasks)

    def run_task(self, task):
        try:
            if AudioUtilities:
                self.set_system_volume(task.volume)
            task.process = subprocess.Popen(task.path)
            logging.info(f"Successfully started task: {task.path} with volume: {task.volume*100}%")
        except Exception as e:
            logging.error(f"Failed to start task: {task.path} - Error: {e}")

    def end_task(self, task):
        try:
            if task.process:
                task.process.terminate()
                task.process = None
            if AudioUtilities:
                self.set_system_volume(0.0)
            logging.info(f"Successfully ended task: {task.path}")
        except Exception as e:
            logging.error(f"Failed to end task: {task.path} - Error: {e}")

    def set_system_volume(self, volume_level):
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            
            # 设置音量
            volume.SetMasterVolumeLevelScalar(max(volume_level, 0.0), None)
            
            # 设置静音状态
            if volume_level == 0:
                volume.SetMute(1, None)
            else:
                volume.SetMute(0, None)
            
            logging.info(f"Volume set to: {volume_level*100}%")
        except Exception as e:
            logging.error(f"Failed to set volume: {e}")


    def minimize_to_tray(self):
        print("Minimizing to tray")
        self.root.withdraw()
        if not self.tray_icon_initialized:
            self.tray_icon.run_detached()
            self.tray_icon_initialized = True
        else:
            self.tray_icon.visible = True

    def show_window(self, icon, item):
        print("Restoring window")
        self.root.deiconify()
        icon.visible = False

    def exit_app(self, icon, item):
        print("Exiting application")
        icon.visible = False
        icon.stop()
        self.root.quit()

    def load_icon(self):
        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, "favicon.ico")
        else:
            icon_path = "favicon.ico"
        return Image.open(icon_path)

    def save_tasks(self):
        tasks_data = [{
            'start_time': task.start_time.strftime("%H:%M"),
            'end_time': task.end_time.strftime("%H:%M"),
            'path': task.path,
            'days': task.days,
            'volume': task.volume
        } for task in self.tasks]

        with open('tasks.json', 'w', encoding='utf-8') as f:
            json.dump(tasks_data, f, ensure_ascii=False, indent=4)

    def load_tasks(self):
        try:
            with open('tasks.json', 'r', encoding='utf-8') as f:
                tasks_data = json.load(f)
                for task_data in tasks_data:
                    start_time = datetime.datetime.strptime(task_data['start_time'], "%H:%M").time()
                    end_time = datetime.datetime.strptime(task_data['end_time'], "%H:%M").time()
                    task = Task(start_time, end_time, task_data['path'], task_data['days'], task_data['volume'])
                    self.tasks.append(task)
                    days_str = ','.join([day for day, var in zip(["周一", "周二", "周三", "周四", "周五", "周六", "周日"], task.days) if var])
                    if not days_str:
                        days_str = "一次性"
                    self.task_list.insert(tk.END, f"{start_time} - {end_time} - {os.path.basename(task_data['path'])} - 音量: {task_data['volume']*100}% - {days_str}")
        except FileNotFoundError:
            pass

if __name__ == "__main__":
    root = tk.Tk()
    app = SchedulerApp(root)
    root.mainloop()