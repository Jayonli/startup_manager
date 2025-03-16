import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class StartupManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("自启动管理器")
        self.root.geometry("400x300")

        # 创建UI组件
        self.create_widgets()

    def create_widgets(self):
        # 欢迎信息
        tk.Label(self.root, text="拖拽或选择需要自启动的软件", font=("Arial", 12)).pack(pady=10)

        # 文件列表框
        self.listbox = tk.Listbox(self.root, width=50, height=10)
        self.listbox.pack(pady=10)

        # 按钮区域
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="选择文件", command=self.select_file).grid(row=0, column=0, padx=5)
        tk.Button(button_frame, text="删除选中", command=self.remove_selected).grid(row=0, column=1, padx=5)
        tk.Button(button_frame, text="保存到自启动", command=self.save_to_startup).grid(row=0, column=2, padx=5)

    def select_file(self):
        # 打开文件选择对话框
        file_path = filedialog.askopenfilename(title="选择需要自启动的软件", filetypes=[("可执行文件", "*.exe")])
        if file_path:
            self.listbox.insert(tk.END, file_path)

    def remove_selected(self):
        # 删除选中的文件路径
        selected_index = self.listbox.curselection()
        if selected_index:
            self.listbox.delete(selected_index)

    def save_to_startup(self):
        # 获取Windows的启动文件夹路径
        startup_folder = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')

        # 遍历列表框中的所有文件路径
        for index in range(self.listbox.size()):
            file_path = self.listbox.get(index)
            if os.path.exists(file_path):
                # 创建快捷方式名称
                shortcut_name = os.path.basename(file_path) + ".lnk"
                shortcut_path = os.path.join(startup_folder, shortcut_name)

                # 创建快捷方式
                self.create_shortcut(file_path, shortcut_path)

        messagebox.showinfo("成功", "所有选中的软件已添加到自启动！")

    def create_shortcut(self, target_path, shortcut_path):
        # 使用Windows的快捷方式创建工具
        from win32com.client import Dispatch
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.save()

if __name__ == "__main__":
    root = tk.Tk()
    app = StartupManagerApp(root)
    root.mainloop()