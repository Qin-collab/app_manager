import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import winreg
import subprocess
import os
from PIL import Image, ImageTk
import json

class AddAppDialog:
    def __init__(self, parent):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("添加应用程序")
        self.dialog.geometry("400x300")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 创建表单
        ttk.Label(self.dialog, text="应用程序名称:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.name_entry = ttk.Entry(self.dialog, width=40)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(self.dialog, text="发布者:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.publisher_entry = ttk.Entry(self.dialog, width=40)
        self.publisher_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(self.dialog, text="程序路径:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.path_frame = ttk.Frame(self.dialog)
        self.path_frame.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        self.path_entry = ttk.Entry(self.path_frame, width=32)
        self.path_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(self.path_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT)
        
        ttk.Label(self.dialog, text="卸载命令:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.uninstall_entry = ttk.Entry(self.dialog, width=40)
        self.uninstall_entry.grid(row=3, column=1, padx=5, pady=5)
        
        # 按钮框架
        button_frame = ttk.Frame(self.dialog)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="确定", command=self.confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        self.result = None
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="选择应用程序",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if filename:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, filename)
            
    def confirm(self):
        name = self.name_entry.get().strip()
        publisher = self.publisher_entry.get().strip()
        path = self.path_entry.get().strip()
        uninstall = self.uninstall_entry.get().strip()
        
        if not name or not path:
            messagebox.showwarning("警告", "应用程序名称和路径不能为空！")
            return
            
        self.result = {
            'name': name,
            'publisher': publisher,
            'path': path,
            'uninstall': uninstall
        }
        self.dialog.destroy()

class AppManager:
    def __init__(self, root):
        self.root = root
        self.root.title("应用程序管理器")
        
        # 配置根窗口的网格权重
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置主框架的网格权重
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # 创建搜索框和添加按钮的框架
        top_frame = ttk.Frame(self.main_frame)
        top_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        top_frame.grid_columnconfigure(0, weight=1)
        
        # 创建搜索框
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(top_frame, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        self.search_var.trace('w', self.filter_apps)
        
        # 添加"添加应用"按钮
        ttk.Button(top_frame, text="添加应用", command=self.add_app).grid(row=0, column=1)
        
        # 创建应用列表
        self.tree = ttk.Treeview(self.main_frame, columns=('名称', '发布者', '路径', '卸载命令'), show='headings')
        self.tree.heading('名称', text='名称')
        self.tree.heading('发布者', text='发布者')
        self.tree.heading('路径', text='路径')
        self.tree.heading('卸载命令', text='卸载命令')
        self.tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 设置列宽
        self.tree.column('名称', width=200)
        self.tree.column('发布者', width=150)
        self.tree.column('路径', width=250)
        self.tree.column('卸载命令', width=200)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 创建按钮框架
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        # 添加按钮
        ttk.Button(button_frame, text="启动", command=self.launch_app).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="卸载", command=self.uninstall_app).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新", command=self.refresh_apps).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除", command=self.remove_app).pack(side=tk.LEFT, padx=5)
        
        # 初始化自定义应用列表
        self.custom_apps_file = "custom_apps.json"
        self.custom_apps = self.load_custom_apps()
        
        # 加载应用列表
        self.refresh_apps()
        
        # 更新窗口大小以适应内容
        self.root.update_idletasks()
        width = self.tree.winfo_reqwidth() + scrollbar.winfo_reqwidth() + 30
        height = self.tree.winfo_reqheight() + top_frame.winfo_reqheight() + button_frame.winfo_reqheight() + 50
        self.root.geometry(f"{width}x{height}")
    
    def load_custom_apps(self):
        try:
            if os.path.exists(self.custom_apps_file):
                with open(self.custom_apps_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载自定义应用列表失败: {str(e)}")
        return []
    
    def save_custom_apps(self):
        try:
            with open(self.custom_apps_file, 'w', encoding='utf-8') as f:
                json.dump(self.custom_apps, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存自定义应用列表失败: {str(e)}")
    
    def add_app(self):
        dialog = AddAppDialog(self.root)
        self.root.wait_window(dialog.dialog)
        if dialog.result:
            self.custom_apps.append(dialog.result)
            self.save_custom_apps()
            self.refresh_apps()
    
    def remove_app(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请先选择一个应用程序")
            return
            
        app_name = self.tree.item(selected_item[0])['values'][0]
        if messagebox.askyesno("确认", f"确定要从列表中删除 {app_name} 吗？"):
            # 只能删除自定义添加的应用
            self.custom_apps = [app for app in self.custom_apps if app['name'] != app_name]
            self.save_custom_apps()
            self.refresh_apps()
    
    def get_installed_apps(self):
        apps = []
        keys = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        
        for key_path in keys:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ)
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name)
                        try:
                            display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                            publisher = winreg.QueryValueEx(subkey, "Publisher")[0]
                            try:
                                install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                            except:
                                install_location = ""
                            uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                            apps.append({
                                'name': display_name,
                                'publisher': publisher,
                                'path': install_location,
                                'uninstall': uninstall_string
                            })
                        except WindowsError:
                            pass
                        finally:
                            winreg.CloseKey(subkey)
                    except WindowsError:
                        continue
                winreg.CloseKey(key)
            except WindowsError:
                continue
                
        # 添加自定义应用
        apps.extend(self.custom_apps)
        return apps
    
    def refresh_apps(self):
        # 清空现有列表
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 获取并显示应用列表
        apps = self.get_installed_apps()
        for app in apps:
            self.tree.insert('', tk.END, values=(
                app['name'],
                app['publisher'],
                app.get('path', ''),
                app['uninstall']
            ))
    
    def filter_apps(self, *args):
        search_term = self.search_var.get().lower()
        self.tree.delete(*self.tree.get_children())
        apps = self.get_installed_apps()
        for app in apps:
            if search_term in app['name'].lower():
                self.tree.insert('', tk.END, values=(
                    app['name'],
                    app['publisher'],
                    app.get('path', ''),
                    app['uninstall']
                ))
    
    def launch_app(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请先选择一个应用程序")
            return
            
        app_path = self.tree.item(selected_item[0])['values'][2]
        if not app_path:
            messagebox.showerror("错误", "无法找到应用程序路径")
            return
            
        try:
            if os.path.isfile(app_path):
                subprocess.Popen([app_path])
            else:
                subprocess.Popen(f'explorer.exe "{app_path}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法启动应用程序: {str(e)}")
    
    def uninstall_app(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请先选择一个应用程序")
            return
            
        app_name = self.tree.item(selected_item[0])['values'][0]
        uninstall_cmd = self.tree.item(selected_item[0])['values'][3]
        
        if messagebox.askyesno("确认", f"确定要卸载 {app_name} 吗？"):
            try:
                subprocess.Popen(uninstall_cmd)
            except Exception as e:
                messagebox.showerror("错误", f"卸载失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppManager(root)
    root.mainloop() 