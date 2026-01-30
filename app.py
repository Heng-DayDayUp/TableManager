import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import subprocess
import winreg
import threading
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import pystray
from PIL import Image, ImageDraw
import sys
import functools

class AppWindowManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows窗口管理器")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)
        
        # 窗口居中
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # 设置主题
        style = ttkb.Style(theme="cosmo")
        
        # 数据存储
        self.data_file = "app_data.json"
        self.applications = []
        self.combinations = []
        
        # 加载数据
        self.load_data()
        
        # 初始化UI
        self.create_ui()
        
        # 自动检测软件
        self.detect_applications()
        
        # 创建系统托盘
        self.create_system_tray()
    
    def create_ui(self):
        # 创建主框架
        main_frame = ttkb.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题区域
        title_frame = ttkb.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 主标题
        title_label = ttkb.Label(
            title_frame, 
            text="Windows窗口管理器", 
            font=("微软雅黑", 20, "bold"),
            bootstyle="primary"
        )
        title_label.pack(side=tk.LEFT)
        
        # 版本信息
        version_label = ttkb.Label(
            title_frame, 
            text="v1.0.0", 
            font=("微软雅黑", 12),
            bootstyle="secondary"
        )
        version_label.pack(side=tk.RIGHT, padx=10)
        
        # 创建标签页
        tab_control = ttkb.Notebook(main_frame, bootstyle="primary")
        
        # 应用管理标签页
        app_tab = ttkb.Frame(tab_control)
        tab_control.add(app_tab, text="应用管理")
        
        # 组合管理标签页
        combo_tab = ttkb.Frame(tab_control)
        tab_control.add(combo_tab, text="组合管理")
        
        tab_control.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 应用管理界面
        self.create_app_tab(app_tab)
        
        # 组合管理界面
        self.create_combo_tab(combo_tab)
    
    def create_app_tab(self, parent):
        # 应用列表框架
        app_frame = ttkb.LabelFrame(
            parent, 
            text="已检测应用", 
            padding="15",
            bootstyle="primary"
        )
        app_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 工具栏
        toolbar_frame = ttkb.Frame(app_frame)
        toolbar_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 应用数量
        self.app_count_var = tk.StringVar(value="应用数量: 0")
        app_count_label = ttkb.Label(
            toolbar_frame, 
            textvariable=self.app_count_var,
            font=("微软雅黑", 11),
            bootstyle="info"
        )
        app_count_label.pack(side=tk.LEFT)
        
        # 搜索框
        search_var = tk.StringVar()
        search_frame = ttkb.Frame(toolbar_frame)
        search_frame.pack(side=tk.RIGHT)
        
        ttkb.Label(search_frame, text="搜索:", font=("微软雅黑", 11)).pack(side=tk.LEFT, padx=5)
        search_entry = ttkb.Entry(
            search_frame, 
            textvariable=search_var, 
            width=30,
            bootstyle="primary"
        )
        search_entry.pack(side=tk.LEFT, padx=5)
        
        search_btn = ttkb.Button(
            search_frame, 
            text="搜索", 
            command=lambda: self.search_apps(search_var.get()),
            bootstyle="primary-outline"
        )
        search_btn.pack(side=tk.LEFT, padx=5)
        
        reset_btn = ttkb.Button(
            search_frame, 
            text="重置", 
            command=lambda: self.reset_search(search_var),
            bootstyle="secondary-outline"
        )
        reset_btn.pack(side=tk.LEFT, padx=5)
        
        # 应用列表树
        columns = ("name", "path")
        self.app_tree = ttkb.Treeview(
            app_frame, 
            columns=columns, 
            show="headings",
            bootstyle="primary"
        )
        
        # 配置列
        self.app_tree.heading("name", text="应用名称")
        self.app_tree.heading("path", text="应用路径")
        self.app_tree.column("name", width=200, anchor=tk.W)
        self.app_tree.column("path", width=450, anchor=tk.W)
        
        # 配置样式
        style = ttkb.Style()
        style.configure("Treeview", font=("微软雅黑", 11), rowheight=40)
        style.configure("Treeview.Heading", font=("微软雅黑", 12, "bold"))
        style.configure("TNotebook.Tab", font=("微软雅黑", 11))
        style.configure("TButton", font=("微软雅黑", 11))
        style.configure("TLabel", font=("微软雅黑", 11))
        
        # 滚动条
        scrollbar = ttkb.Scrollbar(
            app_frame, 
            orient=tk.VERTICAL, 
            command=self.app_tree.yview,
            bootstyle="primary"
        )
        self.app_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 水平滚动条
        hscrollbar = ttkb.Scrollbar(
            app_frame, 
            orient=tk.HORIZONTAL, 
            command=self.app_tree.xview,
            bootstyle="primary"
        )
        self.app_tree.configure(xscroll=hscrollbar.set)
        hscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.app_tree.pack(fill=tk.BOTH, expand=True)
        
        # 操作按钮框架
        btn_frame = ttkb.Frame(parent)
        btn_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 刷新按钮
        refresh_btn = ttkb.Button(
            btn_frame, 
            text="刷新应用列表", 
            command=self.detect_applications,
            bootstyle="success"
        )
        refresh_btn.pack(side=tk.LEFT, padx=10)
        
        # 状态信息
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttkb.Label(
            btn_frame, 
            textvariable=self.status_var, 
            font=("微软雅黑", 11, "italic"),
            bootstyle="info"
        )
        status_label.pack(side=tk.RIGHT, padx=10)
    
    def create_combo_tab(self, parent):
        # 组合列表
        combo_frame = ttkb.LabelFrame(parent, text="应用组合", padding="10")
        combo_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 组合列表工具栏
        toolbar_frame = ttkb.Frame(combo_frame)
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # 组合数量显示
        self.combo_count_var = tk.StringVar(value="组合数量: 0")
        combo_count_label = ttkb.Label(
            toolbar_frame, 
            textvariable=self.combo_count_var,
            font=("微软雅黑", 11),
            bootstyle="info"
        )
        combo_count_label.pack(side=tk.LEFT)
        
        # 组合列表树
        columns = ("name", "apps", "count")
        self.combo_tree = ttkb.Treeview(
            combo_frame, 
            columns=columns, 
            show="headings",
            bootstyle="primary"
        )
        self.combo_tree.heading("name", text="组合名称")
        self.combo_tree.heading("apps", text="包含应用")
        self.combo_tree.heading("count", text="应用数量")
        self.combo_tree.column("name", width=250, anchor=tk.W)
        self.combo_tree.column("apps", width=500, anchor=tk.W)
        self.combo_tree.column("count", width=100, anchor=tk.CENTER)
        
        # 滚动条
        scrollbar = ttkb.Scrollbar(
            combo_frame, 
            orient=tk.VERTICAL, 
            command=self.combo_tree.yview,
            bootstyle="primary"
        )
        self.combo_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 水平滚动条
        hscrollbar = ttkb.Scrollbar(
            combo_frame, 
            orient=tk.HORIZONTAL, 
            command=self.combo_tree.xview,
            bootstyle="primary"
        )
        self.combo_tree.configure(xscroll=hscrollbar.set)
        hscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.combo_tree.pack(fill=tk.BOTH, expand=True)
        
        # 操作按钮框架
        btn_frame = ttkb.Frame(parent)
        btn_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 按钮容器
        button_container = ttkb.Frame(btn_frame)
        button_container.pack(side=tk.LEFT)
        
        # 创建组合按钮
        add_combo_btn = ttkb.Button(
            button_container, 
            text="创建组合", 
            command=self.create_combination,
            bootstyle="success"
        )
        add_combo_btn.pack(side=tk.LEFT, padx=10)
        
        # 编辑组合按钮
        edit_combo_btn = ttkb.Button(
            button_container, 
            text="编辑组合", 
            command=self.edit_combination,
            bootstyle="warning"
        )
        edit_combo_btn.pack(side=tk.LEFT, padx=10)
        
        # 删除组合按钮
        delete_combo_btn = ttkb.Button(
            button_container, 
            text="删除组合", 
            command=self.delete_combination,
            bootstyle="danger"
        )
        delete_combo_btn.pack(side=tk.LEFT, padx=10)
        
        # 运行组合按钮
        run_combo_btn = ttkb.Button(
            button_container, 
            text="运行组合", 
            command=self.run_combination,
            bootstyle="primary"
        )
        run_combo_btn.pack(side=tk.LEFT, padx=10)
        
        # 状态信息
        self.combo_status_var = tk.StringVar(value="就绪")
        combo_status_label = ttkb.Label(
            btn_frame, 
            textvariable=self.combo_status_var, 
            font=("微软雅黑", 11, "italic"),
            bootstyle="info"
        )
        combo_status_label.pack(side=tk.RIGHT, padx=10)
    
    def detect_applications(self):
        """检测系统中的应用"""
        def detect():
            # 检查是否已经有应用数据，避免重复检测
            if not self.applications:
                self.applications = []
                
                # 更新状态
                if hasattr(self, 'status_var'):
                    self.status_var.set("正在检测应用...")
                
                # 检测开始菜单中的应用
                self._detect_start_menu_apps()
                
                # 检测注册表中的应用
                self._detect_registry_apps()
                
                # 更新应用列表
                self.root.after(0, self.update_app_tree)
                
                # 保存数据
                self.save_data()
                
                # 更新状态
                if hasattr(self, 'status_var'):
                    self.status_var.set("应用检测完成")
            else:
                # 更新状态
                if hasattr(self, 'status_var'):
                    self.status_var.set("应用列表已存在")
                # 更新应用列表
                self.root.after(0, self.update_app_tree)
        
        # 在线程中执行检测，避免阻塞UI
        thread = threading.Thread(target=detect)
        thread.daemon = True
        thread.start()
    
    def _detect_start_menu_apps(self):
        """检测开始菜单中的应用"""
        # 开始菜单路径
        start_menu_paths = [
            os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs"),
            os.path.join(os.environ.get("PROGRAMDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs")
        ]
        
        for path in start_menu_paths:
            if os.path.exists(path):
                self._scan_directory_for_apps(path)
    
    def _detect_registry_apps(self):
        """检测注册表中的应用"""
        # 检测注册表中的应用路径
        try:
            # 64位系统
            reg_paths = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            ]
            
            # 预定义过滤关键词
            uninstall_keywords = ["uninstall", "卸载", "remove", "删除", "cleanup", "清理"]
            uninstall_path_keywords = ["uninstall", "unins000", "setup", "install", "remove", "delete"]
            
            for reg_path in reg_paths:
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                    subkey_count = winreg.QueryInfoKey(key)[0]
                    
                    for i in range(subkey_count):
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            subkey = None
                            
                            try:
                                subkey = winreg.OpenKey(key, subkey_name)
                                
                                # 尝试获取应用名称
                                try:
                                    app_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                    
                                    # 过滤掉uninstall相关程序
                                    if any(keyword in app_name.lower() for keyword in uninstall_keywords):
                                        continue
                                    
                                    # 尝试获取应用路径
                                    try:
                                        app_path = winreg.QueryValueEx(subkey, "DisplayIcon")[0]
                                        # 处理路径，移除图标参数
                                        if app_path.endswith(".exe,0"):
                                            app_path = app_path[:-3]
                                        elif "," in app_path:
                                            app_path = app_path.split(",")[0]
                                        
                                        # 过滤掉uninstall路径
                                        if any(keyword in app_path.lower() for keyword in uninstall_path_keywords):
                                            continue
                                        
                                        # 验证路径存在
                                        if os.path.exists(app_path):
                                            self._add_application(app_name, app_path)
                                    except:
                                        pass
                                except:
                                    pass
                            finally:
                                if subkey:
                                    try:
                                        subkey.Close()
                                    except:
                                        pass
                        except:
                            pass
                except:
                    pass
        except Exception as e:
            print(f"注册表检测错误: {e}")
    
    def _scan_directory_for_apps(self, directory):
        """扫描目录中的应用"""
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith(".lnk"):
                    # 快捷方式文件
                    lnk_path = os.path.join(root, file)
                    app_name = os.path.splitext(file)[0]
                    app_path = self._get_shortcut_target(lnk_path)
                    if app_path and os.path.exists(app_path):
                        self._add_application(app_name, app_path)
                elif file.endswith(".exe"):
                    # 可执行文件
                    exe_path = os.path.join(root, file)
                    app_name = os.path.splitext(file)[0]
                    self._add_application(app_name, exe_path)
    
    def _get_shortcut_target(self, lnk_path):
        """获取快捷方式的目标路径"""
        try:
            import winshell
            shortcut = winshell.shortcut(lnk_path)
            return shortcut.path
        except:
            # 如果没有winshell库，尝试其他方法
            try:
                import pythoncom
                from win32com.client import Dispatch
                
                shell = Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(lnk_path)
                return shortcut.Targetpath
            except:
                return None
    
    def _add_application(self, name, path):
        """添加应用到列表"""
        # 去重
        for app in self.applications:
            if app["path"] == path:
                return
        
        self.applications.append({"name": name, "path": path})
    
    def search_apps(self, search_text):
        """搜索应用"""
        search_text = search_text.lower()
        # 清空现有数据
        for item in self.app_tree.get_children():
            self.app_tree.delete(item)
        
        # 添加匹配的应用
        for app in self.applications:
            if search_text in app["name"].lower() or search_text in app["path"].lower():
                self.app_tree.insert("", tk.END, values=(app["name"], app["path"]))
    
    def reset_search(self, search_var):
        """重置搜索"""
        search_var.set("")
        self.update_app_tree()
    
    def update_app_tree(self):
        """更新应用列表树"""
        # 清空现有数据
        for item in self.app_tree.get_children():
            self.app_tree.delete(item)
        
        # 添加应用数据
        for app in self.applications:
            self.app_tree.insert("", tk.END, values=(app["name"], app["path"]))
        
        # 更新应用数量
        if hasattr(self, 'app_count_var'):
            self.app_count_var.set(f"应用数量: {len(self.applications)}")
        
        # 更新状态
        if hasattr(self, 'status_var'):
            self.status_var.set("应用列表已更新")
    
    def update_combo_tree(self):
        """更新组合列表树"""
        # 清空现有数据
        for item in self.combo_tree.get_children():
            self.combo_tree.delete(item)
        
        # 添加组合数据
        for combo in self.combinations:
            app_names = ", ".join([app["name"] for app in combo["apps"]])
            app_count = len(combo["apps"])
            self.combo_tree.insert("", tk.END, values=(combo["name"], app_names, app_count))
        
        # 更新组合数量
        if hasattr(self, 'combo_count_var'):
            self.combo_count_var.set(f"组合数量: {len(self.combinations)}")
        
        # 更新状态
        if hasattr(self, 'combo_status_var'):
            self.combo_status_var.set("组合列表已更新")
    
    def create_combination(self):
        """创建新组合"""
        # 创建对话框
        dialog = ttkb.Toplevel(self.root)
        dialog.title("创建组合")
        dialog.geometry("900x1000")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # 窗口居中
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # 设置对话框样式
        style = ttkb.Style(theme="cosmo")
        
        # 组合名称
        name_frame = ttkb.Frame(dialog, padding="20")
        name_frame.pack(fill=tk.X)
        
        ttkb.Label(
            name_frame, 
            text="组合名称:", 
            font=("微软雅黑", 11)
        ).pack(side=tk.LEFT, padx=10)
        
        name_var = tk.StringVar()
        name_entry = ttkb.Entry(
            name_frame, 
            textvariable=name_var, 
            width=40,
            bootstyle="primary"
        )
        name_entry.pack(side=tk.LEFT, padx=10)
        
        # 应用列表
        app_frame = ttkb.LabelFrame(
            dialog, 
            text="选择应用", 
            padding="20",
            bootstyle="primary"
        )
        app_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 应用列表框
        listbox = tk.Listbox(
            app_frame, 
            selectmode=tk.MULTIPLE, 
            width=70, 
            height=18,
            font=("微软雅黑", 10)
        )
        
        # 添加应用到列表框
        for app in self.applications:
            listbox.insert(tk.END, f"{app['name']} - {app['path']}")
        
        # 滚动条
        scrollbar = ttkb.Scrollbar(
            app_frame, 
            orient=tk.VERTICAL, 
            command=listbox.yview,
            bootstyle="primary"
        )
        listbox.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(fill=tk.BOTH, expand=True)
        
        # 按钮框架
        btn_frame = ttkb.Frame(dialog, padding="20")
        btn_frame.pack(fill=tk.X)
        
        def save():
            combo_name = name_var.get().strip()
            if not combo_name:
                messagebox.showerror("错误", "请输入组合名称")
                return
            
            # 获取选中的应用
            selected_indices = listbox.curselection()
            if not selected_indices:
                messagebox.showerror("错误", "请至少选择一个应用")
                return
            
            selected_apps = [self.applications[i] for i in selected_indices]
            
            # 添加组合
            self.combinations.append({"name": combo_name, "apps": selected_apps})
            
            # 更新组合列表
            self.update_combo_tree()
            
            # 保存数据
            self.save_data()
            
            # 更新状态栏图标右键菜单
            self.update_tray_menu()
            
            dialog.destroy()
            self.combo_status_var.set(f"组合 '{combo_name}' 创建成功")
        
        # 保存按钮
        ttkb.Button(
            btn_frame, 
            text="保存", 
            command=save,
            bootstyle="success",
            width=15
        ).pack(side=tk.RIGHT, padx=10)
        
        # 取消按钮
        ttkb.Button(
            btn_frame, 
            text="取消", 
            command=dialog.destroy,
            bootstyle="secondary",
            width=15
        ).pack(side=tk.RIGHT, padx=10)
    
    def edit_combination(self):
        """编辑组合"""
        selected_item = self.combo_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要编辑的组合")
            return
        
        # 获取选中的组合
        item = selected_item[0]
        values = self.combo_tree.item(item, "values")
        combo_name = values[0]
        
        # 查找组合
        combo = None
        for c in self.combinations:
            if c["name"] == combo_name:
                combo = c
                break
        
        if not combo:
            return
        
        # 创建对话框
        dialog = ttkb.Toplevel(self.root)
        dialog.title("编辑组合")
        dialog.geometry("900x1000")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # 窗口居中
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # 设置对话框样式
        style = ttkb.Style(theme="cosmo")
        
        # 组合名称
        name_frame = ttkb.Frame(dialog, padding="20")
        name_frame.pack(fill=tk.X)
        
        ttkb.Label(
            name_frame, 
            text="组合名称:", 
            font=("微软雅黑", 11)
        ).pack(side=tk.LEFT, padx=10)
        
        name_var = tk.StringVar(value=combo_name)
        name_entry = ttkb.Entry(
            name_frame, 
            textvariable=name_var, 
            width=40,
            bootstyle="primary"
        )
        name_entry.pack(side=tk.LEFT, padx=10)
        
        # 应用列表
        app_frame = ttkb.LabelFrame(
            dialog, 
            text="选择应用", 
            padding="20",
            bootstyle="primary"
        )
        app_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 应用列表框
        listbox = tk.Listbox(
            app_frame, 
            selectmode=tk.MULTIPLE, 
            width=70, 
            height=18,
            font=("微软雅黑", 10)
        )
        
        # 标记已选应用
        selected_app_paths = [app["path"] for app in combo["apps"]]
        for i, app in enumerate(self.applications):
            listbox.insert(tk.END, f"{app['name']} - {app['path']}")
            if app["path"] in selected_app_paths:
                listbox.selection_set(i)
        
        # 滚动条
        scrollbar = ttkb.Scrollbar(
            app_frame, 
            orient=tk.VERTICAL, 
            command=listbox.yview,
            bootstyle="primary"
        )
        listbox.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(fill=tk.BOTH, expand=True)
        
        # 按钮框架
        btn_frame = ttkb.Frame(dialog, padding="20")
        btn_frame.pack(fill=tk.X)
        
        def save():
            new_name = name_var.get().strip()
            if not new_name:
                messagebox.showerror("错误", "请输入组合名称")
                return
            
            # 获取选中的应用
            selected_indices = listbox.curselection()
            if not selected_indices:
                messagebox.showerror("错误", "请至少选择一个应用")
                return
            
            selected_apps = [self.applications[i] for i in selected_indices]
            
            # 更新组合
            combo["name"] = new_name
            combo["apps"] = selected_apps
            
            # 更新组合列表
            self.update_combo_tree()
            
            # 保存数据
            self.save_data()
            
            # 更新状态栏图标右键菜单
            self.update_tray_menu()
            
            dialog.destroy()
            self.combo_status_var.set(f"组合 '{new_name}' 更新成功")
        
        # 保存按钮
        ttkb.Button(
            btn_frame, 
            text="保存", 
            command=save,
            bootstyle="success",
            width=15
        ).pack(side=tk.RIGHT, padx=10)
        
        # 取消按钮
        ttkb.Button(
            btn_frame, 
            text="取消", 
            command=dialog.destroy,
            bootstyle="secondary",
            width=15
        ).pack(side=tk.RIGHT, padx=10)
    
    def delete_combination(self):
        """删除组合"""
        selected_item = self.combo_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要删除的组合")
            return
        
        # 获取选中的组合
        item = selected_item[0]
        values = self.combo_tree.item(item, "values")
        combo_name = values[0]
        
        if messagebox.askyesno("确认", f"确定要删除组合 '{combo_name}' 吗？"):
            # 删除组合
            self.combinations = [c for c in self.combinations if c["name"] != combo_name]
            
            # 更新组合列表
            self.update_combo_tree()
            
            # 保存数据
            self.save_data()
            
            # 更新状态栏图标右键菜单
            self.update_tray_menu()
            
            self.combo_status_var.set(f"组合 '{combo_name}' 删除成功")
    
    def run_combination(self):
        """运行组合"""
        selected_item = self.combo_tree.selection()
        if not selected_item:
            messagebox.showerror("错误", "请选择要运行的组合")
            return
        
        # 获取选中的组合
        item = selected_item[0]
        values = self.combo_tree.item(item, "values")
        combo_name = values[0]
        
        # 查找组合
        combo = None
        for c in self.combinations:
            if c["name"] == combo_name:
                combo = c
                break
        
        if not combo:
            return
        
        # 运行组合中的应用
        def run_apps():
            # 更新状态
            if hasattr(self, 'combo_status_var'):
                self.combo_status_var.set(f"组合 '{combo_name}' 启动中...")
            
            for app in combo["apps"]:
                try:
                    subprocess.Popen(app["path"])
                except Exception as e:
                    print(f"启动应用失败 {app['name']}: {e}")
            
            # 更新状态
            if hasattr(self, 'combo_status_var'):
                self.combo_status_var.set(f"组合 '{combo_name}' 启动完成")
        
        # 在线程中运行，避免阻塞UI
        thread = threading.Thread(target=run_apps)
        thread.daemon = True
        thread.start()
    
    def save_data(self):
        """保存数据到文件"""
        data = {
            "applications": self.applications,
            "combinations": self.combinations
        }
        
        try:
            with open(self.data_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存数据错误: {e}")
    
    def load_data(self):
        """从文件加载数据"""
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.applications = data.get("applications", [])
                    self.combinations = data.get("combinations", [])
        except Exception as e:
            print(f"加载数据错误: {e}")
    
    def create_system_tray(self):
        """创建系统托盘"""
        # 创建图标
        def create_image():
            # 使用更小的图标尺寸，减少内存占用
            width = 32
            height = 32
            image = Image.new('RGB', (width, height), (30, 144, 255))
            draw = ImageDraw.Draw(image)
            draw.rectangle([5, 5, 27, 27], fill=(65, 105, 225), outline=(255, 255, 255), width=1)
            draw.text((10, 12), "WM", fill=(255, 255, 255), font_size=10)
            return image
        
        # 创建菜单
        def create_menu():
            # 创建菜单项列表
            menu_items = [
                pystray.MenuItem("打开窗口", self.show_window)
            ]
            
            # 添加组合到主菜单
            if self.combinations:
                menu_items.append(pystray.Menu.SEPARATOR)
                for combo in self.combinations:
                    # 使用嵌套函数来正确绑定参数
                    def create_callback(combo):
                        def callback(icon, item):
                            self.quick_run_combination(combo)
                        return callback
                    
                    menu_items.append(pystray.MenuItem(
                        combo["name"],
                        create_callback(combo)
                    ))
            
            # 添加退出选项
            menu_items.extend([
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出", self.exit_app)
            ])
            
            menu = pystray.Menu(*menu_items)
            return menu
        
        # 创建系统托盘
        self.tray = pystray.Icon(
            "Windows窗口管理器",
            create_image(),
            "Windows窗口管理器",
            create_menu()
        )
        
        # 在后台运行托盘
        def run_tray():
            try:
                self.tray.run()
            except Exception as e:
                print(f"系统托盘运行错误: {e}")
        
        tray_thread = threading.Thread(target=run_tray)
        tray_thread.daemon = True
        tray_thread.start()
    
    def update_tray_menu(self):
        """更新状态栏图标的右键菜单"""
        if hasattr(self, 'tray') and self.tray:
            try:
                # 创建新的菜单
                def create_menu():
                    # 创建菜单项列表
                    menu_items = [
                        pystray.MenuItem("打开窗口", self.show_window)
                    ]
                    
                    # 添加组合到主菜单
                    if self.combinations:
                        menu_items.append(pystray.Menu.SEPARATOR)
                        for combo in self.combinations:
                            # 使用嵌套函数来正确绑定参数
                            def create_callback(combo):
                                def callback(icon, item):
                                    self.quick_run_combination(combo)
                                return callback
                            
                            menu_items.append(pystray.MenuItem(
                                combo["name"],
                                create_callback(combo)
                            ))
                    
                    # 添加退出选项
                    menu_items.extend([
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem("退出", self.exit_app)
                    ])
                    
                    menu = pystray.Menu(*menu_items)
                    return menu
                
                # 更新菜单
                self.tray.menu = create_menu()
                # 刷新图标以应用新菜单
                if hasattr(self.tray, 'update_menu'):
                    self.tray.update_menu()
            except Exception as e:
                print(f"更新托盘菜单错误: {e}")
    
    def show_window(self):
        """显示主窗口"""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
    
    def quick_run_combination(self, combo):
        """快速运行组合"""
        combo_name = combo["name"]
        
        # 运行组合中的应用
        def run_apps():
            for app in combo["apps"]:
                try:
                    subprocess.Popen(app["path"])
                except Exception as e:
                    print(f"启动应用失败 {app['name']}: {e}")
        
        # 在线程中运行，避免阻塞UI
        thread = threading.Thread(target=run_apps)
        thread.daemon = True
        thread.start()
    
    def exit_app(self):
        """退出应用"""
        self.tray.stop()
        self.root.destroy()
        sys.exit()

if __name__ == "__main__":
    # 使用ttkbootstrap创建根窗口
    root = ttkb.Window(themename="cosmo")
    
    # 修改窗口关闭行为
    def on_closing():
        root.withdraw()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    app = AppWindowManager(root)
    app.update_combo_tree()
    root.mainloop()