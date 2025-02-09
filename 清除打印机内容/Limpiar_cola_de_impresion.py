import win32print
import tkinter as tk
from tkinter import ttk, messagebox
import pywintypes
import ctypes
import logging
from functools import partial
import winreg
from threading import Thread
import sys

# 初始化日志
logging.basicConfig(
    filename='printer_manager.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def require_admin():
    """检查并请求管理员权限"""
    try:
        if not ctypes.windll.shell32.IsUserAnAdmin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", "python.exe", __file__, None, 1)
            sys.exit()
    except Exception as e:
        logging.error(f"管理员权限请求失败: {e}")
        messagebox.showerror("错误", "需要管理员权限运行本程序")
        sys.exit(1)

def get_real_usb_ports():
    """从注册表获取真实USB端口"""
    ports = []
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DEVICEMAP\SERIALCOMM")
        for i in range(winreg.QueryInfoKey(key)[1]):
            _, port, _ = winreg.EnumValue(key, i)
            if "USB" in port.upper():
                ports.append(port)
        return sorted(ports, key=lambda x: int(x[3:]))
    except Exception as e:
        logging.error(f"获取USB端口失败: {e}")
        return []

def get_current_port(printer_name):
    """获取打印机当前端口"""
    try:
        handle = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(handle, 2)
        win32print.ClosePrinter(handle)
        return info['pPortName']
    except pywintypes.error as e:
        logging.error(f"获取端口失败: {e}")
        return ""

def clear_printer_queue(printer_name):
    """带异常处理的清空队列"""
    try:
        handle = win32print.OpenPrinter(printer_name)
        jobs = win32print.EnumJobs(handle, 0, -1, 1)
        
        for job in jobs:
            try:
                win32print.SetJob(handle, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
            except pywintypes.error as e:
                if e.winerror == 5:
                    raise Exception("权限不足，无法删除打印任务")
        
        win32print.ClosePrinter(handle)
        return len(jobs), ""
    except pywintypes.error as e:
        error_msg = f"操作失败: {e.strerror}"
        if e.winerror == 5:
            error_msg = "需要管理员权限"
        return 0, error_msg
    except Exception as e:
        return 0, str(e)

def update_printer_port(printer_name):
    """智能端口切换"""
    original_port = get_current_port(printer_name)
    usb_ports = get_real_usb_ports()
    
    if not usb_ports:
        return False, "未找到可用USB端口"
    
    try:
        handle = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(handle, 2)
        
        # 寻找下一个可用端口
        current_index = usb_ports.index(original_port) if original_port in usb_ports else -1
        new_port = usb_ports[(current_index + 1) % len(usb_ports)] if usb_ports else ""
        
        if new_port and new_port != original_port:
            info['pPortName'] = new_port
            win32print.SetPrinter(handle, 2, info, 0)
            win32print.ClosePrinter(handle)
            return True, f"端口已从 {original_port} 切换至 {new_port}"
        return False, "无可用新端口"
    except pywintypes.error as e:
        error_msg = f"端口切换失败: {e.strerror}"
        if e.winerror == 5:
            error_msg = "需要管理员权限"
        return False, error_msg
    except Exception as e:
        return False, str(e)
    
def get_printer_queue_count(printer_name):
    """安全获取打印队列数量"""
    try:
        handle = win32print.OpenPrinter(printer_name)
        jobs = win32print.EnumJobs(handle, 0, -1, 1)
        win32print.ClosePrinter(handle)
        return len(jobs)
    except pywintypes.error as e:
        logging.error(f"获取队列数量失败: {e}")
        return 0
    except Exception as e:
        logging.error(f"未知错误: {e}")
        return 0

def set_default_printer(printer_name):
    """设置默认打印机"""
    try:
        # 检查打印机是否存在
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        if not any(p[2] == printer_name for p in printers):
            return False, f"打印机 {printer_name} 不存在"
            
        # 设置默认打印机
        if win32print.SetDefaultPrinter(printer_name):
            return True, f"已成功将 {printer_name} 设为默认打印机"
        return False, "设置默认打印机失败"
    except pywintypes.error as e:
        error_msg = f"操作失败: {e.strerror}"
        if e.winerror == 5:
            error_msg = "需要管理员权限"
        return False, error_msg
    except Exception as e:
        return False, str(e)

class PrinterManagerGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("打印机管理工具 v2.0")
        self.window.geometry("500x350")
        self._center_window()
        
        # 样式配置
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#4CAF50")
        self.style.map("TButton", background=[("active", "#45a049")])
        
        self._create_widgets()
        self._update_status()
        
    def _center_window(self):
        """窗口居中"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self):
        """创建界面组件"""
        # 打印机选择
        ttk.Label(self.window, text="选择打印机:").pack(pady=10)
        self.printer_var = tk.StringVar()
        self._update_printer_list()
        
        # 状态显示
        self.status_frame = ttk.LabelFrame(self.window, text="当前状态")
        self.status_frame.pack(pady=10, fill="x", padx=20)
        
        self.status_labels = {
            'printer': ttk.Label(self.status_frame, text=""),
            'queue': ttk.Label(self.status_frame, text=""),
            'port': ttk.Label(self.status_frame, text="")
        }
        for label in self.status_labels.values():
            label.pack(anchor="w")
        
        # 操作按钮
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="刷新状态", command=self._update_status).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="清空队列", command=self._threaded_clear_queue).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="切换端口", command=self._threaded_update_port).grid(row=0, column=2, padx=5)
        ttk.Button(btn_frame, text="设为默认", command=self._threaded_set_default).grid(row=0, column=3, padx=5)
        
    def _update_printer_list(self):
        """更新打印机列表（自动排序IMP1）"""
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        
        # 优先排序逻辑
        sorted_printers = sorted(
            printers,
            key=lambda x: (
                not x[2].upper().startswith("IMP1"),  # IMP1排第一
                x[2] != win32print.GetDefaultPrinter(),  # 默认打印机排第二
                x[2]  # 其他按字母排序
            )
        )
        
        printer_names = [p[2] for p in sorted_printers]
        
        if hasattr(self, 'printer_menu'):
            self.printer_menu['menu'].delete(0, 'end')
            default_printer = win32print.GetDefaultPrinter()
            for name in printer_names:
                fg = "#45a049" if name == default_printer else "#333333"
                font = ('微软雅黑', 9, 'bold' if name == default_printer else 'normal')
                self.printer_menu['menu'].add_command(
                    label=name, 
                    command=partial(self.printer_var.set, name),
                    foreground=fg,
                    font=font
                )
        else:
            self.printer_menu = ttk.OptionMenu(
                self.window,
                self.printer_var,
                printer_names[0],
                *printer_names
            )
            self.printer_menu.pack()
            self.printer_var.trace_add('write', lambda *_: self._update_status())
            
    def _update_status(self, event=None):
        """更新状态信息"""
        printer_name = self.printer_var.get()
        if not printer_name:
            return
        
        queue_count = get_printer_queue_count(printer_name)
        current_port = get_current_port(printer_name)
        
        status_text = {
            'printer': f"当前打印机: {printer_name}",
            'queue': f"待处理任务: {queue_count} 个",
            'port': f"当前端口: {current_port or '未知'}"
        }
        
        for key, label in self.status_labels.items():
            label.config(text=status_text[key])
            label.config(foreground="#333" if key == 'printer' else ("#d32f2f" if queue_count >0 else "#388e3c"))

    def _threaded_operation(self, func, success_msg):
        """通用线程操作"""
        def wrapper():
            try:
                success, result = func()
                if success:
                    messagebox.showinfo("操作成功", f"{success_msg}\n{result}")
                else:
                    messagebox.showerror("操作失败", result)
                self._update_status()
                self._update_printer_list()  # 刷新打印机列表
            except Exception as e:
                messagebox.showerror("错误", str(e))
            finally:
                self.window.config(cursor="")
                self.window.attributes('-disabled', False)
        
        self.window.config(cursor="watch")
        self.window.attributes('-disabled', True)
        Thread(target=wrapper).start()

    def _threaded_clear_queue(self):
        """线程化清空队列"""
        printer_name = self.printer_var.get()
        if not messagebox.askyesno("确认", "确定要清空打印队列吗？"):
            return
        
        def clear_task():
            count, error = clear_printer_queue(printer_name)
            if error:
                return False, error
            return True, f"已清除 {count} 个打印任务"
        
        self._threaded_operation(clear_task, "队列已清空")

    def _threaded_update_port(self):
        """线程化切换端口"""
        def port_task():
            success, msg = update_printer_port(self.printer_var.get())
            return success, msg
        
        self._threaded_operation(port_task, "端口切换成功")
        
    def _threaded_set_default(self):
        """线程化设置默认打印机"""
        def set_default_task():
            return set_default_printer(self.printer_var.get())
        
        self._threaded_operation(set_default_task, "默认打印机设置成功")

def main():
    require_admin()
    try:
        # 启动时自动尝试设置IMP1
        default_printer = win32print.GetDefaultPrinter()
        if "IMP1" not in default_printer.upper():
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            imp1_printers = [p[2] for p in printers if "IMP1" in p[2].upper()]
            if imp1_printers:
                success, msg = set_default_printer(imp1_printers[0])
                if success:
                    logging.info("自动设置IMP1为默认打印机成功")
                else:
                    logging.warning(f"自动设置IMP1失败: {msg}")
        
        app = PrinterManagerGUI()
        app.window.mainloop()
    except Exception as e:
        logging.exception("程序异常退出")
        messagebox.showerror("严重错误", f"程序发生未预期错误:\n{str(e)}")

if __name__ == "__main__":
    main()