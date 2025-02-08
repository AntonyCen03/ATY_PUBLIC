import win32print
import tkinter as tk
from tkinter import messagebox

def clear_printer_queue(printer_name):
    # 如果没有提供打印机名称，则获取默认打印机
    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()

    print(f"正在清除打印机 {printer_name} 的打印队列")

    # 打开打印机
    printer_handle = win32print.OpenPrinter(printer_name)

    # 获取所有打印任务
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)

    # 删除每个打印任务
    job_count = 0
    for job in jobs:
        print(f"正在删除打印机 {printer_name} 的任务 {job['JobId']}")
        win32print.SetJob(printer_handle, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
        job_count += 1

    # 关闭打印机句柄
    win32print.ClosePrinter(printer_handle)

    print("打印队列已清除。")
    return job_count

def select_printer():
    # 创建窗口
    window = tk.Tk()
    window.title("选择打印机")
    window.geometry("400x200")  # 设置窗口大小为400x200

    # 获取屏幕宽度和高度
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # 计算窗口居中的位置
    x = (screen_width / 2) - (400 / 2)
    y = (screen_height / 2) - (200 / 2)

    # 设置窗口位置
    window.geometry(f"400x200+{int(x)}+{int(y)}")

    # 获取所有打印机
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)

    # 创建下拉菜单
    printer_var = tk.StringVar(window)
    preferred_printer = "IMP1"  # 设置首选打印机名称
    fetch_printer = "XP-80C"  # 设置获取打印机名称
    printer_names = [printer[2] for printer in printers]
    if preferred_printer in printer_names:
        printer_var.set(preferred_printer)  # 设置默认选项为首选打印机
    elif fetch_printer in printer_names:
        printer_var.set(fetch_printer)
    else:
        printer_var.set(printer_names[0])  # 如果首选打印机不存在，则设置为第一个打印机

    printer_menu = tk.OptionMenu(window, printer_var, *printer_names)

    # 创建按钮
    def on_clear_queue():
        job_count = clear_printer_queue(printer_var.get())
        messagebox.showinfo("清除完成", f"已清除 {job_count} 个打印任务")
        window.quit()  # 关闭窗口

    button = tk.Button(window, text="清除打印队列", command=on_clear_queue)

    # 布局窗口
    printer_menu.pack(pady=20)
    button.pack(pady=20)

    # 使用place方法将控件居中
    printer_menu.place(relx=0.5, rely=0.3, anchor=tk.CENTER)
    button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # 运行窗口
    window.mainloop()

if __name__ == "__main__":
    select_printer()