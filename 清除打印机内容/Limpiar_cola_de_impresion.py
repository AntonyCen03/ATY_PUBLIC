import win32print
import tkinter as tk
from tkinter import messagebox

def clear_printer_queue(printer_name):
    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()

    print(f"正在清除打印机 {printer_name} 的打印队列")

    printer_handle = win32print.OpenPrinter(printer_name)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)

    job_count = 0
    for job in jobs:
        print(f"正在删除打印机 {printer_name} 的任务 {job['JobId']}")
        win32print.SetJob(printer_handle, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
        job_count += 1

    win32print.ClosePrinter(printer_handle)
    print("打印队列已清除。")
    return job_count

def get_printer_queue_count(printer_name):
    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()

    printer_handle = win32print.OpenPrinter(printer_name)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
    win32print.ClosePrinter(printer_handle)

    return len(jobs)

def select_printer():
    window = tk.Tk()
    window.title("选择打印机")
    window.geometry("400x200")

    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width / 2) - (400 / 2)
    y = (screen_height / 2) - (200 / 2)
    window.geometry(f"400x200+{int(x)}+{int(y)}")

    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    printer_var = tk.StringVar(window)
    preferred_printer = "IMP1"
    fetch_printer = "XP-80C"
    printer_names = [printer[2] for printer in printers]
    if preferred_printer in printer_names:
        printer_var.set(preferred_printer)
    elif fetch_printer in printer_names:
        printer_var.set(fetch_printer)
    else:
        printer_var.set(printer_names[0])

    printer_menu = tk.OptionMenu(window, printer_var, *printer_names)

    def on_clear_queue():
        printer_name = printer_var.get()
        job_count = get_printer_queue_count(printer_name)
        if messagebox.askyesno("确认清除", f"打印队列中有 {job_count} 个文档。是否清除？"):
            job_count = clear_printer_queue(printer_name)
            messagebox.showinfo("清除完成", f"已清除 {job_count} 个打印任务")
        #window.quit()

    button = tk.Button(window, text="清除打印文档", command=on_clear_queue)
    printer_menu.pack(pady=20)
    button.pack(pady=20)
    printer_menu.place(relx=0.5, rely=0.3, anchor=tk.CENTER)
    button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)
    window.mainloop()

if __name__ == "__main__":
    select_printer()