import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import pyperclip


# 获取桌面快捷方式所指向的可执行文件
def get_desktop_shortcuts():
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    software_list = []
    shell = win32com.client.Dispatch("WScript.Shell")
    
    for item in os.listdir(desktop_path):
        if item.endswith(".lnk"):
            shortcut = shell.CreateShortcut(os.path.join(desktop_path, item))
            target_path = shortcut.Targetpath
            if target_path.endswith(".exe"):
                software_list.append((os.path.splitext(item)[0], target_path))
    
    return software_list

# 获取已安装的软件列表
def get_installed_software():
    software_list = []
    try:
        import winreg
        reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
        for i in range(winreg.QueryInfoKey(reg_key)[0]):
            sub_key_name = winreg.EnumKey(reg_key, i)
            sub_key = winreg.OpenKey(reg_key, sub_key_name)
            try:
                display_name = winreg.QueryValueEx(sub_key, "DisplayName")[0]
                install_location = winreg.QueryValueEx(sub_key, "InstallLocation")[0]
                if display_name and install_location:
                    software_list.append((display_name, install_location))
            except FileNotFoundError:
                continue
    except ImportError:
        software_list = [("Example Software 1", "C:\\Program Files\\Example Software 1"),
                         ("Example Software 2", "C:\\Program Files\\Example Software 2")]
    return software_list

# 搜索可执行文件
def find_executable(display_name, install_location):
    exe_files = []
    for root, dirs, files in os.walk(install_location):
        for file in files:
            if file.endswith(".exe"):
                exe_files.append(os.path.join(root, file))

    # 优先考虑与软件名相近的可执行文件
    for exe in exe_files:
        if display_name.lower() in os.path.basename(exe).lower():
            return exe

    # 如果未找到相近的可执行文件，返回所有可执行文件，让用户选择
    return exe_files

# 生成批处理文件
def generate_batch_file(selected_software):
    batch_file_path = filedialog.asksaveasfilename(defaultextension=".bat", filetypes=[("Batch files", "*.bat")])
    if not batch_file_path:
        return
    with open(batch_file_path, 'w') as batch_file:
        batch_file.write("@echo off\n")
        batch_file.write("echo Running as administrator...\n")
        for name, path in selected_software:
            if os.path.isfile(path) and path.endswith(".exe"):
                exe_path = path
            else:
                exe_path = find_executable(name, path)
                if isinstance(exe_path, list):
                    exe_path = prompt_user_to_select_executable(name, exe_path)
            if exe_path:
                batch_file.write(f'cd /d "{os.path.dirname(exe_path)}"\n')
                batch_file.write(f'start "" "{exe_path}"\n')
            else:
                messagebox.showwarning("警告", f"在路径 {path} 中未找到可执行文件。")
    messagebox.showinfo("完成", f"批处理文件已生成：{batch_file_path}")

# 提示用户选择可执行文件
def prompt_user_to_select_executable(display_name, exe_files):
    def on_select():
        selected_index = listbox.curselection()
        if selected_index:
            selected_exe.set(exe_files[selected_index[0]])
            dialog.destroy()
        else:
            messagebox.showwarning("警告", "请先选择一个可执行文件。")

    selected_exe = tk.StringVar()
    dialog = tk.Toplevel()
    dialog.title(f"选择 {display_name} 的可执行文件")
    dialog.geometry("500x400")

    frame = ttk.Frame(dialog, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(frame, selectmode=tk.SINGLE, yscrollcommand=scrollbar.set)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)

    for exe in exe_files:
        listbox.insert(tk.END, exe)

    select_button = ttk.Button(dialog, text="选择", command=on_select)
    select_button.pack(pady=10)

    dialog.wait_window(dialog)

    return selected_exe.get()

# 创建主界面
def create_main_window(software_list):
    def filter_software_list(keyword):
        filtered_list = [(name, path) for name, path in software_list if keyword.lower() in name.lower()]
        update_treeview(filtered_list)

    def update_treeview(filtered_list):
        tree.delete(*tree.get_children())
        for software in filtered_list:
            tree.insert('', 'end', values=(software[0], software[1]))

    def on_search():
        keyword = search_var.get()
        filter_software_list(keyword)


    root = tk.Tk()
    root.title("选择已安装软件")
    root.geometry("800x600")

    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)\
    
    # 添加搜索框和搜索按钮
    search_var = tk.StringVar()
    search_entry = ttk.Entry(root, textvariable=search_var)
    search_entry.pack(padx=10, pady=10, fill=tk.X)

    search_button = ttk.Button(root, text="搜索", command=on_search)
    search_button.pack(pady=5)

    v_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL)
    v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    tree = ttk.Treeview(frame,columns=('Software', 'Path'), show='headings', yscrollcommand=v_scrollbar.set,selectmode='extended')
    tree.heading('Software', text='软件名称')
    tree.heading('Path', text='安装路径')

    tree.column('Software', width=250,stretch=False)
    tree.column('Path', width=550,stretch=True)

    tree.pack(fill=tk.BOTH, expand=True)

    v_scrollbar.config(command=tree.yview)

    for software in software_list:
        tree.insert('', 'end', values=(software[0], software[1]))

    def on_generate():
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要生成批处理文件的软件。")
            return
        
        selected_software = []
        for item in selected_items:
            software_name = tree.item(item, 'values')[0]
            software_path = tree.item(item, 'values')[1]
            selected_software.append((software_name, software_path))
        
        generate_batch_file(selected_software)

    generate_button = ttk.Button(root, text="生成批处理文件", command=on_generate)
    generate_button.pack(pady=10)

    # 创建显示路径的不可编辑 Entry
    def on_select(event):
        selected_item = tree.selection()
        if selected_item:
            path = tree.item(selected_item[0], 'values')[1]
            entry_var.set(path)

    entry_var = tk.StringVar()
    entry = ttk.Entry(root, textvariable=entry_var, state='readonly')
    entry.pack(fill=tk.X, pady=10)
    tree.bind("<<TreeviewSelect>>", on_select)

    # 使 Entry 的内容可复制
    def copy_to_clipboard(event):
        root.clipboard_clear()
        root.clipboard_append(entry_var.get())

    entry.bind("<Button-1>", copy_to_clipboard)

    root.mainloop()

if __name__ == "__main__":
    desktop_shortcuts = get_desktop_shortcuts()
    installed_software = get_installed_software()
    combined_software_list = desktop_shortcuts + installed_software
    create_main_window(combined_software_list)