from requests import get
from threading import Event
from os import remove
import win32api
from json import loads
from tqdm import tqdm
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

def check_version():
    # 获取现版本
    filename = "绩点计算器.exe"
    info = win32api.GetFileVersionInfo(filename, '\\')
    ms = info['FileVersionMS']
    ls = info['FileVersionLS']
    current_version = 'v' + '%d.%d.%d.%d' % (win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls), win32api.LOWORD(ls))

    # 设置GitHub仓库url
    api_url = "https://api.github.com/repos/LTong-g/NUIST_GPA_calculator/releases/latest"
    headers = {"Authorization": "token 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"}

    # 发送请求
    response = get(api_url, headers=headers, timeout=1)

    # 获取数据
    data = loads(response.text)

    # 获取最新版本号
    latest_version = data["tag_name"]

    # 比较版本号,返回值
    return current_version < latest_version

def download_update():
    # 初始化
    handle = {"cancel": 0,
              "finished": 0}
    if check_version():
        result = tk.messagebox.askyesno("Update", "检测到新版本，是否下载？")
        if result:
            save_path = filedialog.asksaveasfilename(initialfile="绩点计算器 Setup.exe", defaultextension=".exe",
                                                     filetypes=[("Application", "*.exe"), ("All files", "*.*")])
            if save_path:
                stop_event = Event()

                def stop_download():
                    stop_event.set()
                    label_ask_update["text"] = "下载取消！"
                    handle["cancel"] = 1
                    if not handle["finished"]:
                        button_1.grid_remove()

                update_window = tk.Tk()
                update_window.title("Update")
                update_window.geometry("600x150")
                update_window.resizable(False, False)
                update_window.columnconfigure(0, weight=1)
                update_window.rowconfigure(0, weight=1)

                label_ask_update = tk.Label(update_window, text="正在准备下载")
                label_ask_update.grid(row=0, column=0, pady=40, sticky="nwe")
                label_tip = tk.Label(update_window, text="请等待下载完成或取消下载再关闭")
                label_tip.grid(row=0, column=0, sticky="s")

                progress_bar = ttk.Progressbar(update_window, orient="horizontal", length=560, mode="determinate")
                progress_bar.grid(row=0, column=0)

                rate = 0.0
                label_rate = tk.Label(update_window, text=rate)
                label_rate.grid(row=0, column=0, pady=30, sticky="se")

                button_1 = tk.Button(update_window, text="取消下载", command=stop_download)
                button_1.grid(row=0, column=0, sticky="se")

                label_ask_update["text"] = "正在下载"
                api_url = "https://api.github.com/repos/LTong-g/NUIST_GPA_calculator/releases/latest"
                headers = {"Authorization": "token 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"}
                response = get(api_url, headers=headers, timeout=1)
                data = loads(response.text)
                file_url = data["assets"][0]["browser_download_url"]
                response = get(file_url, stream=True)
                total_size = int(response.headers.get('content-length', 0))
                progress_bar["maximum"] = total_size

                with open(save_path, "wb") as f:
                    for chunk in response.iter_content(chunk_size=1024):
                        if stop_event.is_set():
                            break
                        if chunk:
                            f.write(chunk)
                            progress_bar["value"] += len(chunk)
                            progress_bar.update()
                            rate = round(progress_bar.cget("value") / progress_bar.cget("maximum") * 100, 2)
                            label_rate["text"] = str(rate) + "%"

                handle["finished"] = 1
                if handle["cancel"]:
                    remove(save_path)
                else:
                    update_window.destroy()

                update_window.mainloop()


