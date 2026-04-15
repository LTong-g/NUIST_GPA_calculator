import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageDraw, ImageFont
import datetime
import camelot
import pandas as pd
import PyPDF2
import os


# handle字典
handle = {"window_help_open?": 0,
          "skin": "day",
          "error_data_included?": 0}

pure_score_relationship = {"优": 92, "优秀": 92, "良": 85, "良好": 85, "中": 75, "中等": 75, "及格": 65, "不及格": 0}

bend_relationship_1 = {100: 5.0, 96: 4.8, 93: 4.5,
                       90: 4.0, 86: 3.8, 83: 3.5,
                       80: 3.0, 76: 2.8, 73: 2.5,
                       70: 2.0, 66: 1.8, 63: 1.5,
                       60: 1.0, 0: 0.0}

bend_relationship_2 = {"优": 4.0, "优秀": 4.0, "良": 3.5, "良好": 3.5, "中": 2.5, "中等": 2.5, "及格": 1.5, "不及格": 0.0}

labels = []  # label列表
entries = []  # entry列表
buttons = []  # button列表
deleted_scores = []  # 删除的内容（用于撤销）

font = ImageFont.truetype(r"C:\Windows\Fonts\MSYH.ttc", size=30)  # 设置字体
fontBD = ImageFont.truetype(r"C:\Windows\Fonts\MSYHBD.ttc", size=30)  # 设置粗体字体
color_root = "#ffe0ff"  # 设置root初始颜色（background）
color_words = "black"  # 设置文字初始颜色（所有foreground）
color_button_bg = "#ffe190"  # 设置按钮初始background
color_button_abg = "#efd180"  # 设置按钮初始active background
color_entry_bg = "#f0f0f0"  # 设置输入框初始background


# 设置默认输入框文字
def set_default_entry_words():
    entry_term.configure(fg="grey")
    entry_term.insert(0, '可不填')
    entry_course_name.configure(fg="grey")
    entry_course_name.insert(0, '可不填')
    entry_nature.configure(fg="grey")
    entry_nature.insert(0, '输入包含“选”字为选修课，否则为必修')


# 重设空行
def reset_blank_row():
    for child in table.get_children():
        if not table.item(child)["values"]:
            table.delete(child)
    if 8 - len(table.get_children()) >= 0:
        for _ in range(8 - len(table.get_children())):
            table.insert("", "end", tags="row", values=())


# 清空输入框
def clean_entries():
    for entry in entries:
        entry.delete(0, 'end')


# 清空计算结果
def delete_label_result():
    global label_result
    if type(label_result) != int:
        for _ in range(len(root.grid_slaves(row=6, column=2))):
            root.grid_slaves(row=6, column=2)[0].grid_forget()


# 设置剪贴板内容
def set_clipboard_content(content):
    root.clipboard_clear()  # 清空剪贴板
    root.clipboard_append(content)  # 设置剪贴板内容
    root.update()  # 更新剪贴板


# 计算课程绩点（对应关系）
def calculate(score):
    # 如果是等级制
    if score in bend_relationship_2.keys():
        return bend_relationship_2[score]
    # 如果是分数值
    else:
        for key, value in bend_relationship_1.items():
            if score >= key:
                return value
            else:
                continue


# 添加成绩（按钮）
def add_score():
    delete_label_result()  # 清空绩点计算结果

    # 从输入框获取输入的成绩信息
    term = entry_term.get()
    course_name = entry_course_name.get()
    nature = entry_nature.get()
    score = entry_score.get()
    credit = entry_credit.get()

    # 验证数据：分数如果不是等级制，就必须能转换成浮点数（再转换成整数），且在0-100之间。学分必须能转换成浮点数，且大于0。
    if score not in ("优", "优秀", "良", "良好", "中", "中等", "及格", "不及格"):
        try:
            score = float(score)
            score //= 1
            score = int(score)
            if score < 0 or score > 100:
                tk.messagebox.showerror("错误", "请输入正确的分数")
                return
        except:
            tk.messagebox.showerror("错误", "请输入正确的分数")
            return
    try:
        credit = float(credit)
        if credit < 0:
            tk.messagebox.showerror("错误", "请输入正确的学分")
            return
    except:
            tk.messagebox.showerror("错误", "请输入正确的学分")
            return

    # 将成绩信息添加到tabel中，学期、课程名称和课程性质可以空着
    if term == "可不填":
        term = "None"
    if course_name == "可不填":
        course_name = "None"
    if nature == "输入包含“选”字为选修课，否则为必修":
        nature = "None"
    table.insert("", "end", tags="row", values=(f" {term} {course_name} {nature} {score} {credit}"))

    clean_entries()  # 清空输入框
    set_default_entry_words()  # 设置默认文字
    reset_blank_row()  # 重设空行


# 计算绩点（按钮）
def calculate_gpa():
    # 处理错误数据函数（检查或删除）
    def deal_error_data(method):
        for child in table.get_children():
            row_data = table.item(child)["values"]
            if row_data:  # 如果不是空行
                if row_data[3] not in ("优", "优秀", "良", "良好", "中", "中等", "及格", "不及格"):
                    try:
                        row_data[3] = float(row_data[3])
                        continue
                    except:
                        if method == "delete":
                            table.delete(child)
                            handle["error_data_included?"] = 0
                        elif method == "check":
                            handle["error_data_included?"] = 1
                        continue
                try:
                    row_data[4] = float(row_data[4])
                    continue
                except:
                    if method == "delete":
                        table.delete(child)
                        handle["error_data_included?"] = 0
                    elif method == "check":
                        handle["error_data_included?"] = 1
                    continue

    delete_label_result()  # 清除计算结果

    # 初始化变量
    total_credit = 0
    total_must_credit = 0
    pure_score = 0
    total_scores = 0
    total_must_scores = 0
    total_gpa = 0
    total_must_gpa = 0
    gpa = 0
    must_gpa = 0
    deal_error_data("check")

    # 存在错误数据则弹框提示请求删除
    if handle["error_data_included?"] == 1:
        result = tk.messagebox.askyesno("错误", "存在错误数据，是否自动删除？", icon="error")
        if result:
            deal_error_data("delete")
        else:
            return

    # 逐行读取table内容，取[2]位为课程性质，[3]位为分数，[4]位为学分，并根据对应关系计算绩点，加权求和
    for child in table.get_children():
        row_data = table.item(child)["values"]
        if row_data:  # 如果不是空行
            nature = str(row_data[2])
            score = row_data[3]
            pure_score = score if type(score) in (int, float) else pure_score_relationship[score]
            credit = float(row_data[4])
            gpa = calculate(score)
            total_credit += credit
            total_scores += float(pure_score) * credit
            total_gpa += gpa * credit
            if "选" not in nature:
                total_must_credit += credit
                total_must_scores += float(pure_score) * credit
                total_must_gpa += gpa * credit


    # 除以总学分，计算出绩点，在窗口底部显示计算结果
    global label_result
    delete_label_result()
    if total_credit > 0:  # 如果总学分大于0
        score = round(total_scores / total_credit, 4)
        gpa = round(total_gpa / total_credit, 4)  # 计算出总绩点
        if total_must_credit > 0:  # 如果总必修学分大于0
            must_score = round(total_must_scores / total_must_credit, 4)  # 计算出必修加权平均分
            must_gpa = round(total_must_gpa / total_must_credit, 4)  # 计算出必修绩点

        if check_selective_involved_variable.get():
            label_result = tk.Label(root, text=f"共获学分：{total_credit}    加权平均分：{score}    平均绩点：{gpa}", fg="red" if handle["skin"] == "day" else "#ffa0a0", bg=color_root)
            label_result.grid(row=6, column=2, sticky="e")
        else:
            label_result = tk.Label(root, text=f"共获学分：{total_must_credit}    加权平均分：{must_score}    平均绩点：{must_gpa}", fg="red" if handle["skin"] == "day" else "#ffa0a0", bg=color_root)
            label_result.grid(row=6, column=2, sticky="e")


# 删除（按钮）
def delete_score():
    deleted_scores.clear()  # 清空删除的成绩列表

    # 删除选中的数据，添加进列表
    for item_id in table.selection():
        deleted_scores.append(table.item(item_id)["values"])
        table.delete(item_id)

    reset_blank_row()  # 重设空行
    delete_label_result()  # 清空计算结果
    clean_entries()  # 清空输入框
    set_default_entry_words()  # 设置默认文字

    # 隐藏删除和覆盖按钮，显示撤销按钮
    button_delete_score.grid_remove()
    button_edit_score.grid_remove()
    button_undo_delete.grid()


# 清空成绩（按钮）
def delete_all_scores():
    deleted_scores.clear()  # 清空被删成绩列表

    # 删除所有成绩，添加进列表
    for child in table.get_children():
        deleted_scores.append(table.item(child)["values"])
        table.delete(child)

    reset_blank_row()  # 重设空行
    delete_label_result()  # 清空计算结果
    clean_entries()  # 清空输入框
    set_default_entry_words()  # 设置默认文字

    # 隐藏删除和覆盖按钮，显示撤销按钮
    button_edit_score.grid_remove()
    button_delete_score.grid_remove()
    button_undo_delete.grid()


# 解析（按钮）
def analyse():
    delete_label_result()  # 清空绩点计算结果

    # 弹框选择文件
    file_path = filedialog.askopenfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

    # 读取内容确认是否是成绩单
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfFileReader(file)
        num_pages = reader.numPages
        for page_num in range(num_pages):
            page = reader.getPage(page_num)
            text = page.extractText()
            if text.find('南京信息工程大学本科生学业成绩表') == -1:
                return tk.messagebox.showerror("错误", "解析失败，请选择正确的成绩单pdf文件")

    # 读取PDF中的所有表格
    tables = camelot.read_pdf(file_path, pages='all')
    score_table = tables[0].df

    # 跳过第一行
    score_table = score_table.iloc[1:]

    # 分别获取四组数据
    group1 = score_table.iloc[:, 0:4]
    group2 = score_table.iloc[:, 4:8]
    group3 = score_table.iloc[:, 8:12]

    # 重置索引
    group2.columns = group1.columns
    group3.columns = group1.columns

    # 拼接这些组
    score_table = pd.concat([group1, group2, group3], ignore_index=True)

    for _, row in score_table.iterrows():
        if row[0] == '以下空白':
            break
        if not row[3]:
            term = row[0]
        else:
            course_name = row[0]
            nature = row[1]
            score = row[3]
            credit = row[2]
            table.insert("", "end", tags="row", values=(f" {term} {course_name} {nature} {score} {credit}"))

    reset_blank_row()  # 重设空行
    clean_entries()  # 清空输入框
    set_default_entry_words()  # 设置默认文字


# 覆盖（按钮）
def edit_score():
    # 将输入框的内容替换选中的数据
    selected_item = table.selection()[0]
    term = entry_term.get()
    course_name = entry_course_name.get()
    nature =entry_nature.get()
    score = entry_score.get()
    credit = entry_credit.get()
    if term == "可不填":
        term = "None"
    if course_name == "可不填":
        course_name = "None"
    if nature == "输入包含“选”字为选修课，否则为必修":
        nature = "None"
    new_data = [term, course_name, nature, score, credit]
    table.item(selected_item, values=new_data)


# 保存（按钮）
def save_image():
    global label_result

    # 创建画布和字体
    hhh = 225
    for i, item_id in enumerate(table.get_children()):
        hhh = i * 50 + 225
    image = Image.new("RGB", (1100, hhh), "white")
    draw = ImageDraw.Draw(image)

    # 添加表头信息到图片中
    x = 50
    y = 50
    draw.rectangle([(x - 25, y - 3), (x + 1025, y + 47)], fill=(228, 235, 252))
    text = table.heading(table["columns"][0])["text"]
    draw.text((x, y), text, font=fontBD, fill=(0, 0, 0))
    x += 100
    text = table.heading(table["columns"][1])["text"]
    draw.text((x, y), text, font=fontBD, fill=(0, 0, 0))
    x += 450
    text = table.heading(table["columns"][2])["text"]
    draw.text((x, y), text, font=fontBD, fill=(0, 0, 0))
    x += 200
    text = table.heading(table["columns"][3])["text"]
    draw.text((x, y), text, font=fontBD, fill=(0, 0, 0))
    x += 200
    text = table.heading(table["columns"][4])["text"]
    draw.text((x, y), text, font=fontBD, fill=(0, 0, 0))
    x += 100

    # 添加成绩信息到图片中
    for i, item_id in enumerate(table.get_children()):
        if table.item(item_id)["values"] != "":
            x = 50
            y += 50
            color = (255, 255, 255) if i % 2 == 0 else (228, 235, 252)
            draw.rectangle([(x-25, y-3), (x+75, y+47)], fill=color)
            draw.text((x, y), str(table.item(item_id)["values"][0]), font=font, fill=(0, 0, 0))
            x += 100
            draw.rectangle([(x-25, y-3), (x+425, y+47)], fill=color)
            draw.text((x, y), str(table.item(item_id)["values"][1]), font=font, fill=(0, 0, 0))
            x += 450
            draw.rectangle([(x-25, y-3), (x+175, y+47)], fill=color)
            draw.text((x, y), str(table.item(item_id)["values"][2]), font=font, fill=(0, 0, 0))
            x += 200
            draw.rectangle([(x-25, y-3), (x+175, y+47)], fill=color)
            draw.text((x, y), str(table.item(item_id)["values"][3]), font=font, fill=(0, 0, 0))
            x += 200
            draw.rectangle([(x-25, y-3), (x+75, y+47)], fill=color)
            draw.text((x, y), str(table.item(item_id)["values"][4]), font=font, fill=(0, 0, 0))
            x += 100

    # 添加绩点到图片中
    x = 50
    y += 50
    if type(label_result) != int:
        draw.rectangle([(x-25, y-3), (x+1025, y+47)], fill=(200, 200, 200))
        draw.text((x, y), label_result.cget("text"), font=font, fill=(255, 0, 0))

    # 选择保存路径并设置默认文件名为”当前时间_绩点“
    file_name = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + "_绩点.jpg"  # 设置文件名
    file_path = filedialog.asksaveasfilename(initialfile=file_name, defaultextension=".jpg", filetypes=[("JPEG files", "*.jpg"), ("All files", "*.*")])

    # 保存图片
    if file_path:
        image.save(file_path)


# 重设颜色
def reset_color():
    global color_root
    global color_words
    global color_button_bg
    global color_button_abg
    global color_entry_bg
    if handle["skin"] == "day":
        color_root = "#ffe0ff"
        color_words = "black"
        color_button_bg = "#ffe190"
        color_button_abg = "#efd180"
        color_entry_bg = "#f0f0f0"
    else:
        color_root = "#202020"
        color_words = "#e0e0e0"
        color_button_bg = "#001e7f"
        color_button_abg = "#000e6f"
        color_entry_bg = "#2f2f2f"


# 切换皮肤（按钮）
def change_skin():
    # 设置skin状态
    if handle["skin"] == "day":
        handle["skin"] = "night"
        reset_color()
        button_change_skin.configure(text="☀", fg="yellow", activeforeground="yellow", bg=color_button_bg, activebackground=color_button_abg)
    else:
        handle["skin"] = "day"
        reset_color()
        button_change_skin.configure(text="🌙", fg="purple", activeforeground="purple", bg=color_button_bg, activebackground=color_button_abg)

    # 设置颜色
    root.configure(bg=color_root)
    table.tag_configure("row", background=color_entry_bg, foreground=color_words)
    check_selective_involved.configure(fg=color_words, activeforeground=color_words, bg=color_root, activebackground=color_root)
    for button in buttons:
        if button != button_change_skin:
            button.configure(fg=color_words, activeforeground=color_words, bg=color_button_bg, activebackground=color_button_abg)
    for label in labels:
        delete_label_result()
        calculate_gpa()
        if type(label) != int:
            label.configure(fg=color_words, bg=color_root)
    for entry in entries:
        entry.configure(fg=color_words, bg=color_entry_bg)

    # 离焦
    on_entry_term_focus_out('<FocusOut>')
    on_entry_course_name_focus_out('<FocusOut>')


# 复制邮箱（按钮）
def copy_email():
    set_clipboard_content("LTongg@qq.com")


# 撤销（按钮）
def undo_delete():
    for content in deleted_scores:
        table.insert("", "end", tags="row", values=tuple(content))
    reset_blank_row()
    deleted_scores.clear()
    button_undo_delete.grid_remove()


# 选中行（绑定函数）
def on_select(event):
    if table.selection():
        button_delete_score.grid()  # 显示删除按钮
        button_edit_score.grid()  # 显示覆盖按钮
        clean_entries()  # 清空输入框

        # 获取选中的行的数据
        selected_item = table.selection()[0]
        selected_data = table.item(selected_item)['values']

        # 如果不是空行
        if selected_data:
            term = selected_data[0]
            course_name = selected_data[1]
            nature = selected_data[2]
            score = selected_data[3]
            credit = selected_data[4]

            # 写入输入框
            entry_term.configure(fg=color_words)
            entry_course_name.configure(fg=color_words)
            entry_nature.configure(fg=color_words)
            entry_score.configure(fg=color_words)
            entry_credit.configure(fg=color_words)
            entry_term.insert(0, term)
            entry_course_name.insert(0, course_name)
            entry_nature.insert(0, nature)
            entry_score.insert(0, score)
            entry_credit.insert(0, credit)


# 聚焦学期输入框（绑定函数）
def on_entry_term_focus_in(event):
    if entry_term.get() == '可不填':
        entry_term.delete(0, tk.END)
    entry_term.configure(fg=color_words)


# 离焦学期输入框（绑定函数）
def on_entry_term_focus_out(event):
    if not entry_term.get():
        entry_term.insert(0, '可不填')
        entry_term.configure(fg="grey")
    if entry_term.get() == '可不填':
        entry_term.configure(fg="grey")


# 聚焦课程名称输入框（绑定函数）
def on_entry_course_name_focus_in(event):
    if entry_course_name.get() == '可不填':
        entry_course_name.delete(0, tk.END)
    entry_course_name.configure(fg=color_words)


# 离焦课程名称输入框（绑定函数）
def on_entry_course_name_focus_out(event):
    if not entry_course_name.get():
        entry_course_name.configure(fg="grey")
        entry_course_name.insert(0, '可不填')
    if entry_course_name.get() == '可不填':
        entry_course_name.configure(fg="grey")


# 聚焦课程性质输入框（绑定函数)
def on_entry_nature_focus_in(event):
    if entry_nature.get() == '输入包含“选”字为选修课，否则为必修':
        entry_nature.delete(0, tk.END)
    entry_nature.configure(fg=color_words)


# 离焦课程性质输入框（绑定函数）
def on_entry_nature_focus_out(event):
    if not entry_nature.get():
        entry_nature.configure(fg="grey")
        entry_nature.insert(0, '输入包含“选”字为选修课，否则为必修')
    if entry_nature.get() == '输入包含“选”字为选修课，否则为必修':
        entry_nature.configure(fg="grey")


# 创建帮助子窗口
def run_help():
    if not handle["window_help_open?"]:
        handle["window_help_open?"] = 1
        window_help = tk.Toplevel()
        window_help.resizable(False, False)
        window_help.title("Help")
        window_help.rowconfigure(0, weight=1)
        window_help.columnconfigure(0, weight=1)
        reset_color()
        text = '''添加成绩：输入相关内容后点”添加新成绩“进行添加。
计算绩点：确认成绩数据无误后点击”计算绩点“按钮，绩点将显示在下方（”清空成绩“按钮旁边）。
删除成绩：选中一行或多行成绩后点击“删除”按钮。
覆盖成绩：选中一行数据后内容将显示在上方输入框，修改后点击“覆盖”按钮即可覆盖原成绩。
清空成绩：点击“清空成绩”按钮清空成绩。
解析：进入南信大本科新教务系统，成绩单/证明打印，下载成绩单pdf文件。点击“解析”按钮，选择下载的成绩单文件。
保存：点击“保存”按钮，将成绩保存为图片。
附：成绩与绩点对照关系和绩点计算公式。
##############################################################
#       #  成绩  #    100    #  99-96  #  95-93  #  92-90  #  89-86  #  85-83  #  82-80  #
#       ##########################################################
#  百  #  绩点  #    5.0    #    4.8    #    4.5    #    4.0    #    3.8    #    3.5    #    3.0     #
#  分  ##########################################################
#  制  #  成绩  #  79-76  #  75-73  #  72-70  #  69-66  #  65-63  #  62-60  #    <60   #
#       ##########################################################
#       #  绩点  #    2.8    #    2.5    #    2.0    #    1.8    #    1.5    #    1.0    #     0.     #
##############################################################
#  等  #  成绩  #     优    #     良     #     中     #   及格   #  不及格  #
#  级  ############################################
#  制  #  绩点  #    4.0    #     3.5    #     2.5    #  1.5     #      0     #
################################################

                         ∑（已修必修课学分 × 课程绩点）
平均分绩点 = ----------------------------------------------
                               ∑ 已修必修课程的学分'''
        label_help = tk.Label(window_help, text=text, justify="left", fg=color_words, bg=color_root)
        label_help.grid(row=0, column=0)

        # 关闭window_help函数
        def close_window_help():
            handle["window_help_open?"] = 0
            window_help.destroy()
        window_help.protocol('WM_DELETE_WINDOW', close_window_help)
        window_help.mainloop()


'''主程序'''
root = tk.Tk()  # 创建一个Tkinter窗口
root.title("绩点计算器")  # 设置窗口标题
root.geometry("600x400")  # 设置窗口大小
root.configure(bg="#ffe0ff")  # 设置窗口底色
root.resizable(False, False)  # 不可改变大小
icon_path = r"image\图标.png"
if os.path.exists(icon_path):
    try:
        root.iconphoto(False, tk.PhotoImage(file=icon_path))  # 设置窗口左上角图标
    except tk.TclError:
        pass

# 创建滚动条
y_scrollbar = tk.Scrollbar(root, orient="vertical")

# 创建表格
table = ttk.Treeview(root, columns=("col1", "col2", "col3", "col4", "col5"), show="headings", yscrollcommand=y_scrollbar.set)
table.tag_configure("row", background=color_entry_bg, foreground=color_words)
table.heading("col1", text="学期")
table.heading("col2", text="课程名称")
table.heading("col3", text="课程性质")
table.heading("col4", text="总评成绩")
table.heading("col5", text="学分")
table.column("col1", width=-50, anchor="center")
table.column("col2", width=140, anchor="center")
table.column("col3", width=-20, anchor="center")
table.column("col4", width=-20, anchor="center")
table.column("col5", width=-50, anchor="center")
reset_blank_row()

# 设置滚动条
y_scrollbar.config(command=table.yview)
y_scrollbar.grid(row=5, column=3, sticky="wns")

# 设置表格所在的列和行都可以扩展
root.columnconfigure(2, weight=1)
root.rowconfigure(5, weight=1)

# 创建标签和输入框
label_author = tk.Label(root, text="by.LTongg")
label_email = tk.Label(root, text="问题反馈/联系方式：LTongg@qq.com")
label_term = tk.Label(root, text="学期：")
entry_term = tk.Entry(root, width=52)
label_course_name = tk.Label(root, text="课程名称：")
entry_course_name = tk.Entry(root, width=52)
label_nature = tk.Label(root, text="课程性质：")
entry_nature = tk.Entry(root, width=52)
label_score = tk.Label(root, text="总评成绩：")
entry_score = tk.Entry(root, width=52)
label_credit = tk.Label(root, text="学分：")
entry_credit = tk.Entry(root, width=52)
label_result = 0

# 添加标签和输入框到对应列表
labels.append(label_author)
labels.append(label_email)
labels.append(label_term)
labels.append(label_course_name)
labels.append(label_nature)
labels.append(label_score)
labels.append(label_credit)
labels.append(label_result)
entries.append(entry_term)
entries.append(entry_course_name)
entries.append(entry_nature)
entries.append(entry_score)
entries.append(entry_credit)

# 设置标签和输入框颜色
for label in labels:
    if type(label) != int:
        label.configure(bg=color_root)
for entry in entries:
    entry.configure(bg=color_entry_bg)

# 创建“选修参与计算”勾选框
check_selective_involved_variable = tk.IntVar()
check_selective_involved = tk.Checkbutton(root, text="选修参与计算", fg=color_words, activeforeground=color_words, bg=color_root, activebackground=color_root, variable=check_selective_involved_variable)

# 创建按钮
button_add_score = tk.Button(root, width=9, text="添加新成绩", command=add_score)  # 添加成绩按钮
button_calculate_gpa = tk.Button(root, width=9, text="计算绩点", command=calculate_gpa)  # 计算绩点按钮
button_save_image = tk.Button(root, width=6, text="保存", command=save_image)  # 保存图片按钮
button_delete_score = tk.Button(root, width=6, text="删除", command=delete_score)  # 删除按钮
button_delete_all_scores = tk.Button(root, width=9, text="清空成绩", command=delete_all_scores)  # 清空成绩按钮
button_analyse = tk.Button(root, width=6, text="解析", command=analyse)  # 解析按钮
button_edit_score = tk.Button(root, width=6, text="覆盖", command=edit_score)  # 覆盖按钮
button_help = tk.Button(root, width=9, text="使用说明", command=run_help)  # 使用说明
button_change_skin = tk.Button(root, width=3, text="🌛", fg="purple", activeforeground="purple", bg="#ffe190", activebackground="#efd180", relief="groove", command=change_skin)  # 切换皮肤
button_copy_email = tk.Button(root, width=9, text="复制邮箱", command=copy_email)  # 复制邮箱
button_undo_delete = tk.Button(root, width=6, text="撤销", command=undo_delete)  # 撤销删除

# 添加按钮到对应列表
buttons.append(button_add_score)
buttons.append(button_calculate_gpa)
buttons.append(button_save_image)
buttons.append(button_delete_score)
buttons.append(button_delete_all_scores)
buttons.append(button_analyse)
buttons.append(button_edit_score)
buttons.append(button_help)
buttons.append(button_change_skin)
buttons.append(button_copy_email)
buttons.append(button_undo_delete)

# 设置按钮样式
for button in buttons:
    if button != button_change_skin:
        button.configure(fg=color_words, activeforeground=color_words, bg=color_button_bg, activebackground=color_button_abg, relief="groove")

# 设置表格位置
table.grid(row=5, column=2, sticky="nsew")

# 设置标签和输入框位置
label_author.grid(row=7, column=0, columnspan=4, sticky="w")
label_email.grid(row=7, column=0, columnspan=4, padx=75, sticky="e")
label_term.grid(row=0, column=2, sticky="w")
entry_term.grid(row=0, column=2)
label_course_name.grid(row=1, column=2, sticky="w")
entry_course_name.grid(row=1, column=2)
label_nature.grid(row=2, column=2, sticky="w")
entry_nature.grid(row=2, column=2)
label_score.grid(row=3, column=2, sticky="w")
entry_score.grid(row=3, column=2)
label_credit.grid(row=4, column=2, sticky="w")
entry_credit.grid(row=4, column=2)

# 设置复选框位置
check_selective_involved.grid(row=2, rowspan=2, column=2, columnspan=2, pady=3, sticky="se")

# 设置按钮位置
button_add_score.grid(row=3, column=2, rowspan=2, sticky="se")
button_calculate_gpa.grid(row=3, column=3, rowspan=2, sticky="se")
button_save_image.grid(row=5, column=3, sticky="se")
button_delete_score.grid(row=5, column=3, pady=30, sticky="ne")
button_delete_all_scores.grid(row=6, column=3, sticky="e")
button_analyse.grid(row=5, column=3, sticky="ne")
button_edit_score.grid(row=5, column=3, pady=60, sticky="ne")
button_help.grid(row=0, column=3, rowspan=2, sticky="ne")
button_change_skin.grid(row=0, column=2, rowspan=2, sticky="ne")
button_copy_email.grid(row=7, column=2, columnspan=2, sticky="se")
button_undo_delete.grid(row=5, column=3, pady=30, sticky="se")

# 隐藏删除、覆盖、撤销按钮
button_delete_score.grid_remove()
button_edit_score.grid_remove()
button_undo_delete.grid_remove()

# 绑定事件
table.bind("<<TreeviewSelect>>", on_select)
entry_term.bind('<FocusIn>', on_entry_term_focus_in)
entry_term.bind('<FocusOut>', on_entry_term_focus_out)
entry_course_name.bind('<FocusIn>', on_entry_course_name_focus_in)
entry_course_name.bind('<FocusOut>', on_entry_course_name_focus_out)
entry_nature.bind('<FocusIn>', on_entry_nature_focus_in)
entry_nature.bind('<FocusOut>', on_entry_nature_focus_out)

# 设置默认文字
set_default_entry_words()

# 启动窗口循环
root.mainloop()
