import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont


# 计算课程绩点
def calculate(score):
    if score in ("优", "优秀", "良", "良好", "中", "中等", "及格", "不及格"):
        if score in ("优", "优秀"):
            return 4.0
        elif score in ("良", "良好"):
            return 3.5
        elif score in ("中", "中等"):
            return 2.5
        elif score == "及格":
            return 1.5
        else:
            return 0.0
    else:
        if score >= 100:
            return 5.0
        elif score >= 96:
            return 4.8
        elif score >= 93:
            return 4.5
        elif score >= 90:
            return 4.0
        elif score >= 86:
            return 3.8
        elif score >= 83:
            return 3.5
        elif score >= 80:
            return 3.0
        elif score >= 76:
            return 2.8
        elif score >= 73:
            return 2.5
        elif score >= 70:
            return 2.0
        elif score >= 66:
            return 1.8
        elif score >= 63:
            return 1.5
        elif score >= 60:
            return 1.0
        else:
            return 0.0


# 添加成绩
def add_score():
    # 清空绩点计算结果
    try:
        for i in range(len(root.grid_slaves(row=6, column=2))):
            root.grid_slaves(row=6, column=2)[0].grid_forget()
    except:
        pass
    # 获取输入的成绩信息
    course_name = entry_course_name.get()
    term = entry_term.get()
    score = entry_score.get()
    credit = entry_credit.get()
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
                tk.messagebox.showerror("错误", "学分必须大于等于0")
                return
        except:
            tk.messagebox.showerror("错误", "请输入正确的学分")
            return

    # 将成绩信息添加到tabel中
    if not course_name:
        course_name = "None"
    if not term:
        term = "None"
    table.insert("", "end", values=(f"{course_name} {term} {score} {credit}"))

    # 清空输入框
    entry_course_name.delete(0, 'end')
    entry_term.delete(0, 'end')
    entry_score.delete(0, 'end')
    entry_credit.delete(0, 'end')


# 计算总绩点
def calculate_gpa():
    try:
        # 清除计算结果
        try:
            for i in range(len(root.grid_slaves(row=6, column=2))):
                root.grid_slaves(row=6, column=2)[0].grid_forget()
        except:
            pass
        total_credit = 0
        total_gpa = 0
        for child in table.get_children():
            row_data = table.item(child)["values"]
            score = row_data[2]
            credit = float(row_data[3])
            gpa = calculate(score)
            total_credit += credit
            total_gpa += gpa * credit

        try:
            gpa = round(total_gpa / total_credit, 4)
        except:
            return

        # 在窗口底部显示计算结果
        global label_result
        label_result = tk.Label(root, text=f"总学分：{total_credit}        绩点：{gpa}", fg="red")
        label_result.grid(row=6, column=2, sticky="e")
    except ValueError:
        tk.messagebox.showerror("错误", "有错误数据，请仔细检查！")
        result = tk.messagebox.askyesno("提示", "是否应用自动修正？")
        if result:
            for child in table.get_children():
                row_data = table.item(child)["values"]
                if row_data[2] not in ("优", "优秀", "良", "良好", "中", "中等", "及格", "不及格"):
                    try:
                        row_data[2] = float(row_data[2])
                        continue
                    except:
                        table.delete(child)
                        continue
                try:
                    row_data[3] = float(row_data[3])
                    continue
                except:
                    table.delete(child)
                    continue
            return calculate_gpa()


# 删除成绩
def delete_score():
    for item_id in table.selection():
        table.delete(item_id)
    try:
        for i in range(len(root.grid_slaves(row=6, column=2))):
            root.grid_slaves(row=6, column=2)[0].grid_forget()
    except:
        pass
    button_delete_score.grid_remove()


# 删除所有成绩
def delete_all_scores():
    for child in table.get_children():
        table.delete(child)
    try:
        for i in range(len(root.grid_slaves(row=6, column=2))):
            root.grid_slaves(row=6, column=2)[0].grid_forget()
    except:
        pass


# 解析粘贴成绩
def analyse():
    # 清空绩点计算结果
    try:
        for i in range(len(root.grid_slaves(row=6, column=2))):
            root.grid_slaves(row=6, column=2)[0].grid_forget()
    except:
        pass

    # 获取粘贴的成绩信息
    content = entry_paste.get()
    content = content.split("\n")
    for i in range(len(content)):
        content_child = content[i].split("\t")
        course_name = content_child[2]
        term = content_child[1]
        score = content_child[4]
        credit = content_child[7]

        # 将成绩信息添加到tabel中
        if not course_name:
            course_name = "None"
        if not term:
            term = "None"
        table.insert("", "end", values=(f"{course_name} {term} {score} {credit}"))

    # 清空输入框
    entry_paste.delete(0, 'end')


# 定义保存图片函数
def save_image():
    global label_result

    # 创建画布和字体
    hhh = 0
    for i, item_id in enumerate(table.get_children()):
        hhh = (i + 4) * 50
    image = Image.new("RGB", (1850, hhh), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(r"C:\Windows\Fonts\simsun.ttc", size=32)

    # 添加表头信息到图片中
    x = 50
    y = 50
    for col in table["columns"]:
        text = table.heading(col)["text"]
        draw.text((x, y), text, font=font, fill=(0, 0, 0))
        x += 450

    # 添加成绩信息到图片中
    for i, item_id in enumerate(table.get_children()):
        x = 50
        y += 50
        color = (255, 255, 255) if i % 2 != 0 else (228, 235, 252)
        for value in table.item(item_id)["values"]:
            draw.rectangle([(x-25, y-10), (x+425, y+40)], fill=color)
            draw.text((x, y), str(value), font=font, fill=(0, 0, 0))
            x += 450
    y += 50
    if type(label_result) != int:
        draw.rectangle([(1100, y - 10), (1700, y + 40)], fill=(200, 200, 200))
        draw.text((1150, y), label_result.cget("text"), font=font, fill=(255, 0, 0))

    # 选择保存路径和文件名
    file_path = filedialog.asksaveasfilename(defaultextension=".jpg", filetypes=[("JPEG files", "*.jpg"), ("All files", "*.*")])

    # 保存图片
    if file_path:
        image.save(file_path)


# 行列标签（调试用）
def set_labels():
    tk.Label(root, text="0").grid(row=0, column=0)
    tk.Label(root, text="1").grid(row=0, column=1)
    tk.Label(root, text="2").grid(row=0, column=2)
    tk.Label(root, text="3").grid(row=0, column=3)
    tk.Label(root, text="4").grid(row=0, column=4)

    tk.Label(root, text="1").grid(row=1, column=0)
    tk.Label(root, text="2").grid(row=2, column=0)
    tk.Label(root, text="3").grid(row=3, column=0)
    tk.Label(root, text="4").grid(row=4, column=0)
    tk.Label(root, text="5").grid(row=5, column=0)
    tk.Label(root, text="6").grid(row=6, column=0)
    tk.Label(root, text="7").grid(row=7, column=0)
    tk.Label(root, text="8").grid(row=8, column=0)

    tk.Label(root, text="1").grid(row=1, column=4)
    tk.Label(root, text="2").grid(row=2, column=4)
    tk.Label(root, text="3").grid(row=3, column=4)
    tk.Label(root, text="4").grid(row=4, column=4)
    tk.Label(root, text="5").grid(row=5, column=4)
    tk.Label(root, text="6").grid(row=6, column=4)
    tk.Label(root, text="7").grid(row=7, column=4)
    tk.Label(root, text="8").grid(row=8, column=4)

    tk.Label(root, text="1").grid(row=8, column=1)
    tk.Label(root, text="2").grid(row=8, column=2)
    tk.Label(root, text="3").grid(row=8, column=3)
    tk.Label(root, text="4").grid(row=8, column=4)


root = tk.Tk()  # 创建一个Tkinter窗口
root.title("分数绩点计算器")  # 设置窗口标题
root.geometry("600x400")  # 设置窗口大小
root.resizable(False, False)  # 不可改变大小


# 创建滚动条
scrollbar = ttk.Scrollbar(root)
scrollbar.grid(row=5, column=3, sticky="wns")

# 创建表格
table = ttk.Treeview(root, columns=("col1", "col2", "col3", "col4"), show="headings", yscrollcommand=scrollbar.set)
table.heading("col1", text="课程名称")
table.heading("col2", text="开设学期")
table.heading("col3", text="课程分数")
table.heading("col4", text="课程学分")
table.column("col1", width=90, anchor="center")
table.column("col2", width=50, anchor="center")
table.column("col3", width=50, anchor="center")
table.column("col4", width=50, anchor="center")

# 设置滚动条
scrollbar.config(command=table.yview)

# 设置表格所在的列和行都可以扩展
root.columnconfigure(2, weight=1)
root.rowconfigure(5, weight=1)

# 创建标签和输入框
label_author = tk.Label(root, text="by.LTongg")
label_00 = tk.Label(root, text=" ")
label_course_name = tk.Label(root, text="课程名称：")
entry_course_name = tk.Entry(root, width=50)
label_term = tk.Label(root, text="开设学期：")
entry_term = tk.Entry(root, width=50)
label_score = tk.Label(root, text="课程分数：")
entry_score = tk.Entry(root, width=50)
label_credit = tk.Label(root, text="课程学分：")
entry_credit = tk.Entry(root, width=50)
label_paste = tk.Label(root, text="粘贴：")
entry_paste = tk.Entry(root, width=65)
entry_paste.focus()
label_result = 0

# 创建按钮
button_add_score = tk.Button(root, text="添加新成绩", command=add_score)  # 添加成绩按钮
button_calculate_gpa = tk.Button(root, text="计算绩点", command=calculate_gpa)  # 计算绩点按钮
button_save_image = tk.Button(root, text="保存", command=save_image)  # 保存图片按钮
button_delete_score = tk.Button(root, text="删除", command=delete_score)  # 删除按钮
button_delete_all_scores = tk.Button(root, text="清空成绩", command=delete_all_scores)  # 清空成绩按钮
button_analyse = tk.Button(root, text="解析", command=analyse)  # 解析按钮

# 设置表格位置
table.grid(row=5, column=2, sticky="nsew")

# 设置标签和输入框位置
label_00.grid(row=0, column=0)
label_author.grid(row=0, column=0, columnspan=4, sticky="e")
label_course_name.grid(row=1, column=2, sticky="w")
entry_course_name.grid(row=1, column=2)
label_term.grid(row=2, column=2, sticky="w")
entry_term.grid(row=2, column=2)
label_score.grid(row=3, column=2, sticky="w")
entry_score.grid(row=3, column=2)
label_credit.grid(row=4, column=2, sticky="w")
entry_credit.grid(row=4, column=2)
label_paste.grid(row=7, column=2, sticky="w")
entry_paste.grid(row=7, column=2, sticky="e")

# 设置按钮位置
button_add_score.grid(row=4, column=2, sticky="e")
button_calculate_gpa.grid(row=4, column=3)
button_save_image.grid(row=5, column=3, sticky="se")
button_delete_score.grid(row=5, column=3, sticky="ne")
button_delete_all_scores.grid(row=6, column=3)
button_analyse.grid(row=7, column=3, sticky="we")

# 隐藏删除按钮
button_delete_score.grid_remove()


# 选中行时显示删除按钮
def on_select(event):
    if table.selection():
        button_delete_score.grid()


table.bind("<<TreeviewSelect>>", on_select)


root.mainloop()
