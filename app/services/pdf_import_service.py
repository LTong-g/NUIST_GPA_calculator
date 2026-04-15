import camelot
import pandas as pd
import PyPDF2


def is_valid_score_pdf(file_path, title_keyword):
    # 读取内容确认是否是成绩单
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfFileReader(file)
        num_pages = reader.numPages
        for page_num in range(num_pages):
            page = reader.getPage(page_num)
            text = page.extractText()
            if text.find(title_keyword) == -1:
                return False
    return True


def parse_score_rows(file_path, stop_keyword):
    # 读取PDF中的所有表格
    tables = camelot.read_pdf(file_path, pages="all")
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

    rows = []
    for _, row in score_table.iterrows():
        if row[0] == stop_keyword:
            break
        if not row[3]:
            term = row[0]
        else:
            course_name = row[0]
            nature = row[1]
            score = row[3]
            credit = row[2]
            rows.append((f" {term} {course_name} {nature} {score} {credit}"))
    return rows

