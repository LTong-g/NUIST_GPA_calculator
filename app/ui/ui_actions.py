def reset_blank_rows(table, min_rows=8, tag="row"):
    for child in table.get_children():
        if not table.item(child)["values"]:
            table.delete(child)
    if min_rows - len(table.get_children()) >= 0:
        for _ in range(min_rows - len(table.get_children())):
            table.insert("", "end", tags=tag, values=())


def clear_entries(entries):
    for entry in entries:
        entry.delete(0, "end")


def clear_result_label(root, label_result, row=6, column=2):
    if type(label_result) != int:
        for _ in range(len(root.grid_slaves(row=row, column=column))):
            root.grid_slaves(row=row, column=column)[0].grid_forget()


def set_clipboard_content(root, content):
    root.clipboard_clear()
    root.clipboard_append(content)
    root.update()
