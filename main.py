import os
import re
import xlsxwriter
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

def find_line_in_file(file_name, TAG_value):
    if not os.path.exists(file_name):
        print(f"File {file_name} does not exist")
        return

    with open(file_name, 'r', encoding='utf-8') as file:
        lines = file.readlines()
        for line_num, line in enumerate(lines, 1):
            if re.search(rf"{TAG_value}=", line):
                if line_num < len(lines) - 1 and re.search(r"instances_counter=\d+", lines[line_num]):
                    print(f"Success: {TAG_value} line followed by instances_counter")


def extract_data():
    file_name = file_entry.get()  # Получаем имя файла из текстового поля
    output_name=file_name
    file_name+='.hoi4'
    if not os.path.exists(file_name):
        messagebox.showerror("File Error", f"File {file_name} does not exist.")
        return

    with open('Settings.txt', 'r', encoding='utf-8') as file:
        lines = file.readlines()
        TAG_values = [line.strip().split(':')[1].split(';')[0] for line in lines if ':' in line]
        NAME_values = [line.strip().split(':')[1].split(';')[1].split('#')[0] for line in lines if '#' in line]
        print(NAME_values)
        TAG_values = [value.upper() for value in TAG_values]

    workbook = xlsxwriter.Workbook(f'{output_name}.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Tag Name')
    worksheet.write(0, 1, 'Name')
    worksheet.write(0, 2, 'GDP')
    worksheet.write(0, 3, 'GDP per capita')
    worksheet.write(0, 4, 'Effective civilian factories')
    worksheet.write(0, 5, 'Effective military factories')
    worksheet.write(0, 6, 'Effective_naval_factories')
    worksheet.write(0, 7, 'Equipment operational cost')
    worksheet.write(0, 8, 'Overall productivity')
    worksheet.write(0, 9, 'Interest rate')
    worksheet.write(0, 10, 'defence_gain')
    row = 1
    i = 0
    for TAG_value in TAG_values:
        find_line_in_file(file_name, TAG_value)
        with open(file_name, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            for line_num, line in enumerate(lines, 1):
                if re.search(rf"{TAG_value}=", line):
                    if line_num < len(lines) - 1 and (re.search(r"instances_counter=\d+", lines[line_num]) or re.search(r"gdpc_converging_var=\d+", lines[line_num])):
                        for next_line_num, next_line in enumerate(lines[line_num+1:], line_num + 2):
                            if re.search(r"defence_gain=([0-9.,]+)", next_line):
                                value_defence_gain = float(re.search(r"defence_gain=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found defence_gain: {value_defence_gain}")
                            elif re.search(r"interest_rate=([0-9.,]+)", next_line):
                                value_interest_rate = float(re.search(r"interest_rate=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found interest_rate: {value_interest_rate}")
                            elif re.search(r"effective_civilian_factories=([0-9.,]+)", next_line):
                                value_effective_civilian_factories = float(re.search(r"effective_civilian_factories=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found effective_civilian_factories: {value_effective_civilian_factories}")
                            elif re.search(r"effective_military_factories=([0-9.,]+)", next_line):
                                value_effective_military_factories = float(re.search(r"effective_military_factories=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found effective_military_factories: {value_effective_military_factories}")
                            elif re.search(r"effective_naval_factories=([0-9.,]+)", next_line):
                                value_effective_naval_factories = float(re.search(r"effective_naval_factories=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found effective_naval_factories: {value_effective_naval_factories}")
                            elif re.search(r"equipment_operative_cost=([0-9.,]+)", next_line):
                                value_equipment_operative_cost = float(re.search(r"equipment_operative_cost=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found equipment_operative_cost=: {value_equipment_operative_cost}")
                            elif re.search(r"gdp_per_capita=([0-9.,]+)", next_line):
                                value_gdp_per_capita = float(re.search(r"gdp_per_capita=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found gdp_per_capita: {value_gdp_per_capita}")
                            elif re.search(r"gdp_total=([0-9.,]+)", next_line):
                                value_gdp = float(re.search(r"gdp_total=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} Found Found gdp_total: {value_gdp}")
                            elif re.search(r"overall_productivity=([0-9.,]+)", next_line):
                                value_overall_productivity = float(re.search(r"overall_productivity=([0-9.,]+)", next_line).group(1))
                                print(f"{TAG_value} overall_productivity: {value_overall_productivity}")
                                break

        if value_gdp is not None:
            worksheet.write(row, 0, TAG_value)
            worksheet.write(row, 1, NAME_values[i])
            worksheet.write(row, 2, value_gdp)
            worksheet.write(row, 3, value_gdp_per_capita)
            worksheet.write(row, 4, value_effective_civilian_factories)
            worksheet.write(row, 5, value_effective_military_factories)
            worksheet.write(row, 6, value_effective_naval_factories)
            worksheet.write(row, 7, value_equipment_operative_cost)
            worksheet.write(row, 8, value_overall_productivity)
            worksheet.write(row, 9, value_interest_rate)
            worksheet.write(row, 10, value_defence_gain)
            i += 1
            row += 1

    workbook.close()
    messagebox.showinfo("Extraction complete", f"{output_name}.xlsx saved")


def settings():
    settings_window = tk.Tk()
    settings_window.title("Settings")

    settings_treeview = ttk.Treeview(settings_window)
    settings_treeview["columns"] = ("Tag", "Name")
    settings_treeview.column("#0", width=0)
    settings_treeview.column("Tag", anchor="w")
    settings_treeview.column("Name", anchor="w")
    settings_treeview.heading("#0", text="", anchor="w")
    settings_treeview.heading("Tag", text="Tag", anchor="w")
    settings_treeview.heading("Name", text="Name", anchor="w")

    settings_treeview.pack()

    with open('Settings.txt', 'r', encoding='utf-8') as file:
        lines = file.readlines()
        TAG_values = []
        Names = []
        for line in lines:
            if ':' in line:
                TAG_value = line.strip().split(':')[1].split(';')[0]
                TAG_values.append(TAG_value)
                Name = line.strip().split(';')[1].strip('#')
                Names.append(Name)

    for row_num, TAG_value in enumerate(TAG_values):
        settings_treeview.insert('', 'end', values=(TAG_value, Names[row_num]))

    def save_settings():
        TAG_values = []
        Names = []
        for row in settings_treeview.get_children():
            TAG_value = settings_treeview.item(row)['values'][0]
            Name = settings_treeview.item(row)['values'][1]
            TAG_values.append(f":{TAG_value};")
            Names.append(f"{Name}#")
        with open('Settings.txt', 'w', encoding='utf-8') as file:
            file.write('\n'.join([line.strip() + Names[row_num] + '\n' for row_num, line in enumerate(TAG_values)]))
        messagebox.showinfo("Settings saved", "Changes saved successfully")

    def on_double_click(event):
        item_id = event.widget.selection()[0]
        item_values = event.widget.item(item_id)['values']
        new_value_tag = simpledialog.askstring("New value", "Enter new tag", initialvalue=item_values[0])
        new_value_name = simpledialog.askstring("New value", "Enter new name", initialvalue=item_values[1])
        if new_value_tag is not None and new_value_name is not None:
            event.widget.item(item_id, values=(new_value_tag, new_value_name))

    def add_tag(event=None):
        new_tag = simpledialog.askstring("New tag", "Enter new tag")
        new_name = simpledialog.askstring("New name", "Enter new name")
        if new_tag is not None and new_name is not None:
            settings_treeview.insert('', 'end', values=(new_tag, new_name))
            messagebox.showinfo("Added tag", "New tag added successfully")

    def delete_tag():
        item_id = settings_treeview.selection()[0]
        settings_treeview.delete(item_id)

    settings_treeview.bind("<Double-Button-1>", on_double_click)
    settings_treeview.bind("<Button-3>", add_tag)
    settings_treeview.bind("<Delete>", delete_tag)

    save_button = tk.Button(settings_window, text="Save", command=save_settings)
    add_button = tk.Button(settings_window, text="Добавить", command=add_tag)
    delete_button = tk.Button(settings_window, text="Удалить", command=delete_tag)
    save_button.pack(side=tk.LEFT)
    add_button.pack(side=tk.LEFT)
    delete_button.pack(side=tk.LEFT)

    settings_window.mainloop()


root = tk.Tk()
root.title("Excel Extractor")

tk.Label(root, text="Enter the file name:").pack()
file_entry = tk.Entry(root)
file_entry.pack()
file_entry.insert(0, "")  # Предустановленное имя файла

button_extract = tk.Button(root, text="Extract", command=extract_data)
button_extract.pack()

button_settings = tk.Button(root, text="Settings", command=settings)
button_settings.pack()

root.mainloop()