import os
import re
import xlsxwriter
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog, scrolledtext

current_file = ''

# Цветовая схема
COLORS = {
    'bg': '#2b2b2b',
    'fg': '#ffffff',
    'accent': '#4CAF50',
    'accent_hover': '#45a049',
    'secondary': '#2196F3',
    'secondary_hover': '#0b7dda',
    'danger': '#f44336',
    'danger_hover': '#da190b',
    'card_bg': '#3c3f41',
    'entry_bg': '#515151',
    'text_light': '#cccccc'
}

# Глобальная переменная для текстового поля вывода
output_text = None


def setup_styles():
    style = ttk.Style()

    # Современная тема
    style.theme_use('clam')

    # Настройка стилей для кнопок
    style.configure('Accent.TButton',
                    background=COLORS['accent'],
                    foreground=COLORS['fg'],
                    borderwidth=0,
                    focuscolor='none',
                    font=('Arial', 10, 'bold'))

    style.map('Accent.TButton',
              background=[('active', COLORS['accent_hover']),
                          ('pressed', COLORS['accent_hover'])])

    style.configure('Secondary.TButton',
                    background=COLORS['secondary'],
                    foreground=COLORS['fg'],
                    borderwidth=0,
                    focuscolor='none',
                    font=('Arial', 10))

    style.map('Secondary.TButton',
              background=[('active', COLORS['secondary_hover']),
                          ('pressed', COLORS['secondary_hover'])])

    style.configure('Danger.TButton',
                    background=COLORS['danger'],
                    foreground=COLORS['fg'],
                    borderwidth=0,
                    focuscolor='none',
                    font=('Arial', 10))

    style.map('Danger.TButton',
              background=[('active', COLORS['danger_hover']),
                          ('pressed', COLORS['danger_hover'])])

    # Стиль для Treeview
    style.configure('Custom.Treeview',
                    background=COLORS['card_bg'],
                    foreground=COLORS['fg'],
                    fieldbackground=COLORS['card_bg'],
                    borderwidth=0)

    style.configure('Custom.Treeview.Heading',
                    background=COLORS['secondary'],
                    foreground=COLORS['fg'],
                    borderwidth=0,
                    font=('Arial', 11, 'bold'))

    style.map('Custom.Treeview.Heading',
              background=[('active', COLORS['secondary_hover'])])

    # Стиль для вкладок
    style.configure('Custom.TNotebook',
                    background=COLORS['bg'],
                    borderwidth=0)

    style.configure('Custom.TNotebook.Tab',
                    background=COLORS['card_bg'],
                    foreground=COLORS['text_light'],
                    borderwidth=0,
                    padding=[15, 5],
                    font=('Arial', 10))

    style.map('Custom.TNotebook.Tab',
              background=[('selected', COLORS['accent']),
                          ('active', COLORS['accent_hover'])],
              foreground=[('selected', COLORS['fg'])])


def create_modern_button(parent, text, command, style='Accent.TButton', width=15):
    return ttk.Button(parent, text=text, command=command, style=style, width=width)


def log_message(message):
    """Вывод сообщения в текстовое поле вместо консоли"""
    if output_text:
        output_text.insert(tk.END, message + "\n")
        output_text.see(tk.END)  # Автопрокрутка к новому сообщению
        output_text.update_idletasks()


def find_line_in_file(file_name, TAG_value):
    if not os.path.exists(file_name):
        log_message(f"File {file_name} does not exist")
        return False

    tag_pattern = f"{TAG_value}="
    counter_pattern = "instances_counter="

    with open(file_name, 'r', encoding='utf-8') as file:
        prev_line_has_tag = False

        for line in file:
            if prev_line_has_tag and counter_pattern in line:
                log_message(f"Success: {TAG_value} line followed by instances_counter")
                return True

            prev_line_has_tag = tag_pattern in line

    return False


def extract_data():
    if not os.path.exists(current_file):
        messagebox.showerror("Error", f"File {current_file} not found!")
        return

    file_name = 'Settings.txt'
    if not os.path.exists(file_name):
        messagebox.showerror("Error", "Settings.txt not found!")
        return

    try:
        with open(file_name, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            entries = []
            for line in lines:
                if ':' in line and ';' in line and '@' in line and '#' in line:
                    parts = line.strip().split(':')[1].split(';')
                    tag = parts[0]
                    remaining = parts[1].split('@')
                    name = remaining[0]
                    extract_flag = remaining[1].split('#')[0]
                    entries.append((tag, name, extract_flag))

            TAG_values = [entry[0] for entry in entries if entry[2] == '1']
            NAME_values = [entry[1] for entry in entries if entry[2] == '1']
            TAG_values = [value.upper() for value in TAG_values]

        if not TAG_values:
            messagebox.showwarning("Warning", "No tags selected for extraction!")
            return

        # Очистка предыдущих сообщений
        if output_text:
            output_text.delete(1.0, tk.END)

        log_message("Starting data extraction...")
        log_message(f"Processing {len(TAG_values)} tags")

        # Показать прогресс
        progress_window = tk.Toplevel()
        progress_window.title("Extracting Data")
        progress_window.geometry("300x100")
        progress_window.configure(bg=COLORS['bg'])
        progress_window.transient(root)
        progress_window.grab_set()

        tk.Label(progress_window, text="Extracting data...",
                 bg=COLORS['bg'], fg=COLORS['fg'], font=('Arial', 11)).pack(pady=10)

        progress = ttk.Progressbar(progress_window, orient='horizontal',
                                   length=250, mode='determinate')
        progress.pack(pady=10)
        progress_window.update()

        workbook = xlsxwriter.Workbook('output.xlsx')
        worksheet = workbook.add_worksheet()

        headers = [
            'Tag Name', 'Name', 'GDP', 'GDP p/c',
            'Civilian factories', 'Military factories',
            'Naval factories', 'Equipment operational cost',
            'Productivity', 'Debt', 'Interest rate', 'Globalism',
            'Investments', 'Bonds held'
        ]

        # Форматирование заголовков с переносом текста
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4CAF50',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True  # Включаем перенос текста
        })

        # Устанавливаем ширину колонок (можно настроить под ваши нужды)
        column_widths = {
            'Tag Name': 8,
            'Name': 12,
            'GDP': 10,
            'GDP p/c': 8,
            'Civilian factories': 10,
            'Military factories': 10,
            'Naval factories': 10,
            'Equipment operational cost': 12,
            'Productivity': 10,
            'Debt': 8,
            'Interest rate': 10,
            'Globalism': 8,
            'Investments': 10,
            'Bonds held': 10
        }

        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
            # Устанавливаем ширину колонок
            width = column_widths.get(header, 10)
            worksheet.set_column(col, col, width)

        # Устанавливаем высоту строки для заголовков (опционально)
        worksheet.set_row(0, 30)  # Высота 30 пунктов для заголовков

        row = 1
        total_tags = len(TAG_values)

        for i, TAG_value in enumerate(TAG_values):
            progress['value'] = (i / total_tags) * 100
            progress_window.update()

            log_message(f"Processing tag: {TAG_value}")
            find_line_in_file(current_file, TAG_value)

            value_debt = value_interest_rate = 0
            value_gdp = value_gdp_per_capita = value_effective_civilian_factories = value_effective_military_factories = \
                value_effective_naval_factories = value_equipment_operative_cost = value_overall_productivity = \
                value_globalism = value_int_investments = value_bonds_held = None

            with open(current_file, 'r', encoding='utf-8') as file:
                lines = file.readlines()
                for line_num, line in enumerate(lines, 1):
                    if re.search(rf"{TAG_value}=", line):
                        if line_num < len(lines) - 1 and (
                                re.search(r"instances_counter=\d+", lines[line_num]) or re.search(
                            r"gdpc_converging_var=\d+", lines[line_num])):
                            for next_line_num, next_line in enumerate(lines[line_num + 1:], line_num + 2):
                                if re.search(r"debt=([0-9.,]+)", next_line):
                                    value_debt = float(re.search(r"debt=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"imp_total_bonds_held=([0-9.,]+)", next_line):
                                    value_bonds_held = float(
                                        re.search(r"imp_total_bonds_held=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"int_investments=([0-9.,]+)", next_line):
                                    value_int_investments = float(
                                        re.search(r"int_investments=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"interest_rate=([0-9.,]+)", next_line):
                                    value_interest_rate = float(
                                        re.search(r"interest_rate=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"effective_civilian_factories=([0-9.,]+)", next_line):
                                    value_effective_civilian_factories = float(
                                        re.search(r"effective_civilian_factories=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"effective_military_factories=([0-9.,]+)", next_line):
                                    value_effective_military_factories = float(
                                        re.search(r"effective_military_factories=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"effective_naval_factories=([0-9.,]+)", next_line):
                                    value_effective_naval_factories = float(
                                        re.search(r"effective_naval_factories=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"equipment_operative_cost=([0-9.,]+)", next_line):
                                    value_equipment_operative_cost = float(
                                        re.search(r"equipment_operative_cost=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"gdp_per_capita=([0-9.,]+)", next_line):
                                    value_gdp_per_capita = float(
                                        re.search(r"gdp_per_capita=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"gdp_total=([0-9.,]+)", next_line):
                                    value_gdp = float(re.search(r"gdp_total=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"globalism_value=([0-9.,]+)", next_line):
                                    value_globalism = float(
                                        re.search(r"globalism_value=([0-9.,]+)", next_line).group(1))
                                elif re.search(r"overall_productivity=([0-9.,]+)", next_line):
                                    value_overall_productivity = float(
                                        re.search(r"overall_productivity=([0-9.,]+)", next_line).group(1))
                                    break

            if value_gdp is not None:
                worksheet.write(row, 0, TAG_value)
                worksheet.write(row, 1, NAME_values[i])
                worksheet.write(row, 2, round(value_gdp) if value_gdp is not None else 0)
                worksheet.write(row, 3, round(value_gdp_per_capita) if value_gdp_per_capita is not None else 0)
                worksheet.write(row, 4,
                                round(
                                    value_effective_civilian_factories) if value_effective_civilian_factories is not None else 0)
                worksheet.write(row, 5,
                                round(
                                    value_effective_military_factories) if value_effective_military_factories is not None else 0)
                worksheet.write(row, 6,
                                round(
                                    value_effective_naval_factories) if value_effective_naval_factories is not None else 0)
                worksheet.write(row, 7,
                                round(value_equipment_operative_cost,
                                      2) if value_equipment_operative_cost is not None else 0)
                worksheet.write(row, 8,
                                round(value_overall_productivity) if value_overall_productivity is not None else 0)
                worksheet.write(row, 9, round(value_debt) if value_debt is not None else 0)
                worksheet.write(row, 10, round(value_interest_rate, 1) if value_interest_rate is not None else 0)
                worksheet.write(row, 11, value_globalism if value_globalism is not None else 0)
                worksheet.write(row, 12, round(value_int_investments) if value_int_investments is not None else 0)
                worksheet.write(row, 13, round(value_bonds_held) if value_bonds_held is not None else 0)
                row += 1
                log_message(f"✓ Data extracted for {TAG_value}")
            else:
                log_message(f"✗ No data found for {TAG_value}")

        workbook.close()
        progress_window.destroy()

        log_message(f"\nExtraction completed!")
        log_message(f"Processed {row - 1} countries")
        log_message("Data saved to output.xlsx")

        messagebox.showinfo("Extraction Complete",
                            f"Data successfully extracted to output.xlsx\n\nProcessed {row - 1} countries")

    except Exception as e:
        log_message(f"ERROR: {str(e)}")
        messagebox.showerror("Error", f"An error occurred during extraction:\n{str(e)}")

def choose_file():
    global current_file
    file_path = filedialog.askopenfilename(
        title="Select HOI4 File",
        filetypes=[("HOI4 Files", "*.hoi4"), ("All Files", "*.*")]
    )
    if file_path:
        current_file = file_path
        file_label.config(text=f"Selected: {os.path.basename(current_file)}")
        log_message(f"Selected file: {current_file}")


def save_settings_to_file(tree):
    try:
        with open('Settings.txt', 'w', encoding='utf-8') as file:
            for item in tree.get_children():
                values = tree.item(item)['values']
                tag = values[0]
                name = values[1]
                extract_flag = '1' if "✓" in values[2] else '0'
                file.write(f":{tag};{name}@{extract_flag}#\n")
        log_message("Settings saved successfully")
    except Exception as e:
        log_message(f"ERROR saving settings: {str(e)}")
        messagebox.showerror("Error", f"Failed to save settings:\n{str(e)}")


def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children()]

    if col == "Extract":
        l.sort(key=lambda t: t[0] == "✓", reverse=reverse)
    else:
        try:
            l.sort(reverse=reverse)
        except:
            l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)

    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

    for c in tv["columns"]:
        if c == col:
            tv.heading(c, text=f"{c} {'↑' if reverse else '↓'}")
        else:
            tv.heading(c, text=c)


def settings():
    settings_window = tk.Toplevel()
    settings_window.title("Tag Settings")
    settings_window.geometry("600x600")
    settings_window.configure(bg=COLORS['bg'])
    settings_window.transient(root)
    settings_window.grab_set()

    # Header
    header_frame = tk.Frame(settings_window, bg=COLORS['secondary'], height=60)
    header_frame.pack(fill='x', padx=10, pady=10)
    header_frame.pack_propagate(False)

    tk.Label(header_frame, text="Tag Management",
             bg=COLORS['secondary'], fg=COLORS['fg'],
             font=('Arial', 16, 'bold')).pack(expand=True)

    # Main content
    content_frame = tk.Frame(settings_window, bg=COLORS['bg'])
    content_frame.pack(fill='both', expand=True, padx=10, pady=10)

    # Create treeview with custom style
    tree_frame = tk.Frame(content_frame, bg=COLORS['bg'])
    tree_frame.pack(fill='both', expand=True)

    tree = ttk.Treeview(tree_frame, style="Custom.Treeview")
    tree["columns"] = ("Tag", "Name", "Extract")
    tree.column("#0", width=0)
    tree.column("Tag", width=120, anchor="w")
    tree.column("Name", width=180, anchor="w")
    tree.column("Extract", width=80, anchor="center")

    for col in ("Tag", "Name", "Extract"):
        tree.heading(col, text=col, anchor="center",
                     command=lambda c=col: treeview_sort_column(tree, c, False))

    tree.tag_configure('checked', foreground='#4CAF50', font=('Arial', 12, 'bold'))
    tree.tag_configure('unchecked', foreground='#f44336', font=('Arial', 12, 'bold'))

    # Scrollbar for treeview
    tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll.set)

    tree.pack(side='left', fill='both', expand=True)
    tree_scroll.pack(side='right', fill='y')

    # Button frame
    button_frame = tk.Frame(content_frame, bg=COLORS['bg'])
    button_frame.pack(fill='x', pady=10)

    create_modern_button(button_frame, "Add Tag", lambda: add_tag(tree),
                         style='Accent.TButton', width=12).pack(side='left', padx=5)
    create_modern_button(button_frame, "Delete Selected", lambda: delete_tag(tree),
                         style='Danger.TButton', width=15).pack(side='left', padx=5)
    create_modern_button(button_frame, "Select All", lambda: select_all(tree),
                         style='Secondary.TButton', width=12).pack(side='left', padx=5)
    create_modern_button(button_frame, "Deselect All", lambda: deselect_all(tree),
                         style='Secondary.TButton', width=12).pack(side='left', padx=5)
    create_modern_button(button_frame, "Close", settings_window.destroy,
                         style='Secondary.TButton', width=10).pack(side='right', padx=5)

    # Load settings
    load_settings(tree)

    # Bind events
    tree.bind("<Button-1>", lambda e: on_cell_click(e, tree))
    tree.bind("<Double-1>", lambda e: on_double_click(e, tree))
    tree.bind("<Delete>", lambda e: delete_tag(tree))


def load_settings(tree):
    if os.path.exists('Settings.txt'):
        try:
            with open('Settings.txt', 'r', encoding='utf-8') as file:
                lines = file.readlines()
                for line in lines:
                    if ':' in line and ';' in line and '@' in line and '#' in line:
                        parts = line.strip().split(':')[1].split(';')
                        tag = parts[0]
                        remaining = parts[1].split('@')
                        name = remaining[0]
                        extract_flag = remaining[1].split('#')[0]

                        if extract_flag == '1':
                            tree.insert('', 'end', values=(tag, name, "✓"), tags=('checked',))
                        else:
                            tree.insert('', 'end', values=(tag, name, "✗"), tags=('unchecked',))
            log_message("Settings loaded successfully")
        except Exception as e:
            log_message(f"ERROR loading settings: {str(e)}")
            messagebox.showerror("Error", f"Failed to load settings:\n{str(e)}")


def on_cell_click(event, tree):
    region = tree.identify_region(event.x, event.y)
    if region == "cell":
        column = tree.identify_column(event.x)
        item = tree.identify_row(event.y)

        if column == '#3':  # Extract column
            values = tree.item(item)['values']
            current_state = "✓" in values[2]

            if current_state:
                tree.item(item, values=(values[0], values[1], "✗"), tags=('unchecked',))
            else:
                tree.item(item, values=(values[0], values[1], "✓"), tags=('checked',))

            save_settings_to_file(tree)


def on_double_click(event, tree):
    item_id = tree.identify_row(event.y)
    column = tree.identify_column(event.x)

    if column in ('#1', '#2') and item_id:
        item_values = tree.item(item_id)['values']
        if column == '#1':  # Tag column
            new_value = simpledialog.askstring("Edit Tag", "Enter new tag:", initialvalue=item_values[0])
            if new_value is not None and new_value.strip():
                tree.item(item_id, values=(new_value.strip(), item_values[1], item_values[2]),
                          tags=('checked' if "✓" in item_values[2] else 'unchecked'))
                save_settings_to_file(tree)
        elif column == '#2':  # Name column
            new_value = simpledialog.askstring("Edit Name", "Enter new name:", initialvalue=item_values[1])
            if new_value is not None and new_value.strip():
                tree.item(item_id, values=(item_values[0], new_value.strip(), item_values[2]),
                          tags=('checked' if "✓" in item_values[2] else 'unchecked'))
                save_settings_to_file(tree)


def add_tag(tree):
    new_tag = simpledialog.askstring("New Tag", "Enter new tag:")
    if new_tag and new_tag.strip():
        new_name = simpledialog.askstring("New Name", "Enter new name:")
        if new_name and new_name.strip():
            tree.insert('', 'end', values=(new_tag.strip(), new_name.strip(), "✓"), tags=('checked',))
            log_message(f"Added new tag: {new_tag.strip()} - {new_name.strip()}")
            save_settings_to_file(tree)


def delete_tag(tree):
    selected_items = tree.selection()
    if selected_items:
        if messagebox.askyesno("Confirm Delete", f"Delete {len(selected_items)} selected item(s)?"):
            for item in selected_items:
                values = tree.item(item)['values']
                log_message(f"Deleted tag: {values[0]} - {values[1]}")
                tree.delete(item)
            save_settings_to_file(tree)


def select_all(tree):
    for item in tree.get_children():
        values = tree.item(item)['values']
        if "✗" in values[2]:
            tree.item(item, values=(values[0], values[1], "✓"), tags=('checked',))
    log_message("All tags selected for extraction")
    save_settings_to_file(tree)


def deselect_all(tree):
    for item in tree.get_children():
        values = tree.item(item)['values']
        if "✓" in values[2]:
            tree.item(item, values=(values[0], values[1], "✗"), tags=('unchecked',))
    log_message("All tags deselected for extraction")
    save_settings_to_file(tree)


def clear_output():
    """Очистка текстового поля вывода"""
    if output_text:
        output_text.delete(1.0, tk.END)


# Создание главного окна
root = tk.Tk()
root.title("HOI4 Data Extractor")
root.geometry("600x700")  # Увеличил высоту для текстового поля
root.configure(bg=COLORS['bg'])

# Настройка стилей
setup_styles()

# Header
header_frame = tk.Frame(root, bg=COLORS['secondary'], height=80)
header_frame.pack(fill='x', padx=15, pady=15)
header_frame.pack_propagate(False)

tk.Label(header_frame, text="HOI4 Data Extractor",
         bg=COLORS['secondary'], fg=COLORS['fg'],
         font=('Arial', 20, 'bold')).pack(expand=True)

# Main content
main_frame = tk.Frame(root, bg=COLORS['bg'])
main_frame.pack(fill='both', expand=True, padx=20, pady=20)

# File selection section
file_frame = tk.Frame(main_frame, bg=COLORS['card_bg'], relief='ridge', bd=1)
file_frame.pack(fill='x', pady=10)

tk.Label(file_frame, text="Current File:",
         bg=COLORS['card_bg'], fg=COLORS['fg'],
         font=('Arial', 11, 'bold')).pack(anchor='w', padx=10, pady=5)

file_label = tk.Label(file_frame, text=f"Selected: {os.path.basename(current_file)}",
                      bg=COLORS['card_bg'], fg=COLORS['text_light'],
                      font=('Arial', 10))
file_label.pack(anchor='w', padx=10, pady=(0, 10))

create_modern_button(file_frame, "Change File", choose_file,
                     style='Secondary.TButton', width=15).pack(pady=10)

# Buttons section
buttons_frame = tk.Frame(main_frame, bg=COLORS['bg'])
buttons_frame.pack(fill='x', pady=10)

create_modern_button(buttons_frame, "Extract Data", extract_data,
                     style='Accent.TButton', width=20).pack(pady=5)
create_modern_button(buttons_frame, "Tag Settings", settings,
                     style='Secondary.TButton', width=20).pack(pady=5)

# Output section
output_frame = tk.Frame(main_frame, bg=COLORS['bg'])
output_frame.pack(fill='both', expand=True, pady=10)

tk.Label(output_frame, text="Output Log:",
         bg=COLORS['bg'], fg=COLORS['fg'],
         font=('Arial', 11, 'bold')).pack(anchor='w')

# Текстовое поле для вывода
output_text_frame = tk.Frame(output_frame, bg=COLORS['card_bg'])
output_text_frame.pack(fill='both', expand=True, pady=5)

output_text = scrolledtext.ScrolledText(
    output_text_frame,
    wrap=tk.WORD,
    bg=COLORS['entry_bg'],
    fg=COLORS['fg'],
    font=('Consolas', 10),
    height=10
)
output_text.pack(fill='both', expand=True, padx=2, pady=2)

# Кнопка очистки вывода
create_modern_button(output_frame, "Clear Output", clear_output,
                     style='Secondary.TButton', width=12).pack(anchor='e', pady=5)

# Status bar
status_frame = tk.Frame(root, bg=COLORS['card_bg'], height=30)
status_frame.pack(fill='x', side='bottom')
status_frame.pack_propagate(False)

status_label = tk.Label(status_frame, text="Ready",
                        bg=COLORS['card_bg'], fg=COLORS['text_light'],
                        font=('Arial', 9))
status_label.pack(side='left', padx=10)

# Приветственное сообщение
log_message("HOI4 Data Extractor started")
log_message("Select a file and configure tags to begin")

root.mainloop()
