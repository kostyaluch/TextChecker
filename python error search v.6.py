import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel
import pandas as pd
import threading
import os
import time
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

RULES_FILE = 'translation_rules.txt'
IGNORE_FILE = 'ignore_rules.txt'
BLACKLIST_FILE = 'blacklist_terms.txt'

DEFAULT_RULES = """# Формат: російський_корінь = українські_корені_через_кому
# Наприклад:
питани = харчуванн
лук = цибул
настройк = настройк, настроюванн
вид = вигляд
пользовател = користувач
"""

DEFAULT_IGNORES = """# Додайте сюди фрази (українською), які є ПРАВИЛЬНИМИ винятками.
# Якщо програма знайде цю фразу у реченні, вона проігнорує помилку.
# Кожна фраза з нового рядка:
збалансоване харчування
дитяче харчування
здорове харчування
зелена цибуля
ріпчаста цибуля
"""

DEFAULT_BLACKLIST = """# Додайте сюди слова та фрази (будь-якою мовою), які є ПІДОЗРІЛИМИ або ЗАБОРОНЕНІ.
# Якщо програма знайде цю фразу в українському описі, вона позначить рядок як помилковий.
# Кожна фраза з нового рядка:
www.
http://
@gmail.
e-mail:
копія
репліка
"""

class TextEditor(Toplevel):
    def __init__(self, parent, title, file_path, default_content):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title(title)
        self.geometry("650x450")
        self.file_path = file_path
        
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill='both', expand=True)

        # ФІКС: Спочатку розміщуємо панель з кнопками внизу
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', side='bottom', pady=(10, 0))

        self.btn_save = ttk.Button(btn_frame, text="Зберегти", command=self.save_and_close)
        self.btn_save.pack(side='right', padx=10)
        
        self.btn_cancel = ttk.Button(btn_frame, text="Скасувати", command=self.destroy)
        self.btn_cancel.pack(side='right')

        # ФІКС: Тепер розміщуємо текстове поле, воно займе лише місце, що залишилося
        self.txt_content = tk.Text(main_frame, wrap='word', font=('Consolas', 11))
        self.txt_content.pack(fill='both', expand=True, side='top')
        
        # Створюємо файл, якщо його не існує
        if not os.path.exists(file_path):
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(default_content)
                
        # Завантажуємо текст
        with open(file_path, 'r', encoding='utf-8') as f:
            self.txt_content.insert('1.0', f.read())

    def save_and_close(self):
        content = self.txt_content.get('1.0', tk.END).strip()
        with open(self.file_path, 'w', encoding='utf-8') as f:
            f.write(content + "\n")
        self.destroy()


class SpellCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Аналізатор Перекладу Описів v8")
        self.root.geometry("850x600")
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')

        self.file_path = ""
        self.create_widgets()

        # Ініціалізуємо файли при старті
        if not os.path.exists(RULES_FILE):
            with open(RULES_FILE, 'w', encoding='utf-8') as f: f.write(DEFAULT_RULES)
        if not os.path.exists(IGNORE_FILE):
            with open(IGNORE_FILE, 'w', encoding='utf-8') as f: f.write(DEFAULT_IGNORES)
        if not os.path.exists(BLACKLIST_FILE):
            with open(BLACKLIST_FILE, 'w', encoding='utf-8') as f: f.write(DEFAULT_BLACKLIST)

    def create_widgets(self):
        file_frame = ttk.LabelFrame(self.root, text="1. Вибір файлу та колонок", padding="10")
        file_frame.pack(fill='x', padx=10, pady=5)

        file_select_row = ttk.Frame(file_frame)
        file_select_row.pack(fill='x', pady=5)
        
        self.btn_select_file = ttk.Button(file_select_row, text="📂 Обрати Excel-файл", command=self.select_file)
        self.btn_select_file.pack(side='left', padx=5)

        self.lbl_file_status = ttk.Label(file_select_row, text="Файл не обрано", foreground="gray")
        self.lbl_file_status.pack(side='left', padx=5, fill='x', expand=True)

        columns_row = ttk.Frame(file_frame)
        columns_row.pack(fill='x', pady=10)

        ttk.Label(columns_row, text="Стовпчик RU:", font=('Arial', 9, 'bold')).pack(side='left', padx=5)
        self.combo_ru = ttk.Combobox(columns_row, state="disabled", width=25)
        self.combo_ru.pack(side='left', padx=5)

        ttk.Label(columns_row, text="Стовпчик UA:", font=('Arial', 9, 'bold')).pack(side='left', padx=15)
        self.combo_ua = ttk.Combobox(columns_row, state="disabled", width=25)
        self.combo_ua.pack(side='left', padx=5)

        control_frame = ttk.LabelFrame(self.root, text="2. Керування та Словники", padding="10")
        control_frame.pack(fill='x', padx=10, pady=5)

        self.btn_start_analysis = ttk.Button(control_frame, text="▶ Почати перевірку", command=self.start_analysis, state='disabled')
        self.btn_start_analysis.pack(side='left', padx=5, fill='x', expand=True)

        self.btn_edit_rules = ttk.Button(control_frame, text="✎ Словник помилок", command=lambda: self.open_editor("Словник помилок", RULES_FILE, DEFAULT_RULES))
        self.btn_edit_rules.pack(side='left', padx=5)

        self.btn_edit_ignores = ttk.Button(control_frame, text="🚫 Словник винятків", command=lambda: self.open_editor("Словник винятків", IGNORE_FILE, DEFAULT_IGNORES))
        self.btn_edit_ignores.pack(side='left', padx=5)

        self.btn_edit_blacklist = ttk.Button(control_frame, text="⛔ Чорний список", command=lambda: self.open_editor("Чорний список термінів", BLACKLIST_FILE, DEFAULT_BLACKLIST))
        self.btn_edit_blacklist.pack(side='left', padx=5)

        progress_frame = ttk.Frame(self.root, padding="5")
        progress_frame.pack(fill='x', padx=10, pady=5)

        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate', length=100)
        self.progress_bar.pack(fill='x', expand=True)

        self.lbl_progress_status = ttk.Label(progress_frame, text="Очікування файлу...", anchor='center')
        self.lbl_progress_status.pack(fill='x', pady=2)

        self.txt_log = tk.Text(self.root, height=10, state='disabled', wrap='word', bg="#f0f0f0", font=('Consolas', 9))
        self.txt_log.pack(fill='both', expand=True, padx=10, pady=10)
        
        log_scroll = ttk.Scrollbar(self.txt_log, orient='vertical', command=self.txt_log.yview)
        self.txt_log['yscrollcommand'] = log_scroll.set
        log_scroll.pack(side='right', fill='y')

    def log_message(self, message):
        self.txt_log.config(state='normal')
        timestamp = time.strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] {message}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    def open_editor(self, title, file_path, default_content):
        editor = TextEditor(self.root, title, file_path, default_content)
        self.root.wait_window(editor)
        self.log_message(f"Файл '{title}' оновлено.")

    def select_file(self):
        path = filedialog.askopenfilename(title="Оберіть файл Excel", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        if path:
            try:
                df_preview = pd.read_excel(path, nrows=0)
                columns = df_preview.columns.tolist()
                self.file_path = path
                self.lbl_file_status.config(text=os.path.basename(path), foreground="green")
                self.combo_ru.config(values=columns, state="readonly")
                self.combo_ua.config(values=columns, state="readonly")
                ru_col = next((c for c in columns if 'ru' in c.lower() or 'рус' in c.lower()), '')
                ua_col = next((c for c in columns if 'ua' in c.lower() or 'uk' in c.lower() or 'укр' in c.lower()), '')
                if ru_col: self.combo_ru.set(ru_col)
                if ua_col: self.combo_ua.set(ua_col)
                self.log_message("Файл завантажено. Можна починати перевірку.")
                self.btn_start_analysis.config(state='normal')
            except Exception as e:
                messagebox.showerror("Помилка файлу", f"Не вдалося зчитати файл:\n{e}")

    def set_ui_state(self, is_running):
        state = 'disabled' if is_running else 'normal'
        self.btn_select_file.config(state=state)
        self.btn_edit_rules.config(state=state)
        self.btn_edit_ignores.config(state=state)
        self.btn_edit_blacklist.config(state=state)
        self.btn_start_analysis.config(state=state)
        self.combo_ru.config(state="disabled" if is_running else "readonly")
        self.combo_ua.config(state="disabled" if is_running else "readonly")

    def parse_rules(self):
        rules = {}
        with open(RULES_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.split('#')[0].strip()
                if '=' in line:
                    ru, ua_list = line.split('=', 1)
                    ua_stems = [ua.strip() for ua in ua_list.split(',')]
                    rules[ru.strip()] = ua_stems
        return rules

    def parse_ignores(self):
        ignores = []
        with open(IGNORE_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.split('#')[0].strip()
                if line:
                    ignores.append(line.lower())
        return ignores

    def parse_blacklist(self):
        blacklist = []
        with open(BLACKLIST_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.split('#')[0].strip()
                if line:
                    blacklist.append(line.lower())
        return blacklist

    def extract_error_sentence(self, text, error_word_stem):
        text_clean = re.sub(r'(?i)(</li>|<br\s*/?>|</p>|</div>)', '. ', text)
        text_clean = re.sub(r'<[^>]+>', ' ', text_clean)
        text_clean = re.sub(r'\s+', ' ', text_clean).strip()
        
        sentences = re.split(r'(?<=[.!?\n])\s+', text_clean)
        for s in sentences:
            # Додано \b для точного пошуку початку слова
            if re.search(r'(?i)\b' + re.escape(error_word_stem), s):
                return s.strip(' .-;:,\t')
                
        # Додано \b у резервний пошук
        match = re.search(r'(.{0,50}\b' + re.escape(error_word_stem) + r'.{0,50})', text_clean, re.IGNORECASE)
        if match:
            return "..." + match.group(1).strip() + "..."
        return ""

    def run_analysis(self, save_path, col_ru, col_ua):
        try:
            self.log_message("Початок перевірки...")
            self.lbl_progress_status.config(text="Зчитування даних...")
            
            rules = self.parse_rules()
            ignores = self.parse_ignores()
            blacklist = self.parse_blacklist()
            
            df = pd.read_excel(self.file_path)
            total_rows = len(df)
            self.progress_bar.config(maximum=total_rows, value=0)
            
            errors_column = []
            rows_with_errors = set()

            for i, row in df.iterrows():
                text_ru = str(row.get(col_ru, ''))
                text_ua = str(row.get(col_ua, ''))
                row_errors = []

                # Перевірка чорного списку термінів
                text_ua_lower = text_ua.lower()
                for blacklisted_term in blacklist:
                    if blacklisted_term in text_ua_lower:
                        # Знайти контекст терміну для відображення
                        idx = text_ua_lower.find(blacklisted_term)
                        start = max(0, idx - 30)
                        end = min(len(text_ua), idx + len(blacklisted_term) + 30)
                        context = text_ua[start:end].strip()
                        if start > 0:
                            context = "..." + context
                        if end < len(text_ua):
                            context = context + "..."
                        row_errors.append(f"[ЗАБОРОНЕНИЙ ТЕРМІН] Знайдено '{blacklisted_term}' у тексті: \"{context}\"")

                for ru_stem, ua_stems in rules.items():
                    # ВИПРАВЛЕНО: тепер шукає тільки з початку слова (\b), щоб уникнути "фо-ток-ореспондент"
                    match_ru = re.search(r'(?i)\b' + re.escape(ru_stem) + r'[\w-]*\b', text_ru)
                    
                    if match_ru:
                        ru_full_word = match_ru.group(0)
                        
                        for ua_stem in ua_stems:
                            # ВИПРАВЛЕНО: додано \b для українського слова
                            if re.search(r'(?i)\b' + re.escape(ua_stem), text_ua):
                                error_sentence = self.extract_error_sentence(text_ua, ua_stem)
                                
                                is_ignored = False
                                for ig in ignores:
                                    if ig in error_sentence.lower():
                                        is_ignored = True
                                        break
                                
                                if not is_ignored and error_sentence:
                                    row_errors.append(f"[ПОМИЛКА] Значення '{ru_full_word}' хибно перекладено. Речення: \"{error_sentence}\"")

                errors_column.append("\n".join(row_errors))
                if row_errors:
                    rows_with_errors.add(i)

                if i % 10 == 0:
                    self.progress_bar['value'] = i + 1
                    self.lbl_progress_status.config(text=f"Обробка: {i+1}/{total_rows}")

            df['Помилки перекладу'] = errors_column
            
            self.lbl_progress_status.config(text="Збереження файлу...")
            wb = Workbook()
            ws = wb.active
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)

            fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            for row_idx in rows_with_errors:
                for cell in ws[row_idx + 2]:  
                    cell.fill = fill

            wb.save(save_path)
            self.progress_bar['value'] = total_rows
            self.lbl_progress_status.config(text="Готово!")
            self.log_message(f"Завершено. Знайдено помилок у {len(rows_with_errors)} рядках.")
            self.root.after(0, lambda: messagebox.showinfo("Успіх", f"Файл збережено:\n{save_path}"))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Помилка", str(e)))
        finally:
            self.root.after(0, lambda: self.set_ui_state(False))

    def start_analysis(self):
        col_ru = self.combo_ru.get()
        col_ua = self.combo_ua.get()
        if not col_ru or not col_ua:
            messagebox.showwarning("Увага", "Оберіть стовпчики для обох мов!")
            return
        self.set_ui_state(True)
        base, ext = os.path.splitext(self.file_path)
        save_path = f"{base}_checked{ext}"
        threading.Thread(target=self.run_analysis, args=(save_path, col_ru, col_ua), daemon=True).start()

if __name__ == "__main__":
    main_root = tk.Tk()
    app = SpellCheckerApp(main_root)
    main_root.mainloop()