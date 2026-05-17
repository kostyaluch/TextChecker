import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel
import pandas as pd
import threading
import os
import time
import re
from html import unescape
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
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
# Також програма автоматично видалить речення або пункт списку, де знайдено заборонений термін.
# Кожна фраза з нового рядка:
www.
http://
@gmail.
e-mail:
копія
репліка
"""

UKRAINIAN_MARKERS = {
    'і', 'та', 'це', 'для', 'що', 'від', 'під', 'при', 'через', 'згідно', 'наявність',
    'зручність', 'технічні', 'обслуговування', 'попередження', 'увага', 'примітка', 'модель'
}
RUSSIAN_MARKERS = {
    'и', 'или', 'это', 'для', 'что', 'при', 'через', 'согласно', 'наличие',
    'удобство', 'технические', 'обслуживание', 'внимание', 'примечание', 'модель'
}
UKRAINIAN_UNIQUE_CHARS = set('іїєґ')
RUSSIAN_UNIQUE_CHARS = set('ыэёъ')
CHINESE_CHAR_PATTERN = re.compile(r'[\u4e00-\u9fff]')
TECHNICAL_HTML_PATTERNS = [
    r'(?is)<style\b[^>]*>.*?</style>',
    r'(?is)<script\b[^>]*>.*?</script>',
    r'(?is)<div\b[^>]*class="[^"]*ssd-module-wrap[^"]*"[^>]*>.*?</div>',
    r'(?is)<div\b[^>]*class=["\']?ssd-module[^>]*>.*?</div>',
    r'(?is)<div\b[^>]*(cssurl|skucode|skudesign|id="zbViewModulesH"|id="zbViewModulesHeight")[^>]*>.*?</div>',
    r'(?is)<input\b[^>]*(zbViewModulesHeight)[^>]*>',
]


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

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', side='bottom', pady=(10, 0))

        self.btn_save = ttk.Button(btn_frame, text="Зберегти", command=self.save_and_close)
        self.btn_save.pack(side='right', padx=10)

        self.btn_cancel = ttk.Button(btn_frame, text="Скасувати", command=self.destroy)
        self.btn_cancel.pack(side='right')

        self.txt_content = tk.Text(main_frame, wrap='word', font=('Consolas', 11))
        self.txt_content.pack(fill='both', expand=True, side='top')

        if not os.path.exists(file_path):
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(default_content)

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
        self.root.title("Аналізатор Перекладу Описів v13")
        self.root.geometry("850x600")
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')

        self.file_path = ""
        self.create_widgets()

        if not os.path.exists(RULES_FILE):
            with open(RULES_FILE, 'w', encoding='utf-8') as f:
                f.write(DEFAULT_RULES)
        if not os.path.exists(IGNORE_FILE):
            with open(IGNORE_FILE, 'w', encoding='utf-8') as f:
                f.write(DEFAULT_IGNORES)
        if not os.path.exists(BLACKLIST_FILE):
            with open(BLACKLIST_FILE, 'w', encoding='utf-8') as f:
                f.write(DEFAULT_BLACKLIST)

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

    def detect_default_columns(self, columns):
        normalized = {col: str(col).strip().lower() for col in columns}
        ru_priority = ['описание;1', 'описание', 'description;1', 'description ru', 'ru', 'рус']
        ua_priority = ['описание (ua);1', 'опис (ua);1', 'описание ua;1', 'описание (ua)', 'description (ua);1', 'description ua', 'ua', 'uk', 'укр']

        def find_by_priority(priority_list):
            for target in priority_list:
                for original, lowered in normalized.items():
                    if lowered == target:
                        return original
            return ''

        ru_col = find_by_priority(ru_priority)
        ua_col = find_by_priority(ua_priority)

        if not ru_col:
            ru_col = next((c for c in columns if any(token in normalized[c] for token in ['описание', 'description']) and '(ua)' not in normalized[c] and 'ua' not in normalized[c] and 'uk' not in normalized[c]), '')
        if not ua_col:
            ua_col = next((c for c in columns if any(token in normalized[c] for token in ['ua', 'uk', 'укр', '(ua)'])), '')

        return ru_col, ua_col

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
                ru_col, ua_col = self.detect_default_columns(columns)
                if ru_col:
                    self.combo_ru.set(ru_col)
                if ua_col:
                    self.combo_ua.set(ua_col)
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

    def normalize_text(self, text):
        if pd.isna(text):
            return ""
        return str(text)

    def strip_technical_html(self, text):
        text = self.normalize_text(text)
        text = unescape(text)
        text = re.sub(r'\[/?html\]', '', text, flags=re.IGNORECASE)
        for pattern in TECHNICAL_HTML_PATTERNS:
            text = re.sub(pattern, ' ', text)
        text = re.sub(r'(?is)<!--.*?-->', ' ', text)
        return text

    def clean_html_text(self, text):
        text = self.strip_technical_html(text)
        text = re.sub(r'(?is)<style\b[^>]*>.*?</style>', ' ', text)
        text = re.sub(r'(?is)<script\b[^>]*>.*?</script>', ' ', text)
        text = re.sub(r'(?i)(</li>|<br\s*/?>|</p>|</div>)', '. ', text)
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'http[s]?://\S+', ' ', text)
        text = re.sub(r'\b[a-z0-9_-]+\s*\{[^{}]*\}', ' ', text, flags=re.IGNORECASE)
        text = re.sub(r'\s+', ' ', text).strip(' .')
        return text

    def contains_chinese(self, text):
        return bool(CHINESE_CHAR_PATTERN.search(self.normalize_text(text)))

    def is_meaningful_text(self, text):
        cleaned = self.clean_html_text(text)
        if not cleaned:
            return False
        if self.contains_chinese(cleaned):
            return False
        words = re.findall(r'[A-Za-zА-Яа-яІіЇїЄєҐґЁёЪъЫыЭэ0-9]+', cleaned)
        if len(words) <= 1:
            return False
        if len(cleaned) < 12:
            return False
        if 'ssd-module-wrap' in cleaned.lower() or 'background-image' in cleaned.lower():
            return False
        return True

    def extract_error_sentence(self, text, error_word_stem):
        text_clean = self.clean_html_text(text)
        sentences = re.split(r'(?<=[.!?\n])\s+', text_clean)
        for s in sentences:
            if re.search(r'(?i)\b' + re.escape(error_word_stem), s):
                return s.strip(' .-;:,\t')
        match = re.search(r'(.{0,50}\b' + re.escape(error_word_stem) + r'.{0,50})', text_clean, re.IGNORECASE)
        if match:
            return "..." + match.group(1).strip() + "..."
        return ""

    def sentence_contains_blacklist(self, sentence, blacklist):
        sentence_lower = sentence.lower()
        return any(term in sentence_lower for term in blacklist)

    def convert_markdown_bold_to_html(self, text):
        return re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)

    def normalize_label(self, label):
        label = re.sub(r'\s+', ' ', label).strip()
        return label[:-1].strip() if label.endswith(':') else label

    def extract_paragraphs(self, source):
        paragraphs = re.findall(r'(?is)<p\b[^>]*>(.*?)</p>', source)
        normalized_paragraphs = []
        for paragraph in paragraphs:
            paragraph = re.sub(r'(?is)<img\b[^>]*>', ' ', paragraph)
            paragraph = self.convert_markdown_bold_to_html(paragraph)
            paragraph = re.sub(r'<br\s*/?>', ' ', paragraph, flags=re.IGNORECASE)
            paragraph = re.sub(r'\s+', ' ', paragraph).strip()
            paragraph_text = unescape(re.sub(r'<[^>]+>', '', paragraph)).strip()
            if paragraph_text:
                normalized_paragraphs.append((paragraph, paragraph_text))
        return normalized_paragraphs

    def classify_paragraph(self, paragraph_html, paragraph_text):
        strong_match = re.match(r'^<strong>(.+?)</strong>\s*(.*)$', paragraph_html, flags=re.IGNORECASE | re.DOTALL)
        if strong_match:
            raw_label = strong_match.group(1).strip()
            value = strong_match.group(2).strip()
            label = self.normalize_label(raw_label)
            if value:
                return {'type': 'labeled_value', 'label': label, 'value': value, 'text': paragraph_text}
            return {'type': 'heading', 'label': label, 'text': paragraph_text}

        colon_match = re.match(r'^([^:]{1,80}):\s*(.+)$', paragraph_text)
        if colon_match:
            label = self.normalize_label(colon_match.group(1))
            value = colon_match.group(2).strip()
            if label and value:
                return {'type': 'plain_labeled_value', 'label': label, 'value': value, 'text': paragraph_text}

        return {'type': 'text', 'text': paragraph_text}

    def detect_language(self, text):
        cleaned = self.clean_html_text(text).lower()
        if not cleaned:
            return 'unknown'
        if self.contains_chinese(cleaned):
            return 'zh'

        ua_score = sum(cleaned.count(ch) for ch in UKRAINIAN_UNIQUE_CHARS)
        ru_score = sum(cleaned.count(ch) for ch in RUSSIAN_UNIQUE_CHARS)

        words = re.findall(r'[а-яіїєґёъыэ]+', cleaned, flags=re.IGNORECASE)
        for word in words:
            if word in UKRAINIAN_MARKERS:
                ua_score += 2
            if word in RUSSIAN_MARKERS:
                ru_score += 2

        if ua_score == 0 and ru_score == 0:
            return 'unknown'
        if ua_score > ru_score:
            return 'ua'
        if ru_score > ua_score:
            return 'ru'
        return 'unknown'

    def format_description(self, text, blacklist):
        source = self.strip_technical_html(text)
        paragraphs = self.extract_paragraphs(source)
        if not paragraphs:
            return ""

        result = []
        current_list_items = []
        current_list_labels = set()

        def flush_list():
            nonlocal current_list_items, current_list_labels
            if current_list_items:
                result.append('<ul>')
                result.extend(current_list_items)
                result.append('</ul>')
                current_list_items = []
                current_list_labels = set()

        for paragraph_html, paragraph_text in paragraphs:
            if self.sentence_contains_blacklist(paragraph_text, blacklist):
                continue
            if self.contains_chinese(paragraph_text):
                continue

            item = self.classify_paragraph(paragraph_html, paragraph_text)
            item_type = item['type']

            if item_type == 'heading':
                flush_list()
                result.append(f"<p><strong>{item['label']}:</strong></p>")
                continue

            if item_type in ('labeled_value', 'plain_labeled_value'):
                label = item['label']
                value = item['value']

                if self.contains_chinese(label) or self.contains_chinese(value):
                    continue

                if label.lower() in ('увага', 'примітка', 'внимание', 'примечание'):
                    flush_list()
                    result.append(f"<p><strong>{label}:</strong> {value}</p>")
                    continue

                if label not in current_list_labels:
                    current_list_items.append(f"<li><strong>{label}:</strong> {value}</li>")
                    current_list_labels.add(label)
                else:
                    flush_list()
                    result.append(f"<p><strong>{label}:</strong> {value}</p>")
                continue

            flush_list()
            result.append(f"<p>{item['text']}</p>")

        flush_list()
        formatted = '\n'.join(result).strip()
        return formatted if self.is_meaningful_text(formatted) else ""

    def get_unique_column_name(self, columns, base_name):
        if base_name not in columns:
            return base_name
        counter = 2
        while f"{base_name}_{counter}" in columns:
            counter += 1
        return f"{base_name}_{counter}"

    def filter_content_errors(self, error_text):
        """Filter out [КОНТЕНТ] errors about technical HTML from error text"""
        filtered_errors = []
        for error_line in error_text.split('\n'):
            if error_line and not error_line.startswith('[КОНТЕНТ]'):
                filtered_errors.append(error_line)
        return '\n'.join(filtered_errors)

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
            formatted_ru_descriptions = []
            formatted_ua_descriptions = []
            rows_with_errors = set()
            rows_with_language_mismatch = set()

            for i, row in df.iterrows():
                text_ru = self.normalize_text(row.get(col_ru, ''))
                text_ua = self.normalize_text(row.get(col_ua, ''))
                row_errors = []

                ru_detected_lang = self.detect_language(text_ru)
                ua_detected_lang = self.detect_language(text_ua)

                if text_ru.strip() and ru_detected_lang in ('ua', 'zh'):
                    if ru_detected_lang == 'zh':
                        row_errors.append("[МОВА] У стовпчику RU виявлено китайський текст. Потрібен переклад російською.")
                    else:
                        row_errors.append("[МОВА] У стовпчику RU виявлено український текст. Потрібен переклад російською.")
                    formatted_ru_text = ""
                    rows_with_errors.add(i)
                    rows_with_language_mismatch.add(i)
                else:
                    formatted_ru_text = self.format_description(text_ru, blacklist)
                    if text_ru.strip() and not formatted_ru_text:
                        row_errors.append("[КОНТЕНТ] У стовпчику RU не виявлено придатного текстового опису або знайдено технічний HTML/китайський контент.")
                        rows_with_errors.add(i)
                        rows_with_language_mismatch.add(i)

                if text_ua.strip() and ua_detected_lang in ('ru', 'zh'):
                    if ua_detected_lang == 'zh':
                        row_errors.append("[МОВА] У стовпчику UA виявлено китайський текст. Потрібен переклад українською.")
                    else:
                        row_errors.append("[МОВА] У стовпчику UA виявлено російський текст. Потрібен переклад українською.")
                    formatted_ua_text = ""
                    rows_with_errors.add(i)
                    rows_with_language_mismatch.add(i)
                else:
                    formatted_ua_text = self.format_description(text_ua, blacklist)
                    if text_ua.strip() and not formatted_ua_text:
                        row_errors.append("[КОНТЕНТ] У стовпчику UA не виявлено придатного текстового опису або знайдено технічний HTML/китайський контент.")
                        rows_with_errors.add(i)
                        rows_with_language_mismatch.add(i)

                text_ua_lower = text_ua.lower()
                for blacklisted_term in blacklist:
                    if blacklisted_term in text_ua_lower:
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
                    match_ru = re.search(r'(?i)\b' + re.escape(ru_stem) + r'[\w-]*\b', text_ru)
                    if match_ru:
                        ru_full_word = match_ru.group(0)
                        for ua_stem in ua_stems:
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
                formatted_ru_descriptions.append(formatted_ru_text)
                formatted_ua_descriptions.append(formatted_ua_text)

                if row_errors:
                    rows_with_errors.add(i)

                if i % 10 == 0:
                    self.progress_bar['value'] = i + 1
                    self.lbl_progress_status.config(text=f"Обробка: {i + 1}/{total_rows}")

            errors_col_name = self.get_unique_column_name(df.columns, 'Помилки')
            checked_ru_col_name = self.get_unique_column_name(df.columns, f'{col_ru};1_checked')
            checked_ua_col_name = self.get_unique_column_name(df.columns, f'{col_ua};1_checked')
            status_col_name = self.get_unique_column_name(df.columns, 'Статус')

            # Create status column - only errors that remain unfixed (exclude Technical HTML errors)
            status_column = []
            for row_error_text in errors_column:
                status_column.append(self.filter_content_errors(row_error_text))

            df[errors_col_name] = errors_column
            df[checked_ru_col_name] = formatted_ru_descriptions
            df[checked_ua_col_name] = formatted_ua_descriptions
            df[status_col_name] = status_column

            # Reorder columns: Помилки, Описание;1_checked, Описание (ua);1_checked, Статус
            # Get all original columns except the new ones we just added
            original_cols = [col for col in df.columns if col not in [errors_col_name, checked_ru_col_name, checked_ua_col_name, status_col_name]]
            # Create new column order with our specific columns first
            new_column_order = [errors_col_name, checked_ru_col_name, checked_ua_col_name, status_col_name] + original_cols
            df = df[new_column_order]

            self.lbl_progress_status.config(text="Збереження файлу...")
            wb = Workbook()
            ws = wb.active
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)

            # Apply formatting to all cells
            bold_font = Font(bold=True)
            wrap_alignment = Alignment(wrap_text=True, vertical='top')
            yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
            red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

            # Find column indices for the checked columns
            checked_ru_col_idx = None
            checked_ua_col_idx = None
            for idx, cell in enumerate(ws[1], start=1):
                if cell.value == checked_ru_col_name:
                    checked_ru_col_idx = idx
                elif cell.value == checked_ua_col_name:
                    checked_ua_col_idx = idx
                # Early exit if both columns found
                if checked_ru_col_idx and checked_ua_col_idx:
                    break

            # Format header row (row 1): height 30, bold, wrap text
            ws.row_dimensions[1].height = 30
            for cell in ws[1]:
                cell.font = bold_font
                cell.alignment = wrap_alignment

            # Format data rows: height 15, wrap text
            for row_num in range(2, ws.max_row + 1):
                ws.row_dimensions[row_num].height = 15
                for cell in ws[row_num]:
                    cell.alignment = wrap_alignment

            # Apply highlighting only to cells with errors in _checked columns
            for row_idx in rows_with_errors:
                excel_row_num = row_idx + 2  # +2 because Excel is 1-indexed and we have a header row
                fill_to_use = red_fill if row_idx in rows_with_language_mismatch else yellow_fill
                
                # Only highlight the _checked columns if they have content
                if checked_ru_col_idx and ws.cell(row=excel_row_num, column=checked_ru_col_idx).value:
                    ws.cell(row=excel_row_num, column=checked_ru_col_idx).fill = fill_to_use
                if checked_ua_col_idx and ws.cell(row=excel_row_num, column=checked_ua_col_idx).value:
                    ws.cell(row=excel_row_num, column=checked_ua_col_idx).fill = fill_to_use

            wb.save(save_path)
            self.progress_bar['value'] = total_rows
            self.lbl_progress_status.config(text="Готово!")
            self.log_message(
                f"Завершено. Знайдено помилок у {len(rows_with_errors)} рядках. "
                f"Критичних рядків: {len(rows_with_language_mismatch)}. "
                f"Створено колонки: '{checked_ru_col_name}', '{checked_ua_col_name}' та '{status_col_name}'."
            )
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
