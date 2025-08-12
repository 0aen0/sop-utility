# -*- coding: UTF-8 -*-

"""
SOP File Utility GUI

Графический интерфейс для работы с SOP-файлами.
Позволяет:
- Извлекать JSON-файлы из SOP архива
- Создавать SOP архивы из JSON файлов 
- Извлекать все файлы из архива
- Просматривать содержимое архива
- Просматривать содержимое JSON файлов
"""

# Импорт необходимых библиотек
import os         # Для работы с путями и директориями
import zipfile    # Для работы с ZIP архивами
import zlib       # Для сжатия/распаковки данных
import tkinter as tk                    # Для создания GUI
from tkinter import ttk                 # Для современных виджетов
from tkinter import filedialog         # Для диалогов выбора файлов
from tkinter import messagebox         # Для диалоговых окон сообщений
import json

# Проверяем наличие библиотеки tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    ENABLE_DND = True
except ImportError:
    ENABLE_DND = False
    print("Библиотека tkinterdnd2 не найдена. Функционал drag-n-drop отключен.")

version = "2.0.4"

class ToolTip:
    def __init__(self, widget, msg):
        self.widget = widget
        self.msg = msg
        self.tip_window = None
        self.id = None
        self.x = self.y = 0
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        "Показать подсказку"
        if self.tip_window or not self.msg:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(tw, text=self.msg, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, 
                        borderwidth=1, font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        "Скрыть подсказку"
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None

class SettingsWindow(tk.Toplevel):
    """Окно настроек программы"""
    def __init__(self, parent, settings):
        super().__init__(parent)
        self.title("Настройки")
        self.geometry("400x350")
        self.settings = settings
        self.parent = parent
        
        # Делаем окно модальным
        self.transient(parent)
        self.grab_set()
        
        # Создаем основной контейнер
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Секция форматирования JSON
        format_frame = ttk.LabelFrame(main_frame, text="Форматирование JSON", padding="10")
        format_frame.pack(fill=tk.X, pady=5)
        
        # Основная опция форматирования
        self.format_json = tk.BooleanVar(value=self.settings.get('format_json', True))
        format_check = ttk.Checkbutton(format_frame, 
                                     text="Форматировать JSON при извлечении",
                                     variable=self.format_json)
        format_check.pack(anchor=tk.W)
        ToolTip(format_check, "Включает/выключает форматирование при извлечении JSON из SOP")
        
        # Настройки отступов
        indent_frame = ttk.Frame(format_frame)
        indent_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(indent_frame, text="Размер отступа:").pack(side=tk.LEFT)
        self.indent_size = tk.IntVar(value=self.settings.get('indent_size', 4))
        indent_spin = ttk.Spinbox(indent_frame, 
                                 from_=0, 
                                 to=8, 
                                 width=3,
                                 textvariable=self.indent_size)
        indent_spin.pack(side=tk.LEFT, padx=5)
        ToolTip(indent_spin, "Количество пробелов для отступов в форматированном JSON")
        
        # Дополнительные параметры форматирования
        options_frame = ttk.Frame(format_frame)
        options_frame.pack(fill=tk.X, pady=5)
        
        self.ensure_ascii = tk.BooleanVar(value=self.settings.get('ensure_ascii', False))
        ensure_ascii_check = ttk.Checkbutton(options_frame,
                                            text="Преобразовывать UTF-8 в \\u-последовательности",
                                            variable=self.ensure_ascii)
        ensure_ascii_check.pack(anchor=tk.W)
        ToolTip(ensure_ascii_check, "Преобразует Unicode символы в escape-последовательности")
        
        self.sort_keys = tk.BooleanVar(value=self.settings.get('sort_keys', False))
        sort_keys_check = ttk.Checkbutton(options_frame,
                                        text="Сортировать ключи в алфавитном порядке",
                                        variable=self.sort_keys)
        sort_keys_check.pack(anchor=tk.W)
        ToolTip(sort_keys_check, "Сортирует ключи JSON объектов по алфавиту")
        
        # Секция настроек для SOP архивов
        sop_frame = ttk.LabelFrame(main_frame, text="Настройки создания SOP", padding="10")
        sop_frame.pack(fill=tk.X, pady=5)
        
        self.minify_json = tk.BooleanVar(value=self.settings.get('minify_json', False))
        minify_check = ttk.Checkbutton(sop_frame,
                                     text="Минимизировать JSON при создании SOP",
                                     variable=self.minify_json)
        minify_check.pack(anchor=tk.W)
        ToolTip(minify_check, "Удаляет все пробелы и переносы строк для минимального размера файла")
        
        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        save_button = ttk.Button(button_frame, 
                               text="Сохранить", 
                               command=self.save_settings)
        save_button.pack(side=tk.RIGHT, padx=5)
        
        cancel_button = ttk.Button(button_frame,
                                 text="Отмена",
                                 command=self.destroy)
        cancel_button.pack(side=tk.RIGHT)
        
        self.center_window(parent)
    
    def center_window(self, parent):
        """Центрирует окно относительно родительского"""
        self.update_idletasks()
        
        # Получаем геометрию родительского окна
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        
        # Получаем размеры текущего окна
        width = self.winfo_width()
        height = self.winfo_height()
        
        # Вычисляем новую позицию
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        # Устанавливаем позицию
        self.geometry(f"+{x}+{y}")
    
    def save_settings(self):
        """Сохраняет настройки и закрывает окно"""
        # Сохраняем все значения настроек
        self.settings['format_json'] = self.format_json.get()
        self.settings['indent_size'] = self.indent_size.get()
        self.settings['ensure_ascii'] = self.ensure_ascii.get()
        self.settings['sort_keys'] = self.sort_keys.get()
        self.settings['minify_json'] = self.minify_json.get()
        
        # Применяем настройки в главном окне
        if hasattr(self.parent, 'apply_settings'):
            self.parent.apply_settings(self.settings)
        
        # Закрываем окно
        self.destroy()
		
class PackageInfoWindow(tk.Toplevel):
    """Окно для отображения информации о пакете"""
    # Дефолтные значения для маппинга приложений
    DEFAULT_APP_MAPPINGS = [
        {"pack_application_id": "155931135900000002", "AppName": "Simple"},
        {"pack_application_id": "156950344216170038", "AppName": "ITSM"},
        {"pack_application_id": "158517002917848381", "AppName": "Personal Schedule"},
        {"pack_application_id": "163881580010333018", "AppName": "Auth"},
        {"pack_application_id": "166452097111650278", "AppName": "CRM"},
        {"pack_application_id": "168519538400157974", "AppName": "ITAM"},
        {"pack_application_id": "169089649104163836", "AppName": "SDLC"},
        {"pack_application_id": "170489310418077474", "AppName": "HRMS"}
    ]
    
    def __init__(self, parent, json_data):
        super().__init__(parent)
        self.title("Информация о пакете")
        self.geometry("500x320")
        self.minsize(500, 320)  # Устанавливаем минимальный размер окна
        
        # Делаем окно модальным
        self.transient(parent)
        self.grab_set()
        
        # Загружаем маппинги приложений
        self.app_mappings = self.load_app_mappings()
        
        # Создаем фрейм с отступами
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настраиваем расширение
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)  # Даем вес строке с описанием
        
        # Получаем данные из JSON
        name = json_data.get('name', '')
        description = json_data.get('description', '')
        pack_application_id = json_data.get('pack_application_id', '')
        records = json_data.get('records', [])
        records_count = len(records) if isinstance(records, list) else 0
        
        # Ищем AppName в массиве соответствий
        app_name = self.find_app_name(pack_application_id)
        
        # Создаем и размещаем метки с информацией
        ttk.Label(main_frame, text="Название:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Label(main_frame, text=name).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Создаем фрейм для описания с прокруткой
        desc_frame = ttk.Frame(main_frame)
        desc_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        desc_frame.columnconfigure(0, weight=1)
        desc_frame.rowconfigure(0, weight=1)
        
        # Метка "Описание"
        ttk.Label(desc_frame, text="Описание:").grid(row=0, column=0, sticky=tk.W, pady=(0,5))
        
        # Создаем текстовое поле с прокруткой
        desc_text = tk.Text(desc_frame, wrap=tk.WORD, height=6)  # Увеличиваем высоту на 2 строки
        desc_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Добавляем вертикальную прокрутку
        desc_scroll = ttk.Scrollbar(desc_frame, orient="vertical", command=desc_text.yview)
        desc_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        desc_text.configure(yscrollcommand=desc_scroll.set)
        
        # Вставляем описание и делаем поле только для чтения
        desc_text.insert('1.0', description)
        desc_text.config(state='disabled')
        
        # ID приложения
        ttk.Label(main_frame, text="ID приложения:").grid(row=2, column=0, sticky=tk.W, pady=5)
        id_text = f"{pack_application_id} ({app_name})" if app_name else pack_application_id
        ttk.Label(main_frame, text=id_text).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Количество записей
        ttk.Label(main_frame, text="Количество записей:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Label(main_frame, text=str(records_count)).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Кнопка закрытия
        ttk.Button(main_frame, text="Закрыть", command=self.destroy).grid(
            row=4, column=0, columnspan=2, pady=20)
            
        # Центрируем окно относительно родительского окна
        self.center_window(parent)
        
    def center_window(self, parent):
        """Центрирует окно относительно родительского окна"""
        # Ждем, пока окно будет создано
        self.update_idletasks()
        
        # Получаем размеры и позицию родительского окна
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        
        # Получаем размеры текущего окна
        width = self.winfo_width()
        height = self.winfo_height()
        
        # Вычисляем позицию для центрирования
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        # Устанавливаем позицию окна
        self.geometry(f"+{x}+{y}")
        
    def load_app_mappings(self):
        """Загружает маппинги приложений из файла или возвращает дефолтные значения"""
        try:
            # Получаем путь к файлу app_mappings.json относительно текущего скрипта
            script_dir = os.path.dirname(os.path.abspath(__file__))
            mappings_file = os.path.join(script_dir, 'app_mappings.json')
            
            if os.path.exists(mappings_file):
                with open(mappings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                # Если файл не существует, создаем его с дефолтными значениями
                with open(mappings_file, 'w', encoding='utf-8') as f:
                    json.dump(self.DEFAULT_APP_MAPPINGS, f, indent=4, ensure_ascii=False)
                return self.DEFAULT_APP_MAPPINGS
        except Exception as e:
            print(f"Ошибка при загрузке маппингов приложений: {str(e)}")
            return self.DEFAULT_APP_MAPPINGS
            
    def find_app_name(self, pack_application_id):
        """Поиск AppName по pack_application_id в загруженном массиве соответствий"""
        for mapping in self.app_mappings:
            if mapping.get("pack_application_id") == pack_application_id:
                return mapping.get("AppName")
        return ""  # Если соответствие не найдено

class JsonViewerWindow(tk.Toplevel):
    """Окно для просмотра JSON с подсветкой синтаксиса"""
    def __init__(self, parent, json_data):
        super().__init__(parent)
        self.title("Просмотр JSON")
        self.geometry("800x600")
        self.minsize(600, 400)
        
        # Делаем окно модальным
        self.transient(parent)
        self.grab_set()
        
        # Создаем фрейм с отступами
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настраиваем расширение
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Создаем текстовое поле с прокруткой
        self.text = tk.Text(main_frame, wrap=tk.NONE, font=('Courier New', 10))
        self.text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Добавляем полосы прокрутки
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=self.text.yview)
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=self.text.xview)
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Кнопка закрытия
        ttk.Button(main_frame, text="Закрыть", command=self.destroy).grid(
            row=2, column=0, columnspan=2, pady=10)
        
        # Форматируем и вставляем JSON
        formatted_json = json.dumps(json_data, indent=4, ensure_ascii=False)
        self.text.insert('1.0', formatted_json)
        
        # Делаем текст только для чтения
        self.text.config(state='disabled')
        
        # Настраиваем расширение
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Центрируем окно
        self.center_window(parent)
        
    def center_window(self, parent):
        """Центрирует окно относительно родительского окна"""
        self.update_idletasks()
        
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        
        width = self.winfo_width()
        height = self.winfo_height()
        
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        self.geometry(f"+{x}+{y}")

class AboutWindow(tk.Toplevel):
    """Окно информации о программе"""
    def __init__(self, parent):
        super().__init__(parent)
        self.title("О программе")
        self.geometry("400x320")
        self.minsize(400, 320)
        self.resizable(False, False)  # Запрещаем изменение размера окна
        
        # Делаем окно модальным
        self.transient(parent)
        self.grab_set()
        
        # Создаем фрейм с отступами
        main_frame = ttk.Frame(self, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настраиваем расширение
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Заголовок программы
        title_label = ttk.Label(main_frame, text="SOP File Utility", 
                              font=('Helvetica', 14, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Версия
        version_label = ttk.Label(main_frame, text=f"Версия {version}")
        version_label.grid(row=1, column=0, pady=(0, 20))
        
        # Описание программы
        description = """Утилита для работы с SOP-файлами.

Позволяет:
- Просматривать содержимое файла SOP
- Извлекать JSON файлы из файла SOP
- Извлекать все файлы из SOP
- Создавать файл SOP из JSON файлов
- Просматривать содержимое JSON файлов
- Поддерживается Drag-n-Drop"""
        
        desc_text = tk.Text(main_frame, wrap=tk.WORD, height=8, width=40)
        desc_text.grid(row=2, column=0, pady=(0, 20))
        desc_text.insert('1.0', description)
        desc_text.config(state='disabled')
        
        # Автор
        author_label = ttk.Label(main_frame, text="© AEN 2024-2025")
        author_label.grid(row=3, column=0, pady=(0, 20))
        
        # Кнопка закрытия
        ttk.Button(main_frame, text="Закрыть", command=self.destroy).grid(
            row=4, column=0)
            
        # Центрируем окно относительно родительского окна
        self.center_window(parent)
        
    def center_window(self, parent):
        """Центрирует окно относительно родительского окна"""
        self.update_idletasks()
        
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        
        width = self.winfo_width()
        height = self.winfo_height()
        
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        self.geometry(f"+{x}+{y}")

class SopUtilityGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"SOP File Utility v{version}")
        
        # Настройки по умолчанию
        self.settings = {
            'format_json': True,
            'minify_json': True,
            'indent_size': 4,
            'ensure_ascii': False,
            'sort_keys': False
        }

        # Создаем меню
        self.create_menu()
        
        # Настраиваем главное окно для изменения размера
        self.root.minsize(618, 400)  # Минимальный размер окна
        self.root.geometry("618x400")  # Начальный размер окна
        
        # Регистрируем поддержку drag-n-drop только если библиотека доступна
        if ENABLE_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.handle_drop)
        
        # Разрешаем изменение размера главного окна
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        
        # Создаем основной контейнер
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настраиваем расширение для main_frame
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)  # Основной вес для области с таблицей и логами
        
        # Создаем фрейм для пути к файлу
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # Добавляем метку и поле для входного файла
        ttk.Label(file_frame, text="Входной файл:").grid(row=0, column=0, padx=(0,5))
        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path)
        self.file_entry.grid(row=0, column=1, padx=5, sticky=tk.W+tk.E)
        
        # Добавляем кнопку выбора файла
        ttk.Button(file_frame, text="Обзор...", command=self.browse_file).grid(row=0, column=2, padx=5)

        # Добавляем метку и поле для выходного каталога
        ttk.Label(file_frame, text="Выходной каталог:").grid(row=1, column=0, padx=(0,5), pady=(5,0))
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(file_frame, textvariable=self.output_path)
        self.output_entry.grid(row=1, column=1, padx=5, pady=(5,0), sticky=tk.W+tk.E)
        
        # Добавляем кнопку выбора выходного каталога
        ttk.Button(file_frame, text="Обзор...", command=self.browse_output_dir).grid(row=1, column=2, padx=5, pady=(5,0))
        
        # Настраиваем расширение колонок
        file_frame.columnconfigure(1, weight=1)
        
        # Создаем фрейм для кнопок операций с центрированием
        self.buttons_frame = ttk.Frame(main_frame)
        self.buttons_frame.grid(row=1, column=0, padx=5, pady=5, sticky=tk.EW)
        
        # Добавляем пустой фрейм слева для центрирования
        self.left_spacer = ttk.Frame(self.buttons_frame)
        self.left_spacer.grid(row=0, column=0, sticky=tk.EW)
        
        # Создаем центральный фрейм для кнопок
        self.center_buttons_frame = ttk.Frame(self.buttons_frame)
        self.center_buttons_frame.grid(row=0, column=1, sticky=tk.EW)
        
        # Добавляем пустой фрейм справа для центрирования
        self.right_spacer = ttk.Frame(self.buttons_frame)
        self.right_spacer.grid(row=0, column=2, sticky=tk.EW)
        
        # Создаем кнопки операций
        self.extract_json_button = ttk.Button(self.center_buttons_frame, text="Извлечь JSON файл", 
                                            command=self.extract_json_dialog, width=21)
        self.create_sop_button = ttk.Button(self.center_buttons_frame, text="Создать SOP файл", 
                                          command=self.create_sop_dialog, width=21)
        self.extract_all_button = ttk.Button(self.center_buttons_frame, text="Извлечь все файлы", 
                                           command=self.extract_all_dialog, width=21)
        self.package_info_button = ttk.Button(self.center_buttons_frame, text="Информация о пакете",
                                            command=self.show_package_info, width=21)
        self.show_json_button = ttk.Button(self.center_buttons_frame, text="Просмотр JSON",
                                         command=self.view_json, width=21)
                                           
        # Добавляем чекбокс для включения всех файлов
        self.include_all_files = tk.BooleanVar()
        self.include_all_checkbox = ttk.Checkbutton(self.center_buttons_frame, 
                                                   text="Добавить все файлы",
                                                   variable=self.include_all_files)
        
        # Настраиваем веса колонок для центрирования
        self.buttons_frame.columnconfigure(0, weight=1)  # Левый отступ
        self.buttons_frame.columnconfigure(1, weight=0)  # Центральная часть не растягивается
        self.buttons_frame.columnconfigure(2, weight=1)  # Правый отступ
        
        # Создаем фрейм для таблицы и логов
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настраиваем расширение для bottom_frame
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.rowconfigure(0, weight=1)
        
        # Создаем разделитель для таблицы и логов
        separator = ttk.PanedWindow(bottom_frame, orient=tk.VERTICAL)
        separator.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Создаем фрейм для таблицы
        tree_frame = ttk.Frame(separator)
        separator.add(tree_frame, weight=3)  # Таблица занимает 3/4 пространства
        
        # Создаем таблицу
        self.tree = ttk.Treeview(tree_frame, columns=("directory", "filename", "size"), show="headings")
        
        # Настраиваем заголовки
        self.tree.heading("directory", text="Каталог", command=lambda: self.sort_tree("directory"))
        self.tree.heading("filename", text="Имя файла", command=lambda: self.sort_tree("filename"))
        self.tree.heading("size", text="Размер (байт)", command=lambda: self.sort_tree("size", numeric=True))
        
        # Настраиваем колонки
        self.tree.column("directory", width=160, anchor="w")
        self.tree.column("filename", width=300, anchor="w")
        self.tree.column("size", width=80, anchor="e")
        
        # Добавляем полосы прокрутки для таблицы
        tree_vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)
        
        # Размещаем элементы таблицы
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tree_hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Настраиваем расширение для фрейма таблицы
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Создаем фрейм для логов
        log_frame = ttk.LabelFrame(separator, text="Логи")
        separator.add(log_frame, weight=1)  # Логи занимают 1/4 пространства
        
        # Создаем текстовое поле для логов с уменьшенной высотой
        self.output_text = tk.Text(log_frame, height=4, wrap=tk.WORD)  # Уменьшаем высоту до 4 строк
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Добавляем полосу прокрутки для логов
        log_vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.output_text.yview)
        log_vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.output_text.configure(yscrollcommand=log_vsb.set)
        
        # Настраиваем расширение для фрейма логов
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Настраиваем расширение колонок в главном фрейме
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)  # Позволяет таблице расширяться
        
        # Привязываем обновление кнопок к изменению пути файла
        self.file_path.trace_add("write", self.update_buttons)
        
        # Скрываем все кнопки изначально
        self.hide_all_buttons()

    def create_menu(self):
        """Создает главное меню программы"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Меню "Файл"
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Выход", command=self.root.quit)
        
        # Меню Настройки
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Настройки...", command=self.show_settings)
        menubar.add_cascade(label="Настройки", menu=settings_menu)
        
        # Меню Справка
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="О программе", command=self.show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def show_settings(self):
        """Показывает окно настроек"""
        SettingsWindow(self.root, self.settings)
    
    def apply_settings(self, new_settings):
        """Применяет новые настройки"""
        self.settings = new_settings
        self.log_message("Настройки применены")
    
    def show_about(self):
        """Показывает окно 'О программе'"""
        AboutWindow(self.root)

    def sort_tree(self, col, numeric=False):
        """Сортирует данные в таблице по указанной колонке"""
        # Получаем все элементы
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children("")]
        
        # Сортируем
        items.sort(key=lambda x: int(x[0]) if numeric and x[0].isdigit() else x[0].lower())
        
        # Если уже отсортировано в прямом порядке, сортируем в обратном
        if hasattr(self, "_sort_reverse") and not self._sort_reverse:
            items.reverse()
            self._sort_reverse = True
        else:
            self._sort_reverse = False
            
        # Переупорядочиваем элементы
        for index, (_, item) in enumerate(items):
            self.tree.move(item, "", index)

    def clear_tree(self):
        """Очищает таблицу"""
        for item in self.tree.get_children():
            self.tree.delete(item)

    def hide_all_buttons(self):
        """Скрывает все кнопки операций"""
        self.extract_json_button.grid_remove()
        self.create_sop_button.grid_remove()
        self.extract_all_button.grid_remove()
        self.include_all_checkbox.grid_remove()
        self.package_info_button.grid_remove()
        self.show_json_button.grid_remove()

    def show_sop_buttons(self):
        """Показывает кнопки для работы с SOP файлами"""
        self.package_info_button.grid(row=0, column=0, padx=5, pady=2)
        self.show_json_button.grid(row=0, column=1, padx=5, pady=2)
        self.extract_json_button.grid(row=0, column=2, padx=5, pady=2)
        self.extract_all_button.grid(row=0, column=3, padx=5, pady=2)
        self.create_sop_button.grid_remove()
        self.include_all_checkbox.grid_remove()

    def show_json_buttons(self):
        """Показывает кнопки для работы с JSON файлами"""
        self.show_json_button.grid(row=0, column=0, padx=5, pady=2)  # Добавляем кнопку просмотра
        self.create_sop_button.grid(row=0, column=1, padx=5, pady=2)
        self.include_all_checkbox.grid(row=0, column=3, padx=5, pady=2)  # Показываем чекбокс рядом с кнопкой
        self.extract_json_button.grid_remove()
        self.extract_all_button.grid_remove()
        self.package_info_button.grid_remove()

    def update_buttons(self, *args):
        """Обновляет видимость кнопок в зависимости от типа файла"""
        file_path = self.file_path.get()
        if not file_path:
            self.hide_all_buttons()
            self.clear_tree()  # Очищаем таблицу если файл не выбран
            return

        if file_path.lower().endswith('.sop'):
            self.show_sop_buttons()
            self.list_archive(file_path)  # Автоматически показываем содержимое SOP архива
        elif file_path.lower().endswith('.json'):
            self.show_json_buttons()
            self.clear_tree()  # Очищаем таблицу для JSON файла
        else:
            self.hide_all_buttons()
            self.clear_tree()  # Очищаем таблицу для неподдерживаемых файлов

    def validate_file_path(self, file_path):
        """Проверяет корректность пути к файлу"""
        if not file_path:
            messagebox.showerror("Ошибка", "Путь к файлу не указан")
            return False
            
        # Нормализуем путь для безопасности
        try:
            norm_path = os.path.normpath(os.path.abspath(file_path))
        except Exception:
            messagebox.showerror("Ошибка", "Некорректный путь к файлу")
            return False

        # Проверяем существование файла
        if not os.path.exists(norm_path):
            messagebox.showerror("Ошибка", f"Файл не существует: {file_path}")
            return False
            
        # Проверяем что это файл, а не директория
        if not os.path.isfile(norm_path):
            messagebox.showerror("Ошибка", f"Указанный путь не является файлом: {file_path}")
            return False
            
        # Проверяем доступ на чтение
        if not os.access(norm_path, os.R_OK):
            messagebox.showerror("Ошибка", f"Нет доступа на чтение файла: {file_path}")
            return False
            
        return True
        
    def validate_output_dir(self, output_dir):
        """Проверяет корректность пути для выходного каталога"""
        if not output_dir:
            messagebox.showerror("Ошибка", "Выходной каталог не указан")
            return False
            
        # Нормализуем путь для безопасности
        try:
            norm_path = os.path.normpath(os.path.abspath(output_dir))
        except Exception:
            messagebox.showerror("Ошибка", "Некорректный путь для выходного каталога")
            return False
            
        # Проверяем возможность создания каталога
        try:
            if not os.path.exists(norm_path):
                os.makedirs(norm_path)
            elif not os.path.isdir(norm_path):
                messagebox.showerror("Ошибка", f"Указанный путь не является каталогом: {output_dir}")
                return False
                
            # Проверяем права на запись
            test_file = os.path.join(norm_path, "test_write_access")
            try:
                with open(test_file, 'w') as f:
                    f.write("")
                os.remove(test_file)
            except Exception:
                messagebox.showerror("Ошибка", f"Нет прав на запись в каталог: {output_dir}")
                return False
                
            return True
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать каталог {output_dir}: {str(e)}")
            return False
            
    def validate_json_file(self, file_path):
        """Проверяет что файл является допустимым JSON файлом"""
        if not file_path.lower().endswith('.json'):
            messagebox.showerror("Ошибка", "Файл должен иметь расширение .json")
            return False
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                json.load(f)
            return True
        except json.JSONDecodeError:
            messagebox.showerror("Ошибка", "Файл содержит некорректный JSON")
            return False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при чтении JSON файла: {str(e)}")
            return False
            
    def validate_sop_file(self, file_path):
        """Проверяет что файл является допустимым SOP архивом"""
        if not file_path.lower().endswith('.sop'):
            messagebox.showerror("Ошибка", "Файл должен иметь расширение .sop")
            return False
            
        try:
            if not zipfile.is_zipfile(file_path):
                messagebox.showerror("Ошибка", "Файл не является допустимым ZIP архивом")
                return False
                
            with zipfile.ZipFile(file_path, 'r') as zip:
                # Проверяем целостность архива
                if zip.testzip() is not None:
                    messagebox.showerror("Ошибка", "Архив поврежден")
                    return False
                    
                # Проверяем наличие хотя бы одного .data файла
                has_data_file = any(name.endswith('.data') for name in zip.namelist())
                if not has_data_file:
                    messagebox.showerror("Ошибка", "Архив не содержит .data файлов")
                    return False
                    
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при проверке SOP файла: {str(e)}")
            return False

    def browse_file(self):
        """Открывает диалог выбора файла"""
        filetypes = [
            ("Все поддерживаемые", "*.sop;*.json"),
            ("SOP файлы", "*.sop"),
            ("JSON файлы", "*.json"),
            ("Все файлы", "*.*")
        ]
        filename = filedialog.askopenfilename(
            title="Выберите файл",
            filetypes=filetypes
        )
        if filename:
            self.file_path.set(filename)
            # Кнопки обновятся автоматически через trace

    def browse_output_dir(self):
        """Открывает диалог выбора выходного каталога"""
        dirname = filedialog.askdirectory(
            title="Выберите выходной каталог"
        )
        if dirname:
            self.output_path.set(dirname)
            
    def log_message(self, message):
        """Добавляет сообщение в текстовое поле"""
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        
    def create_output_directory(self, output_dir):
        """Проверяет наличие выходного каталога и создает его при отсутствии"""
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                self.log_message(f"Создан каталог: {output_dir}")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать каталог {output_dir}: {str(e)}")
            return False

    def extract_json_dialog(self):
        """Диалог для извлечения JSON файлов"""
        input_file = self.file_path.get()
        if not input_file:
            input_file = filedialog.askopenfilename(
                title="Выберите SOP файл",
                filetypes=[("SOP files", "*.sop"), ("All files", "*.*")]
            )
            if input_file:
                self.file_path.set(input_file)
        
        if input_file:
            # Валидация входного файла
            if not self.validate_file_path(input_file) or not self.validate_sop_file(input_file):
                return
                
            output_dir = self.output_path.get()
            if not output_dir:
                output_dir = filedialog.askdirectory(title="Выберите папку для сохранения")
                if output_dir:
                    self.output_path.set(output_dir)
            
            # Валидация выходного каталога
            if output_dir and self.validate_output_dir(output_dir):
                self.extract_sop(input_file, output_dir)

    def create_sop_dialog(self):
        """Диалог для создания SOP архива"""
        input_file = self.file_path.get()
        if not input_file:
            input_file = filedialog.askopenfilename(
                title="Выберите JSON файл",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if input_file:
                self.file_path.set(input_file)
        
        if input_file:
            # Валидация входного файла
            if not self.validate_file_path(input_file) or not self.validate_json_file(input_file):
                return
                
            output_dir = self.output_path.get()
            if not output_dir:
                output_dir = filedialog.askdirectory(title="Выберите папку для сохранения")
                if output_dir:
                    self.output_path.set(output_dir)
            
            # Валидация выходного каталога
            if output_dir and self.validate_output_dir(output_dir):
                self.create_sop(input_file, output_dir)

    def extract_all_dialog(self):
        """Диалог для извлечения всех файлов"""
        input_file = self.file_path.get()
        if not input_file:
            input_file = filedialog.askopenfilename(
                title="Выберите SOP файл",
                filetypes=[("SOP files", "*.sop"), ("All files", "*.*")]
            )
            if input_file:
                self.file_path.set(input_file)
        
        if input_file:
            # Валидация входного файла
            if not self.validate_file_path(input_file) or not self.validate_sop_file(input_file):
                return
                
            output_dir = self.output_path.get()
            if not output_dir:
                output_dir = filedialog.askdirectory(title="Выберите папку для сохранения")
                if output_dir:
                    self.output_path.set(output_dir)
            
            # Валидация выходного каталога
            if output_dir and self.validate_output_dir(output_dir):
                self.extract_archive(input_file, output_dir)

    def list_archive(self, input_file):
        """Показывает список файлов в SOP архиве в виде таблицы"""
        try:
            if zipfile.is_zipfile(input_file):
                # Очищаем старые данные
                self.clear_tree()
                
                with zipfile.ZipFile(input_file, "r") as zip:
                    for file in zip.infolist():
                        # Разделяем путь на каталог и имя файла
                        path_parts = os.path.split(file.filename)
                        directory = path_parts[0] if path_parts[0] else "-"
                        filename = path_parts[1]
                        
                        # Добавляем строку в таблицу
                        self.tree.insert("", "end", values=(
                            directory,
                            filename,
                            file.file_size
                        ))
                
            else:
                messagebox.showerror("Ошибка", "Неверный формат файла")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def extract_sop(self, input_file, output_dir):
        """Извлекает JSON файлы из SOP архива с настраиваемым форматированием"""
        try:
            if not self.create_output_directory(output_dir):
                return

            if zipfile.is_zipfile(input_file):
                with zipfile.ZipFile(input_file, "r") as zip:
                    for file in zip.namelist():
                        if file.endswith(".data"):
                            out_file = file.replace('.data', '.json')
                            if os.name == 'nt':
                                out_file = out_file.replace(':', '_')
                            
                            out_path = os.path.join(output_dir, out_file)
                            os.makedirs(os.path.dirname(out_path), exist_ok=True)
                            
                            with zip.open(file) as compressed_data:
                                decompressed_data = zlib.decompress(compressed_data.read())
                                
                                try:
                                    json_data = json.loads(decompressed_data.decode('utf-8'))
                                    
                                    if self.settings['format_json']:
                                        output_data = json.dumps(
                                            json_data,
                                            indent=self.settings['indent_size'],
                                            ensure_ascii=self.settings['ensure_ascii'],
                                            sort_keys=self.settings['sort_keys']
                                        )
                                        mode = 'w'
                                        with open(out_path, mode, encoding='utf-8') as f:
                                            f.write(output_data)

                                        self.log_message(f"Извлечение с форматированием: {out_file}")
                                    else:
                                        with open(out_path, 'wb') as f:
                                            f.write(decompressed_data)
                                        self.log_message(f"Извлечение без форматирования: {out_file}")
                                    
                                except json.JSONDecodeError:
                                    with open(out_path, 'wb') as f:
                                        f.write(decompressed_data)
                                    self.log_message(f"Извлечение (битый JSON, без форматирования): {out_file}")
                
                self.log_message("Извлечение завершено успешно")
            else:
                messagebox.showerror("Ошибка", "Неверный формат файла")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при извлечении: {str(e)}")

    def create_sop(self, input_file, output_dir):
        """Создает SOP архив из JSON файла"""
        try:
            if not self.create_output_directory(output_dir):
                return
                
            input_filename = os.path.basename(input_file)
            if input_filename.endswith(".json"):
                data_file = input_filename.replace('.json', '.data')
                zip_out = input_filename.replace('.json', '.sop')
                zip_path = os.path.join(output_dir, zip_out)

                # Получаем абсолютный путь к входному файлу для корректного сравнения
                input_file_abs = os.path.abspath(input_file)

                # Считываем и сжимаем JSON данные
                with open(input_file, 'rb') as fi:
                    # Применяем настройки минификации
                    if self.settings['minify_json']:
                        json_data = json.load(fi)

                        # Функция для дополнительного экранирования слэшей
                        def escape_slashes(obj):
                            if isinstance(obj, str):
                                return obj.replace('/', '\\/')  # Заменяем '/' на '\/'
                            elif isinstance(obj, dict):
                                return {k: escape_slashes(v) for k, v in obj.items()}
                            elif isinstance(obj, list):
                                return [escape_slashes(item) for item in obj]
                            return obj

                        json_str = escape_slashes(json.dumps(json_data, separators=(',', ':'), 
                                            ensure_ascii=True,
                                            sort_keys=False))

                        # Экранирование служебных символов и преобразование в UTF-8
                        # json_bytes = json_str.encode('utf-8', errors='backslashreplace')
                        # compressed_data = zlib.compress(json_bytes, level=5)
                        compressed_data = zlib.compress(json_str.encode('utf-8'), level=5)
                    else:
                        compressed_data = zlib.compress(fi.read(), level=5)

                    self.log_message(f"Создание архива {zip_out}")
                    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=5) as zip:
                        zip.writestr(data_file, compressed_data)
                        
                        # Если выбран чекбокс, добавляем все файлы из каталога
                        if self.include_all_files.get():
                            source_dir = os.path.dirname(input_file)
                            self.log_message(f"Добавление файлов из каталога {source_dir}")
                            self.log_message(f"Исходный файл {input_filename} будет пропущен")
                            
                            # Рекурсивно обходим все файлы в каталоге
                            for root, dirs, files in os.walk(source_dir):
                                for file in files:
                                    file_path = os.path.join(root, file)
                                    # Получаем абсолютный путь для корректного сравнения
                                    file_path_abs = os.path.abspath(file_path)
                                    if file_path_abs != input_file_abs:  # Сравниваем абсолютные пути
                                        # Получаем относительный путь для архива
                                        rel_path = os.path.relpath(file_path, source_dir)
                                        self.log_message(f"Добавление файла {rel_path}")
                                        zip.write(file_path, rel_path)
                    
                self.log_message("Архив создан успешно")
                
                # Меняем путь к файлу на созданный SOP архив
                self.file_path.set(zip_path)
                # Кнопки и содержимое обновятся автоматически через trace
                
            else:
                messagebox.showerror("Ошибка", "Выберите JSON файл")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def extract_archive(self, input_file, output_dir):
        """Извлекает все файлы из SOP архива, кроме .data файлов, и сохраняет декомпрессированный JSON"""
        try:
            if not self.create_output_directory(output_dir):
                return
                
            if zipfile.is_zipfile(input_file):
                with zipfile.ZipFile(input_file, "r") as zip:
                    self.log_message(f"Извлечение файлов из {input_file}")
                    
                    # Получаем список всех файлов в архиве
                    file_list = zip.namelist()
                    
                    # Сначала обрабатываем .data файлы и сохраняем их как JSON
                    for file in file_list:
                        if file.endswith(".data"):
                            # Создаем имя для JSON файла
                            json_file = file.replace('.data', '.json')
                            if os.name == 'nt':
                                json_file = json_file.replace(':','_')
                                
                            json_path = os.path.join(output_dir, json_file)
                            
                            # Создаем подкаталоги если они есть в пути
                            os.makedirs(os.path.dirname(json_path), exist_ok=True)
                            
                            # Распаковываем и декомпрессируем .data в JSON
                            with zip.open(file) as compressed_data:
                                decompressed_data = zlib.decompress(compressed_data.read())
                                self.log_message(f"Извлечение и декомпрессия {json_file}")
                                with open(json_path, 'wb') as f:
                                    f.write(decompressed_data)
                    
                    # Затем извлекаем все остальные файлы, кроме .data
                    for file in file_list:
                        if not file.endswith(".data"):
                            # Извлекаем файл в выходной каталог
                            zip.extract(file, output_dir)
                            self.log_message(f"Извлечение файла {file}")
                
                self.log_message("Извлечение завершено успешно")
            else:
                messagebox.showerror("Ошибка", "Неверный формат файла")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def handle_drop(self, event):
        """Обработка перетаскивания файлов и каталогов"""
        # Получаем путь из события
        data = event.data
        
        # В Windows путь может быть в фигурных скобках и содержать несколько файлов
        if data.startswith('{'):
            files = data.strip('{}').split('} {')
        else:
            files = [data]
            
        # Обрабатываем только первый перетащенный элемент
        path = files[0]
        
        if os.path.isfile(path):
            # Если это файл, просто устанавливаем его путь
            self.file_path.set(path)
        elif os.path.isdir(path):
            # Если это каталог, ищем в нем JSON файлы
            json_files = []
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.lower().endswith('.json'):
                        json_files.append(os.path.join(root, file))
                        
            if len(json_files) == 0:
                messagebox.showwarning("Предупреждение", "В каталоге не найдено JSON файлов")
            elif len(json_files) > 1:
                messagebox.showerror("Ошибка", "В каталоге найдено больше одного JSON файла")
            else:
                # Найден ровно один JSON файл
                self.file_path.set(json_files[0])

    def show_package_info(self):
        """Показывает окно с информацией о пакете"""
        try:
            input_file = self.file_path.get()
            if not input_file or not zipfile.is_zipfile(input_file):
                messagebox.showerror("Ошибка", "Выберите корректный SOP файл")
                return

            with zipfile.ZipFile(input_file, 'r') as zip:
                # Ищем .data файл
                data_files = [f for f in zip.namelist() if f.endswith('.data')]
                if not data_files:
                    messagebox.showerror("Ошибка", "SOP файл не содержит .data файлов")
                    return

                # Берем первый .data файл
                with zip.open(data_files[0]) as data_file:
                    # Распаковываем и декодируем JSON
                    json_data = zlib.decompress(data_file.read())
                    json_data = json.loads(json_data.decode('utf-8'))

                    # Создаем окно с информацией
                    PackageInfoWindow(self.root, json_data)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при чтении информации о пакете: {str(e)}")

    def view_json(self):
        """Показывает содержимое JSON из текущего файла (SOP или JSON)"""
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("Ошибка", "Файл не выбран")
            return
        
        if file_path.lower().endswith('.sop'):
            self.view_json_from_sop()
        elif file_path.lower().endswith('.json'):
            self.view_json_file(file_path)
        else:
            messagebox.showerror("Ошибка", "Неподдерживаемый формат файла")

    def view_json_from_sop(self):
        """Показывает JSON из SOP файла"""
        try:
            input_file = self.file_path.get()
            if not input_file or not zipfile.is_zipfile(input_file):
                messagebox.showerror("Ошибка", "Выберите корректный SOP файл")
                return

            with zipfile.ZipFile(input_file, 'r') as zip:
                # Ищем .data файл
                data_files = [f for f in zip.namelist() if f.endswith('.data')]
                if not data_files:
                    messagebox.showerror("Ошибка", "SOP файл не содержит .data файлов")
                    return

                # Берем первый .data файл
                with zip.open(data_files[0]) as data_file:
                    # Распаковываем и декодируем JSON
                    json_data = zlib.decompress(data_file.read())
                    json_data = json.loads(json_data.decode('utf-8'))
                    
                    # Создаем окно для просмотра JSON
                    self.show_json_viewer(json_data)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при чтении JSON: {str(e)}")

    def view_json_file(self, file_path):
        """Показывает содержимое JSON файла в отдельном окне"""
        try:
            if not self.validate_json_file(file_path):
                return
            
            with open(file_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
                self.show_json_viewer(json_data)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при чтении JSON файла: {str(e)}")

    def show_json_viewer(self, json_data):
        """Показывает окно просмотра JSON"""
        JsonViewerWindow(self.root, json_data)

if __name__ == '__main__':
    if ENABLE_DND:
        root = TkinterDnD.Tk()  # Используем TkinterDnD вместо обычного Tk
    else:
        root = tk.Tk()  # Используем обычный Tk если tkinterdnd2 недоступен
    app = SopUtilityGUI(root)
    root.mainloop()
