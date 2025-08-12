# -*- coding: UTF-8 -*-

"""
SOP File Utility

Утилита командной строки для работы с SOP-файлами (специальный формат архива).
Позволяет:
- Извлекать JSON-файлы из SOP архива
- Создавать SOP архивы из JSON файлов 
- Извлекать все файлы из архива
- Просматривать содержимое архива

Использование:
    sop-utility.py [-h] [-v] {e,a,x,l} ...

Аргументы:
    -h, --help     Показать справку
    -v, --version  Показать версию программы
    
Команды:
    e    Извлечь только JSON файлы
    a    Создать SOP архив из JSON
    x    Извлечь все файлы из архива 
    l    Показать список файлов в архиве
"""

import sys
import argparse
import os
import datetime
import zipfile
import zlib
import json
import traceback

version = "1.0.0"

# Дефолтные значения для app_mappings
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

class SOPError(Exception):
    """Базовый класс для ошибок утилиты SOP"""
    pass

class FileOperationError(SOPError):
    """Ошибка при работе с файлом"""
    pass

class DirectoryOperationError(SOPError):
    """Ошибка при работе с директорией"""
    pass

class InvalidFileFormatError(SOPError):
    """Некорректный формат файла"""
    pass

def normalize_path(path):
    """
    Нормализует путь, преобразуя его в абсолютный путь и корректно обрабатывая разделители.
    
    Args:
        path: Путь для нормализации
        
    Returns:
        Нормализованный абсолютный путь с правильными разделителями для текущей ОС
        
    Raises:
        DirectoryOperationError: Если путь не может быть нормализован
    """
    try:
        return os.path.normpath(os.path.abspath(path))
    except Exception as e:
        raise DirectoryOperationError(f"Error normalizing path '{path}': {str(e)}")

def normalize_archive_path(path):
    """
    Нормализует путь внутри архива, заменяя обратные слеши на прямые
    и удаляя лишние разделители.
    
    Args:
        path: Путь для нормализации
        
    Returns:
        Нормализованный путь для использования в архиве
    """
    # Заменяем обратные слеши на прямые
    path = path.replace('\\', '/')
    # Убираем повторяющиеся слеши
    while '//' in path:
        path = path.replace('//', '/')
    # Убираем начальный слеш
    if path.startswith('/'):
        path = path[1:]
    return path

def ensure_directory_exists(dir_path):
    """
    Создает директорию, если она не существует
    
    Args:
        dir_path: Путь к директории
        
    Raises:
        DirectoryOperationError: Если директория не может быть создана
    """
    try:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
        elif not os.path.isdir(dir_path):
            raise DirectoryOperationError(f"Path '{dir_path}' exists but is not a directory")
    except Exception as e:
        raise DirectoryOperationError(f"Error creating directory '{dir_path}': {str(e)}")

def load_app_mappings():
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
                json.dump(DEFAULT_APP_MAPPINGS, f, indent=4, ensure_ascii=False)
            return DEFAULT_APP_MAPPINGS
    except Exception as e:
        print(f"Warning: Could not load app_mappings.json: {str(e)}", file=sys.stderr)
        return DEFAULT_APP_MAPPINGS
        
# Загружаем маппинги при старте программы
app_mappings = load_app_mappings()

def get_app_name(pack_application_id):
    """Возвращает имя приложения по его ID"""
    for mapping in app_mappings:
        if str(mapping.get('pack_application_id')) == str(pack_application_id):
            return mapping.get('AppName', '')
    return ''

def print_json_fields(data, prefix=""):
    if isinstance(data, dict):
        for key, value in data.items():
            new_prefix = f"{prefix}.{key}" if prefix else key
            print(new_prefix)  # Выводим текущее поле
            print_json_fields(value, new_prefix)  # Рекурсивно обходим вложенные поля
    elif isinstance(data, list):
        for i, item in enumerate(data):
            print_json_fields(item, f"{prefix}[{i}]")  # Обрабатываем элементы списка

def show_info(args):
    """
    Показывает информацию о SOP файле
    
    Args:
        args: Аргументы командной строки
        
    Raises:
        FileOperationError: При ошибках работы с файлами
        InvalidFileFormatError: При неверном формате файла
    """
    try:
        # Нормализуем входной путь
        input_path = normalize_path(args.input)
        
        # Проверяем существование входного файла
        if not os.path.exists(input_path):
            raise FileOperationError(f"Input file '{input_path}' does not exist")
        if not os.path.isfile(input_path):
            raise FileOperationError(f"'{input_path}' is not a file")

        # Проверяем, является ли файл ZIP архивом
        if not zipfile.is_zipfile(input_path):
            raise InvalidFileFormatError(f"File '{input_path}' is not a SOP file")

        try:
            with zipfile.ZipFile(input_path, "r") as zip:
                # Ищем .data файл
                data_files = [f for f in zip.namelist() if f.endswith('.data')]
                if not data_files:
                    raise InvalidFileFormatError("No .data file found in SOP archive")
                
                # Берем первый .data файл
                data_file = data_files[0]
                
                try:
                    # Читаем и декомпрессируем данные
                    with zip.open(data_file) as compressed_data:
                        decompressed_data = zlib.decompress(compressed_data.read())
                        json_data = json.loads(decompressed_data.decode('utf-8'))

                        # Извлекаем нужные поля
                        name = json_data.get('name', 'N/A')
                        description = json_data.get('description', 'N/A')
                        pack_application_id = json_data.get('pack_application_id', 'N/A')
                        records = json_data.get('records', [])
                        records_count = len(records) if isinstance(records, list) else 0
                        
                        # Получаем имя приложения
                        app_name = get_app_name(pack_application_id)
                        
                        # Выводим информацию
                        print(f"SOP File Info: {input_path}\n{'='*80}")
                        print(f"{'Name:':<20} {name}")
                        print(f"{'Description:':<20} {description}")
                        print(f"{'Application ID:':<20} {pack_application_id}")
                        print(f"{'Application Name:':<20} {app_name}")
                        print(f"{'Records count:':<20} {records_count}")
                        print("="*80)
                        
                except zlib.error as e:
                    raise FileOperationError(f"Decompression error: {str(e)}")
                except json.JSONDecodeError as e:
                    raise InvalidFileFormatError(f"Invalid JSON format in .data file: {str(e)}")
                except Exception as e:
                    raise FileOperationError(f"Error processing .data file: {str(e)}")
        
        except zipfile.BadZipFile:
            raise InvalidFileFormatError(f"File '{input_path}' is corrupted or not a SOP file")
        except Exception as e:
            raise FileOperationError(f"Error working with SOP file '{input_path}': {str(e)}")

    except SOPError:
        raise
    except Exception as e:
        raise SOPError(f"Unknown error while showing info: {str(e)}")

def extract_sop(args):
    """
    Извлекает JSON файлы из SOP архива
        
    Args:
        args: Аргументы командной строки

        
    Raises:
        FileOperationError: При ошибках работы с файлами
        DirectoryOperationError: При ошибках работы с директориями
        InvalidFileFormatError: При неверном формате файла
    """
    try:
        # Нормализуем входной путь
        input_path = normalize_path(args.input)
        
        # Проверяем существование входного файла
        if not os.path.exists(input_path):
            raise FileOperationError(f"Input file '{input_path}' does not exist")
        if not os.path.isfile(input_path):
            raise FileOperationError(f"'{input_path}' is not a file")

        # Определяем директорию для выходных файлов
        output_dir = normalize_path(args.output) if args.output else os.path.dirname(input_path)
        
        try:
            ensure_directory_exists(output_dir)
        except DirectoryOperationError as e:
            raise DirectoryOperationError(f"Failed to prepare output directory: {str(e)}")

        # Проверяем, является ли файл ZIP архивом
        if not zipfile.is_zipfile(input_path):
            raise InvalidFileFormatError(f"File '{input_path}' is not a SOP file")

        try:
            with zipfile.ZipFile(input_path, "r") as zip:
                # Перебираем все файлы в архиве
                for file in zip.namelist():
                    # Обрабатываем только .data файлы
                    if file.endswith(".data"):
                        try:
                            # Заменяем расширение .data на .json
                            out_file = file.replace('.data', '.json')
                            # Для Windows заменяем двоеточие на подчеркивание
                            if os.name == 'nt':
                                out_file = out_file.replace(':', '_')
                            
                            # Формируем полный путь выходного файла
                            out_file_path = os.path.join(output_dir, out_file)

                            # Распаковываем и декомпрессируем данные
                            with zip.open(file) as compressed_data:
                                try:
                                    decompressed_data = zlib.decompress(compressed_data.read())
                                    print(f"Extracting JSON file to {out_file_path}")
                                    
                                    try:
                                        with open(out_file_path, 'wb') as f:
                                            f.write(decompressed_data)
                                    except Exception as e:
                                        raise FileOperationError(f"Error writing to file '{out_file_path}': {str(e)}")
                                
                                except zlib.error as e:
                                    raise FileOperationError(f"Decompression error for file '{file}': {str(e)}")
                                except Exception as e:
                                    raise FileOperationError(f"Error reading file '{file}' from archive: {str(e)}")
                        
                        except Exception as e:
                            print(f"Error processing file '{file}': {str(e)}", file=sys.stderr)
                            continue
        
        except zipfile.BadZipFile:
            raise InvalidFileFormatError(f"File '{input_path}' is corrupted or not a SOP file")
        except Exception as e:
            raise FileOperationError(f"Error working with SOP file '{input_path}': {str(e)}")

    except SOPError:
        raise
    except Exception as e:
        raise SOPError(f"Unknown error during JSON extraction: {str(e)}")

def create_sop(args):
    """
    Создает SOP архив из JSON файла
    
    Args:
        args: Аргументы командной строки

        
    Raises:
        FileOperationError: При ошибках работы с файлами
        DirectoryOperationError: При ошибках работы с директориями
        InvalidFileFormatError: При неверном формате файла
    """
    try:
        # Нормализуем входной путь
        input_path = normalize_path(args.input)
        
        # Проверяем существование входного файла
        if not os.path.exists(input_path):
            raise FileOperationError(f"Input file '{input_path}' does not exist")
        if not os.path.isfile(input_path):
            raise FileOperationError(f"'{input_path}' is not a file")

        # Проверяем расширение файла
        if not input_path.endswith(".json"):
            raise InvalidFileFormatError(f"File '{input_path}' must have .json extension")

        # Определяем директорию для выходных файлов
        output_dir = normalize_path(args.output) if args.output else os.path.dirname(input_path)
        
        try:
            ensure_directory_exists(output_dir)
        except DirectoryOperationError as e:
            raise DirectoryOperationError(f"Failed to prepare output directory: {str(e)}")

        # Получаем имя входного файла
        input_file = os.path.basename(input_path)

        # Формируем имена выходных файлов
        data_file = input_file.replace('.json', '.data')
        zip_out = input_file.replace('.json', '.sop')
        zip_out_path = os.path.join(output_dir, zip_out)

        # Проверяем, не существует ли уже выходной файл
        if os.path.exists(zip_out_path):
            raise FileOperationError(f"Output file '{zip_out_path}' already exists")

        try:
            # Читаем и сжимаем данные
            with open(input_path, 'rb') as fi:
                try:
                    compressed_data = zlib.compress(fi.read(), level=5)
                except Exception as e:
                    raise FileOperationError(f"Data compression error: {str(e)}")

            print(f"Creating SOP file: {zip_out_path}")
            
            try:
                with zipfile.ZipFile(zip_out_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=5) as zip:
                    # Добавляем сжатый JSON как .data файл
                    zip.writestr(data_file, compressed_data)
                    
                    # Если указан параметр --add_files, добавляем все файлы из директории
                    if hasattr(args, 'add_files') and args.add_files:
                        input_dir = os.path.dirname(input_path)
                        for root, dirs, files in os.walk(input_dir):
                            for file in files:
                                file_path = os.path.join(root, file)
                                # Исключаем исходный JSON файл
                                if file_path != input_path:                                    # Создаем относительный путь для архива
                                    rel_path = os.path.relpath(file_path, input_dir)
                                    # Нормализуем путь для архива
                                    archive_path = normalize_archive_path(rel_path)
                                    print(f"Adding file: {archive_path}")
                                    zip.write(file_path, archive_path)
            except Exception as e:
                raise FileOperationError(f"Error creating SOP file: {str(e)}")
        
        except Exception as e:
            # Удаляем частично созданный файл при ошибке
            if os.path.exists(zip_out_path):
                try:
                    os.remove(zip_out_path)
                except:
                    pass
            raise

    except SOPError:
        raise
    except Exception as e:
        raise SOPError(f"Unknown error during SOP creation: {str(e)}")

def extract_archive(args):
    """
    Извлекает все файлы из SOP архива
    
    Args:
        args: Аргументы командной строки

        
    Raises:
        FileOperationError: При ошибках работы с файлами
        DirectoryOperationError: При ошибках работы с директориями
        InvalidFileFormatError: При неверном формате файла
    """
    try:
        # Нормализуем входной путь
        input_path = normalize_path(args.input)
        
        # Проверяем существование входного файла
        if not os.path.exists(input_path):
            raise FileOperationError(f"Input file '{input_path}' does not exist")
        if not os.path.isfile(input_path):
            raise FileOperationError(f"'{input_path}' is not a file")

        # Определяем директорию для выходных файлов
        output_dir = normalize_path(args.output) if args.output else os.path.dirname(input_path)
        
        try:
            ensure_directory_exists(output_dir)
        except DirectoryOperationError as e:
            raise DirectoryOperationError(f"Failed to prepare output directory: {str(e)}")

        # Проверяем, является ли файл ZIP архивом
        if not zipfile.is_zipfile(input_path):
            raise InvalidFileFormatError(f"File '{input_path}' is not a SOP file")

        try:
            print(f"Extracting all files from archive {input_path}")
            with zipfile.ZipFile(input_path, "r") as zip:
                if args.format == 'json':
                    # Режим извлечения с конвертацией .data в .json
                    for file in zip.namelist():
                        if file.endswith(".data"):
                            try:                                # Заменяем расширение .data на .json и нормализуем путь
                                out_file = normalize_archive_path(file.replace('.data', '.json'))
                                if os.name == 'nt':
                                    out_file = out_file.replace(':', '_')
                                
                                out_file_path = os.path.join(output_dir, out_file)

                                # Распаковываем и декомпрессируем данные
                                with zip.open(file) as compressed_data:
                                    try:
                                        decompressed_data = zlib.decompress(compressed_data.read())
                                        print(f"Extracting JSON file to {out_file_path}")
                                        
                                        with open(out_file_path, 'wb') as f:
                                            f.write(decompressed_data)
                                    except zlib.error as e:
                                        raise FileOperationError(f"Decompression error for file '{file}': {str(e)}")
                            except Exception as e:
                                print(f"Error processing file '{file}': {str(e)}", file=sys.stderr)
                                continue
                        else:
                            # Обычные файлы извлекаем как есть
                            zip.extract(file, output_dir)
                else:
                    # Режим извлечения как есть (--format=data)
                    try:
                        zip.extractall(output_dir)
                    except Exception as e:
                        raise FileOperationError(f"Error extracting files from archive: {str(e)}")
        
        except zipfile.BadZipFile:
            raise InvalidFileFormatError(f"File '{input_path}' is corrupted or not a SOP file")
        except Exception as e:
            raise FileOperationError(f"Error working with SOP file '{input_path}': {str(e)}")

    except SOPError:
        raise
    except Exception as e:
        raise SOPError(f"Unknown error during archive extraction: {str(e)}")

def list_archive(args):
    """
    Показывает список файлов в SOP архиве
    
    Args:
        args: Аргументы командной строки

        
    Raises:
        FileOperationError: При ошибках работы с файлами
        InvalidFileFormatError: При неверном формате файла
    """
    try:
        # Нормализуем входной путь
        input_path = normalize_path(args.input)
        
        # Проверяем существование входного файла
        if not os.path.exists(input_path):
            raise FileOperationError(f"Input file '{input_path}' does not exist")
        if not os.path.isfile(input_path):
            raise FileOperationError(f"'{input_path}' is not a file")

        # Проверяем, является ли файл ZIP архивом
        if not zipfile.is_zipfile(input_path):
            raise InvalidFileFormatError(f"File '{input_path}' is not a SOP file")

        try:
            with zipfile.ZipFile(input_path, "r") as zip:
                print(f"Files in archive: {input_path}\n{'='*80}")
                try:
                    print(f"{'Filename':<48} {'Size (bytes)':>10} {'Date':>18}")
                    print('=' * 80)
                    
                    # Print each file as a table row
                    for file in zip.infolist():
                        date = datetime.datetime(*file.date_time)
                        print(f"{file.filename:<50} {file.file_size:>10} {date.strftime('%Y.%m.%d %H:%M'):>18}")
                except Exception as e:
                    raise FileOperationError(f"Error reading archive contents: {str(e)}")
        
        except zipfile.BadZipFile:
            raise InvalidFileFormatError(f"File '{input_path}' is corrupted or not a SOP file")
        except Exception as e:
            raise FileOperationError(f"Error working with SOP file '{input_path}': {str(e)}")

    except SOPError:
        raise
    except Exception as e:
        raise SOPError(f"Unknown error while viewing archive: {str(e)}")

def createParser():
    """
    Создает парсер аргументов командной строки
    
    Returns:
        parser: Объект парсера
    """
    parser = argparse.ArgumentParser(
        description = 'SOP file utility',
        epilog = '(c) AEN 2024-2025. The author of the program, as always, does not bear any responsibility for anything.'
        )
    parser.add_argument ('-v', '--version', action='version', help = 'Output the version number', version='%(prog)s {}'.format(version))

    subparsers = parser.add_subparsers(title='commands', description='valid commands')
    
    # Парсер для извлечения JSON
    extract_sop_parser = subparsers.add_parser('e', help = 'extract json only')
    extract_sop_parser.add_argument ('-i', '--input', required=True, help = 'file to process', metavar = 'filename.sop')
    extract_sop_parser.add_argument ('-o', '--output', default='', help = 'path to write the result', metavar = 'path/to/extracted')
    extract_sop_parser.set_defaults(func=extract_sop)

    # Парсер для создания архива
    create_sop_parser = subparsers.add_parser('a', help = 'arhive json to sop')
    create_sop_parser.add_argument ('-i', '--input', required=True, help = 'file to process', metavar = 'filename.json')
    create_sop_parser.add_argument ('-o', '--output', default='', help = 'path to write the result', metavar = 'path/to/output')
    create_sop_parser.add_argument ('--add-files', action='store_true', help = 'include all files from input directory')
    create_sop_parser.set_defaults(func=create_sop)

    # Парсер для извлечения всего архива
    extract_archive_parser = subparsers.add_parser('x', help = 'extract all files from sop')
    extract_archive_parser.add_argument ('-i', '--input', required=True, help = 'file to process', metavar = 'filename.sop')
    extract_archive_parser.add_argument ('-o', '--output', default='', help = 'path to write the result', metavar = 'path/to/extracted')
    extract_archive_parser.add_argument('--format', choices=['data', 'json'], default='json',
                                      help='extraction format: "data" (as-is) or "json" (convert .data to .json)')
    extract_archive_parser.set_defaults(func=extract_archive)

    # Парсер для просмотра содержимого
    list_archive_parser = subparsers.add_parser('l', help = 'list of all files in sop')
    list_archive_parser.add_argument ('-i', '--input', required=True, help = 'file to process', metavar = 'filename.sop')
    list_archive_parser.set_defaults(func=list_archive)

    # Парсер для показа информации
    info_parser = subparsers.add_parser('i', help = 'show SOP file information')
    info_parser.add_argument ('-i', '--input', required=True, help = 'file to process', metavar = 'filename.sop')
    info_parser.set_defaults(func=show_info)
        
    return parser

def main():
    """Основная функция для обработки ошибок верхнего уровня"""
    try:
        parser = createParser()
        args = parser.parse_args(sys.argv[1:])

        if not vars(args):
            parser.print_help()
            sys.exit(0)
        
        args.func(args)
    
    except SOPError as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {str(e)}", file=sys.stderr)
        if '--debug' in sys.argv:
            traceback.print_exc()
        sys.exit(2)

if __name__ == '__main__':
    main()
