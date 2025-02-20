import subprocess
from pathlib import Path
import xlwings as xw
from PIL import Image
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def extract_frame(video_path: Path, output_frame_path: Path, time="00:00:01"):
    """
    Извлекает кадр из видео в указанный момент времени.
    Параметры:
    - video_path: путь к видеофайлу (объект Path)
    - output_frame_path: путь к выходному изображению (объект Path)
    - time: время, на котором берётся кадр (формат HH:MM:SS)
    """
    cmd = [
        "ffmpeg",
        "-i", str(video_path),
        "-ss", time,
        "-frames:v", "1",
        "-y", str(output_frame_path)
    ]
    subprocess.run(cmd, stdout=subprocess.PIPE,
                   stderr=subprocess.PIPE, check=True)


def resize_image(image_path: Path, max_width: int = 480, max_height: int = 270):
    """
    Масштабирует изображение, чтобы оно помещалось в указанные размеры,
    сохраняя пропорции.
    """
    try:
        with Image.open(image_path) as img:
            # Конвертируем в RGB, если изображение в другом формате
            if img.mode != 'RGB':
                img = img.convert('RGB')

            # Получаем текущие размеры
            width, height = img.size

            # Вычисляем коэффициент масштабирования
            ratio = min(max_width / width, max_height / height)

            # Новые размеры
            new_width = int(width * ratio)
            new_height = int(height * ratio)

            # Масштабируем
            resized_img = img.resize((new_width, new_height), Image.LANCZOS)
            resized_img.save(image_path)

            return new_width, new_height  # Возвращаем новые размеры
    except Exception as e:
        print(f"Ошибка при масштабировании изображения {image_path}: {str(e)}")
        return None, None


def select_folder():
    """Открывает диалоговое окно для выбора папки"""
    root = tk.Tk()
    root.withdraw()  # Скрываем основное окно
    folder_path = filedialog.askdirectory(
        title="Выберите папку с видеофайлами")
    return Path(folder_path) if folder_path else None


def select_excel_file():
    """Открывает диалоговое окно для выбора Excel файла со списком камер"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Выберите Excel файл со списком камер",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return Path(file_path) if file_path else None


def load_camera_list(excel_path: Path):
    """
    Загружает список камер из Excel файла и возвращает их в нижнем регистре.
    Предполагается, что список находится на первом листе в первом столбце.
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=0)
        # Преобразуем в нижний регистр и удаляем дубликаты
        return set(str(camera).lower() for camera in df.iloc[:, 0])
    except Exception as e:
        print(f"Ошибка при чтении файла со списком камер: {str(e)}")
        return set()


def clean_camera_folder_names(input_folder: Path):
    """
    Переименовывает папки камер, оставляя в названии только первую часть до первого пробела.
    Например: "PHP-URSK-ShV1-K5 Склад кислот Оси В, 21-22 отм. +1,100" → "PHP-URSK-ShV1-K5".
    Если папка с новым именем уже существует, выводится предупреждение.
    """
    for subfolder in input_folder.iterdir():
        if subfolder.is_dir():
            new_name = subfolder.name.split(' ')[0]
            if subfolder.name != new_name:
                new_path = subfolder.parent / new_name
                if new_path.exists():
                    print(
                        f"Невозможно переименовать {subfolder.name} в {new_name}, так как папка {new_path} уже существует.")
                    continue
                print(f"Переименование папки {subfolder.name} -> {new_name}")
                subfolder.rename(new_path)


def create_excel_with_images(input_folder: Path, excel_path: Path, valid_cameras: set):
    """
    Создаёт Excel-файл с кадрами из всех видеофайлов в папке.
    Параметры:
    - input_folder: путь к папке с видеофайлами (Path)
    - excel_path: путь к выходному excel-файлу (Path)
    - valid_cameras: множество допустимых названий камер в нижнем регистре
    """
    # Находим все видеофайлы в папке и подпапках
    video_extensions = ('.mkv', '.mp4', '.avi')
    video_files = []
    for ext in video_extensions:
        video_files.extend(input_folder.rglob(f'*{ext}'))

    if not video_files:
        print("Видеофайлы не найдены в указанной папке!")
        return

    # Создаём новый Workbook
    with xw.App(visible=False) as app:
        wb = app.books.add()
        ws = wb.sheets[0]
        ws.name = 'Камеры'

        # Создаем вспомогательный лист для списка значений "Да/Нет"
        ws_validation = wb.sheets.add('Validation', after=ws.name)
        ws_validation.range('A1').value = 'Да'
        ws_validation.range('A2').value = 'Нет'

        # Заполняем шапку таблицы
        headers = [
            "Имя камеры", "Кадр с камеры", "Камеру в СОВА?",
            "СИЗ-перчатки", "СИЗ-перчатки-описание",
            "СИЗ-очки", "СИЗ-очки-описание",
            "СИЗ-распиратор", "СИЗ-распиратор-описание",
            "СИЗ-газоанализатор", "СИЗ-газоан.-описание",
            "СИЗ-самоспасатель", "СИЗ-самосп.-описание",
            "Опасная зона", "Опасная зона-описание"
        ]
        range_A1 = ws.range("A1")

        range_A1.value = headers
        range_A1.column_width = 20

        # Установка горизонтального выравнивания по центру
        range_A1.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

        # Установка вертикального выравнивания по центру
        range_A1.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter

        # Обрабатываем каждый видеофайл
        current_row = 2
        frame_paths = []

        for video_path in video_files:
            # Получаем имя родительской папки, оставляем только первую часть до пробела
            parent_folder_name = video_path.parent.name.split(' ')[0].lower()

            # Проверяем, есть ли камера в списке допустимых
            if parent_folder_name not in valid_cameras:
                print(
                    f"Камера {parent_folder_name} не найдена в списке, пропускаем...")
                continue

            print(f"Обработка файла: {video_path}")

            # Создаём уникальное имя для кадра, используя имя родительской папки
            parent_folder_name_for_filename = video_path.parent.name.replace(
                ' ', '_')
            frame_filename = f"frame_{parent_folder_name_for_filename}_{video_path.stem}.png"
            frame_path = video_path.parent / frame_filename
            frame_paths.append(frame_path)

            try:
                # Извлекаем кадр
                extract_frame(video_path, frame_path, time="00:00:03")

                # Масштабируем изображение и получаем новые размеры
                new_width, new_height = resize_image(frame_path)

                if new_width is None or new_height is None:
                    continue

                # Записываем имя камеры
                ws.range(f"A{current_row}").value = video_path.parent.name.split(
                    ' ')[0]

                # Настраиваем размеры ячейки для изображения
                ws.range(f"B{current_row}").row_height = 280
                ws.range(f"B{current_row}").column_width = 88

                # Добавляем картинку с использованием реальных размеров изображения
                left = ws.range(f"B{current_row}").left
                top = ws.range(f"B{current_row}").top
                ws.pictures.add(str(frame_path),
                                name=f"Frame_{parent_folder_name_for_filename}_{video_path.stem}",
                                left=left,
                                top=top,
                                width=new_width,
                                height=new_height)

                current_row += 1

            except Exception as e:
                print(f"Ошибка при обработке {video_path}: {str(e)}")
                continue

        if current_row == 2:  # Если не добавлено ни одной камеры
            print("Не найдено подходящих камер для обработки!")
            wb.close()
            return

        # Создаем умную таблицу
        table = ws.tables.add(source=ws['A1'].expand(),
                              name="ТаблицаКамер",
                              has_headers=True,
                              destination=None,
                              table_style_name='TableStyleMedium15')
        # Убираем автофильтры
        table.show_autofilter = False

        # Добавляем проверку данных (выпадающие списки "Да/Нет" со ссылкой на лист Validation)
        validation_columns = ['C', 'D', 'F', 'H', 'J', 'L', 'N']
        for col in validation_columns:
            col_range = f"{col}2:{col}{current_row-1}"
            ws.range(col_range).api.Validation.Add(
                Type=3,  # xlValidateList
                AlertStyle=1,  # xlValidAlertStop
                Operator=1,  # xlBetween
                Formula1="=Validation!$A$1:$A$2"
            )

        # переходим на главный лист
        ws.activate()

        # скрываем лист с выпадающим списком
        ws_validation.visible = False

        # Сохраняем файл
        wb.save(excel_path)
        wb.close()

        # Удаляем временные файлы с кадрами
        for frame_path in frame_paths:
            try:
                frame_path.unlink(missing_ok=True)
            except Exception as e:
                print(f"Ошибка при удалении {frame_path}: {str(e)}")


if __name__ == "__main__":
    # Выбор папки с видеофайлами через диалоговое окно
    input_folder = select_folder()
    if not input_folder:
        print("Папка не выбрана!")
        exit()

    if not input_folder.exists() or not input_folder.is_dir():
        print("Указанная папка не существует или не является папкой!")
        exit()

    # Выбор Excel файла со списком камер
    camera_list_file = select_excel_file()
    if not camera_list_file:
        print("Файл со списком камер не выбран!")
        exit()

    # Загружаем список допустимых камер
    valid_cameras = load_camera_list(camera_list_file)
    if not valid_cameras:
        print("Список камер пуст или не удалось загрузить!")
        exit()

    # Переименовываем папки камер, если в их названиях есть лишние символы
    clean_camera_folder_names(input_folder)

    # Путь к выходному Excel-файлу
    output_excel_path = input_folder.parent / "results.xlsx"

    # Создаём Excel-файл с кадрами
    create_excel_with_images(input_folder, output_excel_path, valid_cameras)
    print(f"Excel-файл создан: {output_excel_path}")
