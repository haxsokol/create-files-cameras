""" 
тут смотрел, сколько у нас есть видео с нормальными названиями, чтобы
можно было автоматом сделать скрин и прикрепить его в Excel.
По итогу оказалось, что камер 276, а надо ~1300
"""

from pathlib import Path
import pandas as pd
import re
from tqdm import tqdm

# Введите путь к папке, в которой нужно искать видео
folder_input = input("Введите путь к папке: ").strip('" ')
base_path = Path(folder_input)

# Задайте расширение видео, которое нужно искать (например, '.avi')
pat_videos = ["*.avi", '*.mkv', '*.mp4']  # можно заменить на нужное расширение

# Список для хранения названий папок (на один уровень выше, где находится видео)
folder_names = []
p1 = Path(
    "S:\Северсталь Диджитал\СОВА-платформа ОТиПБ_Общий сетевой ресурс\Передача видео")
p2 = Path("S:\Северсталь Диджитал\СОВА-платформа ОТиПБ_Общий сетевой ресурс\Выгрузка видеороликов для обучения")
p3 = Path("S:\Череповец\Северсталь Менеджмент\ОТиПБ Цифровизация\Эксперимент с СИЗ")

# Рекурсивно обходим все файлы и папки в base_path
for the_path in tqdm([i for i in p3.glob('*') if i.is_dir()]):
    for pattern in pat_videos:
        for file in the_path.rglob(pattern):
            # Проверяем, что найденный объект является файлом и имеет нужное расширение
            if file.is_file() and bool(re.match(r'^[A-Za-z]', file.parent.name)):
                # Добавляем название родительской папки (на один уровень выше файла)
                folder_names.append(file.parent.name.split(' ')[0].lower())

# Если нужно, можно оставить только уникальные названия:
folder_names = list(set(folder_names))

# Создаем DataFrame из списка папок
df = pd.DataFrame({'folder': folder_names})
df.to_excel("СпискоКамерЭкспериментыСИЗ.xlsx", index=False)

l = []
for j in ['СпискоКамерДляОбучения.xlsx', 'СпискоКамерПередачаВидео.xlsx', 'СпискоКамерЭкспериментыСИЗ.xlsx']:
    df_ex = pd.read_excel(j)
    l.extend(df_ex.folder.to_list())
pd.DataFrame({'CamName': list(set(l))}).to_excel(
    'ВсеКамерыИзКоторыхМожноВзятьСкрины.xlsx', index=False)
