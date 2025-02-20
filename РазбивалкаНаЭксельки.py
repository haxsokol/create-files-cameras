""" делим Excel файл со всеми 1299 камерами с ЛСР на отдельные Excel-файлы,
чтобы потом делать списки скринов
"""

import pandas as pd
from pathlib import Path

p = Path('КамерыЛСРпоПроизводствам')
p.absolute() / "1.xlsx"


df_path = Path(r"C:\work\Скрипты python\1299 камер с ЛСР.xlsx")
df = pd.read_excel(df_path, sheet_name='Список камер').convert_dtypes()

df.Цех.value_counts()

df.Цех = df.apply(lambda x: x["Производство"] if x["Цех"] ==
                  "Участки(Ур-нь цеха не задан)" else x["Цех"], axis=1)
df["ПроизвЦех"] = df.apply(lambda x: x["Производство"] + "_"+x["Цех"], axis=1)
df["ПроизвЦех"] = df["ПроизвЦех"].replace(r'[<>:"/\\|?*]', '', regex=True)

for i in df.ПроизвЦех.value_counts().index.to_list():
    df_split = df.query(f"ПроизвЦех == '{i}'")[["Имя камеры"]]
    path_excel = p.absolute() / f"{i}.xlsx"
    df_split.to_excel(path_excel, index=False)
