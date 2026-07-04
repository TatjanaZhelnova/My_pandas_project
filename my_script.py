import pandas as pd
from pathlib import Path
from datetime import datetime
import glob


# 1. Пути к папкам с выгрузками

sbis_folder = Path('Входящие')
aptiki_folder = Path('Аптеки/csv/correct')

today = datetime.today().strftime('%Y-%m-%d') # полная сегодняшняя дата
result_folder = Path('Результат') / today # папка для сохранения результатов
result_folder.mkdir(parents=True, exist_ok=True) # создаем папку для сохранения результатов, если ее нет


# 2. Обработка выгрузки из СБИС

sbis_dfs = []

sbis_files = glob.glob('Входящие/*.csv')


for file in sbis_files:
    
    try: 
        df = pd.read_csv(file, encoding='1251', sep=';', decimal=',')
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(file,  encoding="utf-8", sep=';', decimal=',')
        except Exception:
            print(f"Не удалось прочитать {file} — пропускаем")
            continue
    
    sbis_dfs.append(df)

sbis = pd.concat(sbis_dfs, ignore_index=True)

# Присваиваем столбцам полученного датафрейма требуемые названия
sbis.columns = [
    "Дата",
    "Номер",
    "Сумма",
    "Статус",
    "Примечание",
    "Комментарий",
    "Контрагент",
    "ИНН/КПП",
    "Организация",
    "ИНН/КПП",
    "Тип документа",
    "Имя файла",
    "Дата",
    "Номер 1",
    "Сумма 1",
    "Сумма НДС",
    "Ответственный",
    "Подразделение",
    "Код",
    "Дата",
    "Время",
    "Тип пакета",
    "Идентификатор пакета",
    "Запущено в обработку",
    "Получено контрагентом",
    "Завершено",
    "Увеличение суммы",
    "НДC",
    "Уменьшение суммы",
    "НДС",
]
sbis.columns = sbis.columns.str.replace(' ', '_') # заменяем пробелы в названиях столбцов на подчеркивания

# фильтруем по типу документа и удаляем дубликаты по номеру документа, оставляя только первую запись

need_types = [
    "СчФктр",
    "УпдДоп",
    "УпдСчфДоп",
    "ЭДОНакл"
    ]
    
sbis = (
    sbis[sbis['Тип_документа'].isin(need_types)]
    .drop_duplicates(subset='Номер', keep='first')
    )


# 3. Обработка выгрузки из АПТЕКИ 

apteka_files = glob.glob('Аптеки/csv/correct/*.csv')

for file in apteka_files:
    try: 
        apteka = pd.read_csv(file, encoding='1251', sep=';', decimal=',').fillna('')
    except UnicodeDecodeError:
        try:
            apteka = pd.read_csv(file,  encoding='utf-8', sep=';', decimal=',').fillna('')
        except Exception:
            print(f"Не удалось прочитать {file} — пропускаем")
            continue

    # новые столбцы
    
    apteka["Номер счет-фактуры"] = ""
    apteka["Сумма счет-фактуры"] = ""
    apteka["Дата счет-фактуры"] = ""
    apteka["Сравнение дат"] = ""

    
    mask = apteka['Поставщик'].astype(str).str.contains('ЕАПТЕКА', case=False, na=False)
    apteka.loc[mask, 'Номер накладной'] = apteka.loc[mask, 'Номер накладной'].astype(str).str.strip() + '/15'

    # Объединяем

    apteka = apteka.merge(
        sbis[['Номер', 'Сумма', 'Дата']],
        left_on='Номер накладной',
        right_on='Номер',
        how='left'
    )

    # Переименовываем найденные поля

    apteka = apteka.rename(columns={
        'Номер': 'Номер счет-фактуры',
        'Сумма': 'Сумма счет-фактуры',
        'Дата': 'Дата счет-фактуры'
    })

    # Форматируем дату счет-фактуры

    # убираем дублированные столбцы после merge, оставляя последнюю версию
    apteka = apteka.loc[:, ~apteka.columns.duplicated(keep='last')]

    apteka['Дата счет-фактуры'] = pd.to_datetime(
    apteka['Дата счет-фактуры'],
    format='%d.%m.%y',
    errors='coerce'
    ).dt.strftime('%d.%m.%Y').fillna('')

    # Проверяем совпадение дат

    date_invoice = pd.to_datetime(
        apteka['Дата накладной'],
        dayfirst=True,
        errors='coerce'
    )

    date_sf = pd.to_datetime(
        apteka['Дата счет-фактуры'],
        dayfirst=True,
        errors='coerce'
    )


    mismatch = (
        date_invoice.notna()
        & (
            date_sf.isna()
            | (date_invoice.dt.normalize() != date_sf.dt.normalize())
        )
    )

    apteka.loc[mismatch, 'Сравнение дат'] = 'Не совпадает!'

    # Итоговые столбцы

    result_columns = [
        "№ п/п",
        "Штрих-код партии",
        "Наименование товара",
        "Поставщик",
        "Дата приходного документа",
        "Номер приходного документа",
        "Дата накладной",
        "Номер накладной",
        "Номер счет-фактуры",
        "Сумма счет-фактуры",
        "Кол-во",
        "Сумма в закупочных ценах без НДС",
        "Ставка НДС поставщика",
        "Сумма НДС",
        "Сумма в закупочных ценах с НДС",
        "Дата счет-фактуры",
        "Сравнение дат"
    ]

    # 4. Сохраняем результат в Excel

    apteka = apteka[result_columns]

    output_file = result_folder / f"{Path(file).stem} - результат.xlsx"

    apteka.to_excel(output_file, index=False)

print("Обработка завершена.")