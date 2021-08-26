import pandas as pd
import os


def parcer():
    # создание датафреймов в зависимости от формата файлов
    # файлы должны называться old и new
    if 'old.xlsx' in os.listdir():
        new_df = pd.read_excel('new.xlsx')
        old_df = pd.read_excel('old.xlsx')
    else:
        new_df = pd.read_excel('new.xls')
        old_df = pd.read_excel('old.xls')

    # Все колонки с заглавной буквы
    new_df.columns = map(str.capitalize, new_df.columns)
    old_df.columns = map(str.capitalize, old_df.columns)

    # Индексируем по столбцу 'Vin'
    new_df.index = new_df['Vin']
    old_df.index = old_df['Vin']

    # Удаляем пустые значения в столбце 'Vin'
    new_df = new_df.dropna(subset=['Vin'])
    old_df = old_df.dropna(subset=['Vin'])

    # Заполняем NaN у цен
    new_df['Актуальная цена для сайта'].fillna(new_df['Розничная цена без скидок'], inplace=True)
    old_df['Актуальная цена для сайта'].fillna(old_df['Розничная цена без скидок'], inplace=True)

    # Преобразование типов
    new_df[['Актуальная цена для сайта', 'Розничная цена без скидок']] = new_df[
        ['Актуальная цена для сайта', 'Розничная цена без скидок']].astype('int')
    old_df[['Актуальная цена для сайта', 'Розничная цена без скидок']] = old_df[
        ['Актуальная цена для сайта', 'Розничная цена без скидок']].astype('int')

    # добавляем в csv файл датафрейм с vin-номерами, которые нужно добавить
    add_vins = list(set(new_df['Vin']) - set(old_df['Vin']))
    add_df = new_df.copy()
    add_df = add_df.query('Vin == @add_vins').drop(['Vin'], axis=1).sort_values(by=['Vin']).to_csv('Добавить.csv')

    # добавляем датафрейм с vin-номерами для удаления в csv файл
    del_vins = list(set(old_df['Vin']) - set(new_df['Vin']))
    delete_vins = pd.DataFrame(del_vins).to_csv('Удалить.csv')

    # Фильтрация
    np = new_df.query('Vin != @add_vins')
    op = old_df.query('Vin != @del_vins')

    # Фильтрация ненужных строк и значений в столбце 'Vin'
    np = np.loc[lambda np: np['Vin'] != 'НЕ ОПРЕДЕЛЕНО']
    op = op.loc[lambda op: op['Vin'] != 'НЕ ОПРЕДЕЛЕНО']

    # Сортировка
    np = np.sort_index()
    op = op.sort_index()

    # Выявление разницы цен
    orp = list(op['Розничная цена без скидок'])
    oap = list(op['Актуальная цена для сайта'])
    np['rd'] = np['Розничная цена без скидок'] - orp
    np['ad'] = np['Актуальная цена для сайта'] - oap

    np = np[(np['rd'] != 0) | (np['ad'] != 0)].drop(['Vin', 'rd', 'ad'], axis=1).to_csv('Поменять цены.csv')


try:
    parcer()
except FileNotFoundError:
    print('Rename files to \'old\' and \'new\'')
