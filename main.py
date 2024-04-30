#This code is designed to perform
# a specific kind of data transformation, which involves cross-referencing or mapping
# spare parts numbers between different brands based on input from an Excel file

import pandas as pd
import os
from tkinter import filedialog, Tk

def load_file():
    # Создаем скрытое окно Tkinter
    root = Tk()
    root.withdraw()  # Скрываем окно

    # Запрос файла у пользователя
    file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        print("Файл не выбран.")
        return None, None
    return file_path, pd.read_excel(file_path)

def transform_data_universal_primary(data, primary_brand):
    transformed_rows = []
    brands = data.columns.tolist()  # Получаем список брендов

    # Перебираем каждую строку в данных
    for index, row in data.iterrows():
        # Список брендов и артикулов
        articles = {brand: str(row[brand]).split(',') for brand in brands}

        # Берем артикулы только от основного бренда и создаем кроссировки с другими брендами
        for article_i in articles[primary_brand]:
            article_i = article_i.strip()
            if article_i == 'nan' or article_i == '':  # Игнорируем пустые артикулы и строки только с пробелами
                continue
            for j in range(1, len(brands)):  # Перебираем другие бренды, начиная со второго столбца
                brand_j = brands[j]
                for article_j in articles[brand_j]:
                    article_j = article_j.strip()
                    if article_j == 'nan' or article_j == '':  # Игнорируем пустые артикулы и строки только с пробелами
                        continue
                    # Добавляем кроссировку в список
                    transformed_rows.append([primary_brand, article_i, brand_j, article_j])

    # Создаем DataFrame из списка
    transformed_data = pd.DataFrame(transformed_rows, columns=['Brand', 'Article', 'Cross Brand', 'Cross Article'])
    return transformed_data

def save_transformed_data(transformed_data, original_file_path):
    # Сохраняем переработанные данные в новый файл в той же директории
    directory = os.path.dirname(original_file_path)
    output_file_path = os.path.join(directory, 'Transformed_Data.xlsx')
    transformed_data.to_excel(output_file_path, index=False)
    print(f"Сохраненный файл: {output_file_path}")

# Главная функция
def main():
    file_path, data = load_file()
    if data is not None:
        primary_brand = data.columns[0]  # Первый столбец считаем основным брендом
        transformed_data = transform_data_universal_primary(data, primary_brand)
        save_transformed_data(transformed_data, file_path)

if __name__ == '__main__':
    main()
