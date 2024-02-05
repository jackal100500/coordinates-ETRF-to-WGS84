import pandas as pd
import pyproj

# Создаем объекты для ETRF и WGS84
etrf2000 = pyproj.Proj('EPSG:3046')  # ETRF
wgs84 = pyproj.Proj('EPSG:4326')  # WGS84

# Загрузка данных из Excel файла
file_path = 'D:/OneDrive - IDC DOO/Кординаты/S1 GMS1.xlsx'
output_file_path = 'D:/OneDrive - IDC DOO/Кординаты/S1 GMS1 с WGS84.xlsx'  # Новый файл для сохранения результатов
df = pd.read_excel(file_path)

# Создаем новые пустые колонки для результатов
df['Широта'] = None
df['Долгота'] = None

# Проходим по строкам и выполняем преобразование координат
for index, row in df.iterrows():
    try:
        x_etrf2000 = row['N']  # Значение из колонки C
        y_etrf2000 = row['E']  # Значение из колонки B

        x_wgs84, y_wgs84 = pyproj.transform(etrf2000, wgs84, x_etrf2000, y_etrf2000)

        # Записываем результаты в соответствующие колонки
        df.at[index, 'Широта'] = x_wgs84
        df.at[index, 'Долгота'] = y_wgs84
    except Exception as e:
        print(f"Ошибка при обработке строки {index}: {e}")

# Сохраняем результаты в новый Excel файл
df.to_excel(output_file_path, index=False)
