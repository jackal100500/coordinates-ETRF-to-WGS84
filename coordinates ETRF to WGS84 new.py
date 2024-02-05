import pandas as pd
import pyproj

# Используем Transformer для преобразования координат
transformer = pyproj.Transformer.from_crs("EPSG:3046", "EPSG:4326", always_xy=True)

# Загрузка данных из Excel файла
file_path = 'D:/OneDrive - IDC DOO/Кординаты/преобразование координат/S1 GMS1.xlsx'
output_file_path = 'D:/OneDrive - IDC DOO/Кординаты/преобразование координат/S1 GMS1 с WGS84.xlsx'  # Новый файл для сохранения результатов
df = pd.read_excel(file_path)

# Создаем новые пустые колонки для результатов
df['Широта'] = None
df['Долгота'] = None

# Проходим по строкам и выполняем преобразование координат
for index, row in df.iterrows():
    try:
        x_etrf2000 = row['E']  # Предполагаем, что 'E' обозначает долготу (восток)
        y_etrf2000 = row['N']  # Предполагаем, что 'N' обозначает широту (север)

        # Преобразование координат с учетом порядка: долгота (x), широта (y)
        x_wgs84, y_wgs84 = transformer.transform(x_etrf2000, y_etrf2000)

        # Корректировка порядка присвоения для соответствия общим стандартам: широта (y), долгота (x)
        df.at[index, 'Широта'] = y_wgs84
        df.at[index, 'Долгота'] = x_wgs84
    except Exception as e:
        print(f"Ошибка при обработке строки {index}: {e}")

# Сохраняем результаты в новый Excel файл
df.to_excel(output_file_path, index=False)
