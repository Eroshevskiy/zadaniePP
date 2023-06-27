import cv2
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# Путь к видео
video_path = r'C:\Users\erosh\PycharmProjects\pythonProject1\video\IMG_0635.MOV'
# Путь к таблице Excel
excel_path = r'C:\Users\erosh\PycharmProjects\pythonProject1\данные.xlsx'
# Путь к файлу с данными
data_file_path = r'C:\Users\erosh\PycharmProjects\pythonProject1\зарплата.xlsx'

# Создание нового файла Excel
workbook = Workbook()
sheet = workbook.active
sheet.title = 'Присутствие'
sheet['A1'] = 'Дата и время'
sheet['B1'] = 'Статус'

# Загрузка видео
video = cv2.VideoCapture(video_path)

# Переменная для отслеживания предыдущего времени
prev_time = None

while True:
    # Считывание кадра из видео
    ret, frame = video.read()

    if not ret:
        # Достигнут конец видео
        break

    # Распознавание прямоугольников

    gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    contours, _ = cv2.findContours(gray_frame, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        area = cv2.contourArea(contour)

        # Проверка условия, что найденный объект является прямоугольником нужного цвета
        if area > 100 and w > 10 and h > 10:
            # Сотрудник присутствует
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # Проверка различного времени
            if current_time != prev_time:
                print("Сотрудник присутствовал в данное время:", current_time)

                # Запись данных в таблицу Excel
                row_data = [current_time, "Присутствовал"]
                sheet.append(row_data)

                # Обновление предыдущего времени
                prev_time = current_time

    # Отображение даты и времени на окне с видео
    current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cv2.putText(frame, f"DateTime: {current_datetime}", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
    cv2.imshow('Video', frame)

    # Выход по нажатию клавиши 'q'
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Закрытие видео и окна
video.release()
cv2.destroyAllWindows()

# Сохранение таблицы Excel
workbook.save(excel_path)

# Чтение данных из таблицы Excel и вывод в консоль
df = pd.read_excel(data_file_path)
print("Данные из таблицы Excel:")
print(df)
