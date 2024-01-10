import sys
import requests
from bs4 import BeautifulSoup
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel, QFileDialog, QLineEdit
from decimal import Decimal, InvalidOperation

score_data = []
data = []
placed = []


class MainWindow(QMainWindow):

    def __init__(self,parent = None):
        super().__init__(parent)
        self.setupUi()

    def setupUi(self):
        self.setWindowTitle('Квиз, плиз! HTML to EXCEL Parser')
        self.setGeometry(100, 100, 400, 250)

        self.label = QLabel('Введите ссылку на сайт:')
        self.url_input = QLineEdit()

        self.parse_button = QPushButton('Парсить')
        self.parse_button.clicked.connect(self.parse)

        self.export_button = QPushButton('Экспортировать сырую таблицу')
        self.export_button.clicked.connect(self.export)

        self.calculate_and_export_button = QPushButton('Эспортировать отсортированную таблицу')
        self.calculate_and_export_button.clicked.connect(self.calculate_and_export)

        self.export_fancy_way_button = QPushButton('Волшебный экспорт в формате команда - количество занятых мест')
        self.export_fancy_way_button.clicked.connect(self.export_fancy_way)

        self.import_data_button = QPushButton('Импортировать данные из Excel файла')
        self.import_data_button.clicked.connect(self.import_data)

        self.remove_data_button = QPushButton('Очистить данные')
        self.remove_data_button.clicked.connect(self.remove_data)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.url_input)
        layout.addWidget(self.parse_button)
        layout.addWidget(self.export_button)
        layout.addWidget(self.calculate_and_export_button)
        layout.addWidget(self.import_data_button)
        layout.addWidget(self.export_fancy_way_button)
        layout.addWidget(self.remove_data_button)


        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def export_fancy_way(self):
        df_placed = pd.DataFrame(placed, columns=['Команда', '1 место', '2 место', '3 место'])

        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        excel_file, _ = QFileDialog.getSaveFileName(self, "Сохранить файл Excel", "", "Excel Files (*.xlsx)",
                                                    options=options)

        if excel_file:
            df_placed.to_excel(excel_file, index=False, engine='openpyxl')
            print(f'Данные о местах записаны в файл {excel_file}')

    def parse(self):
        url = self.url_input.text()
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        table = soup.find('table', {'class': 'game-table'})

        for row in table.find_all('tr'):
            row_data = []
            for cell in row.find_all('td'):
                row_data.append(cell.text)
            data.append(row_data)

        temp_data = []
        rows = table.find_all('tr')
        second_row = rows[1]
        max_numeric_value = None
        max_index = None
        cells = second_row.find_all('td')
        count = 0
        for i, cell in enumerate(cells):
            if cell.text.strip():
                try:
                    value = Decimal(cell.text.strip())
                    if max_numeric_value is None or value > max_numeric_value:
                        max_numeric_value = value
                        max_index = i
                except (ValueError, InvalidOperation):
                    count += 1
                    continue

        for row in rows[1:]:
            cells = row.find_all('td')
            if (len(cells) > 11):
                team_name = cells[3].text.strip()
            else:
                team_name = cells[2].text.strip()
            score = Decimal(cells[max_index].text.strip())
            temp_data.append((team_name, score))

        for name, score in temp_data:
            found = False
            for i, (existing_name, existing_score) in enumerate(score_data):
                if name == existing_name:
                    if score != existing_score:
                        score_data[i] = (existing_name, existing_score + score)
                    found = True
                    break
            if not found:
                score_data.append((name, score))

    def export(self):
        df = pd.DataFrame(data[1:], columns=data[0])
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        excel_file, _ = QFileDialog.getSaveFileName(self, "Сохранить файл Excel", "", "Excel Files (*.xlsx)",
                                                    options=options)

        if excel_file:
            df.to_excel(excel_file, index=False, engine='openpyxl')
            print(f'Данные записаны в файл {excel_file}')

    def calculate_and_export(self):
        sorted_data = sorted(score_data, key=lambda x: x[1], reverse=True)
        df = pd.DataFrame(sorted_data, columns=['Имя команды', 'Количество баллов'])

        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        excel_file, _ = QFileDialog.getSaveFileName(self, "Сохранить файл Excel", "", "Excel Files (*.xlsx)",
                                                    options=options)

        if excel_file:
            df.to_excel(excel_file, index=False, engine='openpyxl')
            print(f'Данные записаны в файл {excel_file}')

    def import_data(self):
        score_data.clear()
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        excel_file, _ = QFileDialog.getOpenFileName(self, "Импортировать файл Excel", "", "Excel Files (*.xlsx)",
                                                    options=options)

        if excel_file:
            try:

                df = pd.read_excel(excel_file, dtype={'Количество баллов': float})

                imported_data = df[['Имя команды', 'Количество баллов']].values
                for name, score in imported_data:
                
                    if pd.notna(name) and pd.notna(score):
                        score_data.append((name, score))

                print(f'Данные успешно импортированы из файла {excel_file}')
            except Exception as e:
                print(f'Ошибка при импорте данных: {str(e)}')

    def remove_data(self):
        score_data.clear()
        data.clear()
        print("Данные успешно очищены.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    print("1. Добавлять ссылки только на те игры, для которых уже есть таблицы, иначе прога вылетит. \n2. \"Экспортировать сырую таблицу\" - программа экспортит таблицы одна за одной в эксель файл. \n3. \"Экспортировать отсортированную таблицу\" - программа экспортит данные в формате (Имя команды, количество баллов) за все запаршенные таблицы. \n4. \"Импортировать таблицу\" - программа принимает на вход файлы, полученные из 3 пункта, при этом вся информация, полученная до этого, стирается, так что импортируйте сначала файл, а потом делайте какую либо работу. \nПример работы программы:  \n1.Парсите 10 игр, экспортите их сплошняком или форматом команда-количество баллов. \n2.Импортируете уже готовую таблицу(перед этим создав ее, как описано в пункте 1) и добавляете в нее данные.\ntg: arakelov_aa")
    sys.exit(app.exec_())