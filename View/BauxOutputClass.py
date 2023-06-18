import os
import json
import tempfile
import plotly.io as pio
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QDialogButtonBox, QTableWidgetItem, QVBoxLayout, QLabel, QPushButton, QWidget, QFormLayout, QLineEdit, QCheckBox, QHeaderView, QScrollArea, QSizePolicy, QTableWidget, QGridLayout, QToolBar, QFileDialog, QApplication, QMessageBox, QAction
from matplotlib.backends.backend_template import FigureCanvas
from PyQt5.QtGui import QPixmap, QFont, QIcon
from PyQt5.QtCore import QStandardPaths
import numpy as np
from matplotlib import pyplot as plt
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from pycel import ExcelCompiler
import pandas as pd
import matplotlib.patheffects as path_effects
import plotly.graph_objects as go
from plotly.offline import plot
from PyQt5.QtWebEngineWidgets import QWebEngineView
import sys
import importlib
import shutil
try:
    module = importlib.import_module('View.BauxOutput2')
except ImportError:
    module = importlib.import_module('BauxOutput2')

ui_class = getattr(module, 'Ui_AlumiaWindow')

class BauxOutputWindow(QMainWindow, ui_class):
    def __init__(self, start_page = 0):
        super().__init__()
        self.setupUi(self)
        self.showMaximized()
        if start_page != 0:
            self.start_page = start_page
        self.tabWidget_5.setCurrentIndex(0)
        self.tabWidget_2.setCurrentIndex(0)
        self.tabWidget_3.setCurrentIndex(0)
        self.tabWidget_4.setCurrentIndex(0)
        self.tabWidget_7.setCurrentIndex(0)
        self.tabWidget_8.setCurrentIndex(0)
        self.tabWidget_9.setCurrentIndex(0)
        
        for table in [self.tableBoxite13,self.tableLimestone13,self.tableWhiteMud13,self.tableCarbonMud13,self.tableSoda13,self.tableRecycledSolution13,self.tablePulp13,self.tableNaLosses13,self.tablePulp14,self.tableCrushedSpeck14,self.tableSinteringLosses14,self.tableCrushingLosses14,self.tableMud15,self.tableAluminateSolution15,self.tableSoda15,self.tableCarbonMud15,self.tableAlkalineSolution15,self.tableLosses15,self.tableCarbonMud16,self.tableAlkalineSolution16,self.tableWater16,self.tableCarbonMudStageTwo16,self.tableCarbonMudCharge16,self.tableAlkalineSolutionFromMud16,self.tablePromwater16,self.tableMud17,self.tableAluminateSolution17,self.tableSodaSolution17,self.tablePromwater17,self.tableCarbonMud17,self.tableAlkalineSolution17,self.tableAlkalineSolutionLeaching17,self.tableLosses17,self.tableCrushedSpeck18,self.tableAlkalineSolution18,self.tableWater18,self.tableRecycledSolution18,self.tablePromwater18,self.tableAluminateSolution18,self.tableAluminateSolution19,self.tableGrindWhiteMud19,self.tableAluminateSolutionWithMud19,self.tableWhiteMud19,self.tableSolution19,self.tableLosses19,self.tableAluminateSolutionStageTwo19,self.tableWhiteMud20,self.tableSolution20,self.tableRecycledMud20,self.tablePromwater20,self.tableWater20,self.tableAluminateSolutionStageOne21,self.tablePromwater21,self.tableCarbonMud21,self.tableCarbonMilk21,self.tableHydrogranate21,self.tableAluminateSolutionMud21,self.tableAluminateSolutionCarbon21,self.tableAluminateSolutionStageTwo22,self.tableAluminumHydroxide22,self.tableAlkalineSolution22,self.tableAluminumHydroxideS22,self.tableAlkalineSolutionS22,self.tableProdAluminumHydroxide22,self.tableProdAlkalineSolution22,self.tableLosses22,self.tableAluminumHydroxide23,self.tableAlkalineSolution23,self.tableWater23,self.tableProdAluminumHydroxide23,self.tablePromwater23,self.tablePromwaterS23,self.tableAluminumHydroxide24,self.tablePromwater24,self.tableAlumina24,self.tableLosses24,self.tableAlkalineSolution25,self.tableAlkalineSolution26,self.tablePromwater26,self.tableRecycledSolution26,self.tableLosses26,self.tableBoxite13_2,self.tableLimestone13_2,self.tableWhiteMud13_2,self.tableCarbonMud13_2,self.tableSoda13_2,self.tableRecycledSolution13_2,self.tablePulp13_2,self.tableNaLosses13_2,self.tablePulp14_2,self.tableCrushedSpeck14_2,self.tableSinteringLosses14_2,self.tableCrushingLosses14_2,self.tableMud15_2,self.tableAluminateSolution15_2,self.tableSoda15_2,self.tableCarbonMud15_2,self.tableAlkalineSolution15_2,self.tableLosses15_2,self.tableCarbonMud16_2,self.tableAlkalineSolution16_2,self.tableWater16_2,self.tableCarbonMudStageTwo16_2,self.tableCarbonMudCharge16_2,self.tableAlkalineSolutionFromMud16_2,self.tablePromwater16_2,self.tableMud17_2,self.tableAluminateSolution17_2,self.tableSodaSolution17_2,self.tablePromwater17_2,self.tableCarbonMud17_2,self.tableAlkalineSolution17_2,self.tableAlkalineSolutionLeaching17_2,self.tableLosses17_2,self.tableCrushedSpeck18_2,self.tableAlkalineSolution18_2,self.tableWater18_2,self.tableRecycledSolution18_2,self.tablePromwater18_2,self.tableAluminateSolution18_2,self.tableAluminateSolution19_2,self.tableGrindWhiteMud19_2,self.tableAluminateSolutionWithMud19_2,self.tableWhiteMud19_2,self.tableSolution19_2,self.tableLosses19_2,self.tableAluminateSolutionStageTwo19_2,self.tableWhiteMud20_2,self.tableSolution20_2,self.tableRecycledMud20_2,self.tablePromwater20_2,self.tableWater20_2,self.tableAluminateSolutionStageOne21_2,self.tablePromwater21_2,self.tableCarbonMud21_2,self.tableCarbonMilk21_2,self.tableHydrogranate21_2,self.tableAluminateSolutionMud21_2,self.tableAluminateSolutionCarbon21_2,self.tableAluminateSolutionStageTwo22_2,self.tableAluminumHydroxide22_2,self.tableAlkalineSolution22_2,self.tableAluminumHydroxideS22_2,self.tableAlkalineSolutionS22_2,self.tableProdAluminumHydroxide22_2,self.tableProdAlkalineSolution22_2,self.tableLosses22_2,self.tableAluminumHydroxide23_2,self.tableAlkalineSolution23_2,self.tableWater23_2,self.tableProdAluminumHydroxide23_2,self.tablePromwater23_2,self.tablePromwaterS23_2,self.tableAluminumHydroxide24_2,self.tablePromwater24_2,self.tableAlumina24_2,self.tableLosses24_2,self.tableAlkalineSolution25_2,self.tableAlkalineSolution26_2,self.tablePromwater26_2,self.tableRecycledSolution26_2,self.tableLosses26_2, self.tableInit]:
            header = table.horizontalHeader()
            table.horizontalHeader().setStretchLastSection(False)
            font = QFont()
            font.setPointSize(11)
            table.setFont(font)
            table.horizontalHeader().setStyleSheet(f"QHeaderView::section {{ font-size: {11}pt; }}")
            # table.setStyleSheet(f"QTableView {{ font-size: {12}pt; }}")
            header.setMinimumSectionSize(70)
            # Устанавливаем режим изменения размера для каждой колонки
            if(header.count() > 5):
                for section in range(header.count()):
                    header.setSectionResizeMode(section, QHeaderView.Stretch)

        
            # Соединяем сигнал currentChanged с пользовательским слотом
        self.tableInit.setStyleSheet("padding-top: 10px")
        
        # Создание панели инструментов
        toolbar = QToolBar("My main toolbar")
        self.addToolBar(toolbar)

        # Создание кнопки "Сохранить"
        save_button = QAction(QIcon(), 'Сохранить', self)
        save_button.triggered.connect(self.save_file)
        toolbar.addAction(save_button)

        # Создание кнопки "О программе"
        about_button = QAction(QIcon(), 'О приложении', self)
        about_button.triggered.connect(self.about_app)
        toolbar.addAction(about_button)

        self.condition_for_saving_two_files = False
    
    def save_file(self):
        # Задание начальной директории в "Документы"
        initial_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        default_file_name = "Отчет о боксите.xlsx"  # Название файла по умолчанию
        target_file_name, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", initial_dir + "/" + default_file_name)

        # Сохранение файла
        if target_file_name:
            source_file_name = "Excels\Производство глинозема1.xlsx"  # Укажите здесь путь к вашему существующему файлу
            shutil.copyfile(source_file_name, target_file_name)

            # Если условие выполняется, сохраняем второй файл
            if self.condition_for_saving_two_files:
                default_file_name2 = "Отчет о втором боксите.xlsx"
                second_file_name, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", initial_dir + "/" + default_file_name2)
                if second_file_name:
                    second_source_file_name = "Excels\Производство глинозема2.xlsx"  # Укажите здесь путь к вашему второму файлу
                    shutil.copyfile(second_source_file_name, second_file_name)

    
    def about_app(self):
        # Отображение информации о программе
        QMessageBox.about(self, "О приложении", "Студент Ыдырыс Олжас, обучающийся в Национальном исследовательском технологическом университете \"МИСиС\" в группе БМТ-19-2, разработал программное решение в 2023 году. Для получения дополнительной информации о данной программе, возможностях ее использования или для задания вопросов по ее функциональности, можно обратиться к Ыдырысу Олжасу через мессенджер Telegram по следующему адресу: @im01zhas.")

        
    def copy_row_to_other_table(self, table1, table2, target_row, source_row=0):
        # Получить имена столбцов для обеих таблиц
        table1_columns = [table1.horizontalHeaderItem(i).text() for i in range(table1.columnCount())]
        table2_columns = [table2.horizontalHeaderItem(i).text() for i in range(table2.columnCount())]

        # Найти общие столбцы
        common_columns = set(table1_columns) & set(table2_columns)

        sum = float(table1.item(source_row, table1_columns.index("Итого")).text())
        
        # Перебрать каждый общий столбец
        for column in common_columns:
            # Получить индекс столбца в каждой таблице
            table1_column_index = table1_columns.index(column)
            table2_column_index = table2_columns.index(column)

            # Копировать данные из столбца первой таблицы в столбец второй таблицы
            item = table1.item(source_row, table1_column_index)
            if item is not None:
                item = float(item.text())
                per = format(round(item/sum * 100, 2), '.2f')
                table2.setItem(target_row, table2_column_index, QTableWidgetItem(per))

    def add_histogram_tab2(self):
        
        self.condition_for_saving_two_files = True
        
        scroll_area = QScrollArea()
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        
        # Your DataFrame generation code here...
        dfs = self.get_graphs()
        list_of_dfs = [dfs[0], dfs[1]]
        df = pd.concat(list_of_dfs)
        fig = self.plot_dataframe(df)

        # Convert the figure to a PNG image
        image_bytes = pio.to_image(fig, format='png')

        # Create a QPixmap from the image data
        pixmap = QPixmap()
        pixmap.loadFromData(image_bytes)

        # Create a QLabel and set the QPixmap as its pixmap
        label = QLabel()
        label.setPixmap(pixmap)

        layout.addWidget(label)
        
        label = QLabel()
        label.setText("Показатели технологий")
        
        label.setStyleSheet("""
            QLabel {
                font: bold 20px;  /* Устанавливаем жирный шрифт размером 16px */
                padding-top: 10px;  /* Устанавливаем отступ сверху 10px */
                padding-bottom: 10px;  /* Устанавливаем отступ снизу 10px */
            }
        """)
        
        layout.addWidget(label)
        layout.addWidget(MyTable2())
        
        df = self.get_koeff()
        
        fig = self.plot_koeff(df)

        # Convert the figure to a PNG image
        image_bytes = pio.to_image(fig, format='png')

        # Create a QPixmap from the image data
        pixmap = QPixmap()
        pixmap.loadFromData(image_bytes)

        # Create a QLabel and set the QPixmap as its pixmap
        label = QLabel()
        label.setPixmap(pixmap)

        self.verticalLayout_56.addWidget(label)
        
        scroll_area.setWidget(scroll_content)
        layout2 = QVBoxLayout(self.graphs)
        layout2.addWidget(scroll_area)
        
        self.add_init2()

    def add_histogram_tab1(self):
        self.condition_for_saving_two_files = True
        
        self.tabWidget_5.setTabText(self.tabWidget_5.indexOf(self.tab_18), "Материальный баланс боксита")
        self.tabWidget_5.setTabText(self.tabWidget_5.indexOf(self.tab_39), "Анализ характеристик боксита") 
        
        scroll_area = QScrollArea()
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        
        # Your DataFrame generation code here...
        dfs = self.get_graphs(["1"])
        list_of_dfs = [dfs[0], dfs[1]]
        df = pd.concat(list_of_dfs)
        fig = self.plot_dataframe(df, ["1"])

        # Convert the figure to a PNG image
        image_bytes = pio.to_image(fig, format='png')

        # Create a QPixmap from the image data
        pixmap = QPixmap()
        pixmap.loadFromData(image_bytes)

        # Create a QLabel and set the QPixmap as its pixmap
        label = QLabel()
        label.setPixmap(pixmap)

        layout.addWidget(label)
        
        label = QLabel()
        label.setText("Показатели технологий")
        
        label.setStyleSheet("""
            QLabel {
                font: bold 20px;  /* Устанавливаем жирный шрифт размером 16px */
                padding-top: 10px;  /* Устанавливаем отступ сверху 10px */
                padding-bottom: 10px;  /* Устанавливаем отступ снизу 10px */
            }
        """)
        
        layout.addWidget(label)
        layout.addWidget(MyTable1())

        df = self.get_koeff(["1"])
        
        fig = self.plot_koeff(df, ["1"])

        # Convert the figure to a PNG image
        image_bytes = pio.to_image(fig, format='png')

        # Create a QPixmap from the image data
        pixmap = QPixmap()
        pixmap.loadFromData(image_bytes)

        # Create a QLabel and set the QPixmap as its pixmap
        label = QLabel()
        label.setPixmap(pixmap)

        self.verticalLayout_56.addWidget(label)
        
        scroll_area.setWidget(scroll_content)
        layout2 = QVBoxLayout(self.graphs)
        layout2.addWidget(scroll_area)
        
        self.add_init1()
    


    def add_init1(self):
        self.tableInit.removeRow(7)
        self.tableInit.removeRow(5)
        self.tableInit.removeRow(3)
        self.tableInit.removeRow(1)

        self.copy_row_to_other_table(self.tableBoxite13, self.tableInit, 0)
        self.copy_row_to_other_table(self.tableLimestone13, self.tableInit, 1)
        self.copy_row_to_other_table(self.tableSoda13, self.tableInit, 2)
        self.copy_row_to_other_table(self.tableWhiteMud13, self.tableInit, 3)
        
        for row in range(self.tableInit.rowCount()):
            for column in range(self.tableInit.columnCount()):
                item = self.tableInit.item(row, column)
                if item is None or not item.text():
                    self.tableInit.setItem(row, column, QTableWidgetItem('0.00'))
                    
    def add_init2(self):
        
        self.copy_row_to_other_table(self.tableBoxite13, self.tableInit, 0)
        self.copy_row_to_other_table(self.tableBoxite13_2, self.tableInit, 1)
        self.copy_row_to_other_table(self.tableLimestone13, self.tableInit, 2)
        self.copy_row_to_other_table(self.tableLimestone13_2, self.tableInit, 3)
        self.copy_row_to_other_table(self.tableSoda13, self.tableInit, 4)
        self.copy_row_to_other_table(self.tableSoda13_2, self.tableInit, 5)
        self.copy_row_to_other_table(self.tableWhiteMud13, self.tableInit, 6)
        self.copy_row_to_other_table(self.tableWhiteMud13_2, self.tableInit, 7)
        
        for row in range(self.tableInit.rowCount()):
            for column in range(self.tableInit.columnCount()):
                item = self.tableInit.item(row, column)
                if item is None or not item.text():
                    self.tableInit.setItem(row, column, QTableWidgetItem('0.00'))
                    
    
    def plot_dataframe(self, df, types = ['1', '2']):
        stages = ['Спекание', 'Выщелачивание', 'Обескремнивание', 'Карбонизация', 'Кальцинация']
        
        fig = go.Figure()

        # ширина баров на графике
        width = 0.35
        
            
        for i, type in enumerate(types):
            for j, stage in enumerate(stages):
                al_key = f'Содержание Al2O3, кг/т глинозема{type}'
                mass_key = f'Масса материального потока, кг/т глинозема {type}'
                
                # x-координаты для группы столбцов
                x = np.arange(len(stages)) + i * width - (len(types) - 1) * width / 2
                
                
                if len(types) == 1:
                    name1 = 'Боксит'
                else:
                    name1 = f'Боксит {type}'
                
                fig.add_trace(go.Bar(
                    x=[x[j]],
                    y=[df.loc[mass_key, stage]/1000],
                    width=width,
                    name=name1 if j == 0 else None,  # Добавляем имя только к первому трейсу
                    marker_color=['lightsalmon', 'lightblue'][i],
                    text=round(df.loc[mass_key, stage])/1000,  # Округляем значения
                    textposition='auto',
                    showlegend=j == 0  # Показываем легенду только для первого трейса
                ))

                fig.add_trace(go.Bar(
                    x=[x[j]],
                    y=[df.loc[al_key, stage]/1000],
                    width=width,
                    name=None,  # Скрываем имя
                    marker_color=['indianred', 'darkblue'][i],
                    text=round(df.loc[al_key, stage])/1000,  # Округляем значения
                    textposition='auto',
                    showlegend=False  # Скрываем легенду
                ))

        fig.update_layout(
            title='Распределение материальных потоков (количество Al₂O₃)<br>по основным переделам технологий спекания<br>',
            xaxis=dict(
                tickmode='array',
                tickvals=np.arange(len(stages)),
                ticktext=stages,
                title='Переделы технологии производства глинозема методом спекания',
                tickfont_size=16,
                tickangle=0,
            ),
            yaxis=dict(
                title='Масса материального потока на переделе<br>Cодержание Al₂O₃ в потоке, т/т товарного глинозема',
                # titlefont_size=20,
                tickfont_size=20,
            ),
            legend=dict(
                x=0,
                y=1.0,
                bgcolor='rgba(255, 255, 255, 0)',
                bordercolor='rgba(255, 255, 255, 0)'
            ),
            barmode='overlay',
            bargap=0.10,
            bargroupgap=0.05,
            autosize=False,
            width=875,  # Увеличиваем ширину графика
            height=575,  # Увеличиваем высоту графика
            margin=dict(
                l=0,
                r=0,
                b=0,
                t=50,
                pad=0
            )
        )
        
        fig.update_layout(
        font=dict(
            size=13,  # Set the font size here
            color="black"
        )
        )
        return fig
    
    

    def plot_koeff(self, df, tab = ["1", "2"]):
        stages = ['Боксит', 'Известняк', 'Белый шлам', 'Карбонатный шлам', 'Сода', 'Оборотный раствор']
        types = []

        for i in tab:
            types.append(f'Боксит {i}')
        
        fig = go.Figure()

        # ширина баров на графике
        width = 0.35

        for i, type in enumerate(types):
            for j, stage in enumerate(stages):
                mass_key = type
                    
                # x-координаты для группы столбцов
                x = np.arange(len(stages)) + i * width - (len(types) - 1) * width / 2
                
                if len(types) == 1:
                    name1 = 'Боксит'
                else:
                    name1 = f'Боксит {i+1}'
                
                fig.add_trace(go.Bar(
                    x=[x[j]],
                    y=[df.loc[mass_key, stage]],
                    width=width,
                    name=name1 if j == 0 else None,  # Добавляем имя только к первому трейсу
                    marker_color=['lightsalmon', 'lightblue'][i],
                    text=round(df.loc[mass_key, stage]),  # Округляем значения
                    textposition='auto',
                    showlegend=j == 0  # Показываем легенду только для первого трейса
                ))

        fig.update_layout(
            title='Вид материала (промежуточные продукты технологий)',
            xaxis=dict(
                tickmode='array',
                tickvals=np.arange(len(stages)),
                ticktext=stages,
                title='Переделы технологии производства глинозема методом спекания',
                tickfont_size=16,
                tickangle=0,
            ),
            yaxis=dict(
                title='Расходные коэффиициенты, кг/т товарного глинозема',
                tickfont_size=16,
            ),
            legend=dict(
                x=0,
                y=1.0,
                bgcolor='rgba(255, 255, 255, 0)',
                bordercolor='rgba(255, 255, 255, 0)'
            ),
            barmode='overlay',
            bargap=0.10,
            bargroupgap=0.05,
            autosize=False,
            width =925,  # Увеличиваем ширину графика
            height=550,  # Увеличиваем высоту графика
            margin=dict(
                l=0,
                r=0,
                b=0,
                t=50,
                pad=0
            )
        )
        
        fig.update_layout(
        font=dict(
            size=14,  # Set the font size here
            color="black"
        )
        )
        return fig


    
    def get_cell_value(self, file_name, cell_coordinates):
        # Создание компилятора Excel
        compiler = ExcelCompiler(filename=file_name)
        
        # Разделение координат на столбец и строку
        column_letter = cell_coordinates[0]
        row = int(cell_coordinates[1:])
        
        # Определение листа и значения ячейки
        sheet_name = "Таблицы"
        cell_value = compiler.evaluate(f'{sheet_name}!{column_letter}{row}')
        
        return cell_value
    
    def get_graphs(self, tab = ["1" , "2"]):
        # Преобразование JSON в словарь
        with open("./JSON/graph.json", "r", encoding="utf8") as file:
            data = json.load(file)
        # Имя файла Excel
        file_name = "Excels\Производство глинозема"
        # Создание списка DataFrame'ов
        dfs = []
        # Обработка каждой категории в JSON-объекте
        for category, inner_dict in data.items():
            df = pd.DataFrame()
            for i in tab:
                # Обработка каждого значения внутри текущей категории
                for value_name, cell_coordinates in inner_dict.items():
                    # Получение значения из ячейки
                    cell_value = self.get_cell_value(file_name + i + ".xlsx", cell_coordinates)
                    
                    # Добавление значения в DataFrame
                    df.loc[category + i ,value_name] = cell_value
                
                # Добавление DataFrame в список
            dfs.append(df)
        return dfs
    
    def get_koeff(self, tab = ["1", "2"]):
        
        nice = {
            "Боксит": "N6",
            "Известняк": "N7",
            "Белый шлам": "N8",
            "Карбонатный шлам": "N9",
            "Сода": "N10",
            "Оборотный раствор": "N11"
            }
        
        # Имя файла Excel
        file_name = "Excels\Производство глинозема"

        # Обработка каждой категории в JSON-объекте
        df = pd.DataFrame()
        for i in tab:
            # Обработка каждого значения внутри текущей категории
            for value_name, cell_coordinates in nice.items():
                # Получение значения из ячейки
                cell_value = self.get_cell_value(file_name + i + ".xlsx", cell_coordinates)
                
                # Добавление значения в DataFrame
                df.loc["Боксит " + i ,value_name] = cell_value
            
            # Добавление DataFrame в список
        return df
        
    
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'canvas'):
            self.canvas.setGeometry(self.tabWidget.geometry())



        # Список пар виджетов для синхронизации
        self.tabs = [(self.tabWidget, self.tabWidget_6),
                        (self.tabWidget_2, self.tabWidget_7),
                        (self.tabWidget_3, self.tabWidget_8),
                        (self.tabWidget_4, self.tabWidget_9)]

        for tab_pair in self.tabs:
            tab_pair[0].currentChanged.connect(self.sync_tabs)
            tab_pair[1].currentChanged.connect(self.sync_tabs)



    def sync_tabs(self, index):
        sender = self.sender()
        for tab_pair in self.tabs:
            if sender == tab_pair[0]:
                tab_pair[1].setCurrentIndex(index)
                break
            elif sender == tab_pair[1]:
                tab_pair[0].setCurrentIndex(index)
                break
    
    
    def closeEvent(self, event):
        self.start_page.show()  # Show the MainWindow when NewWindow closes
        event.accept()


class MyTable2(QTableWidget):
    def __init__(self, r = 5, c = 3):
        super().__init__(r, c)
        self.init_ui()

    def init_ui(self):
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.setFixedHeight(250)
        
        self.horizontalHeader().setVisible(False) # hiding column names
        self.verticalHeader().setVisible(False) # hiding row names

        self.setSpan(0, 0, 2, 1)
        self.setSpan(0, 1, 1, 2)
        
        font1 = QFont()
        font1.setBold(True)
        font1.setPointSize(10)
        
        font = QFont()
        font.setPointSize(10)
        
        def createItem(text):
            item = QTableWidgetItem(text)
            item.setFont(font1)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            return item

        self.setItem(0, 0, createItem("Показатель"))
        
        item = QTableWidgetItem("Характеристика")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(0, 1, item)

        self.setItem(2, 0, createItem("Марка глинозема"))
        self.setItem(3, 0, createItem("Товарный выход, %"))
        self.setItem(4, 0, createItem("Кол-во отвального шлама, кг/т тов. гл."))

        item = QTableWidgetItem("Боксит 1")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(1, 1, item)
        
        item = QTableWidgetItem("Боксит 2")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(1, 2, item)

        item = QTableWidgetItem("Г-000")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(2, 1, item)
        
        item = QTableWidgetItem("Г-000")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(2, 2, item)

        item = QTableWidgetItem("87,0")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(3, 1, item)
        
        item = QTableWidgetItem("87,0")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(3, 2, item)

        item = QTableWidgetItem("1700")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(4, 1, item)
        
        item = QTableWidgetItem("1600")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(4, 2, item)

        
from PyQt5.QtCore import Qt
class MyTable1(QTableWidget):
    def __init__(self, r = 4, c = 2):
        super().__init__(r, c)
        self.init_ui()

    def init_ui(self):
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.setFixedHeight(250)
        
        self.horizontalHeader().setVisible(False) # hiding column names
        self.verticalHeader().setVisible(False) # hiding row names

        font1 = QFont()
        font1.setBold(True)
        font1.setPointSize(12)
        
        font = QFont()
        font.setPointSize(10)
        
        # Creating QTableWidgetItem with desired settings
        def createItem(text):
            item = QTableWidgetItem(text)
            item.setFont(font1)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            return item

        self.setItem(0, 0, createItem("Показатель"))
        
        item = QTableWidgetItem("Характеристика")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(0, 1, item)

        self.setItem(1, 0, createItem("Марка глинозема"))
        self.setItem(2, 0, createItem("Товарный выход, %"))
        self.setItem(3, 0, createItem("Кол-во отвального шлама, кг/т тов. гл."))

        item = QTableWidgetItem("Г-000")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(1, 1, item)
        
        item = QTableWidgetItem("87,0")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(2, 1, item)

        item = QTableWidgetItem("1700")
        item.setFont(font)
        item.setTextAlignment(Qt.AlignHCenter |  Qt.AlignVCenter)
        self.setItem(3, 1, item)

        
        

import qdarkstyle

if __name__ == "__main__":
    

    
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet(qdarkstyle.DarkPalette))
    AlumiaWindow = BauxOutputWindow()
    AlumiaWindow.add_histogram_tab2()
    AlumiaWindow.show()
    sys.exit(app.exec_())