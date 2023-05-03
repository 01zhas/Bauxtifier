from PyQt5.QtWidgets import QApplication, QMainWindow, QHeaderView
from View.BauxOutput import Ui_AlumiaWindow

class BauxOutputWindow(QMainWindow, Ui_AlumiaWindow):
    def __init__(self, start_page):
        super().__init__()
        self.setupUi(self)
        self.showMaximized()
        self.start_page = start_page
        self.tabWidget_5.setCurrentIndex(0)
        self.tabWidget_2.setCurrentIndex(0)
        self.tabWidget_3.setCurrentIndex(0)
        self.tabWidget_4.setCurrentIndex(0)
        self.tabWidget_7.setCurrentIndex(0)
        self.tabWidget_8.setCurrentIndex(0)
        self.tabWidget_9.setCurrentIndex(0)
        for table in [self.tableBoxite13,self.tableLimestone13,self.tableWhiteMud13,self.tableCarbonMud13,self.tableSoda13,self.tableRecycledSolution13,self.tablePulp13,self.tableNaLosses13,self.tablePulp14,self.tableCrushedSpeck14,self.tableSinteringLosses14,self.tableCrushingLosses14,self.tableMud15,self.tableAluminateSolution15,self.tableSoda15,self.tableCarbonMud15,self.tableAlkalineSolution15,self.tableLosses15,self.tableCarbonMud16,self.tableAlkalineSolution16,self.tableWater16,self.tableCarbonMudStageTwo16,self.tableCarbonMudCharge16,self.tableAlkalineSolutionFromMud16,self.tablePromwater16,self.tableMud17,self.tableAluminateSolution17,self.tableSodaSolution17,self.tablePromwater17,self.tableCarbonMud17,self.tableAlkalineSolution17,self.tableAlkalineSolutionLeaching17,self.tableLosses17,self.tableCrushedSpeck18,self.tableAlkalineSolution18,self.tableWater18,self.tableRecycledSolution18,self.tablePromwater18,self.tableAluminateSolution18,self.tableAluminateSolution19,self.tableGrindWhiteMud19,self.tableAluminateSolutionWithMud19,self.tableWhiteMud19,self.tableSolution19,self.tableLosses19,self.tableAluminateSolutionStageTwo19,self.tableWhiteMud20,self.tableSolution20,self.tableRecycledMud20,self.tablePromwater20,self.tableWater20,self.tableAluminateSolutionStageOne21,self.tablePromwater21,self.tableCarbonMud21,self.tableCarbonMilk21,self.tableHydrogranate21,self.tableAluminateSolutionMud21,self.tableAluminateSolutionCarbon21,self.tableAluminateSolutionStageTwo22,self.tableAluminumHydroxide22,self.tableAlkalineSolution22,self.tableAluminumHydroxideS22,self.tableAlkalineSolutionS22,self.tableProdAluminumHydroxide22,self.tableProdAlkalineSolution22,self.tableLosses22,self.tableAluminumHydroxide23,self.tableAlkalineSolution23,self.tableWater23,self.tableProdAluminumHydroxide23,self.tablePromwater23,self.tablePromwaterS23,self.tableAluminumHydroxide24,self.tablePromwater24,self.tableAlumina24,self.tableLosses24,self.tableAlkalineSolution25,self.tableAlkalineSolution26,self.tablePromwater26,self.tableRecycledSolution26,self.tableLosses26,self.tableBoxite13_2,self.tableLimestone13_2,self.tableWhiteMud13_2,self.tableCarbonMud13_2,self.tableSoda13_2,self.tableRecycledSolution13_2,self.tablePulp13_2,self.tableNaLosses13_2,self.tablePulp14_2,self.tableCrushedSpeck14_2,self.tableSinteringLosses14_2,self.tableCrushingLosses14_2,self.tableMud15_2,self.tableAluminateSolution15_2,self.tableSoda15_2,self.tableCarbonMud15_2,self.tableAlkalineSolution15_2,self.tableLosses15_2,self.tableCarbonMud16_2,self.tableAlkalineSolution16_2,self.tableWater16_2,self.tableCarbonMudStageTwo16_2,self.tableCarbonMudCharge16_2,self.tableAlkalineSolutionFromMud16_2,self.tablePromwater16_2,self.tableMud17_2,self.tableAluminateSolution17_2,self.tableSodaSolution17_2,self.tablePromwater17_2,self.tableCarbonMud17_2,self.tableAlkalineSolution17_2,self.tableAlkalineSolutionLeaching17_2,self.tableLosses17_2,self.tableCrushedSpeck18_2,self.tableAlkalineSolution18_2,self.tableWater18_2,self.tableRecycledSolution18_2,self.tablePromwater18_2,self.tableAluminateSolution18_2,self.tableAluminateSolution19_2,self.tableGrindWhiteMud19_2,self.tableAluminateSolutionWithMud19_2,self.tableWhiteMud19_2,self.tableSolution19_2,self.tableLosses19_2,self.tableAluminateSolutionStageTwo19_2,self.tableWhiteMud20_2,self.tableSolution20_2,self.tableRecycledMud20_2,self.tablePromwater20_2,self.tableWater20_2,self.tableAluminateSolutionStageOne21_2,self.tablePromwater21_2,self.tableCarbonMud21_2,self.tableCarbonMilk21_2,self.tableHydrogranate21_2,self.tableAluminateSolutionMud21_2,self.tableAluminateSolutionCarbon21_2,self.tableAluminateSolutionStageTwo22_2,self.tableAluminumHydroxide22_2,self.tableAlkalineSolution22_2,self.tableAluminumHydroxideS22_2,self.tableAlkalineSolutionS22_2,self.tableProdAluminumHydroxide22_2,self.tableProdAlkalineSolution22_2,self.tableLosses22_2,self.tableAluminumHydroxide23_2,self.tableAlkalineSolution23_2,self.tableWater23_2,self.tableProdAluminumHydroxide23_2,self.tablePromwater23_2,self.tablePromwaterS23_2,self.tableAluminumHydroxide24_2,self.tablePromwater24_2,self.tableAlumina24_2,self.tableLosses24_2,self.tableAlkalineSolution25_2,self.tableAlkalineSolution26_2,self.tablePromwater26_2,self.tableRecycledSolution26_2,self.tableLosses26_2]:
            header = table.horizontalHeader()
            table.horizontalHeader().setStretchLastSection(False)
            # Устанавливаем режим изменения размера для каждой колонки
            if(header.count() > 5):
                for section in range(header.count()):
                    header.setSectionResizeMode(section, QHeaderView.Stretch)

        
            # Соединяем сигнал currentChanged с пользовательским слотом
            
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
        self.start_page.show()
