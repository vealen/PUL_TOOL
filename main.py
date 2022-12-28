import time


from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTreeWidgetItem, QInputDialog
from mainWindow_ui import Ui_MainWindow
from purpul_be import *
from utils import *

#ROOT = os.path.dirname(sys.executable)
ROOT = os.path.dirname(__file__)


class mainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # self.ui = Ui_MainWindow() #klasa dziediczy po Ui_MainWindow, aiwec nie trtzeba tak definiowac
        self.setupUi(self)
        self.connect_gui_elements()

        print(f'Drajwery w pyodbc to:{pyodbc.drivers()}')
        self.db = None
        self.dest_path = None
        self.raportyGui = []
        self.adr_for_lines = {
            'RDLP': (self.rdlp, 2),
            # find Qline edit, integer to dlugosc jaka dana kategoria powinna miec
            'NADL': (self.ndl, 2),
            'OBR': (self.obr, 1),
            'LCTWO': (self.lctw, 2),
            'ODDZ': (self.oddz, 6),
            'PODODDZ': (self.pododdz, 4),
            'WYDZ': (self.wydz, 2)
        }
        self.create_raports(self.raportTree, RAPORTS)

    def connect_gui_elements(self):
        self.dbButton.clicked.connect(self.get_data_base)
        self.raportButton.clicked.connect(self.get_raport_dest)
        self.run.clicked.connect(self.final)

    def check_chosen_raports(self): #cholera wie co ja tu zrobiłem, ale tr
        count = 0
        wybrane_raporty = []
        for raport_category in RAPORTS:
            for raport_name in RAPORTS[raport_category]:
                sheet_name = RAPORTS[raport_category][raport_name][2]
                raport_gui_element = RAPORTS[raport_category][raport_name][3]
                RAPORTS[raport_category][raport_name][1] = False  # reset statusów
                if raport_gui_element.checkState(0) == Qt.Checked:
                    RAPORTS[raport_category][raport_name][1] = True
                    count += 1
                    wybrane_raporty.append(sheet_name)
        return count, wybrane_raporty

    def create_raports(self, raportTree, raportDict):
        for key in raportDict.keys():
            parent = QTreeWidgetItem(raportTree)
            parent.setText(0, f'{key}')
            parent.setFlags(parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            parent.setFlags(parent.flags() & ~Qt.ItemIsSelectable)
            self.raportyGui.append(parent)
            parent.setDisabled(True)
            for value in raportDict[key]:
                child = QTreeWidgetItem(parent)
                child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                child.setFlags(child.flags() & ~Qt.ItemIsSelectable)
                child.setText(0, f'{value}')
                child.setCheckState(0, Qt.Unchecked)
                if key in ('Lista Kontrolna','DUpa'):
                    raportDict[key][value].append(child)


    def get_data_base(self):
        self.base_path, self.base_ext = QFileDialog.getOpenFileName(caption="Wskaż Bazę", filter="Baza (*.mdb)")
        if self.base_path:
            self.baseLineEdit.setText(self.base_path)
            self.db = connect_db(self.base_path)[0]
            wydz_Data = get_table_data(self.db, SQL['WYDZ'])
            self.adres = wydz_Data['TEMP_ADRESS_FOREST'].values[0]
            adr_for_dict = create_adr_for_dict(self.adres, list=False)
            for raport_Cat in self.raportyGui:
                raport_Cat.setDisabled(False)
            for raport_category in RAPORTS:
                for raport_name in RAPORTS[raport_category]:
                    raport_gui_element = RAPORTS[raport_category][raport_name][3]
                    raport_gui_element.setDisabled(False)

            for keya in self.adr_for_lines:
                for keyb in adr_for_dict:
                    if keya == keyb:
                        self.adr_for_lines[keya][0].setText(adr_for_dict[keyb])
            message('Sukces', f'Połączono z bazą danych {os.path.basename(self.base_path)}',
                    QMessageBox.Information)

    def get_raport_dest(self):
        default_dir = os.path.join(ROOT, 'results')
        path = QFileDialog.getExistingDirectory(caption='Wskaż scieżkę do utworzenia raportu',
                                                directory=default_dir)
        if path:
            self.dest_path = path
            self.raportLineEdit.setText(self.dest_path)

    def construct_adr_for(self):  # dobrze robiacy adr_FOR bardzo fajne rozwiazanie
        adr_for = ''
        for key in self.adr_for_lines:
            qline_content, correct_lenght = self.adr_for_lines[key]
            if len(qline_content.text()) != correct_lenght and len(qline_content.text()) != 0:
                return message('Błąd w długości pola w adr_for',
                               f'Pole {key} ma złą długość, sprawdź i popraw. Uważaj na spacje',
                               QMessageBox.Warning), False
            else:
                if qline_content.text().isspace() or len(qline_content.text()) == 0:
                    adr_for += '%'
                    break
                else:
                    if key != 'WYDZ':
                        adr_for += (qline_content.text() + '-')
                    else:
                        adr_for += qline_content.text()
        adr_for_dict = create_adr_for_dict(adr_for, False)  # create dict by set adr_For
        for keya in self.adr_for_lines:  # ustaw nowe wartosci sprecyzowane przez uzytkownika ( tylko wyglad, nie ma
            # wplywu na ostatefczny adr_For ktory zostanie przekazany do kwerendy
            for keyb in adr_for_dict:
                if keya == keyb:
                    self.adr_for_lines[keya][0].setText(adr_for_dict[keyb])
                    if self.adr_for_lines[keya][0].text() == '%':
                        self.adr_for_lines[keya][0].setText('')

        return adr_for, True

    def raport_name_dialog(self):
        text, result = QInputDialog.getText(self, 'Raport', 'Podaj nazwę raportu:')
        if result:
            return text

    def final(self):
        count, wybrane_raporty = self.check_chosen_raports()

        if not self.db:
            message('Nie widzę bazy', 'Wskaż bazę danych i spróbuj ponownie', QMessageBox.Warning)
        elif not self.dest_path:
            message('Nie widzę ścieżki docelowej', 'Wskaż ścieżkę do utworzenia raportu i spróbuj ponownie',
                    QMessageBox.Warning)
        elif count == 0:
            message('Nie wybrano raportu','Wybierz raporty do utworzenia i spróbuj ponownie', QMessageBox.Warning)
        else:
            adr_for = self.construct_adr_for()
            if adr_for[1]:
                raport_name_user = self.raport_name_dialog()
                if raport_name_user and not raport_name_user.isspace():
                    date_dict = {'//AKTUAL_DATE//': str(datetime.date.today().strftime('%d/%m/%Y'))}

                    dialog, bar = progdialog(0, 'Wysyłam zapytanie SQL',
                                             f'Rysuję tabele...\n Muszę narysować: {count} tabel')
                    bar.setValue(0)
                    bar.setMaximum(100)
                    progress = 0
                    QApplication.processEvents()
                    i = 0  # wartosc i pomocnicza do ustawiania progress bara
                    # adr_for = self.construct_adr_for()
                      # ilosc wybranych raportow

                    ## SPRAWDZANIE KTORY PARENT W MENU JEST WYBRANY -> DO KOPIOWANIA ODPOWIEDNIEGO SCHEMATU
                    # if self.raportyGui[0].checkState(0)==Qt.PartiallyChecked or self.raportyGui[0].checkState(0)==Qt.Checked :
                    #     print(f'Wybrano {self.raportyGui[0].text(0)}')
                    # if self.raportyGui[1].checkState(0)==Qt.PartiallyChecked or self.raportyGui[1].checkState(0)==Qt.Checked :
                    #     print(f'Wybrano {self.raportyGui[1].text(0)}')

                    new_raport = copy_blank_raport(ROOT, self.dest_path, raport_name_user,'SCHEME_LIST_KONTR.xlsx')

                    if not new_raport:
                        message('Lista Kontrolna','Plik o takiej nazwie już istnieje, wskaż inną ',QMessageBox.Warning)
                    else:
                        delete_empty_raports(new_raport, wybrane_raporty)
                        filter = adr_for[0]
                        print(filter)
                        if adr_for[1]:  # must be true

                            for raport_category in RAPORTS:
                                for raport_name in RAPORTS[raport_category]:
                                    sql_ = RAPORTS[raport_category][raport_name][0]
                                    sheet_name = RAPORTS[raport_category][raport_name][2]
                                    if RAPORTS[raport_category][raport_name][1]:  # #teoretycznie moznabyloby sprawdzic check
                                        # tak jak w w
                                        # check_chosen_raports, ale niech będzie, chyba tak jest troszkę szybciej. CHYBA
                                        # print(raport_gui_element, raport_name, sql_)
                                        self.clear_temp_tables(self.db)
                                        if raport_name in ['Porol_1', 'Porol_2']:
                                            POROL_1A_SQL = """SELECT F_SUBAREA.AREA_TYPE_CD, F_SUBAREA.SOIL_PEC_CD, F_ARODES.TEMP_ADRESS_FOREST, F_SUBAREA.ARODES_INT_NUM, F_SUBAREA.SUB_AREA INTO POROL_1At
FROM F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM
WHERE (((F_SUBAREA.SOIL_PEC_CD)='POROL') AND ((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_ARODES.TEMP_ACT_ADRESS)=True));
"""
                                            POROL_1B_SQL = """SELECT F_SUBAREA.AREA_TYPE_CD, F_ARODES.TEMP_ADRESS_FOREST, F_AROD_STAND_PEC.FOREST_PEC_CD, F_SUBAREA.SUB_AREA, F_SUBAREA.ARODES_INT_NUM INTO POROL_1Bt
FROM (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_STAND_PEC ON F_SUBAREA.ARODES_INT_NUM = F_AROD_STAND_PEC.ARODES_INT_NUM
WHERE (((F_SUBAREA.AREA_TYPE_CD)='D-STAN') AND ((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_AROD_STAND_PEC.FOREST_PEC_CD)='POROL') AND ((F_ARODES.TEMP_ACT_ADRESS)=True));
"""
                                            cursor = self.db.cursor()



                                            cursor.execute(POROL_1A_SQL, filter)
                                            self.db.commit()
                                            cursor.execute(POROL_1B_SQL, filter)
                                            self.db.commit()
                                            cursor.close()

                                            data = get_table_data(self.db, sql_, filter=None)
                                        elif raport_name == 'Zlożenie GTD':
                                            self.zlozenie_GTD_create_goals(self.db,filter)
                                            data = get_table_data(self.db,sql_,[filter])
                                        else: #standardowe zapytania sql z filtrem
                                            data = get_table_data(self.db, sql_, [filter])

                                        to_excel(data, new_raport, sheet_name)


                                        # ustaw zmiany na progress barze
                                        i += 1
                                        progress = i / count * 100
                                        bar.setValue(int(progress))
                                        QApplication.processEvents()
                        self.clear_temp_tables(self.db)
                        if progress == 100:
                            replace_by_dict(date_dict, new_raport, 'LANDING_PAGE') #replace tutaj bo zawsze trzeba zmienic
                            time.sleep(0.25)
                            message('Lista Kontrolna','Ukończono tworzenie raportu', QMessageBox.Information)
                            # pytanie = msg_question(f'Sukces {raport_name_user}',
                            #                        'Ukończono tworzenie raportu, zamknąć aplikacje?',
                            #                        self.centralwidget)
                            # if pytanie:
                            #     sys.exit()
    def clear_temp_tables(self,db): #usuwa tymczasowe tabele (do kwerend POROL_1 i POROL_2, ZLOZENIE GTD)
        cursor  =db.cursor()
        temp_tables_names = ('POROL_1At','POROL_1Bt','GOAL1t','GOAL2t','GOAL3t','GOAL4t')
        for name in temp_tables_names:
            try:
                cursor.execute(f"""DROP TABLE {name}""")
                print(f'Usunięto tabelę {name}')
                db.commit()
            except:
                print(f"Brak tabeli {name}, jest OK")
        cursor.close()

    def zlozenie_GTD_create_goals(self,db,adr_filter):
        cursor = db.cursor()
        for goal in range(1,5):
            goal_sql = f"""SELECT F_AROD_GOAL.ARODES_INT_NUM, F_SUBAREA.MOISTURY_CD, F_SUBAREA.SITE_TYPE_CD, F_AROD_GOAL.SPECIES_CD, F_ARODES.TEMP_ADRESS_FOREST, F_AROD_GOAL.GOAL_RANK_ORDER, F_SUBAREA.AREA_TYPE_CD INTO GOAL{str(goal)}t
FROM (F_ARODES INNER JOIN F_SUBAREA ON F_ARODES.ARODES_INT_NUM = F_SUBAREA.ARODES_INT_NUM) INNER JOIN F_AROD_GOAL ON F_SUBAREA.ARODES_INT_NUM = F_AROD_GOAL.ARODES_INT_NUM
WHERE (((F_AROD_GOAL.GOAL_TYPE_FL)='d') AND ((F_ARODES.TEMP_ACT_ADRESS)=True))
GROUP BY F_AROD_GOAL.ARODES_INT_NUM, F_SUBAREA.MOISTURY_CD, F_SUBAREA.SITE_TYPE_CD, F_AROD_GOAL.SPECIES_CD, F_ARODES.TEMP_ADRESS_FOREST, F_AROD_GOAL.GOAL_RANK_ORDER, F_SUBAREA.AREA_TYPE_CD
HAVING (((F_ARODES.TEMP_ADRESS_FOREST) Like ?) AND ((F_AROD_GOAL.GOAL_RANK_ORDER)=?));
"""
            cursor.execute(goal_sql,(adr_filter,str(goal)))
            db.commit()
        cursor.close()






if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = mainWindow()
    window.show()
    app.setStyle('fusion')
    app.exec_()
