import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtCore, QtGui
from dateutil.parser import parse
import pandas as pd
import pymysql

form_class = uic.loadUiType("conversionQuery.ui")[0]

class WindowClass(QMainWindow, form_class) :
    findCSV = False

    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        QDialog().setFixedSize(self.size())

        self.findFileButton.clicked.connect(self.FindFile)
        self.saveButton.clicked.connect(self.SaveFile)
        self.conversionButton.clicked.connect(self.Conversion)
        self.csvTableWidget.cellClicked.connect(self.SetLabel)
        self.csvTableWidget.itemChanged.connect(self.ItemChanged)
        self.dmlComboBox.currentIndexChanged.connect(self.DmlAction)
        self.whereCheck.stateChanged.connect(self.WhereAction)

        self.spinBox.valueChanged.connect(self.SpinChanged)
        self.spinBox.setMinimum(1)
        self.spinBox.setMaximum(1)

        self.whereCheck.hide()
        self.spinBox.hide()


    def FindFile(self):
        global fname
        fname = QFileDialog.getOpenFileName(self, 'Select File', './', 'cell file(*.csv *xlsx);; all file(*)')

        if not fname[0]:
            return

        fnameTmp = fname[0].replace(".", "/")
        fnameSplit = fnameTmp.split('/')
        
        global df
        if fnameSplit[-1] == 'csv':
            df = pd.read_csv(fname[0], encoding='CP949')
        elif fnameSplit[-1] == 'xlsx':
            df = pd.read_excel(fname[0], engine='openpyxl')
        else:
            self.tableNameEdit.setText('엑셀파일 넣으세요!')
            return

        if 'SEQ' in df.columns:
            df.drop(['SEQ'], axis=1, inplace=True)

        self.csvTableWidget.blockSignals(True)

        self.csvTableWidget.clearSelection()

        WindowClass.findCSV = True

        self.csvTableWidget.setColumnCount(len(df.columns))
        self.csvTableWidget.setRowCount(len(df.index))
        self.csvTableWidget.setHorizontalHeaderLabels(df.columns)
        self.csvTableWidget.resizeColumnsToContents()

        for i in range(len(df.index)):
            for j in range(len(df.columns)):
                self.csvTableWidget.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))
                self.csvTableWidget.item(i, j).setTextAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)

        self.csvTableWidget.resizeColumnsToContents()

        self.csvTableWidget.blockSignals(False)

        self.tableNameEdit.setText(fnameSplit[-2])

        labelString = 'row: ' + str(len(df.index)) + ', column: ' + str(len(df.columns))
        self.sizeLabel.setText(labelString)

        self.findFileButton.setText('테이블 명')

        self.spinBox.setMaximum(len(df.columns))

        self.warningLabel.clear()
        self.warningLabel.setStyleSheet("Color : black")

    def SaveFile(self):
        if not WindowClass.findCSV:
            self.warningLabel.setText('먼저 파일부터 찾으세요!')
            return

        FileSave = QFileDialog.getSaveFileName(self, 'Save file', './', 'cell file(*.csv *xlsx);; all file(*)')

        if not FileSave[0]:
            return

        df.to_csv(FileSave[0], sep=',', index=False, encoding='CP949')

    def ItemChanged(self, item):
        df.loc[item.row(), df.columns[item.column()]] = item.text()

    def SetLabel(self, row, column):
        item = self.csvTableWidget.item(row, column)
        header = self.csvTableWidget.horizontalHeaderItem(column).text()
        value = item.text()
        labelString = str(header) + ' : ' + str(value)
        self.valueLabel.setText(labelString)

    def DmlAction(self):
        if self.dmlComboBox.currentIndex() == 0:
            self.csvTableWidget.clearSelection()
            self.csvTableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.csvTableWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
            self.truncateCheck.show()
            self.whereCheck.hide()
            self.spinBox.hide()
        else:
            self.csvTableWidget.clearSelection()
            self.csvTableWidget.setSelectionBehavior(QAbstractItemView.SelectColumns)
            self.whereCheck.show()
            if not self.whereCheck.isChecked():
                self.csvTableWidget.setSelectionMode(QAbstractItemView.NoSelection)
                self.spinBox.show()
                self.csvTableWidget.setRangeSelected(QTableWidgetSelectionRange(0, 0, len(df.index) - 1, self.spinBox.value() - 1), True)

            self.truncateCheck.hide()
    
    def WhereAction(self):
        self.csvTableWidget.clearSelection()
        if self.whereCheck.isChecked():
            self.spinBox.hide()
            self.csvTableWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        else:
            self.spinBox.show()
            self.csvTableWidget.setSelectionMode(QAbstractItemView.NoSelection)
            self.csvTableWidget.setRangeSelected(QTableWidgetSelectionRange(0, 0, len(df.index) - 1, self.spinBox.value() - 1), True)

    def SpinChanged(self):
        self.csvTableWidget.clearSelection()
        if self.dmlComboBox.currentIndex() == 1 and not self.whereCheck.isChecked():
            self.csvTableWidget.setSelectionMode(QAbstractItemView.NoSelection)
            if WindowClass.findCSV:
                self.csvTableWidget.setRangeSelected(QTableWidgetSelectionRange(0, 0, len(df.index) - 1, self.spinBox.value() - 1), True)


    def Conversion(self):
        def DateTimeCast(dateType, data):
            AMPM = ""
            for ampm in ["am", "Am", "AM", "pm", "Pm", "PM"]:
                if ampm in data:
                    data = data.replace(ampm,"")
                    AMPM = ampm

            if data == "0000-00-00 00:00:00" or data == "0000-00-00":
                if dateType == "date":
                    return "\"0000-00-00\""
                else:
                    return "\"0000-00-00 00:00:00\""

            dt = parse(data)
            if AMPM in ["am", "Am", "AM"]:
                if dt.hour == 12:
                    dt = dt.replace(hour = 0)
            elif AMPM in ["pm", "Pm", "PM"]:
                if 0 <= dt.hour <= 11:
                    dt = dt.replace(hour = dt.hour+12)
            if dateType == "date":
                data = "\"" + dt.strftime("%Y-%m-%d") + "\""
            else:
                data = "\"" + dt.strftime("%Y-%m-%d %H:%M:%S") + "\""
            return data

        self.warningLabel.clear()
        self.warningLabel.setStyleSheet("Color : black")

        if not WindowClass.findCSV:
            self.warningLabel.setText('먼저 파일부터 찾으세요!')
            return

        dmlIdx = self.dmlComboBox.currentIndex()

        try:
            db = pymysql.connect(host='localhost', port=3306, user='root', passwd='0000', db='information_schema', charset='utf8')
        except pymysql.err.OperationalError as e:
            self.warningLabel.setText(e.args[1])
            self.warningLabel.setStyleSheet("Color : red")
            return
        sql = "SELECT COUNT(*) FROM columns WHERE NOT column_name = \"SEQ\" AND table_name=" + "\"" + self.tableNameEdit.text() + "\""

        curs = db.cursor()
        curs.execute(sql)
        column_num = curs.fetchone()[0]

        if column_num != len(df.columns):
            self.warningLabel.setText("경고! 컬럼 수가 맞지 않습니다. 신중하게 확인하시길 바랍니다!")
            self.warningLabel.setStyleSheet("Color : red")
            curs.close()
            db.close()
            return

        if column_num == 0:
            self.warningLabel.setText("경고! DB에 데이터가 없습니다. 테이블명을 확인해주세요!")
            self.warningLabel.setStyleSheet("Color : red")
            curs.close()
            db.close()
            return

        sql = "SELECT COLUMN_NAME, DATA_TYPE FROM columns WHERE NOT column_name = \"SEQ\" AND table_name=" + "\"" + self.tableNameEdit.text() + "\""
        curs.execute(sql)
        column_name_types = curs.fetchall()
        selectedRange = self.csvTableWidget.selectedRanges()

        curs.close()
        db.close()

        if dmlIdx == 0:
            selectedRow = []
            for selected in selectedRange:
                if selected.columnCount() == len(df.columns):
                    selectedRow += range(selected.topRow(), selected.bottomRow() + 1)
            selectedRow.sort()

            if selectedRow == []:
                selectedRow = df.index.values.tolist()

            insert_str = ""

            if self.truncateCheck.isChecked() == True:
                insert_str += "TRUNCATE TABLE " + self.tableNameEdit.text() + ";\n"

            insert_str += "INSERT INTO " + self.tableNameEdit.text() + " ("

            for column in column_name_types:
                if column == column_name_types[0]:
                    insert_str += column[0]
                else:
                    insert_str += ", " + column[0]
            insert_str += ") VALUES\n"

            for i, row in df.iterrows():
                if i not in selectedRow:
                    continue

                insert_str += "("
                for j, value in enumerate(row):
                    if value == "\\N":
                        value = "NULL"
                    elif column_name_types[j][1] in ["int", "bigint", "smallint", "tinyint"]:
                        value = str(int(value))
                    elif column_name_types[j][1] in ["char", "varchar"]:
                        value = "\"" + str(value) + "\""
                    elif column_name_types[j][1] in ["date", "datetime"]:
                        value = DateTimeCast(column_name_types[j][1], value)

                    if j == 0:
                        insert_str += value
                    else:
                        insert_str += ", " + value
                insert_str += ")"
                if i != selectedRow[-1]:
                    insert_str += ",\n"
            insert_str += ";"
            self.queryTextEdit.setText(insert_str)

        else:
            selectedColumn = []
            for selected in selectedRange:
                if selected.rowCount() == len(df.index):
                    selectedColumn += range(selected.leftColumn(), selected.rightColumn() + 1)
            if selectedColumn == []:
                self.queryTextEdit.setText("where 조건에 맞는 column을 선택해주세요!")
                return
            selectedColumn.sort()

            update_str = ""

            last_change_word = 0
            for num in reversed(range(len(df.columns))):
                if num not in selectedColumn:
                    last_change_word = num
                    break
            for row in df.itertuples(index=False):
                update_str += "UPDATE " + self.tableNameEdit.text() + " SET "
                for i, value in enumerate(row):
                    if i in selectedColumn:
                        continue
                    if value == "\\N":
                        value = "NULL"
                    elif column_name_types[i][1] in ["int", "bigint", "smallint", "tinyint"]:
                        value = str(int(value))
                    elif column_name_types[i][1] in ["char", "varchar"]:
                        value = "\"" + str(value) + "\""
                    elif column_name_types[i][1] in ["date", "datetime"]:
                        value = DateTimeCast(column_name_types[i][1], value)

                    update_str += column_name_types[i][0] + " = " + value
                    if i != last_change_word:
                        update_str += ", "
                for i, value in enumerate(row):
                    if i not in selectedColumn:
                        continue
                    if i == selectedColumn[0]:
                        update_str += " WHERE "

                    if value == "\\N":
                        value = "NULL"
                    elif column_name_types[i][1] in ["int", "bigint", "smallint", "tinyint"]:
                        value = str(int(value))
                    elif column_name_types[i][1] in ["char", "varchar"]:
                        value = "\"" + str(value) + "\""
                    elif column_name_types[i][1] in ["date", "datetime"]:
                        value = DateTimeCast(column_name_types[i][1], value)

                    update_str += column_name_types[i][0] + " = " + value
                    if i != selectedColumn[-1]:
                        update_str += " AND "

                update_str += ";\n"

            self.queryTextEdit.setText(update_str)

if __name__ == "__main__" :
    app = QApplication(sys.argv) 
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()