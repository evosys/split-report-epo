import sys
import time
import os
import appinfo
import itertools
import subprocess
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui import Ui_MainWindow
from pathlib import Path
from datetime import date
import xlrd
import xlsxwriter
import distutils.dir_util

SHEET1 = 'PO Uploaded'
SHEET2 = 'User Active'
SHEET3 = 'Production List User'
NEWDIR = str(date.today().strftime('%d-%m-%Y'))

# main class
class mainWindow(QMainWindow, Ui_MainWindow) :
    def __init__(self) :
        QMainWindow.__init__(self)
        self.setupUi(self)

        # app icon
        self.setWindowIcon(QIcon(':/resources/icon.png'))

        # centering app
        tr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        tr.moveCenter(cp)
        self.move(tr.topLeft())

        # button Open
        self.btOpen.clicked.connect(self.openXLS)

        # button convert
        self.btCnv.clicked.connect(self.BtnCnv)

        # status bar
        self.statusBar().showMessage('v'+appinfo._version)

        # hide label path
        self.lbPath.hide()
        self.lbPath.clear()



    # PATH FILE
    def openXLS(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","XLS Files (*.xlsx)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")



    # function xlrd
    def funcXLRD(self, SheetName) :
        # PATH file
        pathXLS = self.lbPath.text()

        if len(pathXLS) == 0:

            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)

        else :
            try :
                book = xlrd.open_workbook(pathXLS, ragged_rows=True)
                sheet = book.sheet_by_name(str(SheetName))

                return sheet

            except xlrd.XLRDError as e:
                msg = "Unsupported format, or corrupt file !"
                errorSrv = QMessageBox.critical(self, "Error", msg, QMessageBox.Abort)
                sys.exit(0)



    # function get cell range
    def get_cell_range(self, SheetName, start_col, start_row, end_col, end_row):
        sheet = self.funcXLRD(str(SheetName))
        return [sheet.row_values(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]



    # open file
    def open_file(self, filename):
        if sys.platform == "win32":
            os.startfile(filename)
        else:
            opener ="open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, filename])



    # create directory if not exist
    def CreateDir(self, cDIR, nDir, filename) :

        resPathFile =  Path(os.path.abspath(os.path.join(cDIR, nDir, "{}.xlsx".format(filename))))

        if os.path.exists(resPathFile) :
            os.remove(resPathFile)
        else :
            # os.makedirs(os.path.dirname(resPathFile), exist_ok=True)
            distutils.dir_util.mkpath(os.path.dirname(resPathFile))

        return resPathFile


    def getAllDataSheet(self, SheetName) :

        sheet = self.funcXLRD(str(SheetName))

        curr_row = 0
        row_list = []
        new_row = []

        while curr_row < (sheet.nrows - 1):
            curr_row += 1
            row = sheet.row_values(curr_row)
            row_list.append(row)
            # new_row.append([[ele.value for ele in each] for each in row_list])
        return row_list


    # Get AllData Sheet
    def getAllDataSheet1(self, SheetName) :
        result = []
        data1 = []
        actual = []

        sheet = self.funcXLRD(str(SheetName))
        totalrow = sheet.nrows
        totalcol = sheet.ncols
        print(totalrow)
        print(totalcol)

        for col in range(totalcol) :
            data1.append(sheet.cell_value(0, col))


        for row in range(1, totalrow) :
            elm = {}
            for col in range(totalcol) :
                elm[data1[col]] = sheet.cell_value(row, col)
            result.append(elm)

        for i in result :
            r = list(i.values())
            actual.append(r)

        return actual


    # Get AllData Custom sheet
    def getAllDataCustom(self, FirstROW, SheetName) :
        result = []
        data1 = []
        actual = []

        sheet = self.funcXLRD(str(SheetName))
        totalrow = sheet.nrows - 1
        totalcol = sheet.ncols - 1

        for col in range(totalcol) :
            data1.append(sheet.cell_value(1, col))


        for row in range(FirstROW, totalrow) :
            elm = {}
            for col in range(totalcol) :
                elm[data1[col]] = sheet.cell_value(row, col)
            result.append(elm)

        for i in result :
            r = list(i.values())
            actual.append(r)

        return actual


    # Get custom Header data Sheet
    def getHeaderDataTable(self, FirstROW, SheetName) :
        result = []

        sheet = self.funcXLRD(str(SheetName))
        totalcol = sheet.ncols - 1

        for col in range(totalcol+1) :
            result.append(sheet.cell_value(FirstROW, col))
            # print(totalcol)

        return result


    # get all KAR email on sheet
    def getAllKARSheet(self, SheetName) :
        result = []

        sheet = self.funcXLRD(str(SheetName))

        totalrow = sheet.nrows - 1

        rh = self.get_cell_range(SheetName, 2, 0, 2, totalrow)

        for z in rh :
            filt = filter(None, z)
            for i in filt :
                result.append(i)

        result.remove('KAR Email')
        return result


    # get unique KAR email on sheet
    def getUniqueAllKARSheet(self, SheetName) :
        KAR = self.getAllKARSheet(SheetName)

        # remove duplicate KAR
        result = [KAR[i] for i in range(len(KAR)) if i == KAR.index(KAR[i])]

        return result


    # split by KAR sheet
    def sheetsplit(self, SheetName, Sheet2 = False) :
        result = []

        if Sheet2 :
            dataList = dataList = self.getAllDataCustom(2, SheetName)

        else :
            dataList = self.getAllDataSheet(SheetName)

        for v in self.getUniqueAllKARSheet(SheetName) :
            # c = list(filter(lambda x:x[0]==v, dataList))
            x = self.search_nested_2d(dataList, v)
            result.append(x)

        return result


    def FileName(self, ListMail) :
        result = []

        for i in ListMail:
            x = i.split('@')[0]
            x = x.replace('.', '_')
            x = x.replace('-', '_')
            result.append(x.title())

        return result


    # search and filter 2D
    def search_nested_2d(self, mylist, filtering) :
        result = []
        for i in range(len(mylist)) :
            for j in range(len(mylist[i])) :
            # print i, j
                if mylist[i][j] == filtering :
                    result.append(mylist[i])

        return result

    # search and filter 3D
    def search_nested_3d(self, mylist, filtering) :
        result = []
        for i in range(len(mylist)) :
            for j in range(len(mylist[i])) :
                for k in range(len(mylist[i][j])) :
                # print i, j
                    if mylist[i][j][k] == filtering :
                        result.append(mylist[i])

        em = [e for sl in result for e in sl]

        cleanList = []
        for x in em:
            if x not in cleanList:
                cleanList.append(x)

        return cleanList


    def BtnCnv(self) :
        valSheet1 = self.sheetsplit(SHEET1)
        HeadSheet1 = self.getHeaderDataTable(0, SHEET1)
        HeadSheet2 = self.getHeaderDataTable(1, SHEET2)
        HeadSheet3 = self.getHeaderDataTable(0, SHEET3)
        print(valSheet1)


    def BtnCnv1(self) :
        current_dir = os.getcwd()
        # PATH file
        pathXLS = self.lbPath.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])

        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        # uniqueKAR1 = self.getUniqueAllKARSheet(SHEET1)
        uniqueKAR2 = self.getUniqueAllKARSheet(SHEET2)
        # uniqueKAR3 = self.getUniqueAllKARSheet(SHEET3)
        HeadSheet1 = self.getHeaderDataTable(0, SHEET1)
        HeadSheet2 = self.getHeaderDataTable(1, SHEET2)
        HeadSheet3 = self.getHeaderDataTable(0, SHEET3)
        ResSheet1 = self.sheetsplit(SHEET1)
        ResSheet2 = self.sheetsplit(SHEET2, True)
        ResSheet3 = self.sheetsplit(SHEET3)

        count = 0

        for NameFile in self.FileName(uniqueKAR2) :
            resPathFile = self.CreateDir(current_dir, NEWDIR, NameFile)

            workbook = xlsxwriter.Workbook(resPathFile, {'default_date_format': 'dd-mm-yy', 'strings_to_urls': True})

            # define worksheet
            ws1 = workbook.add_worksheet(SHEET1)
            ws2 = workbook.add_worksheet(SHEET2)
            ws3 = workbook.add_worksheet(SHEET3)

            # formating tab
            ws1.set_tab_color('orange')
            ws2.set_tab_color('green')
            ws3.set_tab_color('blue')


            # define formating cell
            headtable = workbook.add_format({'bold': 1, 'text_wrap': True, 'align': 'left', 'align': 'center', 'valign': 'vcenter', 'bg_color': '#92D14F'})
            bold = workbook.add_format({'bold': 1})
            orange = workbook.add_format({'bg_color': '#FFC000'})
            dateformat = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'num_format': 'dd-mm-yyyy', 'align': 'center'})
            bordering = workbook.add_format({'border': 1, 'border_color': '#a8a8a8'})
            kar_format = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'color': 'blue', 'underline': True, 'text_wrap': True})
            centering = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'align': 'center'})

            # freezing first row worksheet 1
            ws1.freeze_panes(1, 0)
            ws1.hide_gridlines(2)

            # formating diffrent color header worksheet 1
            ws1.conditional_format(0, 5, 0, len(HeadSheet1) - 1, {
                'type': 'cell',
                'criteria': '!=',
                'value': -100000,
                'format': orange})


            # write header worksheet 1
            ws1.write_row(0, 0, HeadSheet1, headtable)

            # formating column wokrsheet 1
            ws1.set_column(0, 0, 7.00, centering) # No
            ws1.set_column(1, 1, 24.29, bordering) # Area
            ws1.set_column(2, 2, 45.43, kar_format) # KAR Email
            ws1.set_column(3, 3, 22.29, centering) # Distributor Code
            ws1.set_column(4, 4, 50.86, bordering) # Distributor Name
            ws1.set_column(5, 5, 22.00, centering) # Store Code
            ws1.set_column(6, 6, 48.14, bordering) # Store Name
            ws1.set_column(7, 7, 21.86, centering) # PO Number
            ws1.set_column(8, 8, 16.43, dateformat) # Upload Date
            ws1.set_column(9, len(HeadSheet1) - 1, 16.43, centering) # PO Per-Mounth / PO Status / Status
            # ws1.set_column(10, 10, 16.43, centering) # PO Status
            # ws1.set_column(11, 11, 16.43, centering) # Status


            # writing data worksheet 1
            x = self.search_nested_3d(ResSheet1, uniqueKAR2[count])
            if x:
                for index, txtData in enumerate(x) :
                    ws1.write_row(index+1, 0, x[index])


            # create auto filter worksheet 1
            lengthData = len(x)
            ws1.autofilter(0, 0, lengthData, len(HeadSheet1) - 1)


            # formating diffrent color header worksheet 1
            ws2.conditional_format(1, 5, 1, len(HeadSheet2) - 1, {
                'type': 'cell',
                'criteria': '!=',
                'value': -100000,
                'format': orange})


            # write header worksheet 2
            ws2.write_row(1, 0, HeadSheet2, headtable)
            ws2.set_column(0, 0, 7.00, centering) # No
            ws2.set_column(1, 1, 24.29, bordering) # Area
            ws2.set_column(2, 2, 45.43, kar_format) # KAR Email
            ws2.set_column(3, 3, 22.29, centering) # Distributor Code
            ws2.set_column(4, 4, 50.86, bordering) # Distributor Name
            ws2.set_column(5, 5, 22.00, centering) # Store Code
            ws2.set_column(6, 6, 14.43, centering) # Store Code
            ws2.set_column(7, 7, 48.14, bordering) # Store Name
            ws2.set_column(8, len(HeadSheet2) - 1, 19.57, bordering)


            # close xls
            workbook.close()
            count+= 1

        reply = QMessageBox.information(self, "Information", "Success!", QMessageBox.Ok)

        if reply == QMessageBox.Ok :
            self.open_file(str(resultPath))



if __name__ == '__main__' :
    app = QApplication(sys.argv)

    # create splash screen
    splash_pix = QPixmap(':/resources/unilever_splash.png')

    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setWindowFlags(QtCore.Qt.FramelessWindowHint)
    splash.setEnabled(False)

    # adding progress bar
    progressBar = QProgressBar(splash)
    progressBar.setMaximum(10)
    progressBar.setGeometry(17, splash_pix.height() - 20, splash_pix.width(), 50)

    splash.show()

    for iSplash in range(1, 11) :
        progressBar.setValue(iSplash)
        t = time.time()
        while time.time() < t + 0.1 :
            app.processEvents()

    time.sleep(1)

    window = mainWindow()
    window.setWindowTitle(appinfo._appname)
    # window.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
    # window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)
    window.show()
    splash.finish(window)
    sys.exit(app.exec_())