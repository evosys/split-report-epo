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
from traceback import format_exception


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

        sys.excepthook = self.excepthook


    def excepthook(self, type_, value, traceback) :

        msg = format_exception(type_, value, traceback)

        errorhandle = len(msg)

        QMessageBox.critical(self, "Error", msg[errorhandle-1], QMessageBox.Abort)
        print(msg)


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

    def insert_position(self, position, list1, list2):
        return list1[:position] + list2 + list1[position:]


    # Get custom Header data Sheet
    def getHeaderDataTable(self, FirstROW, SheetName, Additional = False) :
        tmp = []
        res = []
        sheet = self.funcXLRD(str(SheetName))

        ### Adding '' to event
        # if Additional :
        #     result = list(filter(None, sheet.row_values(FirstROW)))
        #     keep = result[0]

        #     for addHead in range(2, len(result), 2) :
        #         addHead = ''
        #         tmp.append(addHead)

        #     result.remove(result[0])

        #     i = 1
        #     while i < len(result) :
        #         result.insert(i, '')
        #         i += 2

        #     result.insert(0, keep)

        #     return result
        if Additional :
            result = list(filter(None, sheet.row_values(FirstROW)))
            result.remove(result[0])
            return result

        else :
            result = list(filter(None, sheet.row_values(FirstROW)))

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

    # get all distributor on sheet
    def getAllDISTSheet(self, SheetName) :
        result = []

        sheet = self.funcXLRD(str(SheetName))

        totalrow = sheet.nrows - 1

        rh = self.get_cell_range(SheetName, 4, 0, 4, totalrow)

        for z in rh :
            filt = filter(None, z)
            for i in filt :
                result.append(i)

        result.remove('Distributor Name')
        return result


    # get unique KAR email on sheet
    def getUniqueAllKARSheet(self, SheetName) :
        KAR = self.getAllKARSheet(SheetName)

        # remove duplicate KAR
        result = [KAR[i] for i in range(len(KAR)) if i == KAR.index(KAR[i])]
        result = sorted(result)

        return result


    # get unique Distributor on sheet
    def getUniqueAllDISTSheet(self, SheetName) :
        res = []
        oth = []

        DIST = self.getAllDISTSheet(SheetName)
        KAR = self.getAllKARSheet(SheetName)
        # tmpKAR = self.getUniqueAllKARSheet(SheetName)

        doub = [[i, j] for i, j in zip(KAR, DIST)]
        # doub = list(zip(KAR, DIST))
        # print(len(KAR))
        # print(len(DIST))
        sorting = sorted(doub)



        # for key, grp in itertools.groupby(doub, key = lambda x:x[0]) :
            # b = [key, [n for _, n in grp]]
            # res.append(b)

        # print(doub)

        # grouping list by KAR name
        result = [[key, [n for _, n in grp]] for key, grp in itertools.groupby(sorting, key=lambda x: x[0])]

        # remove duplicate DIST
        # result = [doub[i] for i in range(len(doub[1])) if i == doub.index(doub[i])]




        return result


    # split by KAR sheet
    def sheetsplit(self, SheetName, Sheet2 = False) :
        result = []

        dataList = self.getAllDataSheet(SheetName)

        for v in self.getUniqueAllKARSheet(SheetName) :
            # c = list(filter(lambda x:x[0]==v, dataList))
            x = self.search_nested_2d(dataList, v)
            result.append(x)

        return result


    def FileNameKAR(self, ListMail) :
        result = []

        for i in ListMail:
            x = i.split('@')[0]
            x = x.replace('.', '_')
            x = x.replace('-', '_')
            result.append(x.title())

        return result

    def FileNameDist(self, ListDist) :

        KARres = []
        result = []
        # result = [KAR[i] for i in range(len(KAR)) if i == KAR.index(KAR[i])]

        for i in ListDist :
            KARres.append(i[1])

        for k in KARres :
            x = self.removeDuplicatesCustom(k)
            all_but_last = ', '.join(x[:-1])
            last = x[-1]

            j = " & ".join([", ".join(x[:-1]),x[-1]] if len(x) > 2 else x)

            result.append(j)

        return result

    def removeDuplicatesCustom(self, listofElements):

        # Create an empty list to store unique elements
        uniqueList = []

        # Iterate over the original list and for each element
        # add it to uniqueList, if its not already there.
        for elem in listofElements:
            if elem not in uniqueList:
                uniqueList.append(elem)

        # Return the list of unique elements
        return uniqueList


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



    def getEvent(self, firstCol, listHead) :
        result = []
        for mrged in range(firstCol, len(listHead), 2) :
            result.append(mrged)

        return result


    def removeFistData(self, ListData) :
        return ListData.remove(ListData[0])


    def BtnCnv1(self) :
        valSheet1 = self.sheetsplit(SHEET1)
        HeadSheet1 = self.getHeaderDataTable(0, SHEET1)
        HeadSheet2 = self.getHeaderDataTable(1, SHEET2)
        HeadSheet3 = self.getHeaderDataTable(0, SHEET3)
        AllDist = self.getAllDISTSheet(SHEET1)
        AllDistUnique = self.getUniqueAllDISTSheet(SHEET2)
        uniqueKAR2 = self.getUniqueAllKARSheet(SHEET2)
        x = self.FileNameDist(AllDistUnique)


        # print(len(x))
        print(x)
        # print(uniqueKAR2)

    def BtnCnv(self) :
        current_dir = os.getcwd()
        # PATH file
        pathXLS = self.lbPath.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])

        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        # uniqueKAR1 = self.getUniqueAllKARSheet(SHEET1)
        AllKARUnique = self.getUniqueAllKARSheet(SHEET2)
        AllDISTUnique = self.getUniqueAllDISTSheet(SHEET2)
        # uniqueKAR3 = self.getUniqueAllKARSheet(SHEET3)
        HeadSheet1 = self.getHeaderDataTable(0, SHEET1)
        HeadSheet2 = self.getHeaderDataTable(1, SHEET2)
        HeadSheet3 = self.getHeaderDataTable(0, SHEET3)

        ResSheet1 = self.sheetsplit(SHEET1)
        ResSheet2 = self.sheetsplit(SHEET2, True)
        ResSheet3 = self.sheetsplit(SHEET3)

        AdditionalHeadSheet2 = self.getHeaderDataTable(0, SHEET2)
        MergedData = self.getHeaderDataTable(0, SHEET2, True)

        count = 0

        for NameFile in self.FileNameDist(AllDISTUnique) :
            resPathFile = self.CreateDir(current_dir, NEWDIR, NameFile)

            workbook = xlsxwriter.Workbook(resPathFile, {'default_date_format': 'dd-mm-yy', 'strings_to_urls': True})

            # define worksheet
            ws1 = workbook.add_worksheet(SHEET1)
            ws2 = workbook.add_worksheet(SHEET2)
            ws3 = workbook.add_worksheet(SHEET3)

            # formating tab
            ws1.set_tab_color('#9CCC65')
            ws2.set_tab_color('#AED581')
            ws3.set_tab_color('#C5E1A4')


            # define formating cell
            headtable = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'bold': 1, 'text_wrap': True, 'align': 'left', 'align': 'center', 'valign': 'vcenter', 'bg_color': '#87D37C'})
            bold = workbook.add_format({'bold': 1})
            orange = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'bg_color': '#F5D76E'})
            softYellow = workbook.add_format({'bg_color': '#fceb9f'})
            softGreen = workbook.add_format({'bg_color': '#87D37C'})
            softGrey = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'bg_color': '#DADFE1', 'bold': 1, 'text_wrap': True, 'align': 'center'})
            YesGreen = workbook.add_format({'font_color': '#006100', 'bg_color': '#c6efcd'})
            NoRed = workbook.add_format({'font_color': '#9c0006', 'bg_color': '#ffc8ce'})
            dateformat = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'num_format': 'dd-mm-yyyy', 'align': 'center'})
            datetimeformat = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'num_format': 'dd-mm-yyyy hh:mm:ss', 'align': 'center'})
            bordering = workbook.add_format({'border': 1, 'border_color': '#a8a8a8'})
            kar_format = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'color': 'blue', 'underline': True, 'text_wrap': True})
            centering = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'align': 'center'})
            lefting = workbook.add_format({'border': 1, 'border_color': '#a8a8a8', 'align': 'left'})

            # freezing first row worksheet 1
            ws1.freeze_panes(1, 0)
            ws1.hide_gridlines(2)

            # formating diffrent color first header worksheet 1
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
            ws1.set_column(3, 3, 20.00, centering) # Distributor Code
            ws1.set_column(4, 4, 50.86, bordering) # Distributor Name
            ws1.set_column(5, 5, 23.86, lefting) # Store Code
            ws1.set_column(6, 6, 48.14, bordering) # Store Name
            ws1.set_column(7, 7, 21.86, centering) # PO Number
            ws1.set_column(8, 8, 16.43, dateformat) # Upload Date
            ws1.set_column(9, len(HeadSheet1) - 1, 16.43, centering) # PO Per-Mounth / PO Status / Status


            # writing data worksheet 1
            x = self.search_nested_3d(ResSheet1, AllKARUnique[count])
            if x:
                for index, txtData in enumerate(x) :
                    ws1.write_row(index+1, 0, x[index])

                    if len(x) != 0 :
                        ws1.write_number(index+1, 0, index+1)


            # create auto filter worksheet 1
            lengthData = len(x)
            ws1.autofilter(0, 0, lengthData, len(HeadSheet1) - 1)


            #################################################################

            # Worksheet 2

            # ###############################################################

            # freezing first row worksheet 2
            ws2.freeze_panes(2, 0)
            ws2.hide_gridlines(2)

            # formating diffrent color header worksheet 2
            ws2.conditional_format(0, 7, 0, 7, {
                'type': 'cell',
                'criteria': '!=',
                'value': -100000,
                'format': softGreen})


            # formating diffrent color header worksheet 2
            ws2.conditional_format(1, 5, 1, len(HeadSheet2) - 1, {
                'type': 'cell',
                'criteria': '!=',
                'value': -100000,
                'format': orange})


            # write header worksheet 2
            ws2.write_row(0, 7, AdditionalHeadSheet2, headtable)
            # merged Data Additional Head
            MrgEvent = self.getEvent(8, HeadSheet2)
            MrgOdd = self.getEvent(9, HeadSheet2)

            for idx, dta in enumerate(MrgEvent) :
                ws2.merge_range(0, MrgEvent[idx], 0, MrgOdd[idx], MergedData[idx], softGrey)


            ws2.write_row(1, 0, HeadSheet2, headtable)

            # setting column
            ws2.set_column(0, 0, 7.00, centering) # No
            ws2.set_column(1, 1, 24.29, bordering) # Area
            ws2.set_column(2, 2, 45.43, kar_format) # KAR Email
            ws2.set_column(3, 3, 22.29, centering) # Distributor Code
            ws2.set_column(4, 4, 50.86, bordering) # Distributor Name
            ws2.set_column(5, 5, 23.86, centering) # Store Code
            ws2.set_column(6, 6, 14.43, dateformat) # Registered
            ws2.set_column(7, 7, 48.14, bordering) # Store Name
            ws2.set_column(8, len(HeadSheet2) - 1, 19.57, centering)

            for lastLog in range(9, len(HeadSheet2), 2) :
                ws2.set_column(lastLog, lastLog, 19.57, datetimeformat)

            # writing data worksheet 2
            x = self.search_nested_3d(ResSheet2, AllKARUnique[count])
            if x:
                for index, txtData in enumerate(x) :
                    ws2.write_row(index+2, 0, x[index])

                    if len(x) != 0 :
                        ws2.write_number(index+2, 0, index+1)


            ws2.conditional_format(2, 8, len(ResSheet2[0][0]), len(HeadSheet2), {
                'type': 'cell',
                'criteria': '==',
                'value': '"Yes"',
                'format': YesGreen})

            ws2.conditional_format(2, 8, len(ResSheet2[0][0]), len(HeadSheet2), {
                'type': 'cell',
                'criteria': '==',
                'value': '"No"',
                'format': NoRed})

            ws2.conditional_format(2, 6, len(x)+1, 6, {
                'type': 'cell',
                'criteria': '!=',
                'value': -100000,
                'format': softYellow})


            # create auto filter worksheet 2
            lengthData = len(x)
            ws2.autofilter(1, 1, lengthData, len(HeadSheet2) - 1)


            #################################################################

            # Worksheet 3

            # ###############################################################

            # freezing first row worksheet 3
            ws3.freeze_panes(1, 0)
            ws3.hide_gridlines(2)

            # formating diffrent color first header worksheet 1
            ws3.conditional_format(0, 5, 0, len(HeadSheet3) - 1, {
                'type': 'cell',
                'criteria': '!=',
                'value': -100000,
                'format': orange})

            # write header worksheet 1
            ws3.write_row(0, 0, HeadSheet3, headtable)

            # formating column wokrsheet 1
            ws3.set_column(0, 0, 7.00, centering) # No
            ws3.set_column(1, 1, 24.29, bordering) # Area
            ws3.set_column(2, 2, 45.43, kar_format) # KAR Email
            ws3.set_column(3, 3, 20.00, centering) # Distributor Code
            ws3.set_column(4, 4, 50.86, bordering) # Distributor Name
            ws3.set_column(5, 5, 23.86, lefting) # Store Code
            ws3.set_column(6, 6, 48.14, bordering) # Store Name
            ws3.set_column(7, 7, 30.57, centering) # Username Store
            ws3.set_column(8, 8, 16.43, centering) # Password Store
            ws3.set_column(9, 9, 19.43, dateformat) # Registered
            ws3.set_column(10, 10, 22.71, centering) # PO Uploaded
            ws3.set_column(11, 11, 24.57, centering) # Status

            # writing data worksheet 3
            x = self.search_nested_3d(ResSheet3, AllKARUnique[count])
            if x:
                for index, txtData in enumerate(x) :
                    ws3.write_row(index+1, 0, x[index])

                    if len(x) != 0 :
                        ws3.write_number(index+1, 0, index+1)


            # create auto filter worksheet 1
            lengthData = len(x)
            ws3.autofilter(0, 0, lengthData, len(HeadSheet3) - 1)

            ws3.conditional_format(1, 6, len(x), 6, {
                'type': 'cell',
                'criteria': '!=',
                'value': -100000,
                'format': softYellow})


            count+= 1

            try:
                workbook.close()
            except:
                # Handle your exception here.
                reply = QMessageBox.error(self, "Error", "Error opening file for writing", QMessageBox.Ok)
                print("Couldn't create xlsx file")


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