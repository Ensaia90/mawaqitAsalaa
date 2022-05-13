from PyQt5 import QtCore, QtGui, uic
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt, QSize, QSizeF
from PyQt5.QtWidgets import *
from PyQt5 import QtPrintSupport
import sys
import xlsxwriter
import geocoder
import json
import requests
from AboutDialog import AboutDialog

class MawaqitAlsalaa(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(QMainWindow, self).__init__(*args, **kwargs)
        #####################
        self.gridLayout = QGridLayout()
        self.ApiConfigurationHorizontalLayout = QHBoxLayout()
        #####################
        self.setFixedSize(850, 650)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setLocale(QtCore.QLocale(
            QtCore.QLocale.Arabic, QtCore.QLocale.Libya))
        self.setWindowIconText('مواقيت الصللاة')
        self.setWindowTitle('مواقيت الصللاة')
        self.setWindowIcon(QIcon(':mawaqit-alsalaa.ico'))
        self.show()
        #####################
        self.mawaqitAlsalaaOffset = ['-10', '-9', '-8', '-7', '-6',
                                     '-5', '-4', '-3', '-2', '-1', '0', '1', '2', '3', '4', '5',
                                     '6', '7', '8', '9', '10']
        #####################
        #### PRINTER CONFIGURATION
        #####################
        self.dpi = 72
        self.documentWidth = 8.5 * self.dpi
        self.documentHeight = 11 * self.dpi
        #####################
        #### YEAR NUMBER
        #####################
        self.yearNumberHorizontalLayout = QHBoxLayout()
        self.yearNumberGroupBox = QGroupBox()
        self.yearNumberGroupBox.setTitle('رقم السنة')
        # self.yearNumberGroupBox.setFixedWidth(475)
        self.yearNumberComboBox = QComboBox()
        self.yearNumberComboBox.addItems([str(yearNumber) for yearNumber in range(2010, 2051, 1)])
        self.yearNumberHorizontalLayout.addWidget(self.yearNumberComboBox)
        self.yearNumberGroupBox.setLayout(self.yearNumberHorizontalLayout)
        #####################
        #### CALCULATION  METHODS
        #####################
        self.calculationMethodsHorizontalLayout = QHBoxLayout()
        self.calculationMethodsGroupBox = QGroupBox()
        self.calculationMethodsGroupBox.setTitle('طريقة الحساب')
        # self.monthNumberGroupBox.setFixedWidth(475)
        self.calculationMethodsComboBox = QComboBox()
        self.calculationMethodsComboBox.addItems([
            'رابطة العالم الإسلامي',
            'جامعة أم القرى بمكة المكرمة',
            'الهيئة المصرية العامة للمساحة'])
        self.calculationMethodsHorizontalLayout.addWidget(self.calculationMethodsComboBox)
        self.calculationMethodsGroupBox.setLayout(self.calculationMethodsHorizontalLayout)
        #####################
        #### MONTH NUMBER
        #####################
        self.monthNumberHorizontalLayout = QHBoxLayout()
        self.monthNumberGroupBox = QGroupBox()
        self.monthNumberGroupBox.setTitle('رقم الشهر')
        # self.monthNumberGroupBox.setFixedWidth(475)
        self.monthNumberComboBox = QComboBox()
        self.monthNumberComboBox.addItems(
            ['1', '2', '3', '4', '5',
             '6', '7', '8', '9', '10', '11', '12']
        )
        self.monthNumberHorizontalLayout.addWidget(self.monthNumberComboBox)
        self.monthNumberGroupBox.setLayout(self.monthNumberHorizontalLayout)
        #####################
        #### ANNUAL CALENDAR
        #####################
        self.annualCalendarHorizontalLayout = QHBoxLayout()
        self.annualCalendarGroupBox = QGroupBox()
        self.annualCalendarGroupBox.setTitle('تقويم سنوي')
        # self.annualCalendarGroupBox.setFixedWidth(475)
        self.annualCalendarComboBox = QComboBox()
        self.annualCalendarComboBox.addItems(
            ['ﻻ', 'نعم']
        )
        self.annualCalendarHorizontalLayout.addWidget(self.annualCalendarComboBox)
        self.annualCalendarGroupBox.setLayout(self.annualCalendarHorizontalLayout)
        #####################
        #### DISPLAY MAWAQIT BUTTON
        #####################
        self.displayMawaqitAlslaaPushButton = QPushButton('عرض مواقيت الصلاة', self)
        self.displayMawaqitAlslaaPushButton.setIcon(QIcon(':mawaqit-alsalaa.png'))
        self.displayMawaqitAlslaaPushButton.setIconSize(QSize(32, 32))
        ######################
        #### MAWAQIT
        #####################
        self.mawaqitAlsalaaHorizontalLayout = QHBoxLayout()
        self.mawaqitAlsalaaGroupBox = QGroupBox()
        self.mawaqitAlsalaaGroupBox.setTitle('نسبة التصحيح للمواقيت')
        # self.mawaqitAlsalaaGroupBox.setFixedWidth(185)
        #### fajr
        self.fajrVerticalLayout = QVBoxLayout()
        self.fajrLabel = QLabel('صلاة الفجر')
        self.fajrComboBox = QComboBox()
        self.fajrComboBox.addItems(self.mawaqitAlsalaaOffset)
        self.fajrComboBox.setCurrentIndex(10)
        self.fajrComboBox.setFixedWidth(125)
        self.fajrVerticalLayout.addWidget(self.fajrLabel)
        self.fajrVerticalLayout.addWidget(self.fajrComboBox)
        #### sunrise
        self.sunriseVerticalLayout = QVBoxLayout()
        self.sunriseLabel = QLabel('الشروق')
        self.sunriseComboBox = QComboBox()
        self.sunriseComboBox.addItems(self.mawaqitAlsalaaOffset)
        self.sunriseComboBox.setCurrentIndex(10)
        self.sunriseComboBox.setFixedWidth(125)
        self.sunriseVerticalLayout.addWidget(self.sunriseLabel)
        self.sunriseVerticalLayout.addWidget(self.sunriseComboBox)
        #### dhuhr
        self.dhuhrVerticalLayout = QVBoxLayout()
        self.dhuhrLabel = QLabel('صلاة الظهر')
        self.dhuhrComboBox = QComboBox()
        self.dhuhrComboBox.addItems(self.mawaqitAlsalaaOffset)
        self.dhuhrComboBox.setCurrentIndex(10)
        self.dhuhrComboBox.setFixedWidth(125)
        self.dhuhrVerticalLayout.addWidget(self.dhuhrLabel)
        self.dhuhrVerticalLayout.addWidget(self.dhuhrComboBox)
        #### asr
        self.asrVerticalLayout = QVBoxLayout()
        self.asrLabel = QLabel('صلاة العصر')
        self.asrComboBox = QComboBox()
        self.asrComboBox.addItems(self.mawaqitAlsalaaOffset)
        self.asrComboBox.setCurrentIndex(10)
        self.asrComboBox.setFixedWidth(125)
        self.asrVerticalLayout.addWidget(self.asrLabel)
        self.asrVerticalLayout.addWidget(self.asrComboBox)
        #### maghrib
        self.maghribVerticalLayout = QVBoxLayout()
        self.maghribLabel = QLabel('صلاة المغرب')
        self.maghribComboBox = QComboBox()
        self.maghribComboBox.addItems(self.mawaqitAlsalaaOffset)
        self.maghribComboBox.setCurrentIndex(10)
        self.maghribComboBox.setFixedWidth(125)
        self.maghribVerticalLayout.addWidget(self.maghribLabel)
        self.maghribVerticalLayout.addWidget(self.maghribComboBox)
        #### isha
        self.ishaVerticalLayout = QVBoxLayout()
        self.ishaLabel = QLabel('صلاة العشاء')
        self.ishaComboBox = QComboBox()
        self.ishaComboBox.addItems(self.mawaqitAlsalaaOffset)
        self.ishaComboBox.setCurrentIndex(10)
        self.ishaComboBox.setFixedWidth(125)
        self.ishaVerticalLayout.addWidget(self.ishaLabel)
        self.ishaVerticalLayout.addWidget(self.ishaComboBox)
        ######
        self.mawaqitAlsalaaHorizontalLayout.addLayout(self.fajrVerticalLayout)
        self.mawaqitAlsalaaHorizontalLayout.addLayout(self.sunriseVerticalLayout)
        self.mawaqitAlsalaaHorizontalLayout.addLayout(self.dhuhrVerticalLayout)
        self.mawaqitAlsalaaHorizontalLayout.addLayout(self.asrVerticalLayout)
        self.mawaqitAlsalaaHorizontalLayout.addLayout(self.maghribVerticalLayout)
        self.mawaqitAlsalaaHorizontalLayout.addLayout(self.ishaVerticalLayout)
        self.mawaqitAlsalaaGroupBox.setLayout(
            self.mawaqitAlsalaaHorizontalLayout)
        #####################
        #### CITY NAME LONGITUDE LATITUDE
        #####################
        self.coordinationsGroupBox = QGroupBox()
        self.coordinationsGroupBox.setTitle('اﻻحداثيات')
        self.coordinationsHorizontalLayout = QHBoxLayout()
        ## location detection button
        self.locationDetectionPushButton = QPushButton('تحديد الموقع', self)
        self.locationDetectionPushButton.setIcon(QIcon(':geolocation.png'))
        self.locationDetectionPushButton.setIconSize(QSize(32, 32))
        ## city name
        self.cityNameLabel = QLabel('اسم المدينة:')
        self.cityNameLineEdit = QLineEdit()
        self.cityNameHorizontalLayout = QHBoxLayout()
        self.cityNameHorizontalLayout.addWidget(self.cityNameLabel)
        self.cityNameHorizontalLayout.addWidget(self.cityNameLineEdit)
        ## longitude
        self.longitudeLabel = QLabel('خط الطول:')
        self.longitudeLineEdit = QLineEdit()
        self.longitudeHorizontalLayout = QHBoxLayout()
        self.longitudeHorizontalLayout.addWidget(self.longitudeLabel)
        self.longitudeHorizontalLayout.addWidget(self.longitudeLineEdit)
        ## latitude
        self.latitudeLabel = QLabel('خط العرض:')
        self.latitudeLineEdit = QLineEdit()
        self.latitudeHorizontalLayout = QHBoxLayout()
        self.latitudeHorizontalLayout.addWidget(self.latitudeLabel)
        self.latitudeHorizontalLayout.addWidget(self.latitudeLineEdit)
        ####
        self.coordinationsHorizontalLayout.addWidget(
            self.locationDetectionPushButton)
        self.coordinationsHorizontalLayout.addLayout(
            self.cityNameHorizontalLayout)
        self.coordinationsHorizontalLayout.addLayout(
            self.longitudeHorizontalLayout)
        self.coordinationsHorizontalLayout.addLayout(
            self.latitudeHorizontalLayout)
        ####
        self.coordinationsGroupBox.setLayout(self.coordinationsHorizontalLayout)
        #####################
        #### CONTROL BUTTONS
        #####################
        self.controlButtonsHorizontalLayout = QHBoxLayout()
        self.printPreviewPushButton = QPushButton('معاينة طباعة')
        self.printPreviewPushButton.setIcon(QIcon(':preview.png'))
        self.printPreviewPushButton.setIconSize(QSize(32, 32))
        self.printTablePushButton = QPushButton('طباعة')
        self.printTablePushButton.setIcon(QIcon(':print.png'))
        self.printTablePushButton.setIconSize(QSize(32, 32))
        self.exportAsPDFPushButton = QPushButton('pdf تصدير الى ملف')
        self.exportAsPDFPushButton.setIcon(QIcon(':pdf.png'))
        self.exportAsPDFPushButton.setIconSize(QSize(32, 32))
        self.exportAsExcelPushButton = QPushButton('excel تصدير الى ملف')
        self.exportAsExcelPushButton.setIcon(QIcon(':excel.png'))
        self.exportAsExcelPushButton.setIconSize(QSize(32, 32))
        ########################
        self.controlButtonsHorizontalLayout.addWidget(
            self.printPreviewPushButton)
        self.controlButtonsHorizontalLayout.addWidget(
            self.printTablePushButton)
        self.controlButtonsHorizontalLayout.addWidget(
            self.exportAsPDFPushButton)
        self.controlButtonsHorizontalLayout.addWidget(
            self.exportAsExcelPushButton)
        #####################
        #### MAWAQIT ALSALAA TABLE WIDGET
        #####################
        self.model = QStandardItemModel()
        self.mawaqitAlsalaaTableVerticalLayout = QVBoxLayout()
        self.mawaqitAlsalaaTableView = QTableView()
        self.mawaqitAlsalaaTableVerticalLayout.addWidget(
            self.mawaqitAlsalaaTableView)
        #####################
        #### EVENTS
        #####################
        self.printTablePushButton.clicked.connect(self.handlePrint)
        self.printPreviewPushButton.clicked.connect(self.handlePrintPreview)
        self.exportAsPDFPushButton.clicked.connect(self.exportAsPDF)
        self.exportAsExcelPushButton.clicked.connect(self.exportAsExcel)
        ############
        self.locationDetectionPushButton.clicked.connect(self.locationDetection)
        self.displayMawaqitAlslaaPushButton.clicked.connect(self.requestMawaqitAlsalaa)
        #####################
        #### Grid LAYOUT
        #####################
        self.ApiConfigurationHorizontalLayout.addWidget(
            self.calculationMethodsGroupBox)
        self.ApiConfigurationHorizontalLayout.addWidget(
            self.yearNumberGroupBox)
        self.ApiConfigurationHorizontalLayout.addWidget(
            self.monthNumberGroupBox)
        ################
        self.gridLayout.addWidget(self.coordinationsGroupBox, 0, 0)
        self.gridLayout.addLayout(self.ApiConfigurationHorizontalLayout, 1, 0)
        self.gridLayout.addWidget(self.mawaqitAlsalaaGroupBox, 2, 0)
        self.gridLayout.addWidget(self.displayMawaqitAlslaaPushButton, 3, 0, QtCore.Qt.AlignLeft)
        self.gridLayout.addWidget(self.mawaqitAlsalaaTableView, 4, 0)
        self.gridLayout.addLayout(
            self.controlButtonsHorizontalLayout, 5, 0)
        #####################
        #### CENTRAL WIDGET
        #####################
        self.widgetLayout = QWidget(self)
        self.widgetLayout.setLayout(self.gridLayout)
        self.setCentralWidget(self.widgetLayout)
        #####################
        #### METHODS
        #####################
        # self.populateMawaqitAlsalaaTable()
        #####################
        self.createMenu()
    def createMenu(self):
        # Create logout action
        logoutAction = QAction('خروج', self)
        logoutAction.setIcon(QIcon(':logout.png'))
        logoutAction.setShortcut('Ctrl+Q')
        logoutAction.triggered.connect(self.logout)
        # Create about action
        aboutAction = QAction('حول', self)
        aboutAction.setIcon(QIcon(':information.png'))
        aboutAction.setShortcut('Ctrl+I')
        aboutAction.triggered.connect(self.showAboutDialog)
        # Create menubar
        menuBar = self.menuBar()
        menuBar.setNativeMenuBar(False)
        # Create file menu and add actions
        fileMenu = menuBar.addMenu('ملف')
        fileMenu.addAction(logoutAction)
        #
        aboutMenu = menuBar.addMenu('حول')
        aboutMenu.addAction(aboutAction)
    def replaceString(self, string):
        return string.replace('(EET)', '')

    def requestMawaqitAlsalaa(self):
        if self.longitudeLineEdit.text() == '':
            messageBox = QMessageBox()
            messageBox.setIconPixmap(QPixmap(':warning.png'))
            messageBox.setWindowIcon(QIcon(':mawaqit-alsalaa.ico'))
            messageBox.setWindowTitle('تنبيه')
            messageBox.setText('يجب تحديد الموقع اولا للحصول على خط الطول والعرض')
            messageBox.addButton('اغلاق', QMessageBox.YesRole)
            messageBox.exec()
            return
        payload = {
            'latitude': '32.8874',
            'longitude': '13.1873',
            'year': '2017',
            'month': '2',
            'method': '5',
            'annual': 'true'
        }

        method = 0
        if self.calculationMethodsComboBox.currentText() == 'رابطة العالم الإسلامي':
            method = 3
        elif self.calculationMethodsComboBox.currentText() == 'جامعة أم القرى بمكة المكرمة':
            method = 4
        elif self.calculationMethodsComboBox.currentText() == 'الهيئة المصرية العامة للمساحة':
            method = 5

        year = self.yearNumberComboBox.currentText()
        month = self.monthNumberComboBox.currentText()

        latitude = self.latitudeLineEdit.text()
        longitude = self.longitudeLineEdit.text()

        fajrOffset = self.fajrComboBox.currentText()
        sunriseOffset = self.sunriseComboBox.currentText()
        dhuhrOffset = self.dhuhrComboBox.currentText()
        asrOffset = self.asrComboBox.currentText()
        maghribOffset = self.maghribComboBox.currentText()
        ishaOffset = self.ishaComboBox.currentText()

        school = '1'
        annual = 'annual=true'
        calculationMethod = f'method={method}'
        calendar = f'year={year}&month={month}'
        coordinates = f'latitude={latitude}&longitude={longitude}'
        offset = f'tune={0},{fajrOffset},{sunriseOffset},{dhuhrOffset},{asrOffset},{maghribOffset},{0},{ishaOffset},{0}'

        url = f'http://api.aladhan.com/v1/calendar?{coordinates}&{calculationMethod}&{calendar}&{offset}&{annual}'

        openUrl = requests.get(url)
        print(openUrl.url)
        # print(openUrl.text)
        openUrl.encoding = 'utf-8'
        jsonData = openUrl.json()

        monthNumber = self.monthNumberComboBox.currentText()
        mawaqitAlsalaaList = []
        for times in jsonData['data'][str(monthNumber)]:
            mawaqitAlsalaa = []
            dayNumber = str(times['date']['gregorian']['day'])
            mawaqitAlsalaa.append(dayNumber)
            fajr = self.replaceString(times['timings']['Fajr'])
            mawaqitAlsalaa.append(fajr)
            sunrise = self.replaceString(times['timings']['Sunrise'])
            mawaqitAlsalaa.append(sunrise)
            dhuhr = self.replaceString(times['timings']['Dhuhr'])
            mawaqitAlsalaa.append(dhuhr)
            asr = self.replaceString(times['timings']['Asr'])
            mawaqitAlsalaa.append(asr)
            maghrib = self.replaceString(times['timings']['Maghrib'])
            mawaqitAlsalaa.append(maghrib)
            isha = self.replaceString(times['timings']['Isha'])
            mawaqitAlsalaa.append(isha)
            mawaqitAlsalaaList.append(mawaqitAlsalaa)
        self.populateMawaqitAlsalaaTable(mawaqitAlsalaaList)

    def populateMawaqitAlsalaaTable(self, mawaqitAlsalaaList):
        self.model.clear()
        for i, row in enumerate(mawaqitAlsalaaList):
            items = [QStandardItem(item) for item in row]
            self.model.insertRow(i, items)
        self.model.setHorizontalHeaderLabels(
            ['اليوم', 'الفجر', 'الشروق', 'الظهر', 'العصر', 'المغرب', 'العشاء'])
        self.mawaqitAlsalaaTableView.horizontalHeader(
        ).setSectionResizeMode(QHeaderView.Stretch)
        self.mawaqitAlsalaaTableView.SelectionMode(3)
        self.mawaqitAlsalaaTableView.setModel(self.model)
        self.mawaqitAlsalaaTableView.horizontalHeader().setSectionResizeMode(QHeaderView.
                                                                             Stretch)
        self.mawaqitAlsalaaTableView.verticalHeader().setSectionResizeMode(QHeaderView.
                                                                           Stretch)

    def locationDetection(self):
        geocoderInstance = geocoder.ip('me')
        latitude = geocoderInstance.latlng[0]
        longitude = geocoderInstance.latlng[1]
        cityName = geocoderInstance.city
        self.latitudeLineEdit.setText(str(latitude))
        self.longitudeLineEdit.setText(str(longitude))
        self.cityNameLineEdit.setText(str(cityName))

    def handlePrintPreview(self):
        dialog = QtPrintSupport.QPrintPreviewDialog()
        dialog.paintRequested.connect(self.handlePaintRequest)
        dialog.exec_()

    def handlePrint(self):
        dialog = QtPrintSupport.QPrintDialog()
        if dialog.exec_() == QDialog.Accepted:
            self.handlePaintRequest(dialog.printer())

    def handlePaintRequest(self, printer):
        # Configure defaults:
        # Landscape , Portrait
        printer.setOrientation(QtPrintSupport.QPrinter.Portrait)
        printer.setPageSize(QPageSize(QPageSize.A4))
        printer.setPageMargins(
            15, 15, 15, 15, QtPrintSupport.QPrinter.Millimeter)
        ########################
        tableFormat = QTextTableFormat()
        tableFormat.setHeaderRowCount(1)
        tableFormat.setAlignment(Qt.AlignHCenter)
        tableFormat.setAlignment(Qt.AlignVCenter)
        tableFormat.setCellPadding(1.0)
        tableFormat.setCellSpacing(1.0)
        tableFormat.setWidth(QTextLength(QTextLength.PercentageLength, 100))
        ########################
        textOption = QTextOption()
        textOption.setTextDirection(Qt.RightToLeft)
        textOption.setAlignment(Qt.AlignRight | Qt.AlignHCenter)
        ########################
        document = QTextDocument()
        document.setPageSize(QSizeF(
            self.documentWidth, self.documentHeight))
        document.setDefaultFont(QtGui.QFont('Console,Verdana,Arial,Helvetica,sans-serif', 12, QtGui.QFont.Normal))
        document.setDocumentMargin(1.0)
        document.setDefaultTextOption(textOption)
        ########################
        cursor = QTextCursor(document)
        ########################
        titleFormat = QTextCharFormat()
        titleFormat.setFont(QtGui.QFont('Console,Verdana,Arial,Helvetica,sans-serif',12, QtGui.QFont.Normal))
        theText = ' مواقيت الصلاة لشهر '+self.monthNumberComboBox.currentText()
        cursor.insertText(theText,titleFormat)
        cursor.insertText('')
        cursor.insertBlock()
        ########################
        model = self.mawaqitAlsalaaTableView.model()
        rows = model.rowCount()
        columns = model.columnCount()
        table = cursor.insertTable(rows + 1, columns, tableFormat)
        format = table.format()
        format.setHeaderRowCount(1)
        table.setFormat(format)
        format = cursor.blockCharFormat()
        format.setFontWeight(QFont.Bold)
        #############
        for column in range(columns - 1, -1, -1):
            cursor.insertText(str(model.headerData(column, Qt.Horizontal)))
            cursor.movePosition(QtGui.QTextCursor.NextCell)
        ###########
        for row in range(rows):
            for column in range(columns - 1, -1, -1):
                cursor.insertText(str(model.data(model.index(row, column))))
                cursor.movePosition(QtGui.QTextCursor.NextCell)
        ###########
        document.print_(printer)

    def exportAsPDF(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, 'حفظ كملف PDF', QtCore.QDir.homePath(), 'PDF Files (*.pdf)')
        if filename:
            printer = QtPrintSupport.QPrinter(
                QtPrintSupport.QPrinter.HighResolution)
            printer.setPageSize(QtPrintSupport.QPrinter.A4)
            printer.setOutputFormat(QtPrintSupport.QPrinter.PdfFormat)
            printer.setOutputFileName(filename)
            printer.setPageMargins(
                12, 16, 12, 20, QtPrintSupport.QPrinter.Millimeter)
            printer.pageRect(QtPrintSupport.QPrinter.Point)
            self.handlePaintRequest(printer)

    def exportAsExcel(self):
        model = self.mawaqitAlsalaaTableView.model()
        rows = model.rowCount()
        columns = model.columnCount()
        monthNumber = self.monthNumberComboBox.currentText()
        workbook = xlsxwriter.Workbook(
            f'مواقيت الصلاة لشهر {monthNumber}.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.right_to_left()
        setHeaderText = '&C&"Calibri,Bold"&20 مواقيت الصلاة'
        setFooterText = '&LPage &P of &N' + '&CFilename: &F' + '&RSheetname: &A'
        setFooterText = '&Lالصفحة &P من &N'
        worksheet.set_header(setHeaderText)
        worksheet.set_footer(setFooterText)
        worksheet.set_paper(9)  # A4
        worksheet.center_horizontally()
        worksheet.fit_to_pages(1, 0)
        cell_format = workbook.add_format({
            'bold': True,
            'font_name': 'Times New Roman',
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 16,
            'font_color': 'black'
        })
        worksheet.set_column(0, columns, 18, cell_format)
        data = []
        for row in range(rows):
            data.append([])
            data[row].append(str(model.headerData(row, Qt.Vertical)))
            for column in range(columns):
                index = model.index(row, column)
                data[row].append(str(model.data(index)))

        worksheet.add_table(0, 0, rows, columns - 1, {'data': data, 'autofilter': False,
                                                      'columns': [
                                                          {'header': 'اليوم'},
                                                          {'header': 'الفجر'},
                                                          {'header': 'الشروق'},
                                                          {'header': 'الظهر'},
                                                          {'header': 'العصر'},
                                                          {'header': 'المغرب'},
                                                          {'header': 'العشاء'}
                                                      ],
                                                      'style': 'Table Style Light 15'})

        workbook.close()

    def logout(self, event):
        messageBox = QMessageBox()
        messageBox.setIconPixmap(QPixmap(':question.png'))
        messageBox.setWindowIcon(QIcon(':mawaqit-alsalaa.ico'))
        messageBox.setWindowTitle('تنبيه')
        messageBox.setText('هل أنت متأكد ؟')
        messageBox.addButton('لا', QMessageBox.NoRole)
        messageBox.addButton('نعم', QMessageBox.YesRole)
        reply = messageBox.exec()
        if reply == 1:
            self.close()

    def showAboutDialog(self):
        showAboutDialogInstance = AboutDialog()
        showAboutDialogInstance.show()
        showAboutDialogInstance.exec_()
styleSheet = '''
QMainWindow {
background-color: #fff;
font-size:16px;
color:#000;
border:1px solid #ccc;
}
QComboBox{
background-color: #eee;

}

QPushButton {
    border-style: solid;
    border-top-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-right-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(217, 217, 217), stop: 1 rgb(227, 227, 227));
    border-left-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(227, 227, 227), stop: 1 rgb(217, 217, 217));
    border-bottom-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-width: 1px;
    border-radius: 5px;
    color: rgb(0, 0, 0);
    padding: 8px;
    background-color: rgb(255, 255, 255);
    font-size: 16px;
    font-weight: 400;
    line-height: 1.3333333;
}

QPushButton::default {
    border-style: solid;
    border-top-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-right-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(217, 217, 217), stop: 1 rgb(227, 227, 227));
    border-left-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(227, 227, 227), stop: 1 rgb(217, 217, 217));
    border-bottom-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-width: 1px;
    border-radius: 5px;
    color: rgb(0, 0, 0);
    background-color: rgb(255, 255, 255);
    padding: 8px;
    font-size: 16px;
    font-weight: 400;
    line-height: 1.3333333;
}

QPushButton:hover {
    border-style: solid;
    border-top-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(195, 195, 195), stop: 1 rgb(222, 222, 222));
    border-right-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(197, 197, 197), stop: 1 rgb(227, 227, 227));
    border-left-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(227, 227, 227), stop: 1 rgb(197, 197, 197));
    border-bottom-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(195, 195, 195), stop: 1 rgb(222, 222, 222));
    border-width: 1px;
    border-radius: 5px;
    color: rgb(0, 0, 0);
    background-color: rgb(255, 255, 255);
    padding: 8px;
    font-size: 16px;
    font-weight: 400;
    line-height: 1.3333333;
}

QPushButton:pressed {
    border-style: solid;
    border-top-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-right-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(217, 217, 217), stop: 1 rgb(227, 227, 227));
    border-left-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(227, 227, 227), stop: 1 rgb(217, 217, 217));
    border-bottom-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-width: 1px;
    border-radius: 5px;
    color: rgb(0, 0, 0);
    background-color: rgb(142, 142, 142);
    padding: 8px;
    font-size: 16px;
    font-weight: 400;
    line-height: 1.3333333;
}

QPushButton:disabled {
    border-style: solid;
    border-top-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-right-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(217, 217, 217), stop: 1 rgb(227, 227, 227));
    border-left-color: qlineargradient(spread: pad, x1: 0, y1: 0.5, x2: 1, y2: 0.5, stop: 0 rgb(227, 227, 227), stop: 1 rgb(217, 217, 217));
    border-bottom-color: qlineargradient(spread: pad, x1: 0.5, y1: 1, x2: 0.5, y2: 0, stop: 0 rgb(215, 215, 215), stop: 1 rgb(222, 222, 222));
    border-width: 1px;
    border-radius: 5px;
    color: #808086;
    background-color: rgb(142, 142, 142);
    padding: 8px;
    font-size: 16px;
    font-weight: 400;
    line-height: 1.3333333;
}
'''


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(styleSheet)
    window = MawaqitAlsalaa()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
