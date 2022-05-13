from PyQt5 import QtCore, QtGui, uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QTimer, QTime, Qt
from PyQt5.Qt import *
import sys
import resources
class AboutDialog(QDialog):
    def __init__(self, *args, **wkargs):
        super(QDialog, self).__init__(*args, **wkargs)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setLocale(QLocale(QLocale.Arabic, QLocale.Libya))
        self.setFixedSize(350,250)
        self.setStyleSheet('''
        QDialog{
        background-color: #fff;
        
        border:1px solid #ccc;
        }
        QLabel{
        font-size:18px;
        color:#000;
        }
        ''')
        self.setWindowTitle('حول البرنامج')
        self.setWindowIcon(QIcon(':mawaqit-alsalaa.ico'))
        self.gridLayout = QGridLayout()
        #################################################
        labelHtml = '''<p>مواقيت الصلاة هو عبارة عن واجهة مستخدم تتيح لك معرفة اوقات الصللاة الخاصة بمنطقة معينة
         وذلك عن طريق خط الطول والعرض مع امكانيات التعديل على تلك اﻻوقات من حيث التقديم و التأخير.</p>'''
        labelHtml += '''<p>كذلك يتيح امكانية طباعة هذه اﻻوقات او تصديرها الى ملفات مثل (excel,pdf).</p>'''
        labelHtml += '''<p>برمجة - <b> محمد عثمان نصية.</b></p>'''
        self.aboutLabel = QLabel(labelHtml,self)
        self.aboutLabel.setWordWrap(True)
        #################################################
        self.gridLayout.addWidget(self.aboutLabel,0,0)
        #################################################
        self.setLayout(self.gridLayout)

