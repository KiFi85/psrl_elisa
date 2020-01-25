from PyQt5 import QtWidgets
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtWidgets import QLabel, QGridLayout, QWidget, QVBoxLayout, QPushButton, \
    QTextEdit, QHBoxLayout, QTabWidget, QLineEdit, QSizePolicy, \
    QGroupBox, QCheckBox, QProgressBar, QFileDialog
from datetime import datetime
import subprocess
import os
import sys


class PageDataCheck(QWidget):

    def __init__(self, ctx, *args, **kwargs):
        super(PageDataCheck, self).__init__(*args, **kwargs)

        self.ctx = ctx
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)

#         self.layout = QHBoxLayout()
#         self.btn = QPushButton()
#         self.btn2 = QPushButton()
#         self.btn.clicked.connect(self.sub_process)
#         self.btn2.clicked.connect(self.sub2)
#         self.layout.addWidget(self.btn)
#         self.layout.addWidget(self.btn2)
#
#         self.setLayout(self.layout)
#
#     def sub_process(self):
#
#         c = Configuration()
#         print(c.environ)
#
#     def sub2(self):
#
#         w = r"C:\Users\kier_\Documents_Unsynced\Python\fbs_app_3.6\src\main\resources\base\wkhtmltopdf.exe"
#         c = Configuration2(wkhtmltopdf=w)
#         print(c.environ)
#
#
# import os
# import subprocess
# import sys
#
# class Configuration(object):
#     def __init__(self, wkhtmltopdf='', meta_tag_prefix='pdfkit-', environ=''):
#         self.meta_tag_prefix = meta_tag_prefix
#
#         self.wkhtmltopdf = wkhtmltopdf
#
#         if not self.wkhtmltopdf:
#             if sys.platform == 'win32':
#                 self.wkhtmltopdf = subprocess.Popen(
#                     ['where', 'wkhtmltopdf'], stdout=subprocess.PIPE).communicate()[0].strip()
#             else:
#                 self.wkhtmltopdf = subprocess.Popen(
#                     ['which', 'wkhtmltopdf'], stdout=subprocess.PIPE).communicate()[0].strip()
#
#         try:
#             with open(self.wkhtmltopdf) as f:
#                 pass
#         except IOError:
#             raise IOError('No wkhtmltopdf executable found: "%s"\n'
#                           'If this file exists please check that this process can '
#                           'read it. Otherwise please install wkhtmltopdf - '
#                           'https://github.com/JazzCore/python-pdfkit/wiki/Installing-wkhtmltopdf' % self.wkhtmltopdf)
#
#         self.environ = environ
#
#         if not self.environ:
#             self.environ = os.environ
#
#         for key in self.environ.keys():
#             if not isinstance(self.environ[key], str):
#                 self.environ[key] = str(self.environ[key])
#
# class Configuration2(object):
#     def __init__(self, wkhtmltopdf='', meta_tag_prefix='pdfkit-', environ=''):
#         self.meta_tag_prefix = meta_tag_prefix
#
#         self.wkhtmltopdf = wkhtmltopdf
#         command = "start /min " + self.wkhtmltopdf
#
#         if not self.wkhtmltopdf:
#             if sys.platform == 'win32':
#                 self.wkhtmltopdf = subprocess.Popen(command, shell=True,
#                                                     stdout=subprocess.PIPE).communicate()[0].strip()
#             else:
#                 self.wkhtmltopdf = subprocess.Popen(command, shell=True,
#                                                     stdout=subprocess.PIPE).communicate()[0].strip()
#
#         try:
#             with open(self.wkhtmltopdf) as f:
#                 pass
#         except IOError:
#             raise IOError('No wkhtmltopdf executable found: "%s"\n'
#                           'If this file exists please check that this process can '
#                           'read it. Otherwise please install wkhtmltopdf - '
#                           'https://github.com/JazzCore/python-pdfkit/wiki/Installing-wkhtmltopdf' % self.wkhtmltopdf)
#
#         self.environ = environ
#
#         if not self.environ:
#             self.environ = os.environ
#
#         for key in self.environ.keys():
#             if not isinstance(self.environ[key], str):
#                 self.environ[key] = str(self.environ[key])
#
