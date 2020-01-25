from PyQt5 import QtWidgets
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTextBrowser
from PyQt5.QtCore import QUrl
from PyQt5.QtWebEngineWidgets import QWebEngineView
import qrc_resources

class PageHelp(QWidget):

    def __init__(self, ctx, *args, **kwargs):
        super(PageHelp, self).__init__(*args, **kwargs)

        # Background
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)

        layout = QVBoxLayout(self)
        self.text_browser = QTextBrowser()
        layout.addWidget(self.text_browser)
        self.setLayout(layout)
        # self.text_browser.setText("SETT")
        # QUrl(r"C:\Users\kier_\Documents_Unsynced\Python\fbs_app_3.6\src\main\resources\base\html\help.html")
        self.text_browser.setSource(QUrl("qrc:/help.html"))
        # self.browser = QWebEngineView()
        # self.browser.setHtml(r"app/help/index.html")
        # self.text_browser.setSource(QUrl("qrc:/help.html"))
        # self.html_widget = QWebEngineView()
        # self.browser.load(html_file)
        # layout.addWidget(self.browser)
        #self.setCentralWidget(self.browser)