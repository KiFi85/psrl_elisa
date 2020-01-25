from PyQt5 import QtWidgets
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtWidgets import QLabel, QGridLayout, QWidget, QVBoxLayout, QPushButton, \
    QTextEdit, QHBoxLayout, QTabWidget, QLineEdit, QSizePolicy, \
    QGroupBox, QCheckBox, QProgressBar, QFileDialog
from datetime import datetime
import os


class PageSettings(QWidget):

    def __init__(self, ctx, *args, **kwargs):
        super(PageSettings, self).__init__(*args, **kwargs)

        self.ctx = ctx
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)

        self.layout = QGridLayout(self)

        # Get list of widgets
        widget_names = ["qc_path", "curve_path", "trending_path", "f093_path", "master_path", "project_path"]
        widget_labels = ["QC Limits:", "Curve concentrations:", "Trending:", "F093:",
                         "Master save directory:", "Project List (Gantt)"]

        # Add all widgets in loop from names and labels
        self.add_widgets(widget_names, widget_labels)
        # Add blank widget for aesthetics
        self.layout.addWidget(QWidget(), 11, 0, 1, 3)
        
        # Add callbacks
        self.btn_trending_path.clicked.connect(
            lambda: self.add_path(
                self.trending_path, "Select trending file", "CSV Files (*.csv *.CSV)"))
        self.btn_curve_path.clicked.connect(
            lambda: self.add_path(self.curve_path, "Select 007sp concs. file", "CSV Files (*.csv *.CSV)"))
        self.btn_qc_path.clicked.connect(
            lambda: self.add_path(self.qc_path, "Select QC file", "CSV Files (*.csv *.CSV)"))
        self.btn_f093_path.clicked.connect(
            lambda: self.add_path(self.f093_path, "Select F093 File", "xlsm Files (*.xlsm)"))
        self.btn_master_path.clicked.connect(
            lambda: self.add_master_path(self.master_path))
        self.btn_project_path.clicked.connect(
            lambda: self.add_path(self.project_path, "Select Project List File", "CSV Files (*.csv *.CSV)"))

        # self.trending_path.setText(self.ctx.build_settings['trending_loc'])

    def add_widgets(self, names, labels):
        """ Add line edit widgets to page using object name as input and assign to class attribute
            Label widgets are added based on labels corresponding to the object name
            Button widgets are also added. All widgets are added based on the row index gained from enumerating the list
            """

        for row, txt_name in enumerate(names):

            label_name = "lbl_" + txt_name  # Label name for label widget
            label_text = labels[row]  # Get the label text from list of labels
            btn_name = "btn_" + txt_name  # Button name for button widget

            # Set class attribute as QLabel and set text and object name
            self.__dict__[label_name] = QLabel(text=label_text, objectName=label_name)
            # Set class attribute as QPushButton and set text and object name
            self.__dict__[btn_name] = QPushButton(text="Browse", objectName=btn_name)
            # Set class attribute as line edit with object name
            self.__dict__[txt_name] = QLineEdit(objectName=txt_name)

            txt_widget = self.__getattribute__(txt_name)  # Set text widget as object
            lbl_widget = self.__getattribute__(label_name)  # Set label widget as object
            btn_widget = self.__getattribute__(btn_name)  # Set button widget as object

            # Add widgets to layout based on position in list
            self.layout.addWidget(lbl_widget, row, 0)
            self.layout.addWidget(txt_widget, row, 1)
            self.layout.addWidget(btn_widget, row, 2)

            # Set size policy of the line edit widgets
            txt_widget.setSizePolicy(QSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred))

    def add_path(self, widget, title, file_text):
        """ Get a filename and add path to line edit"""

        # Decide default directory
        default_dir = get_default_dir()

        # File picker
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename = QFileDialog.getOpenFileName(self, title, default_dir, file_text, options=options)

        if filename[0]:
            widget.setText(filename[0].replace("/","\\"))

    def add_master_path(self, widget):
        """ Get the path in which to save master study files """

        default_dir = get_default_dir()

        # Save path
        path = QFileDialog.getExistingDirectory(self, 'Select a directory to save master study data files')

        # Check path was chosen
        if path:
            widget.setText(path.replace("/","\\"))


def get_default_dir():
    """ Decide what the default directory should be when selecting files """

    # If S drive exists or C:\Users or C:
    if os.path.isdir("S:"):
        default_dir = r"S:/"
    elif os.path.isdir("C:/Users"):
        default_dir = r"C:/Users"
    else:
        default_dir = r"C:/"

    return default_dir