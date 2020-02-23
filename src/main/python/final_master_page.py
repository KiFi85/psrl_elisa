from PyQt5 import QtWidgets
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtWidgets import QLabel, QGridLayout, QWidget, QComboBox, QApplication, QPushButton, \
    QTextEdit, QHBoxLayout, QTabWidget, QLineEdit, QSizePolicy, \
    QGroupBox, QCheckBox, QProgressBar, QFileDialog
from pathlib import Path
import os


class PageFinalMaster(QWidget):

    def __init__(self, ctx, master_path, *args, **kwargs):
        super(PageFinalMaster, self).__init__(*args, **kwargs)

        self.ctx = ctx
        self.MASTER_PATH = master_path
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)

        # Working file
        self.working_file = ""

        # Layouts
        self.layout = QGridLayout(self)
        self.layout.setSpacing(50)

        # Get list of master files
        self.file_list = self.get_master_files()

        # Master files
        file_label = QLabel(text="Select master file:")
        self.file_combo = QComboBox(objectName="master_combo")
        self.file_combo.setSizePolicy(QSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred))
        self.file_combo.addItems(self.file_list)

        # Working file
        working_label = QLabel(text="Select working file:")
        self.working_btn = QPushButton(text="Browse")
        self.working_btn.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        self.working_btn.clicked.connect(self.btn_working_clicked)

        # Run
        self.run_btn = QPushButton(text="Run")
        self.run_btn.clicked.connect(self.run_clicked)

        # Add widgets to layout
        self.layout.addWidget(file_label, 0, 0)
        self.layout.addWidget(self.file_combo, 0, 1, 1, 2)
        self.layout.addWidget(working_label, 1, 0)
        self.layout.addWidget(self.working_btn,1,1)
        self.layout.addWidget(self.run_btn, 22, 4)


        widget = QWidget()
        # widget.setStyleSheet("QWidget{background-color: white}")
        self.layout.addWidget(widget, 2, 0, 20, 5)

    def btn_working_clicked(self):
        """ Search for and select F007 containing assay details """

        # Go to parent directory of master directory (main ELISA)
        default_dir = os.path.join(Path(self.MASTER_PATH).parent)

        # what to display
        title = 'Choose a working master file'
        file_text = "Excel Files (*.xls *.xlsx *.xlsm)"

        # File picker
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename = QFileDialog.getOpenFileName(self, title, default_dir, file_text, options=options)
        #
        if filename[0]:
            self.working_file = filename[0]
            print(self.working_file)
            
    def run_clicked(self):
        pass

    def get_master_files(self):
        """ Retrieve all CSV files in master study file to populate combobox """

        all_files = [f for f in os.listdir(self.MASTER_PATH) if os.path.isfile(os.path.join(self.MASTER_PATH, f))]
        csv_files = [f for f in all_files if Path(f).suffix.upper() == '.CSV']
        return csv_files


