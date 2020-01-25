from gantt_create import create_gantt
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtWidgets import QLabel, QGridLayout, QWidget, QVBoxLayout, QPushButton, \
    QSizePolicy, QDateEdit, QFileDialog, QLineEdit, QMessageBox, QTextEdit, QApplication
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from elisa_data_page import find_widget
from error_handling import UncaughtHook
import os


class PageGantt(QWidget):

    def __init__(self, ctx, *args, **kwargs):
        super(PageGantt, self).__init__(*args, **kwargs)

        self.colour_map = ctx.colour_map
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)

        self.layout = QGridLayout(self)
        layout_start = QVBoxLayout(self)
        layout_end = QVBoxLayout(self)

        # Date edits
        now = QDate.currentDate()
        start_date = now.addMonths(-1)
        end_date = now.addMonths(12)
        self.date_start = QDateEdit(date=start_date)
        self.date_start.setCalendarPopup(True)
        self.date_end = QDateEdit(date=end_date)
        self.date_end.setCalendarPopup(True)
        self.date_start.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        self.date_end.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)

        # Date pickers and labels
        label_start = QLabel(text="Select Start Date")
        label_end = QLabel(text="Select End Date")
        label_start.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        label_end.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        layout_start.addWidget(label_start)
        layout_start.addWidget(self.date_start)
        layout_end.addWidget(label_end)
        layout_end.addWidget(self.date_end)
        self.layout.addLayout(layout_start, 1, 0)
        self.layout.addLayout(layout_end, 4, 0)

        # Buttons Run and Save Figure
        self.btn_run = QPushButton(text="Plot")
        self.btn_run.clicked.connect(self.plot_gantt)
        self.btn_save = QPushButton(text="Save")
        self.btn_save.clicked.connect(self.save_plot)
        self.btn_run.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        self.btn_save.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        self.layout.addWidget(self.btn_run, 21, 3)
        self.layout.addWidget(self.btn_save, 21, 4)

        widget = QWidget()
        widget.setStyleSheet("QWidget{background-color: white}")
        self.layout.addWidget(widget, 0, 1, 20, 4)

        # Project list path
        self.error_log = self.set_error_log()

    def set_error_log(self):
        """ Get error log as widget """

        # Get settings page
        page = find_widget("elisa_data")

        # Get error log as widget
        error_log = page.findChild(QTextEdit, "error_log")
        return error_log

    def get_project_list(self):
        """ Check that the project list exists """

        # Get settings page
        page = find_widget("path_settings")

        # Get project list file path from text based on object name
        path_text = page.findChild(QLineEdit, "project_path").text()

        # Check if is a file
        if not os.path.isfile(path_text):
            self.write_errors_to_log(["project list not found"])
            self.display_error_box()
            return ""
        else:
            return path_text

    def write_errors_to_log(self, error_list):
        """ Take in a list and write out errors to log """

        # Append list of errors to error log
        for err in error_list:
            self.error_log.setTextColor(QColor(255, 0, 0))  # Red 'Error' message
            self.error_log.append("ERROR:")
            self.error_log.setTextColor(QColor(0, 0, 0))  # Black error description
            self.error_log.append(err)
            self.error_log.append("")

        # Add white space
        self.error_log.append("")
        self.error_log.append("")

    def display_error_box(self, warning="An error has occurred"):
        """ Display a generic error message """

        warning = warning + "\nSee Error Log for more details"

        err_box = QMessageBox(self)
        err_box.setIcon(QMessageBox.Warning)
        err_box.setText(warning)
        err_box.setWindowTitle("Error")
        err_box.exec_()

    def plot_gantt(self):

        project_list = self.get_project_list()
        if not project_list:
            return
        start_date = self.date_start.date().toString(Qt.ISODate)
        end_date = self.date_end.date().toString(Qt.ISODate)
        self.fig = create_gantt(project_list, self.colour_map, start_date, end_date)
        self.canvas = FigureCanvas(self.fig)
        self.layout.addWidget(self.canvas, 0, 1, 20, 4)
        self.canvas.draw()

    def save_plot(self):

        file_choices = "PNG (*.png)|*.png"

        path, ext = QFileDialog.getSaveFileName(self,
                                                'Save file', '',
                                                file_choices)

        self.fig.savefig(path)

