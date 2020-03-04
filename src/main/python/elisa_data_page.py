from PyQt5 import QtWidgets
from PyQt5.QtGui import QPalette, QColor, QFont
from PyQt5.QtCore import QSettings, QObject, pyqtSignal, QRunnable, pyqtSlot, QThreadPool, Qt
from PyQt5.QtWidgets import QMainWindow, QLabel, QGridLayout, QWidget, QVBoxLayout, QPushButton, \
    QTextEdit, QHBoxLayout, QTabWidget, QLineEdit, QSizePolicy, \
    QGroupBox, QCheckBox, QProgressBar, QFileDialog, QApplication, QMessageBox, QComboBox, QTableWidget, \
    QTableWidgetItem, QHeaderView
from win32api import GetSystemMetrics
from datetime import datetime
from settings_page import get_default_dir, PageSettings
from assay import Assay
from elisa_data import ELISAData
from elisa import ELISA
# from error_handling import show_exception_box
from error_handling import RangeNotFoundError
import time
import os
import pandas as pd
from pathlib import Path
import xlwings as xw
import traceback
import win32api
import win32print
import sys

class WorkerSignals(QObject):

    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    obj_result = pyqtSignal(tuple)
    progress = pyqtSignal(int)


class Worker(QRunnable):

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        # Add the callback to our kwargs
        kwargs['progress_callback'] = self.signals.progress

    @pyqtSlot()
    def run(self):
        '''
        Initialise the runner function with passed args, kwargs.
        '''

        # Retrieve args/kwargs here; and fire processing using them
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            # traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)  # Return the result of the processing
        finally:
            self.signals.finished.emit()  # Done

class PageData(QTabWidget):
    """ Create a tabbed widget to store Data and Error tabs """

    def __init__(self, ctx, *args, **kwargs):
        super(PageData, self).__init__(*args, **kwargs)

        self.setDocumentMode(False)
        self.setMovable(False)
        self.setTabPosition(QTabWidget.South)
        self.main = DataTab(ctx)  # Application context argument to data tab to find resources
        self.error = ErrorTab()
        self.addTab(self.main, "Data")
        self.addTab(self.error, "Error Log")

        self.main.error_log = self.error.error_log

    def write_to_log(self, msg):
        self.error.error_log.append(msg)


class DataTab(QWidget):
    """ Main data page tab """

    def __init__(self, ctx, *args, **kwargs):
        super(DataTab, self).__init__(*args, **kwargs)

        self.ctx = ctx  # Application context (for wkhtmltopdf.exe etc)
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)
        layout_main = QVBoxLayout(self)
        layout_files = QGridLayout(self)
        layout_options = QHBoxLayout(self)
        layout_run = QGridLayout(self)

        # ERror log
        self.error_log = None

        # Make amendments boolean
        self.amendments = pd.DataFrame()

        # Layout spacing
        layout_main.setSpacing(20)

        """ F007 and MARS Files browsing, amendments (layout_files) """
        # F007 details
        label = QLabel("Select F007 file:")
        self.txt_f007 = QLineEdit(objectName="txt_f007")
        self.txt_f007.setSizePolicy(QSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred))
        self.txt_f007.setEnabled(False)
        self.btn_f007 = QPushButton(text="Browse", objectName="btn_f007")
        self.btn_f007.clicked.connect(self.btn_f007_clicked)
        self.f007_file = ''

        # Add widgets
        layout_files.addWidget(label, 0, 0)
        layout_files.addWidget(self.txt_f007, 0, 1)
        layout_files.addWidget(self.btn_f007, 0, 2)

        # MARS details
        label = QLabel("Select MARS file(s):")
        self.txt_mars = QLineEdit(objectName="txt_mars")
        self.txt_mars.setSizePolicy(QSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred))
        self.txt_mars.setEnabled(False)
        self.btn_mars = QPushButton(text="Browse", objectName="btn_mars")
        self.btn_mars.clicked.connect(self.btn_mars_clicked)
        self.mars_files = []

        # Amendments
        self.check_amend = QCheckBox(objectName="check_amend", text="Process data with amended plate fails")
        self.check_amend.toggled.connect(lambda: self.amend_checkbox_changed(self.check_amend))

        # Add widgets
        layout_files.addWidget(label, 1, 0)
        layout_files.addWidget(self.txt_mars, 1, 1)
        layout_files.addWidget(self.btn_mars, 1, 2)
        layout_files.addWidget(QWidget(), 2, 0)
        layout_files.addWidget(self.check_amend, 3, 0, 1, 2)
        # layout_files.addWidget(label_check_amend, 2, 1)

        # Parameter group box
        self.group_box = QGroupBox()
        self.group_box.setObjectName("grp_box")
        self.group_box.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)

        # Grid layout for group box
        group_layout = QGridLayout(self.group_box)
        self.group_box.setTitle("Select Parameters")

        """ GROUP BOX WIDGETS"""
        # Options (clinical/validation/custom) dropdown
        self.combo_options = QComboBox(objectName="combo_options")
        [self.combo_options.addItems(x for x in ["Clinical", "Validation", "Custom"])]
        self.combo_options.currentIndexChanged[str].connect(self.combo_changed)


        # OD check boxes
        self.cb_upper_od = QCheckBox(objectName="cb_upper_od")
        self.cb_upper_od.setText("Set upper OD limit for analysis")
        self.cb_upper_od.toggled.connect(
            lambda: self.od_checkbox_changed(self.cb_upper_od, self.txt_od_upper))

        self.cb_lower_od = QCheckBox(objectName="cb_lower_od")
        self.cb_lower_od.setText("Set lower OD limit for analysis")
        self.cb_lower_od.toggled.connect(
            lambda: self.od_checkbox_changed(self.cb_lower_od, self.txt_od_lower))

        # Option check boxes
        self.cb_lloq = QCheckBox(objectName="cb_lloq", text="Apply LLOQ (<0.15)")
        self.cb_print = QCheckBox(objectName="cb_print", text="Print plate data")
        self.cb_print.setEnabled(True)
        self.cb_print.toggled.connect(lambda: self.print_checkbox_changed(self.cb_print))

        # OD text boxes
        self.txt_od_upper = QLineEdit(objectName="txt_upper_od")
        self.txt_od_upper.setMaximumWidth(100)
        self.txt_od_upper.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        # self.txt_od_upper.setEnabled(False)

        self.txt_od_lower = QLineEdit(objectName="txt_lower_od")
        self.txt_od_lower.setMaximumWidth(100)
        self.txt_od_lower.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        # Group box Layout
        group_layout.addWidget(QLabel("Select pre-defined parameters"), 0, 0, 1, 1)
        group_layout.addWidget(self.combo_options, 0, 1, 1, 1)
        group_layout.addWidget(QLabel(""), 1, 0, 1, 2)
        group_layout.addWidget(self.cb_upper_od, 2, 0, 1, 1)
        group_layout.addWidget(self.cb_lower_od, 3, 0, 1, 1)
        group_layout.addWidget(self.cb_lloq, 4, 0, 1, 1)
        group_layout.addWidget(self.cb_print, 5, 0, 1, 1)
        group_layout.addWidget(self.txt_od_upper, 2, 1, 1, 1)
        group_layout.addWidget(self.txt_od_lower, 3, 1, 1, 1)

        # Add groupboxes to layout
        layout_options.addWidget(self.group_box)
        widget = QWidget()
        widget.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.MinimumExpanding)
        layout_options.addWidget(widget)

        # Run button and progress bar
        self.btn_run = QPushButton(objectName="btn_run", text="Run")
        self.btn_run.clicked.connect(self.btn_run_clicked)

        self.progress_layout = QVBoxLayout()
        self.progress_label = QLabel("")
        self.progress_bar = QProgressBar(objectName="prog_bar")
        self.progress_layout.addWidget(self.progress_label)
        self.progress_layout.addWidget(self.progress_bar)
        self.progress_layout.setSpacing(0)
        self.pdf_ctr = 0
        self.set_progress_transparent()
        self.progress_bar.setFormat(u"(%v / %m)")

        # Add to layout
        # layout_run.addWidget(self.progress_bar, 0,0)
        layout_run.addLayout(self.progress_layout, 0, 0, 2, 1)
        # layout_run.addItem(hor_spacer,0,1)
        layout_run.addWidget(self.btn_run, 1, 2)

        # All layout
        layout_main.addLayout(layout_files)
        layout_main.addLayout(layout_options)
        layout_main.addLayout(layout_run)

        # Initialise all properties
        self.xl_id = None
        self.init_parms()

    def print_checkbox_changed(self, cb):
        pass
        # print(self.amendments)

    def init_parms(self):
        """ Initialise all parameters - used at startup and when Run Button clicked again """

        # Data objects and constants
        self.QC_FILE = ''
        self.CURVE_FILE = ''
        self.XL_FILE = ''
        self.TREND_FILE = ''
        self.F093_FILE = ''
        self.MASTER_PATH = ''
        self.assay = None
        self.elisa = None
        self.elisa_data = None
        self.savedir = ''
        self.cut_high_ods = 2
        self.cut_low_ods = 0.1
        self.apply_lloq = True

        # Threadpool
        self.threadpool = QThreadPool()
        self.file_check_errors = ""
        self.object_errors = []
        self.write_files_errors = []
        self.trending_done = False
        self.summary_done = False
        self.master_done = False
        self.f093_done = False
        self.pdf_names = []

    def combo_changed(self, selection):
        """ Function to change checkboxes dependent on combobox settings"""

        if selection == "Validation":

            # Lower OD - Uncheck but enable
            self.cb_lower_od.setChecked(False)
            self.cb_lower_od.setEnabled(True)
            self.txt_od_lower.setEnabled(True)
            # LLOQ - Uncheck but enable
            self.cb_lloq.setChecked(False)
            self.cb_lloq.setEnabled(True)
            # Upper OD - Check and Disable
            self.cb_upper_od.setChecked(True)
            self.cb_upper_od.setEnabled(False)
            self.txt_od_upper.setEnabled(False)

        elif selection == "Clinical":
            for c in self.group_box.children():
                name = c.objectName()

                # If it's a checkbox or text and not print - check and disable
                if name[:2] == "cb" and name[-5:] != "print":
                    c.setEnabled(False)
                    c.setChecked(True)
                elif name[:2] == "tx":
                    c.setEnabled(False)
        else:
            for c in self.group_box.children():
                c.setEnabled(True)

    def btn_run_clicked(self):
        """ Run data processing """

        # Get parameters
        self.cut_high_ods, self.cut_low_ods, self.apply_lloq = self.get_parms()

        app = xw.App(visible=False)
        app.screen_updating = False
        app.display_alerts = False
        self.xl_id = app.pid

        # Show the progress bar and progress label
        self.progress_bar.setValue(0)
        worker = Worker(self.show_bar_prog)  # Any other args, kwargs are passed to the run function
        worker.signals.finished.connect(self.done_show_bar)
        self.threadpool.start(worker)
        # self.threadpool.waitForDone()

        # Check F007 and MARS FILES not empty
        if not self.f007_file or not self.mars_files:
            self.display_error_box()
            self.write_errors_to_log(["F007 or MARS file information missing"])
            return

        # Check that the files exist and add to constants if so
        self.file_check_worker()

        # Create data objects

        # Process plates
        # self.threadpool.waitForDone()
        # print(self.assay.study)

        # print(str(self.threadpool.waitForDone()) + str(time.time() - self.start))
        # print("i waited "+ str(time.time() - self.start))
        # # # Check formatting of required files
        # # file_formats_ok = self.check_file_formats()
        #
        # Process plate data
        # self.process_plates_worker()
        # self.process_plates()
        #
        # # Write to master files (trending, run_details, master study file)
        # self.write_to_files()

    def get_parms(self):
        """ Get the OD limits and application of LLOQ """

        # Apply upper OD limit
        if self.cb_upper_od.isChecked():
            upper_od = float(self.txt_od_upper.text())
        else:
            upper_od = None

        # Apply upper OD limit
        if self.cb_lower_od.isChecked():
            lower_od = float(self.txt_od_lower.text())
        else:
            lower_od = None

        # Apply LLOQ or not
        if self.cb_lloq.isChecked():
            apply_lloq = True
        else:
            apply_lloq = False

        return upper_od, lower_od, apply_lloq

    def percent_progress(self, n):
        """ Percentage progress """

        self.progress_bar.setFormat("%p%")
        self.progress_bar.setValue(n)

    def int_progress(self, n):
        """ Integer progress """

        self.progress_bar.setFormat(u"(%v / %m)")
        self.progress_bar.setValue(n)

    def result_error(self, err_list):
        """ Function to process errors that may be returned from any worker thread """

        if err_list:
            self.progress_label.setText("Operation cancelled")
            self.display_error_box("Returned an error")
            self.write_errors_to_log(err_list)
            self.set_button_states(True)
            self.init_parms()
            self.kill_xl()

    def show_bar_prog(self, progress_callback):

        time.sleep(0.02)
        return

    def done_show_bar(self):
        """ When finished displaying progress bar on btn_run_clicked """

        # Format progress bar
        self.set_button_states(False)  # Buttons disabled
        self.progress_label.setVisible(True)
        self.progress_label.setText("Checking required files...")
        self.set_progress_opaque()

    def done_file_check(self):
        """ When file check has completed """

        self.threadpool.waitForDone()
        self.progress_label.setText("Creating data objects...")
        self.data_object_worker()  # Create assay and elisa_data objects

    def done_create_objects(self):
        """ when data object have been created """

        self.threadpool.waitForDone()
        self.progress_bar.setValue(100)

        # If the assay object has successfully been created
        # Begin processing elisa data
        if self.assay:
            self.process_plates_worker()
        else:
            return

    def done_processing_data(self):
        """ When finished processing elisa objects """

        # Get default printer and check when printing is enabled
        printer = win32print.GetDefaultPrinter()
        do_print = self.cb_print.isChecked()

        # If printing to pdf/XPS are default or print not selected - don't print
        no_print = ['Print to PDF', 'XPS', 'Fax', 'OneNote']
        if any(x.upper() in printer.upper() for x in no_print):
            self.done_print_data()
        elif not do_print:
            self.done_print_data()
        else:
            self.progress_label.setText("Printing pdfs....")
            self.print_worker()  # Begin printing PDFs

    def done_print_data(self):
        """ When PDFs have been printed """

        # Begin writing summary data files (master study, trending, run_details
        self.progress_label.setText("Writing summary data to file...")
        self.write_files_worker()

    def done_write_files(self):
        """ When finished writing files create F093"""

        # If not a repeated assay save data to Excel template
        if self.assay.run_type != "repeats":
            self.elisa_data.f093_to_excel()
        self.progress_bar.setValue(100)
        self.progress_label.setText("Finished")
        self.set_button_states(True)  # Re-enable buttons

        # Initialise parameters and kill Excel process
        self.init_parms()
        self.kill_xl()

    def thread_error(self, exc_info):
        """ When error raised from within thread """

        self.kill_xl()  # Kill Excel process
        self.display_error_box()  # Display default error message

        self.error_log.setTextColor(QColor(255, 0, 0))  # Red 'Error' message
        self.error_log.append("A thread error occurred:\n")

        # Write Python error to log (will occur if any exception is not handled
        self.error_log.setTextColor(QColor(0, 0, 0))
        self.error_log.append('{0}: {1}'.format(exc_info[0].__name__, exc_info[1]))
        self.error_log.append("\n" + exc_info[2] + "\n\n")
        self.progress_label.setText("Operation cancelled")
        self.set_button_states(True)

    def file_check_countdown(self, progress_callback):
        """ A timer to run and update progress bar on file check progress
            unless an error occurs """

        # 0 to 100%
        for n in range(0, 101):

            # If any errors - add to list and return to run_error function
            if self.file_check_errors:
                err_list = self.file_check_errors
                return err_list

            # Otherwise - set to 100%
            elif self.QC_FILE and self.CURVE_FILE and self.TREND_FILE and self.F093_FILE and self.MASTER_PATH:
                progress_callback.emit(100)
                return
            else:
                time.sleep(0.02)
                progress_callback.emit(n)  # emit progress of file check

    def get_object_countdown(self, progress_callback):
        """ A timer to run and update progress bar on creation of data objects
            unless an error occurs """

        # Format progress bar
        self.progress_label.setText("Creating objects...")

        # 0 to 100%
        for n in range(0, 101):

            # If gets to above 85 and still not created objects - slow down
            if n > 85:
                delay = 0.0125
            else:
                delay = 0.005

            if self.object_errors:  # If any error has been picked up when creating objects - return as result
                err_list = self.object_errors
                return err_list
            elif self.assay and self.elisa_data:  # Else if the objects exist - set progress to 100% and return
                progress_callback.emit(100)
                return
            else:
                time.sleep(delay)
                progress_callback.emit(n)

    def write_files_countdown(self, progress_callback):
        """ A timer to run and update progress bar on writing files progress
            unless an error occurs """

        # Delay based on calculated time per file
        delay = (len(self.mars_files) * 0.0474) + 4.023
        delay /= 100

        for n in range(0, 101):

            if self.write_files_errors:  # If errors - return to run_error function
                err_list = self.write_files_errors
                return err_list

            elif self.trending_done and self.summary_done and self.master_done:

                # If created quickly, run remainder of progress
                for m in range(n, 100):
                    time.sleep(delay/12)
                    progress_callback.emit(m)
                return

            else:
                time.sleep(delay)
                progress_callback.emit(n)  # Emit progress

    def file_check_worker(self):
        """ Start the progress worker object.
            Check files and add to data constants if file check ok """

        # COUNTDOWN
        worker = Worker(self.file_check_countdown)  # Check file countdown
        worker.signals.result.connect(self.result_error)  # When the function returns
        worker.signals.error.connect(self.thread_error)  # If an uncaught error occurs
        worker.signals.finished.connect(self.done_file_check)  # Finished file check
        worker.signals.progress.connect(self.percent_progress)  # Emit progress

        # RUN THE FUNCTION
        worker2 = Worker(self.check_files_exist)  # Check required files exist
        worker2.signals.error.connect(self.thread_error)  # If an uncaught error occurs

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)

    def data_object_worker(self):
        """ Create assay and elisa_data objects on worker threads """

        # COUNTDOWN
        worker = Worker(self.get_object_countdown)  # Emit to progress bar
        worker.signals.result.connect(self.result_error)  # When the function returns
        worker.signals.error.connect(self.thread_error)  # If an uncaught error occurs
        worker.signals.finished.connect(self.done_create_objects)  # Finished creating objects
        worker.signals.progress.connect(self.percent_progress)  # Progress as percentage

        # CREATE OBJECTS
        worker2 = Worker(self.create_data_objects)  # Create data objects
        worker2.signals.result.connect(self.result_error)  # When the function returns something
        worker2.signals.error.connect(self.thread_error)  # When an uncaught error occurs

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)

    def process_plates_worker(self):
        """ Process the elisa data as worker threads """

        # Change max int value of progress bar to number of files
        max_cnt = len(self.mars_files)
        self.progress_bar.setMaximum(max_cnt)

        worker = Worker(self.process_plates)  # Process elisa data
        worker.signals.result.connect(self.result_error)  # If the function returns
        worker.signals.error.connect(self.thread_error)  # Uncaught error
        worker.signals.finished.connect(self.done_processing_data)  # Finished processing data
        worker.signals.progress.connect(self.int_progress)  # Emit as integer

        # Execute worker
        self.threadpool.start(worker)

    def write_files_worker(self):
        """ Write summary data - trending, master study data, run_details """

        self.progress_bar.setMaximum(100)

        # COUNTDOWN
        worker = Worker(self.write_files_countdown)  # Countdown for writing files
        worker.signals.result.connect(self.result_error)  # If function returns
        worker.signals.error.connect(self.thread_error)  # If uncaught error in thread
        worker.signals.finished.connect(self.done_write_files)  # Finished writing files
        worker.signals.progress.connect(self.percent_progress)

        # WRITE FILES
        worker2 = Worker(self.write_to_files)  # Write data
        worker2.signals.error.connect(self.thread_error)  # If uncaught error

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)

    def print_worker(self):
        """ Two workers to print pdfs and to detect jobs coming through """

        # Set progress bar to number of files to print
        self.progress_bar.setMaximum(len(self.pdf_names))

        # COUNTDOWN
        worker = Worker(self.count_print_jobs)  # Loop through n jobs
        worker.signals.result.connect(self.result_error)
        worker.signals.error.connect(self.thread_error)
        worker.signals.finished.connect(self.done_print_data)
        worker.signals.progress.connect(self.int_progress)

        # PRINT DATA
        worker2 = Worker(self.print_pdf)  # Print pdfs
        worker2.signals.error.connect(self.thread_error)  # Uncaught error

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)

    def get_data_constants(self):
        """ Get a list of constants to assign to main run script """

        self.savedir = os.path.join(Path(self.mars_files[0]).parent)  # PDF directory
        self.QC_FILE = self.find_required_file("qc_path")
        self.CURVE_FILE = self.find_required_file("curve_path")
        self.TREND_FILE = self.find_required_file("trending_path")
        self.F093_FILE = self.find_required_file("f093_path")
        self.MASTER_PATH = self.find_required_file("master_path")

    def create_data_objects(self, progress_callback):
        """ Create assay and elisa_data objects """

        self.progress_label.setText("Creating data objects...")

        # Assay object
        try:
            self.assay = Assay(self.f007_file, self.QC_FILE, self.CURVE_FILE, self.xl_id, self.mars_files)
            # Master study testing file
            master_str = self.assay.sponsor + "_" + self.assay.study + "_Master.csv"
            master_file = os.path.join(self.MASTER_PATH, master_str)
        except RangeNotFoundError:
            self.object_errors.append("Error creating assay object - please check F007 file")
            return self.object_errors

        # Create ELISA Data Object
        try:
            self.elisa_data = ELISAData(assay=self.assay, savedir=self.savedir,
                                        trend_file=self.TREND_FILE, f093_file=self.F093_FILE,
                                        master_file=master_file, xl_id=self.xl_id, ctx=self.ctx)
        except RangeNotFoundError:
            self.object_errors.append("Error creating elisa data object")
            return self.object_errors

        return []

    def process_plates(self, progress_callback):
        """ Loop through elisa files, create elisa object and process data """

        self.progress_label.setText("Processing plate data")

        # Loop through files in assay list
        n_files = 0
        progress_callback.emit(0)

        for f in self.assay.files:

            # Check if should ignore file
            ignore_file = self.check_ignore_file(f)
            if ignore_file:
                continue

            n_files += 1

            # Create ELISA Object
            self.elisa = ELISA(f, self.assay.first_list, self.assay.repeats_list,
                                   self.assay.qc_limits, self.assay.curve_vals, self.savedir,
                                   self.cut_high_ods, self.cut_low_ods, self.apply_lloq)

            # Add file to list of names for printing
            try:
                self.pdf_names.append(self.elisa.pdf_path)
            except AttributeError:
                continue

            # Check data imported correctly
            if self.elisa.data is None or self.elisa.parameters is None:
                progress_callback.emit(n_files)
                continue

            # Check that the assay and ELISA details match
            f007_ok, f007_details = self.check_f007()
            if not f007_ok:
                self.display_error_box()
                self.write_errors_to_log([f007_details])
                break

            # Create pdf, F093 and get trending data
            self.input_data()
            progress_callback.emit(n_files)

    def input_data(self):
        """ Create the pdf, create F093 if required, get trending data and add to list """

        # Input data to html template and Create pdf
        self.elisa_data.input_plate_data(self.elisa)

        # Input data to dataframe
        if self.assay.run_type != "repeats":
            self.elisa_data.data_to_table(self.elisa)

        # Save trending data to list
        if self.elisa.plate_fail == "R4" or not self.elisa.template:
            return
        else:
            self.elisa_data.get_trend_data(self.elisa)

    def print_pdf(self, progress_callback):
        """ Loop through pdf files and print """

        # Get printer name
        printer = win32print.GetDefaultPrinter()

        # Loop through pdf names
        for p in self.pdf_names:

            # Print pdf
            win32api.ShellExecute(0, "print", p, '/d:"%s"' % printer, ".", 0)

    def count_print_jobs(self, progress_callback):
        """ Search for print jobs and emit counter """

        default = win32print.GetDefaultPrinter()  # Default printer name
        phandle = win32print.OpenPrinter(str(default))  # Get printer handle

        jobs_found = 0  # Counter for number of jobs found
        n_jobs = len(self.pdf_names)  # Total number of jobs to be found
        job_ids = []  # Unique job ids so don't recount

        # Carry on until all jobs found in queue
        while jobs_found < n_jobs:

            jobs = win32print.EnumJobs(phandle, 0, -1, 1)  # Get list of jobs for printer

            for job, user in enumerate(j['pUserName'] for j in jobs):
                if user == os.getlogin():  # If sent by current user - get the job id
                    job_id = jobs[job]['JobId']

                    if job_id not in job_ids:  # If new job - increase number and add job id to list
                        jobs_found += 1
                        progress_callback.emit(jobs_found)  # Emit number of jobs found
                        job_ids.append(job_id)

            time.sleep(0.1)  # Search for new job every 0.1 seconds

    def write_to_files(self, progress_callback):
        """ Write to trending file, run_details and master study file """

        # Trend QC data
        self.elisa_data.update_trending()
        self.trending_done = True

        # Summary table of testing details
        summary_name = os.path.join(os.path.abspath(self.savedir),
                                    'run_details ' + self.assay.f007_ref + '.csv')

        # Check for suspected R35s
        df_blocks, block_list = self.get_block_list()

        # If list of blocks (may not be if repeats) - check for R35s
        if block_list:
            new_warnings = get_r35s(df_blocks, block_list)

            # Add any R35 warnings that may have been returned
            for w in new_warnings:
                self.elisa_data.warnings.append(w)

        # If run_details doesn't exist - create. Else - update
        if not os.path.isfile(summary_name):
            self.elisa_data.create_summary()
        else:
            self.elisa_data.update_summary()
            
        self.summary_done = True

        # If master study file doesn't exist - create. Else - update
        if not os.path.isfile(self.elisa_data.master_file):
            self.elisa_data.create_master()
        else:
            self.elisa_data.update_master()

        self.master_done = True

    def get_block_list(self):
        """ Get a list of blocks to check for R35s """

        df = self.elisa_data.df_plates

        block_list = []
        # Just check sample duplicates (i.e. the same list of samples on different plates)
        colnames = df.columns[1:-1].tolist()
        # Get list of duplicates - keep all
        dup_rows = df.duplicated(subset=colnames, keep=False)
        # Subset df
        df_dups = df[dup_rows]

        # If duplicates
        if not df_dups.empty:
            # Group by duplicate and create comma separated list
            grouped = df_dups.groupby(colnames)
            r_blocks = grouped.agg(lambda x: ','.join(x))['Plate'].tolist()
            [block_list.append(f.split(",")) for f in r_blocks]

            return df_dups, block_list
        else:
            return None

    def check_ignore_file(self, file):
        """ Check that the file in the assay object is not a pdf, xlsm, json or run_details """

        # Split file name to see if run details file
        run_split = file.split("\\")[-1]
        run_split = run_split.split(" ")[0]

        # If file is pdf, xlsm, json or name is run details - ignore
        if Path(file).suffix.upper() == '.PDF' \
                or Path(file).suffix.upper() == '.XLSM' \
                or Path(file).suffix.upper() == '.JSON' \
                or run_split == "run_details":
            return True
        else:
            return False

    def check_f007(self):
        """ Check that the details on the F007 match those obtained from MARS file """

        # Check F007 details matches read time
        if self.assay.tech != self.elisa.barc_tech:
            return False, \
                   "Technician initials in F007 (" + self.assay.tech + ") " \
                   + "don't match those in barcode (" + self.elisa.barc_tech + ") "
        elif self.assay.date != self.elisa.barc_date:
            return False, \
                   "Assay date in F007 (" + self.assay.date + ") " \
                   + "doesn't match that in barcode (" + self.elisa.barc_date + ") "
        else:
            return True, ""

    def check_files_exist(self, progress_callback):
        """ Check that the required files exist """

        # Get settings page
        page = find_widget("path_settings")

        # List of widgets containing filenames/filepaths to check
        widget_dict = {"qc_path": "QC Limits file not found",
                       "curve_path": "IgG Curve concentrations not found",
                       "trending_path": "Trending file not found",
                       "f093_path": "F093 template not found"}

        # List of errors
        err_list = []

        # Loop through file widgets
        for key, val in widget_dict.items():

            # Get file path from text based on object name
            path_text = page.findChild(QLineEdit, key).text()

            # Check if is a file
            if not os.path.isfile(path_text):
                err_list.append(val)

        # Check that the master study directory exists
        path_text = page.findChild(QLineEdit, "master_path").text()
        if not os.path.isdir(path_text):
            err_list.append("Master study data directory not found")

        # WRite to error log and display error message box
        if err_list:
            self.file_check_errors = err_list
            return err_list
        else:
            self.get_data_constants()
            return []

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

    def display_error_box(self, warning="An error has occurred", prompt_error=True):
        """ Display a generic error message """

        if prompt_error:
            warning = warning + "\nSee Error Log for more details"

        err_box = QMessageBox(self)
        err_box.setIcon(QMessageBox.Warning)
        err_box.setText(warning)
        err_box.setWindowTitle("Error")
        err_box.exec_()

    def set_progress_opaque(self):
        """ Make progress bar 'visible' with style """

        self.progress_bar.setStyleSheet("QProgressBar{"
                                        "border: 1px solid grey;"
                                        "text-align: center;"
                                        "color:rgba(0,0,0,255);"
                                        "border-radius: 2px;"
                                        "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:0, y2:0, "
                                        # "stop:0 rgba(0, 0, 0, 0), "
                                        # "stop:1 rgba(0, 0, 0, 0));}"
                                        "stop:0 rgba(255, 255, 255, 255), "
                                        "stop:1 rgba(255, 255, 255, 255));}"
                                        "QProgressBar::chunk{"
                                        "background-color: rgba(185, 202, 141, 255);"
                                        # + style_str +
                                        "margin: 1px;"
                                        "}")

    def set_progress_transparent(self):
        """ Make progress bar invisible """

        self.progress_bar.setStyleSheet("QProgressBar{"
                                        "border: 1px solid transparent;"
                                        "text-align: center;"
                                        "color:rgba(0,0,0,0);"
                                        "border-radius: 5px;"
                                        "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:0, y2:0, "
                                        # "stop:0 rgba(0, 0, 0, 0), "
                                        # "stop:1 rgba(0, 0, 0, 0));}"
                                        "stop:0 rgba(182, 182, 182, 0), "
                                        "stop:1 rgba(209, 209, 209, 0));}"
                                        "QProgressBar::chunk{"
                                        "background-color: rgba(0,0,0,100);"
                                        "}")

    def od_checkbox_changed(self, cbox, line_edit):
        """ If OD checkbox has been changed """

        if cbox.isChecked():
            line_edit.setEnabled(True)
        else:
            line_edit.setEnabled(False)

    def amend_checkbox_changed(self, cbox):
        """ If user selects to make amendments """

        # Determine whether checked or unchecked
        if cbox.isChecked():

            # Check that there is a F007 loaded
            if not self.f007_file or not self.mars_files:
                self.display_error_box(warning="Please select F007 and MARS files\nbefore adding amendments",
                                       prompt_error=False)
                cbox.setChecked(False)
                return

            # Check F007 is valid - if so, load sample table
            sample_table = self.get_sample_table()

            if sample_table is None:  # If not valid F007
                self.display_error_box(warning="This doesn't appear to be a valid F007", prompt_error=False)
                return
            elif sample_table.empty:  # IF valid F007 but no sample table
                self.display_error_box(warning="Sample Table appears to be empty", prompt_error=False)
                return
            else:  # Open amendments page

                self.NewWindow = AmendmentsWindow(f007=self.f007_file, sample_table=sample_table,
                                                  marsfiles=self.mars_files, caller=self)
                self.NewWindow.show()



            # If so, load amendments page


    def get_sample_table(self):
        """ Loose check to see if expected sheets are in F007 and return
            sample table as dataframe """

        # Excel Object
        xl = pd.ExcelFile(self.f007_file)

        # Get all sheet names
        all_sheets = xl.sheet_names

        # Test for these three sheets
        # If detect all, parse sample table
        test_list = ['Plate Layout', 'Sample Table', 'Sponsors and Study IDs']
        if all(x in all_sheets for x in test_list):
            df = xl.parse('Sample Table',
                          dtype={'Sample 1': 'str', 'Sample 2': 'str',
                                 'Sample 3': 'str', 'Sample 4': 'str'})
            return df
        else:
            return None


    def general_popup(self, message):
        """ A general notice popup box"""

        message_box = QMessageBox(self)
        message_box.setIcon(QMessageBox.Information)
        message_box.setText(message)
        message_box.setWindowTitle("Unable to perform operation")
        message_box.exec_()

    def btn_f007_clicked(self):
        """ Search for and select F007 containing assay details """

        self.f007_file = ""
        self.check_amend.setChecked(False)
        settings = QSettings()
        # Get default directory - check if one has been previously saved
        saved_dir = settings.value("f007_dir") or ""

        # If saved_dir isn't empty and it is an existing directory - set as default
        if saved_dir and os.path.isdir(saved_dir) or os.path.isfile(saved_dir):
            default_dir = str(saved_dir)
        else:
            default_dir = get_default_dir()

        # what to display
        title = 'Choose a F007 file'
        file_text = "xlsm Files (*.xlsm)"

        # File picker
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename = QFileDialog.getOpenFileName(self, title, default_dir, file_text, options=options)

        if filename[0]:
            self.f007_file = filename[0].replace("/","\\")
            self.txt_f007.setText(filename[0].replace("/","\\"))

    def btn_mars_clicked(self):
        """ Search for and select MARS files for processing ELISA data """

        self.mars_files = []
        settings = QSettings()

        # Get default directory - check if one has been previously saved
        saved_dir = settings.value("mars_dir") or ""

        # If saved_dir not empty and is existing directory
        if saved_dir and os.path.isdir(saved_dir) or os.path.isfile(saved_dir):
            default_dir = str(saved_dir)
        else:
            default_dir = get_default_dir()

        # what to display
        title = 'Choose MARS file(s)'
        file_text = "CSV Files (*.csv *.CSV)"

        # File picker
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filenames = QFileDialog.getOpenFileNames(self, title, default_dir, file_text, options=options)

        # Save filenames to attribute
        self.mars_files = filenames[0]
        self.mars_files = [f.replace("/","\\") for f in self.mars_files]
        self.mars_files = [x for x in self.mars_files if not self.check_ignore_file(x)]

        # Get filenames only and display in lineedit as comma-separated list
        filenames = [f.split('/')[-1] for f in filenames[0]]
        self.txt_mars.setText(', '.join(filenames))

    def find_required_file(self, name):
        """ Get the text from the settings page for a required file"""

        # Get settings page
        page = find_widget("path_settings")

        # Get file path from text based on object name
        path_text = page.findChild(QLineEdit, name).text()

        return path_text

    def set_button_states(self, enabled):
        """ Enable or disable each of the buttons on the elisa data processing page """

        self.btn_run.setEnabled(enabled)
        self.btn_mars.setEnabled(enabled)
        self.btn_f007.setEnabled(enabled)
        self.group_box.setEnabled(enabled)

    def kill_xl(self):
        """ Kill any Excel processes """

        # If there are no Excel processes
        if not xw.apps:
            return

        # Loop through Excel processes, if process ID matches one working on:
        # Kill and close all books
        for app in xw.apps:
            if app.pid == self.xl_id:
                for book in app.books:
                    book.close()
                app.kill()

class ErrorTab(QWidget):
    def __init__(self, *args, **kwargs):
        super(ErrorTab, self).__init__(*args, **kwargs)

        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)

        layout = QGridLayout(self)

        # Create error log
        self.error_log = QTextEdit()
        self.error_log.setReadOnly(True)
        self.error_log.setObjectName("error_log")
        layout.addWidget(self.error_log, 1, 1, 20, 4)

        # Clear button
        self.clear_log = QPushButton("Clear Log")
        self.clear_log.setObjectName("clear_log")
        self.clear_log.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        # self.clear_log.clicked.connect(self.error_log.clear)
        self.clear_log.clicked.connect(self.clear_log_text)

        # Save button
        self.save_log = QPushButton("Save Log")
        self.save_log.setObjectName("save_log")
        self.save_log.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        self.save_log.clicked.connect(self.log_to_file)

        # Add buttons to layout
        layout.addWidget(self.clear_log, 21, 3)
        layout.addWidget(self.save_log, 21, 4)

    def log_to_file(self):
        """ Save the error log to a text file """

        # Date details
        now = datetime.now()
        date_str = now.strftime("%d%b%y %H%M%S")
        details_str = "Error generated at: " + date_str

        # Error log text
        text = self.error_log.toPlainText()

        # Save path
        path = QFileDialog.getExistingDirectory(self, 'Save error log')

        # Check path was chosen
        if path:
            file_name = "ELISA error_log " + date_str + ".txt"
            save_name = os.path.join(os.path.abspath(path), file_name)

            with open(save_name, 'w') as file:
                file.write(text)

    def clear_log_text(self):
        self.error_log.clear()


def find_widget(obj_name):
    """ Find a widget from the application based on object name """

    # Get list of widgets from application
    widgets = QApplication.allWidgets()
    widget = None

    # If found the widget - break
    for w in widgets:
        if w.objectName() == obj_name:
            widget = w
            break

    return widget

class AmendmentsWindow(QMainWindow):

    def __init__(self, f007, sample_table, marsfiles, caller=None, *args, **kwargs):
        super(AmendmentsWindow, self).__init__(*args, **kwargs)

        self.setWindowTitle("Amendments Page")
        self.setWindowModality(Qt.ApplicationModal)
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(141, 185, 202))
        self.setAutoFillBackground(True)
        self.setPalette(palette)
        self.datapage = caller
        self.df = pd.DataFrame(columns=['Plate', 'Sample', 'Amendment'])

        """ PARAMETERS """
        self.f007 = f007  # F007 file
        self.marsfiles = marsfiles  # MARS FILES LIST
        self.sample_table = sample_table  # SAMPLE TABLE
        self.sample_table.set_index('Plate ID', inplace=True)

        """ DROPDOWN SETTINGS """
        self.plateblock = "Plate"
        self.selection = ""
        self.samples = []
        self.amendment = ""
        self.amend_cnt = 0

        x, y, w, h = get_geometry()  # Get optimal geometry
        self.setFixedSize(w, h)  # No resize

        # Layouts
        # layout_main = QVBoxLayout(self)

        self.layout = QGridLayout(self)
        self.layout.setSpacing(10)

        """ GET PLATE IDs """
        self.plate_list = [get_plate_id(f) for f in self.marsfiles]
        self.selection = self.plate_list[0]

        """ PLATE/BLOCK COMBO """
        label_plateblock = QLabel(text="Choose amendment to\nplate, block or sample", alignment=Qt.AlignBottom)
        self.combo_plateblock = QComboBox(objectName="combo_plateblock")
        [self.combo_plateblock.addItems(x for x in ["Plate", "Block", "Sample"])]
        self.combo_plateblock.currentIndexChanged[str].connect(self.plateblock_changed)

        """ SELECT PLATES/BLOCKS - GIVE A LIST OF PLATES OR BLOCKS DEPENDENT ON OPTION ABOVE"""
        label_selection = QLabel(text="Select plate or block", alignment=Qt.AlignBottom)
        self.combo_selection = QComboBox(objectName="combo_selection")
        self.add_plate_list()
        self.combo_selection.currentIndexChanged[str].connect(self.selection_changed)

        """ SAMPLES """
        label_samples = QLabel(text="Select sample(s)", alignment=Qt.AlignBottom)
        self.combo_samples = QComboBox(objectName="combo_samples")
        self.combo_samples.setEnabled(False)
        self.combo_samples.currentIndexChanged[str].connect(self.samples_changed)

        """ FAIL CODES/SAMPLE TESTED IN ERROR """
        label_amendments = QLabel(text="Select amendment to apply", alignment=Qt.AlignBottom)
        self.combo_amendments = QComboBox(objectName="combo_amendments")
        # Default plate fail codes
        fail_list = ['R1', "R6", "R7", "R13"]
        [self.combo_amendments.addItems(f for f in fail_list)]
        self.amendment = fail_list[0]
        self.combo_amendments.currentIndexChanged[str].connect(self.amendment_changed)

        """ ADD BUTTON """
        self.btn_add = QPushButton(text="Add", objectName="btn_add")
        self.btn_add.clicked.connect(self.add_amendment)

        """ TABLE WITH BLOCK/PLATES FOR AMENDMENTS """
        self.table = QTableWidget()
        self.table.setRowCount(8)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Plate/Block', 'Sample', 'Amendment'])
        self.table.verticalHeader().setVisible(False)
        header_font = QFont()
        header_font.setBold(True)
        self.table.horizontalHeaderItem(0).setFont(header_font)
        self.table.horizontalHeaderItem(1).setFont(header_font)
        self.table.horizontalHeaderItem(2).setFont(header_font)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        """ BUTTON OK AND CANCEL """
        self.btn_ok = QPushButton(text="OK")
        self.btn_ok.clicked.connect(self.ok_pressed)
        self.btn_cancel = QPushButton(text="Cancel")
        self.btn_cancel.clicked.connect(sys.exit)


        """ ADD ITEMS TO LAYOUT """
        label_w = 10
        combo_w = 5  # Combo and label grid width
        btn_w = 2
        list_w = 10

        self.layout.addWidget(label_plateblock, 0, 0, 1, label_w)
        self.layout.addWidget(self.combo_plateblock, 1, 0, 1, combo_w)  # PLATE/BLOCK/SAMPLES

        self.layout.addWidget(label_selection, 2, 0, 1, label_w)
        self.layout.addWidget(self.combo_selection, 3, 0, 1, combo_w)  # SELECTION OF PLATE/BLOCK

        self.layout.addWidget(label_samples, 4, 0, 1, label_w)
        self.layout.addWidget(self.combo_samples, 5, 0, 1, combo_w)  # SELECTION OF SAMPLES

        self.layout.addWidget(label_amendments, 6, 0, 1, label_w)
        self.layout.addWidget(self.combo_amendments, 7, 0, 1, combo_w) # SELECT FROM LIST OF FAIL CODES/AMENDMENTS

        self.layout.addWidget(QWidget(), 8, 0, 1, combo_w)
        self.layout.addWidget(self.btn_add, 9, 0, 1, btn_w, Qt.AlignBottom)  # ADD BUTTON
        self.layout.addWidget(self.table, 10, 0, 10, list_w)  # table
        self.layout.addWidget(self.btn_ok, 20, 5, 1, 2)  # OK BUTTON
        self.layout.addWidget(self.btn_cancel, 20, 8, 1, 2)  # CANCEL BUTTON
        # layout_main.addLayout(self.layout)

        """ SET CENTRAL WIDGET """
        widget = QWidget()
        widget.setLayout(self.layout)
        widget.setContentsMargins(20, 20, 20, 20)
        self.setCentralWidget(widget)

        # Adjust row height for table rows - default to 8 rows
        table_height = self.table.height()
        for row in range(self.table.rowCount()):
            self.table.setRowHeight(row, table_height/7)

    def ok_pressed(self):
        """ Pass amendments back to datapage """

        self.df.drop_duplicates(inplace=True)

        # Check the user wants to commit
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText("Are you sure you want to make these amendments?")
        msg.setWindowTitle("Commit amendments")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

        return_val = msg.exec_()
        if return_val == QMessageBox.Ok:
            self.datapage.amendments = self.df

        self.close()

    def add_plate_list(self):
        """ Add a list of plate IDs to the selection box """

        # Clear the selection combobox and set default selection to first plate
        self.combo_selection.clear()
        [self.combo_selection.addItems(x for x in self.plate_list)]
        self.selection = self.plate_list[0]

    def add_block_list(self):
        """ Add a list of blocks to the selection box """

        # Clear selection box
        self.combo_selection.clear()
        # If first run just add list of blocks
        # If mixed add all single letters and duplicate rows
        # If repeats add all duplicates

        """ FIRST LOOK FOR ANY SINGLE LETTER (FIRST RUN BLOCKS) """
        df = self.sample_table
        block_list = []
        mask = (df.index.str.len() == 1)  # Mask only single letters in sample table
        fr_blocks = df.loc[mask].index.tolist()  # Get row number of single letters in index
        [block_list.append(f) for f in fr_blocks]

        """ REPEATS """

        # If there are more blocks than picked up from first run blocks
        if len(fr_blocks) < df.shape[0]:

            # Colnames to find duplicated samples across plate
            # Index is numeric
            df.reset_index(inplace=True)
            # Only sample columns (not plate IDs)
            colnames = df.columns[1:].tolist()
            # Get list of duplicates - keep all
            dup_rows = df.duplicated(subset=colnames, keep=False)
            # Subset df
            df_dups = df[dup_rows]

            # If duplicates
            if not df_dups.empty:
                # Group by duplicate and create comma separated list
                grouped = df_dups.groupby(colnames)
                r_blocks = grouped.agg(lambda x: ', '.join(x))['Plate ID'].tolist()
                [block_list.append(f) for f in r_blocks]
            else:
                block_list.append('None')

            # Set sample table index to plate ID
            self.sample_table.set_index('Plate ID', inplace=True)

        # Add blocks to selection combo
        [self.combo_selection.addItems(x for x in block_list)]
        self.selection = block_list[0]


    def add_sample_list(self):
        """ Add a list of sample IDs to the selection box based on the selected plate
            in the plate combobox"""

        # First get whether first run/repeats/mixed to determine how to retrieve sample IDs
        # Could 'try' and retrieve the plate ID from the sample list, if not, look for letter
        # Finally - can't find plate in list

        # Get plate ID from dropdown (and/or blocks from sample table)
        plate_id = self.selection

        try:
            self.samples = self.sample_table.loc[plate_id].tolist()
        except KeyError:
            try:
                self.samples = self.sample_table.loc[plate_id[-1]].tolist()
            except KeyError:
                self.combo_samples.clear()
                self.samples = []
                return


        # Clear sample list and replace with samples from plate lookup
        self.combo_samples.clear()
        [self.combo_samples.addItems(str(x) for x in self.samples)]

    def add_amendment(self):
        """ Add the amendment to the table of amendments """

        # Get plate ID(s)
        plate_ids = self.combo_selection.currentText()

        # Get sample (ALL for plate/block amendments)
        if self.combo_plateblock.currentText() != "Sample":
            samples = "ALL"
        else:
            samples = self.combo_samples.currentText()

        # Get fail/amendment
        amendment = self.combo_amendments.currentText()

        # Add to table
        self.table.setItem(self.amend_cnt, 0, QTableWidgetItem(plate_ids))
        self.table.setItem(self.amend_cnt, 1, QTableWidgetItem(samples))
        self.table.setItem(self.amend_cnt, 2, QTableWidgetItem(amendment))

        # Increment amend_cnt
        self.amend_cnt += 1

        # Enter data to dataframe
        if self.combo_plateblock.currentText() == "Block":
            self.update_df_block()
        else:
            self.update_df_platesample()

    def update_df_platesample(self):
        """ Add the amendments to a dataframe to pass to datapage amendments
            if plate or sample """

        # Get next empty row from dataframe
        next_row = len(self.df.index)

        sample = self.combo_samples.currentText()
        if not sample:
            sample = ""

        # Add data
        self.df.loc[next_row] = [self.selection, sample, self.amendment]

    def update_df_block(self):
        """ Update the dataframe from a block """

        next_row = len(self.df.index)
        # Loop through plates and add if correct block

        # If block only length 1, add all plates from that block
        if len(self.selection) == 1:
            for p in self.plate_list:

                # If matches block - add
                if p[-1] == self.selection:
                    self.df.loc[next_row] = [p, "", self.amendment]
                    next_row += 1

        else:

            # Get all plates in block (repeats)
            plate_list = self.selection.split(",")

            for p in plate_list:
                self.df.loc[next_row] = [p, "", self.amendment]
                next_row += 1


    def get_selection_items(self):
        """ Get a list of plates or blocks to add to selection combo
            Can get list of blocks from F007 and plates from MARS files """


    def plateblock_changed(self, plateblock):
        """ Function to get list of blocks or plates """

        # Change plateblock selection
        self.plateblock = plateblock

        if plateblock == "Plate":

            # Disable and clear sample box
            self.combo_samples.clear()
            self.combo_samples.setEnabled(False)

            # Add plate_list to selection combo box
            self.add_plate_list()

        elif plateblock == "Block":  # Get list of blocks

            # Disable sample box
            self.combo_samples.clear()
            self.combo_samples.setEnabled(False)

            # Add list of blocks to combobox
            self.add_block_list()

        else:

            # Enable Sample dropdown
            self.combo_samples.setEnabled(True)
            # Add plate list to selection combo box
            self.add_plate_list()

    def selection_changed(self, selection):
        """ When plate or block has changed - change sample list if 'Sample' selected
            in the combo_plateblock. Change list of amendments - Tested in error/<0.15/QNS for samples
            Various plate fails and dropped plate for plate and R35 for block (maybe management decision
            or technician error as well) """

        # Update selection of plate/block
        self.selection = selection
        # Check the selection in combo_plateblock
        # If Plate - just change possible amendments and disable sample list
        # If block - just change possible amendments and disable sample list
        # If Sample - change sample list (self.samples and dropdown)

        if self.plateblock == "Plate":

            # Add list of possible amendments for plate fail
            self.combo_amendments.clear()
            fail_list = ['R1', "R6", "R7", "R13"]
            [self.combo_amendments.addItems(f for f in fail_list)]

        elif self.plateblock == "Block":

            # Add list of possible amendments for block fail
            self.combo_amendments.clear()
            fail_list = ['R1', "R6", "R7", "R13", "R35"]
            [self.combo_amendments.addItems(f for f in fail_list)]

        elif self.plateblock == "Sample" and self.selection:

            # Clear amendments box
            self.combo_amendments.clear()

            # Get samples associated with plate
            self.add_sample_list()

    def samples_changed(self, sample):
        """ If user selects a samples from dropdown or list is updated """

        # Add list of possible amendments for sample
        self.combo_amendments.clear()

        # If there is a sample ID update amendments box
        if self.combo_samples.currentText():

            self.combo_amendments.addItems(['Tested in error'])

    def amendment_changed(self, amendment):
        """ If amendment to be added has changed """

        self.amendment = amendment

def get_geometry():
    """ Get the screen resolution and return the optimal dimensions on startup """

    # Get screen resolution in pixels
    screen_width = GetSystemMetrics(0)
    screen_height = GetSystemMetrics(1)

    # Get geometry relative to 1920 x 1080 screen
    x = 450 / 1920 * screen_width
    y = 125 / 1080 * screen_height
    width = 500 / 1920 * screen_width
    height = 800 / 1080 * screen_height

    return x, y, width, height


def get_plate_id(filename):
    """ Get technician, date and plate ID from barcode """

    barcode = os.path.split(filename)[-1]

    barcode = barcode.split("_")[1].split(".")[0]
    barcode = barcode + "NO"

    # Check if last character is alpga
    if barcode[0].isalpha():
        bstring = barcode[1:]
    else:
        bstring = barcode

    # Get details - extract portion of barcode

    if bstring[-2:].isalpha():  # If two letters at end
        plate_id = bstring[:-10]

    elif bstring[-1].isalpha():  # If one letter at end
        plate_id = bstring[:-9]

    else:  # If no letters at end
        plate_id = bstring[:-8]

    return plate_id


def get_r35s(df, block_list):
    """ Check for suspected R35s from list of blocks """

    new_warnings = []
    # Subset by block
    for plates in block_list:

        block_df = df[df['Plate'].isin(plates)]

        # Calculate % fail
        percent_fail = block_df['Fail'].notna().sum() / len(block_df.index) * 100

        # If >70% R35
        # Check if all plates have same block ID and return just that if so
        is_r35 = True if percent_fail > 70 else False
        if is_r35:

            # Check if all plates are in block A for example
            whole_block = all(plate[-1] == plates[0][-1] for plate in plates)

            if whole_block:
                block_str = plates[0][-1]
            else:
                block_str = '[' + ','.join(plates) + ']'

            r35_string = 'Suspected R35 on block ' + block_str + ' Please check and amend if necessary'
            new_warnings.append(r35_string)

    return new_warnings
