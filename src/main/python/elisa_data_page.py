from PyQt5 import QtWidgets
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import QSettings, QObject, pyqtSignal, QRunnable, pyqtSlot, QThreadPool
from PyQt5.QtWidgets import QLabel, QGridLayout, QWidget, QVBoxLayout, QPushButton, \
    QTextEdit, QHBoxLayout, QTabWidget, QLineEdit, QSizePolicy, \
    QGroupBox, QCheckBox, QProgressBar, QFileDialog, QApplication, QMessageBox, QComboBox
from datetime import datetime
from settings_page import get_default_dir, PageSettings
from assay import Assay
from elisa_data import ELISAData
from elisa import ELISA
# from error_handling import show_exception_box
from error_handling import RangeNotFoundError
import time
import os
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

    def __init__(self, ctx, *args, **kwargs):
        super(PageData, self).__init__(*args, **kwargs)

        self.setDocumentMode(False)
        self.setMovable(False)
        self.setTabPosition(QTabWidget.South)
        self.main = DataTab(ctx)
        self.error = ErrorTab()
        self.addTab(self.main, "Data")
        self.addTab(self.error, "Error Log")

        self.main.error_log = self.error.error_log

    def write_to_log(self, msg):
        self.error.error_log.append(msg)


class DataTab(QWidget):

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

        # Layout spacing
        layout_main.setSpacing(20)

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

        # Add widgets
        layout_files.addWidget(label, 1, 0)
        layout_files.addWidget(self.txt_mars, 1, 1)
        layout_files.addWidget(self.btn_mars, 1, 2)
        layout_files.addWidget(QWidget(), 2, 0)

        # Parameter group box
        self.group_box = QGroupBox()
        self.group_box.setObjectName("grp_box")
        self.group_box.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        # Grid layout for group box
        group_layout = QGridLayout(self.group_box)
        self.group_box.setTitle("Select Parameters")

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

        # OD text boxes
        self.txt_od_upper = QLineEdit(objectName="txt_upper_od")
        self.txt_od_upper.setMaximumWidth(100)
        self.txt_od_upper.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        # self.txt_od_upper.setEnabled(False)

        self.txt_od_lower = QLineEdit(objectName="txt_lower_od")
        self.txt_od_lower.setMaximumWidth(100)
        self.txt_od_lower.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
        # Layout
        group_layout.addWidget(QLabel("Select pre-defined parameters"), 0, 0, 1, 1)
        group_layout.addWidget(self.combo_options, 0, 1, 1, 1)
        group_layout.addWidget(QLabel(""), 1, 0, 1, 2)
        group_layout.addWidget(self.cb_upper_od, 2, 0, 1, 1)
        group_layout.addWidget(self.cb_lower_od, 3, 0, 1, 1)
        group_layout.addWidget(self.cb_lloq, 4, 0, 1, 1)
        group_layout.addWidget(self.cb_print, 5, 0, 1, 1)
        group_layout.addWidget(self.txt_od_upper, 2, 1, 1, 1)
        group_layout.addWidget(self.txt_od_lower, 3, 1, 1, 1)

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

    def init_parms(self):
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
        """ Percentage progress """

        self.progress_bar.setFormat(u"(%v / %m)")
        self.progress_bar.setValue(n)

    def result_error(self, err_list):

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

        # Format progress bar
        self.set_button_states(False)
        self.progress_label.setVisible(True)
        self.progress_label.setText("Checking required files...")
        self.set_progress_opaque()

    def done_file_check(self):

        # self.progress_label.setText("File check complete!")
        self.threadpool.waitForDone()
        self.progress_label.setText("Creating data objects...")
        self.data_object_worker()
        # self.data_object_worker()  # Create data objects

    def done_create_objects(self):

        self.threadpool.waitForDone()
        self.progress_bar.setValue(100)
        if self.assay:
            self.process_plates_worker()
        else:
            return

    def done_processing_data(self):
        """ When finished processing elisa objects """

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
            self.print_worker()

    def done_print_data(self):

        self.progress_label.setText("Writing summary data to file...")
        self.write_files_worker()

    def done_write_files(self):

        # Save F093

        if self.assay.run_type != "repeats":
            self.elisa_data.f093_to_excel()
        self.progress_bar.setValue(100)
        self.progress_label.setText("Finished")
        self.set_button_states(True)

        self.init_parms()
        self.kill_xl()

    def thread_error(self, exc_info):
        """ When error raised from within thread """

        self.kill_xl()
        self.display_error_box()
        self.error_log.setTextColor(QColor(255, 0, 0))  # Red 'Error' message
        self.error_log.append("A thread error occurred:\n")
        self.error_log.setTextColor(QColor(0, 0, 0))
        self.error_log.append('{0}: {1}'.format(exc_info[0].__name__, exc_info[1]))
        self.error_log.append("\n" + exc_info[2] + "\n\n")
        self.progress_label.setText("Operation cancelled")
        self.set_button_states(True)

    def file_check_countdown(self, progress_callback):
        """ A timer to run and update progress bar unless an error occurs """

        # Format progress bar
        # self.set_button_states(False)
        # self.progress_label.setVisible(True)
        # self.progress_label.setText("Checking required files...")
        # self.set_progress_opaque()

        # WRite to progress unless an error has been found
        # while not self.error_occurred:
        for n in range(0, 101):

            if self.file_check_errors:
                err_list = self.file_check_errors
                return err_list
            elif self.QC_FILE and self.CURVE_FILE and self.TREND_FILE and self.F093_FILE and self.MASTER_PATH:
                progress_callback.emit(100)
                return
            else:
                time.sleep(0.02)
                progress_callback.emit(n)

    def get_object_countdown(self, progress_callback):
        """ A timer to run and update progress bar unless an error occurs """

        # Format progress bar
        self.progress_label.setText("Creating objects...")

        # WRite to progress unless an error has been found
        # while not self.error_occurred:
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
        """ A timer to run and update progress bar unless an error occurs """

        # WRite to progress unless an error has been found
        delay = (len(self.mars_files) * 0.0474) + 4.023
        delay /= 100
        for n in range(0, 101):

            if self.write_files_errors:
                err_list = self.write_files_errors
                return err_list
            elif self.trending_done and self.summary_done and self.master_done:
                for m in range(n, 100):
                    time.sleep(delay/12)
                    progress_callback.emit(m)
                return
            else:
                time.sleep(delay)
                progress_callback.emit(n)

    def file_check_worker(self):
        """ Start the progress worker object. Check files and add to data constants if file check ok """

        # Pass the function to execute
        worker = Worker(self.file_check_countdown)  # Any other args, kwargs are passed to the run function
        worker.signals.result.connect(self.result_error)
        worker.signals.error.connect(self.thread_error)
        worker.signals.finished.connect(self.done_file_check)
        worker.signals.progress.connect(self.percent_progress)

        # Pass the function to execute
        worker2 = Worker(self.check_files_exist)  # Any other args, kwargs are passed to the run function
        # worker2.signals.result.connect(self.result_error)
        worker2.signals.error.connect(self.thread_error)
        # worker2.signals.finished.connect(self.done_file_check)
        # worker2.signals.progress.connect(self.percent_progress)

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)
        # self.check_files_exist()

    def data_object_worker(self):
        """ """

        # Pass the function to execute
        worker = Worker(self.get_object_countdown)  # Any other args, kwargs are passed to the run function
        worker.signals.result.connect(self.result_error)
        worker.signals.error.connect(self.thread_error)
        worker.signals.finished.connect(self.done_create_objects)
        worker.signals.progress.connect(self.percent_progress)

        # Pass the function to execute
        worker2 = Worker(self.create_data_objects)  # Any other args, kwargs are passed to the run function
        worker2.signals.result.connect(self.result_error)
        worker2.signals.error.connect(self.thread_error)
        # worker2.signals.finished.connect(self.done_create_objects)

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)

    def process_plates_worker(self):
        """ """

        max_cnt = len(self.mars_files)
        self.progress_bar.setMaximum(max_cnt)
        worker = Worker(self.process_plates)
        worker.signals.result.connect(self.result_error)
        worker.signals.error.connect(self.thread_error)
        worker.signals.finished.connect(self.done_processing_data)
        worker.signals.progress.connect(self.int_progress)

        self.threadpool.start(worker)

    def write_files_worker(self):

        self.progress_bar.setMaximum(100)
        # Pass the function to execute
        worker = Worker(self.write_files_countdown)  # Any other args, kwargs are passed to the run function
        worker.signals.result.connect(self.result_error)
        worker.signals.error.connect(self.thread_error)
        worker.signals.finished.connect(self.done_write_files)
        worker.signals.progress.connect(self.percent_progress)

        # Pass the function to execute
        worker2 = Worker(self.write_to_files)  # Any other args, kwargs are passed to the run function
        # worker2.signals.result.connect(self.result_error)
        worker2.signals.error.connect(self.thread_error)
        # worker2.signals.finished.connect(self.done_file_check)
        # worker2.signals.progress.connect(self.percent_progress)

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)

    def print_worker(self):
        """ Two workers to print pdfs and to detect jobs coming through """

        self.progress_bar.setMaximum(len(self.pdf_names))
        # Pass the function to execute
        worker = Worker(self.count_print_jobs)  # Any other args, kwargs are passed to the run function
        worker.signals.result.connect(self.result_error)
        worker.signals.error.connect(self.thread_error)
        worker.signals.finished.connect(self.done_print_data)
        worker.signals.progress.connect(self.int_progress)

        # Pass the function to execute
        worker2 = Worker(self.print_pdf)  # Any other args, kwargs are passed to the run function
        # worker2.signals.result.connect(self.result_error)
        worker2.signals.error.connect(self.thread_error)
        # worker2.signals.finished.connect(self.done_file_check)
        # worker2.signals.progress.connect(self.percent_progress)

        # Execute
        self.threadpool.start(worker)
        self.threadpool.start(worker2)

    def get_data_constants(self):
        """ Get a list of constants to assign to main run script """

        # rootdir ??
        self.savedir = os.path.join(Path(self.mars_files[0]).parent)
        self.QC_FILE = self.find_required_file("qc_path")
        self.CURVE_FILE = self.find_required_file("curve_path")
        self.TREND_FILE = self.find_required_file("trending_path")
        self.F093_FILE = self.find_required_file("f093_path")
        self.MASTER_PATH = self.find_required_file("master_path")

    def create_data_objects(self, progress_callback):
        """ Create assay and elisa_data objects """

        # progress_callback.emit(10)
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

        # progress_callback.emit(95)
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
        # Get file
        n_files = 0
        progress_callback.emit(0)
        for f in self.assay.files:

            # Check if should ignore file
            ignore_file = self.check_ignore_file(f)

            if ignore_file:
                continue

            n_files += 1

            # Create ELISA Object for Assay file
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

        # Input data to f093
        if self.assay.run_type != "repeats":
            self.elisa_data.data_to_table(self.elisa)

        # Save trending data to list
        if self.elisa.plate_fail == "R4" or not self.elisa.template:
            return
        else:
            self.elisa_data.get_trend_data(self.elisa)

    def print_pdf(self, progress_callback):
        """ Loop through pdf files and print """

        ctr = 0
        printer = win32print.GetDefaultPrinter()
        for p in self.pdf_names:
            ctr += 1

            win32api.ShellExecute(0, "print", p, '/d:"%s"' % printer, ".", 0)

    def count_print_jobs(self, progress_callback):
        """ Search for print jobs and emit counter """

        default = win32print.GetDefaultPrinter()
        phandle = win32print.OpenPrinter(str(default))

        jobs_found = 0  # Counter
        n_jobs = len(self.pdf_names)  # Total number of jobs to be found
        job_ids = []  # Unique job ids so don't recount

        # Carry on until all jobs found
        while jobs_found < n_jobs:

            jobs = win32print.EnumJobs(phandle, 0, -1, 1)  # Get list of jobs for printer

            for job, user in enumerate(j['pUserName'] for j in jobs):
                if user == os.getlogin() or user == "sejjinn":  # If sent by current user - get the job id
                    job_id = jobs[job]['JobId']

                    if job_id not in job_ids:  # If new job - increase number and add job id to list
                        jobs_found += 1
                        progress_callback.emit(jobs_found)
                        job_ids.append(job_id)

            time.sleep(0.1)  # Search every 0.1 seconds

    def write_to_files(self, progress_callback):
        """ Write to trending file, run_details and master study file """

        # Trend QC data
        start = time.time()
        self.elisa_data.update_trending()
        self.trending_done = True

        # Summary table of testing details
        summary_name = os.path.join(os.path.abspath(self.savedir),
                                    'run_details ' + self.assay.f007_ref + '.csv')

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

    def check_ignore_file(self, file):
        """ Check that the file in the assay object is not a pdf, xlsm, json or run_details """

        # Split file name to see if run details file
        run_split = file.split("\\")[-1]
        run_split = run_split.split(" ")[0]

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

    def check_file_formats(self):
        """ Check that required files have the expected formatting:
            Trending/QC Limits/IgG Curve/F093/F091/ """

        # Check trending

    def check_trending(self):
        """ Check that the trending file in required files page has the correct formatting """

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

    def set_progress_opaque(self):

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

        # If there are any non-visible Excel instances - kill upon Error
        if not xw.apps:
            return

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
