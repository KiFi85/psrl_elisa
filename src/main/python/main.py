from fbs_runtime.application_context import cached_property, is_frozen
from fbs_runtime.application_context.PyQt5 import ApplicationContext
from fbs_runtime.excepthook.sentry import SentryExceptionHandler
from win32api import GetSystemMetrics
from pathlib import Path
import sys
from PyQt5 import QtWidgets
from PyQt5.QtGui import QPalette, QColor, QPixmap
from PyQt5.QtCore import QSettings, QByteArray, Qt, QThreadPool
from PyQt5.QtWidgets import QApplication, QMainWindow, QStackedWidget, QAction, QSizePolicy, QLineEdit, QStyleFactory, QCheckBox, \
    QComboBox, QSplashScreen
from final_master_page import PageFinalMaster
from elisa_data_page import PageData, Worker
from gantt_page import PageGantt
from settings_page import PageSettings
from help_page import PageHelp
import webbrowser
import sys
import time


class AppContext(ApplicationContext):

    def run(self):
        """ Run App Context """

        # Load main window
        self.main_window.show()
        return self.app.exec_()

    @cached_property
    def app(self):
        """ Overridden Application Context"""

        return MyCustomApp(sys.argv)

    @cached_property
    def main_window(self):
        """ Returns Main Window Class"""

        window = MainWindow(self)
        version = self.build_settings['version']
        app_name = self.build_settings['app_name']
        window.setWindowTitle(app_name + " v" + version)
        time.sleep(1.5)
        return window

    @cached_property
    def pdf_exe(self):
        return self.get_resource('wkhtmltopdf.exe')

    @cached_property
    def template(self):
        return self.get_resource('./templates/template.html')

    @cached_property
    def r4_template(self):
        return self.get_resource('./templates/r4template.html')

    @cached_property
    def css(self):
        return self.get_resource('./static/style.css')

    @cached_property
    def colour_map(self):
        return self.get_resource('gantt_color_map.csv')

    @cached_property
    def help(self):
        return self.get_resource('./html/help.html')

    """ Sentry exception handlers """
    @cached_property
    def exception_handlers(self):
        result = super().exception_handlers
        if is_frozen():
            result.append(self.sentry_exception_handler)
        return result

    @cached_property
    def sentry_exception_handler(self):
        return SentryExceptionHandler(
            self.build_settings['sentry_dsn'],
            self.build_settings['version'],
            self.build_settings['environment']
        )

    def get_pixel_dims(self):
        """ Calculate the optimal pixel dimensions for the splash screen """

        # Get screen resolution in pixels
        screen_width = GetSystemMetrics(0)

        pixels = screen_width / 4.5
        return pixels


class MyCustomApp(QApplication):
    """ Subclass the QApplication from ApplicationContext to set details for registry """

    def __init__(self, *args, **kwargs):
        super(MyCustomApp, self).__init__(*args, **kwargs)

        self.setOrganizationName("UCL-ICH")
        self.setApplicationName("psrl_elisa")


class MainWindow(QMainWindow):
    """ Main Window class to hold pages """

    def __init__(self, ctx):
        super(MainWindow, self).__init__()

        self.ctx = ctx  # Application context
        self.setStyle(QStyleFactory.create("Fusion"))

        x, y, w, h = get_geometry()  # Get optimal geometry
        self.setGeometry(x, y, w, h)

        min_h, min_w = get_minsize()  # Get minimum height and width for app

        self.setMinimumHeight(min_h)
        self.setMinimumWidth(min_w)
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        self.setAutoFillBackground(True)
        palette = self.palette()
        self.back_colour = QColor(141 * 0.5, 185 * 0.5, 202 * 0.5)
        palette.setColor(QPalette.Window, self.back_colour)
        self.setPalette(palette)


        # stacked - Main Data window with Tabs
        self.stacked = QStackedWidget()
        # Add pages to stacked widget
        self.stacked.addWidget(PageData(ctx=self.ctx, objectName="elisa_data"))
        self.stacked.addWidget(PageSettings(ctx=self.ctx, objectName="path_settings"))
        self.stacked.addWidget(PageGantt(ctx=self.ctx, objectName="gantt_page"))
        self.stacked.addWidget(PageHelp(ctx=self.ctx, objectName="help_page"))

        # Retrieve previous settings
        self.MASTER_PATH = self.get_saved_paths()
        self.get_saved_print_settings()

        self.stacked.addWidget(PageFinalMaster(ctx=self.ctx, master_path=self.MASTER_PATH, objectName="final_master"))


        # Create menu bar
        status_bar = self.statusBar()
        status_bar.setStyleSheet("QStatusBar{color:white;}")
        main_menu = self.menuBar()
        colour = QColor(217, 217, 217)
        main_menu.setStyleSheet("QWidget { background-color: %s}" % colour.name())

        """ Create menu"""
        file_menu = main_menu.addMenu("&File")
        data_menu = main_menu.addMenu("&Data")
        report_menu = main_menu.addMenu("&Reporting")
        settings_menu = main_menu.addMenu("&Settings")
        help_menu = main_menu.addMenu("&Help")

        """ Create Actions """
        exit_action = self.create_action("&Exit", self.exit_app, "Ctrl+Q", "Exit Application")

        data_action = self.create_action("&ELISA Data", lambda: self.stacked.setCurrentIndex(0),
                                         'Ctrl+E', "ELISA data processing main page")

        gantt_action = self.create_action("&Gantt Chart", lambda: self.stacked.setCurrentIndex(2),
                                          'Ctrl+R', "Create a testing Gantt chart")

        settings_action = self.create_action("&Change File Paths", lambda: self.stacked.setCurrentIndex(1),
                                             'Ctrl+T', "Edit paths to required files")

        help_action = self.create_action("&Help", self.show_help,
                                         'Ctrl+H', tip="Show help documentation")

        final_master_action = self.create_action("&Create Final Master File", lambda: self.stacked.setCurrentIndex(4),
                                              'Ctrl+F', "Generate a Final Master File")

        # Add actions
        self.add_actions(file_menu, [None, exit_action])
        self.add_actions(data_menu, [data_action, final_master_action])
        self.add_actions(report_menu, [gantt_action])
        self.add_actions(settings_menu, [settings_action])
        self.add_actions(help_menu, [help_action])

        # Add pages to central widget
        self.setCentralWidget(self.stacked)

    def create_action(self, text, function, shortcut=None, tip=None):
        """ Create an action for a file menu """

        # Create Action
        action = QAction(text, self)

        if shortcut:  # Assign shortcut if passed
            action.setShortcut(shortcut)
        if tip:  # Assign status tip and tool tip if passed
            action.setToolTip(tip)
            action.setStatusTip(tip)

        # Connect trigger to function
        action.triggered.connect(function)

        return action

    def add_actions(self, target, actions):
        """ Add a list of actions to a main menu item (target). If None, add separator"""

        for action in actions:
            if action:  # If an action has been passed - add to menu target
                target.addAction(action)
            else:  # Add separator if no action
                target.addSeparator()


    def show_help(self):
        """ Open help file in browser """

        # New tab in browser
        new = 2

        # Get help.html and open in browser
        helpfile = self.ctx.help.replace("\\", "/")
        url = "file://" + helpfile
        webbrowser.open(url, new=new)

    def get_saved_paths(self):
        """ Retrieve the file paths of the required files as saved from settings page """

        settings = QSettings()

        # Get settings page
        s_page = self.stacked.findChild(PageSettings, "path_settings")

        # List of widget names
        obj_names = ["qc_path", "curve_path", "trending_path", "f093_path", "master_path", "project_path"]

        # Loop through widget names, find widget in settings page and get saved text
        for obj_name in obj_names:
            saved_text = settings.value(obj_name) or ""
            s_page.findChild(QLineEdit, obj_name).setText(saved_text)

            if obj_name == "master_path":
                master_path = saved_text

        self.restoreGeometry(settings.value("MainWindow/Geometry", QByteArray()))

        return master_path



    def get_saved_print_settings(self):
        """ Get the settings for last options selected on ELISA data processing page """

        settings = QSettings()

        # Get settings page
        data_page = self.stacked.findChild(PageData, "elisa_data")

        # List of widget names
        checkbox_names = ["cb_upper_od", "cb_lower_od", "cb_lloq", "cb_print"]
        textbox_names = ["txt_upper_od", "txt_lower_od"]

        # Get combobox index
        combo_box = data_page.findChild(QComboBox, "combo_options")
        combo_idx = settings.value("combo_options", [], int)
        combo_box.setCurrentIndex(combo_idx) if combo_idx else combo_box.setCurrentIndex(0)

        # Loop through widget names, find widget in data page and get whether checked/enabled or not
        for obj_name in checkbox_names:
            widget = data_page.findChild(QCheckBox, obj_name)
            try:
                saved_bool = settings.value(obj_name, [], type=bool)
                widget.setChecked(saved_bool[0])
                widget.setEnabled(saved_bool[1])
            except:
                widget.setChecked(True)
                if not obj_name == "cb_print":
                    widget.setEnabled(False)

        # Get OD values in textboxes (if setting found - use if not, default)
        for obj_name in textbox_names:
            widget = data_page.findChild(QLineEdit, obj_name)
            try:
                saved_text = settings.value(obj_name, [], type=str)
                widget.setText(saved_text[0])
                text_bool = True if saved_text[1].upper() == "TRUE" else False
                widget.setEnabled(text_bool)
            except:
                # Set default ODs and disable
                widget.setText("2") if obj_name == "txt_upper_od" else widget.setText("0.1")
                widget.setEnabled(False)
                data_page.findChild(QComboBox, "combo_options").setCurrentIndex(0)

    def closeEvent(self, event):
        """ Override the close window event and save user settings to QSettings """

        # Save required file paths
        self.save_paths()

        # Save data settings
        self.save_print_settings()

        # Save F007 and MARS file paths for default browsing
        self.save_data_paths()

        sys.exit()

    def save_data_paths(self):
        """ Save the default directory for F007 and MARS file searching"""

        # Find ELISA Data main page
        settings = QSettings()
        data_page = self.stacked.findChild(PageData, "elisa_data")

        # Get f007 path (if not empty and save 2 dirs above as default search path
        f007_path = data_page.findChild(QLineEdit, "txt_f007").text()
        if f007_path and isinstance(f007_path, str):
            settings.setValue("f007_dir", Path(f007_path).parent)

        # If there are mars files saved as attribute on main Data Page - save parent dir
        if data_page.main.mars_files:
            try:
                mars_file = data_page.main.mars_files[0]
                mars_path = Path(mars_file).parent.parent
                settings.setValue("mars_dir", mars_path)
            except IndexError:
                return

    def save_paths(self):
        """ Save the file paths for required files """

        s_page = self.stacked.findChild(PageSettings, "path_settings")
        settings = QSettings()

        # List of widget names
        obj_names = ["qc_path", "curve_path", "trending_path", "f093_path", "master_path", "project_path"]

        # Loop through widget names, find widget in settings page, get text and save to settings
        for obj_name in obj_names:
            # Get file path from text based on object name
            path_text = s_page.findChild(QLineEdit, obj_name).text()
            # Save to QSettings
            settings.setValue(obj_name, path_text)

        settings.setValue("MainWindow/Geometry", self.saveGeometry())  # Save geometry

    def save_print_settings(self):
        """ Save the options checked after printing ELISA data """

        # Get the data page
        data_page = self.stacked.findChild(PageData, "elisa_data")
        settings = QSettings()

        # Checkbox values
        checkbox_names = ["cb_upper_od", "cb_lower_od", "cb_lloq", "cb_print"]
        textbox_names = ["txt_lower_od", "txt_upper_od"]

        # ComboBox
        combo = data_page.findChild(QComboBox, "combo_options")
        settings.setValue("combo_options", combo.currentIndex())

        # Loop through checkboxes in data page
        for name in checkbox_names:
            widget = data_page.findChild(QCheckBox, name)
            checked = widget.isChecked()  # Get whether checked or not
            enabled = widget.isEnabled()
            settings.setValue(name, [checked, enabled])  # Save checkbox value

        # Loop through textboxes
        for name in textbox_names:
            widget = data_page.findChild(QLineEdit, name)
            od_val = widget.text()
            enabled = widget.isEnabled()
            settings.setValue(name, [od_val, enabled])

    def exit_app(self):
        sys.exit()


def get_geometry():
    """ Get the screen resolution and return the optimal dimensions on startup """

    # Get screen resolution in pixels
    screen_width = GetSystemMetrics(0)
    screen_height = GetSystemMetrics(1)

    # Get geometry relative to 1920 x 1080 screen
    x = 450 / 1920 * screen_width
    y = 125 / 1080 * screen_height
    width = 700 / 1920 * screen_width
    height = 600 / 1080 * screen_height

    return x, y, width, height


def get_minsize():
    """ Get the minimum size for the application """

    # Get screen resolution in pixels
    screen_width = GetSystemMetrics(0)
    screen_height = GetSystemMetrics(1)

    min_height = 600 / 1920 * screen_width
    min_width = 800 / 1080 * screen_height

    return min_height, min_width


def splash(ctx):
    """ Display splash screen """

    img = ctx.get_resource('splash.png')  # Get splash image
    img_size = ctx.get_pixel_dims()  # Get optimal display size
    pixmap = QPixmap(img).scaled(img_size, img_size, Qt.KeepAspectRatio)  # Resize pixels
    splash = QSplashScreen(pixmap)  # Create splash screen
    splash.show()  # Show splash screen
    splash.finish(ctx.main_window)  # Wait for main window to be displayed


if __name__ == '__main__':
    appctxt = AppContext()       # 1. Instantiate ApplicationContext

    threadpool = QThreadPool()  # Start threadpool
    worker = Worker(splash(appctxt))  # Create splash screen
    threadpool.start(worker)  # Start worker

    exit_code = appctxt.run()      # 2. Invoke appctxt.app.exec_()
    sys.exit(exit_code)