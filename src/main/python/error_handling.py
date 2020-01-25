import sys
import xlwings as xw
import traceback
import logging
from PyQt5 import QtCore, QtWidgets

# basic logger functionality
log = logging.getLogger(__name__)
handler = logging.StreamHandler(stream=sys.stdout)
log.addHandler(handler)


def show_exception_box(log_msg):
    """Checks if a QApplication instance is available and shows a messagebox with the exception message.
    If unavailable (non-console application), log an additional notice.
    """
    q_app = QtWidgets.QApplication.instance()
    if q_app is not None:

        # Get list of widgets from application instance
        widget_list = QtWidgets.QApplication.allWidgets()

        # Find error log and append error message
        for w in widget_list:
            if w.objectName() == "error_log":
                w.append("An unexpected error occurred:\n{0}".format(log_msg) + "\n\n")
            elif w.objectName() in ["btn_run", "btn_f007", "btn_mars", "grp_box"]:
                w.setEnabled(True)

        errorbox = QtWidgets.QMessageBox()
        errorbox.setIcon(QtWidgets.QMessageBox.Warning)
        errorbox.setWindowTitle("Error")
        errorbox.setText("An unexpected error occurred\nSee Error Log for more details")
        errorbox.exec_()
    else:
        log.debug("No QApplication instance available.")

    # If there are any non-visible Excel instances - kill upon Error
    app_list = []
    for a in xw.apps:
        if not a.visible:
            app_list.append(a)

    [app.kill() for app in app_list]



class UncaughtHook(QtCore.QObject):
    _exception_caught = QtCore.pyqtSignal(object)

    def __init__(self, *args, **kwargs):
        super(UncaughtHook, self).__init__(*args, **kwargs)

        # this registers the exception_hook() function as hook with the Python interpreter
        sys.excepthook = self.exception_hook

        # connect signal to execute the message box function always on main thread
        self._exception_caught.connect(show_exception_box)

    def exception_hook(self, exc_type, exc_value, exc_traceback):
        """Function handling uncaught exceptions.
        It is triggered each time an uncaught exception occurs.
        """
        if issubclass(exc_type, KeyboardInterrupt):
            # ignore keyboard interrupt to support console applications
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
        else:
            exc_info = (exc_type, exc_value, exc_traceback)
            log_msg = '\n'.join([''.join(traceback.format_tb(exc_traceback)),
                                 '{0}: {1}'.format(exc_type.__name__, exc_value)])
            log.critical("Uncaught exception:\n {0}".format(log_msg), exc_info=exc_info)

            # trigger message box show
            self._exception_caught.emit(log_msg)


class RangeNotFoundError(Exception):
    """ Custom exception when range is not found """
    def __init__(self, data):
        self.data = data

    def __str__(self):
        return repr(self.data)


# create a global instance of our class to register the hook
# def create_hook(text_box):
qt_exception_hook = UncaughtHook()