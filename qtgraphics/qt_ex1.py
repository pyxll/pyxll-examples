import sys

from PyQt5.QtWidgets import (QApplication, QMessageBox)

def get_qt_app():
    """Returns a QApplication instance.
    Must be called before showing any dialogs.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([sys.executable])
    return app

if __name__ == '__main__':
    app = get_qt_app()
    msgBox = QMessageBox()
    msgBox.setText('This is a multi-line\nmessage for display\nto the user.')
    msgBox.setInformativeText("And here is a detailed explanation,"
                              " or at least the closest you will get.")
    msgBox.setWindowTitle('TITLE')
    msgBox.exec_() 

