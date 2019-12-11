import sys

from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QLineEdit,
)


def get_qt_app():
    """Returns a QApplication instance.
    Must be called before showing any dialogs.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([sys.executable])
    return app


class OpDialog(QDialog):
    "A Dialog to set input and output ranges for an optimization."

    def __init__(self, *args, **kwargs):
        "Create a new dialogue instance."
        super().__init__(*args, **kwargs)
        self.setWindowTitle("Optimization Inputs and Output")
        self.gui_init()

    def gui_init(self):
        "Create and establish the widget layout."
        self.in_range = QLineEdit()
        self.out_cell = QLineEdit()

        row_1 = QHBoxLayout()
        row_1.addWidget(QLabel("Input range:"))
        row_1.addWidget(self.in_range)

        row_2 = QHBoxLayout()
        row_2.addWidget(QLabel("Output Cell:"))
        row_2.addWidget(self.out_cell)

        layout = QVBoxLayout()
        layout.addLayout(row_1)
        layout.addLayout(row_2)
        self.setLayout(layout)


if __name__ == "__main__":
    app = get_qt_app()
    msgBox = OpDialog()
    msgBox.exec_()
