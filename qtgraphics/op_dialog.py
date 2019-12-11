import sys

from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
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
    """A Dialog to set input and output ranges for an optimization.

    With added stretch for a more consistent interface."""

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
        row_1.addStretch()
        row_1.addWidget(self.in_range)

        row_2 = QHBoxLayout()
        row_2.addWidget(QLabel("Output Cell:"))
        row_2.addStretch()
        row_2.addWidget(self.out_cell)

        row_3 = QHBoxLayout()
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        row_3.addWidget(self.buttonBox)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.buttonBox)

        layout = QVBoxLayout()
        layout.addLayout(row_1)
        layout.addLayout(row_2)
        layout.addLayout(row_3)
        self.setLayout(layout)


if __name__ == "__main__":
    app = get_qt_app()
    msgBox = OpDialog()
    result = msgBox.exec_()
    print("Input range:", msgBox.in_range.text())
    print("Output cell:", msgBox.out_cell.text())
    print("You clicked", "OK" if result else "Cancel")
