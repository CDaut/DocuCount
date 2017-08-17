import sys
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow
from ui import Ui_root

#initiate qt window
app = QApplication(sys.argv)
window = QMainWindow()
ui = Ui_root(Pres=None)
ui.setupUi(window)
#show window
window.show()


sys.exit(app.exec_())
