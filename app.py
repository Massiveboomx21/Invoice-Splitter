
import sys
from PyQt5.QtWidgets import QApplication


try:
    from ui.main_window import MainWindow
except ImportError as e:
    print(f"Σφάλμα: Δεν ήταν δυνατή η εισαγωγή του 'MainWindow' από το 'ui.main_window'. {e}", file=sys.stderr)
    
    sys.exit(1)




def run_application():
    """
    Δημιουργεί και εκτελεί την κύρια εφαρμογή PyQt.
    """
    
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    main_window = MainWindow()
    
    main_window.show()    
    sys.exit(app.exec_())

if __name__ == '__main__':
    print("Αυτό το αρχείο προορίζεται να εισαχθεί από το main.py, όχι να εκτελεστεί απευθείας.")
    run_application()