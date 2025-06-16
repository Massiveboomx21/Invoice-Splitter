import sys
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = current_dir 
sys.path.insert(0, project_root)

try:
    from app import run_application
except ImportError as e:
    print(f"Σφάλμα: Δεν ήταν δυνατή η εισαγωγή του 'run_application' από το 'app.py'. {e}", file=sys.stderr)
    
    try:
        from PyQt5.QtWidgets import QMessageBox, QApplication
        
        temp_app = QApplication.instance() 
        if temp_app is None:
           temp_app = QApplication(sys.argv) 
        QMessageBox.critical(None, "Σφάλμα Εκκίνησης", f"Αδυναμία εύρεσης 'app.py'. Βεβαιωθείτε ότι το αρχείο υπάρχει.\n{e}")
    except ImportError:
        pass 
    sys.exit(1)

if __name__ == '__main__':
        
    run_application()