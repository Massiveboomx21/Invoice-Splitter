import os
import sys
import subprocess
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                             QPushButton, QLabel, QFileDialog, QLineEdit,
                             QSpinBox, QDoubleSpinBox, QProgressBar, QListWidget,
                             QGroupBox, QCheckBox, QMessageBox, QTabWidget,
                             QTextEdit, QSplitter, QApplication, QComboBox, QStyle,
                             QSizePolicy)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QSettings
from PyQt5.QtGui import QIcon, QFont

def resource_path(relative_path):
    """ Επιστρέφει τη σωστή διαδρομή για αρχεία πόρων (εικονίδια κλπ.),
        είτε τρέχει από κώδικα είτε από πακέτο PyInstaller. """
    try:
        
        base_path = sys._MEIPASS
    except Exception:
        
        base_path = os.path.abspath(".") 

    return os.path.join(base_path, relative_path)


SETTING_THRESHOLD = "settings/threshold"
SETTING_MAX_SPLIT = "settings/maxSplitValue"
SETTING_VALUE_COL = "settings/valueColumn"
SETTING_PROP_COL1 = "settings/propColumn1"
SETTING_PROP_COL2 = "settings/propColumn2"
SETTING_CREATE_BACKUP = "options/createBackup"
SETTING_OVERWRITE = "options/overwrite"
SETTING_OUTPUT_DIR = "paths/outputDir"
SETTING_INTEGER_SPLIT = "mode/integerSplit"  # Προσθήκη της σταθεράς
SETTING_AUTO_NUMBERING = "options/autoNumbering"  # Προσθήκη για αυτόματη αρίθμηση
SETTING_INVOICE_NUM_COL = "settings/invoiceNumColumn"  # Προσθήκη για στήλη αριθμού

try:
    
    
    from modules.excel_processor import ExcelProcessor
    from modules.logger import Logger 
    from modules.file_manager import FileManager
except ImportError as e:
     
     
     
     app = QApplication.instance()
     if app is None:
         app = QApplication(sys.argv)
     QMessageBox.critical(None, "Σφάλμα Εισαγωγής Module",
                          f"Αδυναμία εύρεσης ή εισαγωγής απαραίτητων modules.\n"
                          f"Βεβαιωθείτε ότι υπάρχει ο φάκελος 'modules' και περιέχει τα"
                          f" 'excel_processor.py', 'logger.py', 'file_manager.py'.\n\n"
                          f"Λεπτομέρειες σφάλματος: {e}")
     sys.exit(1) 


class WorkerThread(QThread):
    """Νήμα εργασίας για την επεξεργασία αρχείων σε παρασκήνιο"""
    progress_signal = pyqtSignal(int, int)
    file_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(dict)

    
    def __init__(self, processor, files, output_dir, threshold, value_col, prop_cols, overwrite, max_split_value, split_mode, auto_numbering=False, invoice_num_col=2): 
        super().__init__()
        self.processor = processor
        self.files = files
        self.output_dir = output_dir
        self.threshold = threshold
        self.value_col = value_col
        self.prop_cols = prop_cols
        self.overwrite = overwrite
        self.split_mode = split_mode
        self.max_split_value = max_split_value 
        self.auto_numbering = auto_numbering
        self.invoice_num_col = invoice_num_col 

    def run(self):
        results = {
            'total_files': len(self.files),
            'processed_files': 0,
            'skipped_files': 0,
            'total_rows_processed': 0,
            'total_rows_split': 0,
            'errors': 0,
            'skipped_impossible_splits': 0, 
            'multi_splits_performed': 0,    
            'file_results': {}
        }
        
        for i, input_file in enumerate(self.files):
             if self.isInterruptionRequested():
                 if self.processor and self.processor.logger:
                      self.processor.logger.warning("Η επεξεργασία διακόπηκε από τον χρήστη.")
                 break

             current_file_name = os.path.basename(input_file)
             self.file_signal.emit(current_file_name)
             self.progress_signal.emit(i, len(self.files))

             try:
                 file_name_base, file_ext = os.path.splitext(current_file_name)
                 output_name = f"{file_name_base}_διασπασμένο{file_ext}"
                 output_path = os.path.join(self.output_dir, output_name)
                 
                 file_results = self.processor.process_file(
                     input_file, output_path, self.threshold,
                     self.value_col, self.prop_cols, self.overwrite,
                     self.max_split_value, self.split_mode,
                     self.auto_numbering, self.invoice_num_col
                 )
                 if file_results.get('skipped'):
                     results['skipped_files'] += 1
                 else:
                     if file_results.get('errors', 0) == 0: 
                          results['processed_files'] += 1
                     results['total_rows_processed'] += file_results.get('processed_rows', 0)
                     results['total_rows_split'] += file_results.get('split_rows', 0)
                     
                     results['skipped_impossible_splits'] += file_results.get('skipped_impossible_splits', 0)
                     results['multi_splits_performed'] += file_results.get('multi_splits_performed', 0)

                 results['errors'] += file_results.get('errors', 0)
                 results['file_results'][current_file_name] = file_results

             except Exception as e:
                 results['errors'] += 1
                 error_message = f"Απρόσμενο σφάλμα στο WorkerThread κατά την επεξεργασία του {current_file_name}: {str(e)}"
                 if self.processor and self.processor.logger:
                     self.processor.logger.error(error_message, exc_info=True) 
                 else:
                      print(error_message)
                 results['file_results'][current_file_name] = {'error': True, 'message': error_message, 'skipped': False}

        
        self.progress_signal.emit(len(self.files), len(self.files))
        self.finished_signal.emit(results)

    def requestInterruption(self):
        super().requestInterruption()
        if self.processor and self.processor.logger:
            self.processor.logger.info("Ζητήθηκε διακοπή του νήματος επεξεργασίας.")
            

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.logger = Logger()
        self.processor = ExcelProcessor(self.logger)

        self.setWindowTitle("Invoice Splitter v1.3")
        self.setMinimumSize(850, 650)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        
        try:
            
            
            app_icon_path = resource_path(r"resources\ico\arrows_16382055.ico")

            if os.path.exists(app_icon_path):
                
                self.setWindowIcon(QIcon(app_icon_path))
                if self.logger: self.logger.debug(f"Εικονίδιο εφαρμογής φορτώθηκε: {app_icon_path}")
            else:
                if self.logger: self.logger.warning(f"Δεν βρέθηκε το αρχείο εικονιδίου: {app_icon_path}")

        except Exception as e:
             if self.logger: self.logger.error(f"Σφάλμα φόρτωσης εικονιδίου εφαρμογής: {e}")
        
        self.setMinimumSize(850, 650)
        
        self.create_ui() 

        
        
        self.logger.log_signal.connect(self.handle_log_message_for_ui)

        
        self.load_settings() 

        
        try:
            
            if not os.path.isdir(self.output_dir):
                 default_output = os.path.join(os.path.expanduser("~"), "Documents", "Díaspasména_Timológia")
                 self.logger.warning(f"Ο φάκελος εξόδου '{self.output_dir}' δεν βρέθηκε. Επαναφορά σε '{default_output}'")
                 self.output_dir = default_output
                 self.output_path_edit.setText(self.output_dir) 
            
            os.makedirs(self.output_dir, exist_ok=True)
        except OSError as e:
            self.logger.error(f"Αδυναμία δημιουργίας φακέλου εξόδου '{self.output_dir}': {e}")
            

        self.logger.info("Εκκίνηση εφαρμογής Διάσπασης Τιμολογίων")
        self.worker = None
    

    def create_ui(self):
        """Δημιουργία του γραφικού περιβάλλοντος"""

        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)

        self.split_tab = QWidget()
        self.tabs.addTab(self.split_tab, "Διάσπαση Αρχείων")
        self.settings_tab = QWidget()
        self.tabs.addTab(self.settings_tab, "Ρυθμίσεις")
        self.log_tab = QWidget()
        self.tabs.addTab(self.log_tab, "Καταγραφή")

        self.split_layout = QVBoxLayout(self.split_tab)
        self.settings_layout = QVBoxLayout(self.settings_tab)
        self.log_layout = QVBoxLayout(self.log_tab)
                
        self.file_group = QGroupBox("Αρχεία Excel προς Επεξεργασία")
        self.split_layout.addWidget(self.file_group)
        self.file_layout = QVBoxLayout(self.file_group)
        self.file_buttons_layout = QHBoxLayout()
        self.file_layout.addLayout(self.file_buttons_layout)
        self.select_file_btn = QPushButton(self.style().standardIcon(QStyle.SP_FileIcon), " Επιλογή Αρχείου...")
        self.select_file_btn.setIconSize(QSize(16, 16))
        self.select_file_btn.clicked.connect(self.select_file)
        self.select_file_btn.setToolTip("Επιλογή ενός αρχείου Excel (.xls, .xlsx, .xlsm)")
        self.file_buttons_layout.addWidget(self.select_file_btn)
        self.select_multiple_btn = QPushButton(self.style().standardIcon(QStyle.SP_DirIcon), " Επιλογή Πολλαπλών...")
        self.select_multiple_btn.setIconSize(QSize(16, 16))
        self.select_multiple_btn.clicked.connect(self.select_multiple_files)
        self.select_multiple_btn.setToolTip("Επιλογή πολλαπλών αρχείων Excel (.xls, .xlsx, .xlsm)")
        self.file_buttons_layout.addWidget(self.select_multiple_btn)
        self.file_buttons_layout.addStretch()
        self.clear_files_btn = QPushButton(self.style().standardIcon(QStyle.SP_TrashIcon), " Καθαρισμός Λίστας")
        self.clear_files_btn.setIconSize(QSize(16, 16))
        self.clear_files_btn.clicked.connect(self.clear_files)
        self.clear_files_btn.setToolTip("Αφαίρεση όλων των αρχείων από τη λίστα")
        self.file_buttons_layout.addWidget(self.clear_files_btn)
        self.file_list = QListWidget()
        self.file_list.setToolTip("Λίστα των αρχείων που θα επεξεργαστούν")
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_layout.addWidget(self.file_list)

        self.output_group = QGroupBox("Φάκελος Αποθήκευσης Επεξεργασμένων Αρχείων")
        self.split_layout.addWidget(self.output_group)
        self.output_layout = QHBoxLayout(self.output_group)
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setReadOnly(True)
        self.output_path_edit.setToolTip("Ο φάκελος όπου θα αποθηκευτούν τα νέα αρχεία")
        self.output_layout.addWidget(self.output_path_edit)
        self.select_output_btn = QPushButton(self.style().standardIcon(QStyle.SP_DirOpenIcon), " Επιλογή...")
        self.select_output_btn.setIconSize(QSize(16, 16))
        self.select_output_btn.clicked.connect(self.select_output_dir)
        self.select_output_btn.setToolTip("Επιλογή του φακέλου αποθήκευσης")
        self.output_layout.addWidget(self.select_output_btn)

        self.execution_group = QGroupBox("Εκτέλεση Επεξεργασίας")
        self.split_layout.addWidget(self.execution_group)
        self.execution_layout = QVBoxLayout(self.execution_group)
        self.progress_label = QLabel("Έτοιμο για επεξεργασία.")
        self.execution_layout.addWidget(self.progress_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.execution_layout.addWidget(self.progress_bar)
        self.process_btn = QPushButton(self.style().standardIcon(QStyle.SP_MediaPlay), " Έναρξη Επεξεργασίας")
        self.process_btn.setIconSize(QSize(24, 24))
        font = self.process_btn.font()
        font.setPointSize(11)
        self.process_btn.setFont(font)
        self.process_btn.clicked.connect(self.start_processing)
        self.process_btn.setToolTip("Ξεκινά τη διαδικασία διάσπασης για τα επιλεγμένα αρχεία")
        self.execution_layout.addWidget(self.process_btn)

        self.params_group = QGroupBox("Παράμετροι Διάσπασης")
        self.settings_layout.addWidget(self.params_group)
        self.params_layout = QVBoxLayout(self.params_group)
        
        self.threshold_layout = QHBoxLayout()
        self.threshold_label = QLabel("Όριο ποσού για διάσπαση (€):")
        self.threshold_layout.addWidget(self.threshold_label)
        self.threshold_spinbox = QDoubleSpinBox() 
        self.threshold_spinbox.setRange(0.01, 10000000.00)
        self.threshold_spinbox.setDecimals(2)
        self.threshold_spinbox.setSingleStep(50.00)
        self.threshold_spinbox.setSuffix(" €")
        self.threshold_spinbox.setToolTip("Η εγγραφή θα διασπαστεί αν η αξία στη 'Στήλη Αξίας' είναι ίση ή μεγαλύτερη από αυτό το ποσό")
        self.threshold_spinbox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.threshold_layout.addWidget(self.threshold_spinbox) 
        self.params_layout.addLayout(self.threshold_layout) 
        
        self.max_split_layout = QHBoxLayout()
        self.max_split_label = QLabel("Μέγιστη τιμή μετά τη διάσπαση (€):")
        self.max_split_layout.addWidget(self.max_split_label)
        self.max_split_value_spinbox = QDoubleSpinBox() 
        self.max_split_value_spinbox.setRange(0.01, 10000000.00)
        self.max_split_value_spinbox.setDecimals(2)
        self.max_split_value_spinbox.setSingleStep(50.00)
        self.max_split_value_spinbox.setSuffix(" €")
        self.max_split_value_spinbox.setToolTip("Κάθε μέρος της διάσπασης πρέπει να είναι ΜΙΚΡΟΤΕΡΟ από αυτή την τιμή.")
        self.max_split_value_spinbox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.max_split_layout.addWidget(self.max_split_value_spinbox) 
        self.params_layout.addLayout(self.max_split_layout) 
        
        # Ενεργοποίηση του checkbox για ακέραια διάσπαση
        self.integer_split_check = QCheckBox("Διάσπαση μόνο σε ακέραια ποσά (που λήγουν σε 0 ή 5)")
        self.integer_split_check.setToolTip("Αν επιλεγεί, η διάσπαση θα παράγει μόνο ακέραιους αριθμούς πολλαπλάσια του 5.\nΠΡΟΣΟΧΗ: Τα αρχικά ποσά πρέπει να είναι ήδη πολλαπλάσια του 5.")
        self.params_layout.addWidget(self.integer_split_check)
        
        # Checkbox για αυτόματη αρίθμηση
        self.auto_numbering_check = QCheckBox("Αυτόματη αύξουσα αρίθμηση διασπασμένων τιμολογίων")
        self.auto_numbering_check.setToolTip("Αν επιλεγεί, τα διασπασμένα τιμολόγια θα παίρνουν αύξοντες αριθμούς με βάση το προηγούμενο τιμολόγιο")
        self.params_layout.addWidget(self.auto_numbering_check)
        
        
        self.columns_group = QGroupBox("Αριθμοί Στηλών (π.χ., 1=A, 2=B, 6=F)")
        self.params_layout.addWidget(self.columns_group)
        self.columns_layout_internal = QGridLayout(self.columns_group) 

        
        self.value_col_label = QLabel("Στήλη Βασικής Αξίας:")
        self.columns_layout_internal.addWidget(self.value_col_label, 0, 0)
        self.value_col_spinbox = QSpinBox()
        self.value_col_spinbox.setRange(1, 200)
        self.value_col_spinbox.setToolTip("Ο αριθμός της στήλης με την κύρια αξία (π.χ., 6 για F)")
        self.columns_layout_internal.addWidget(self.value_col_spinbox, 0, 1) 

        
        self.prop_col_label = QLabel("Στήλες Αναλογικής Διάσπασης:")
        self.columns_layout_internal.addWidget(self.prop_col_label, 1, 0, 1, 2) 

        
        self.prop_col1_label = QLabel("Στήλη 1:")
        self.columns_layout_internal.addWidget(self.prop_col1_label, 2, 0)
        self.prop_col1_spinbox = QSpinBox() 
        self.prop_col1_spinbox.setRange(1, 200)
        self.prop_col1_spinbox.setToolTip("Η πρώτη αναλογική στήλη (π.χ., 8 για H)")
        self.columns_layout_internal.addWidget(self.prop_col1_spinbox, 2, 1) 

        
        self.prop_col2_label = QLabel("Στήλη 2:")
        self.columns_layout_internal.addWidget(self.prop_col2_label, 3, 0)
        self.prop_col2_spinbox = QSpinBox() 
        self.prop_col2_spinbox.setRange(1, 200)
        self.prop_col2_spinbox.setToolTip("Η δεύτερη αναλογική στήλη (π.χ., 19 για S)")
        self.columns_layout_internal.addWidget(self.prop_col2_spinbox, 3, 1) 
        
        # Προσθήκη πεδίου για στήλη αριθμού τιμολογίου
        self.invoice_num_label = QLabel("Στήλη Αριθμού Τιμολογίου:")
        self.columns_layout_internal.addWidget(self.invoice_num_label, 4, 0)
        self.invoice_num_spinbox = QSpinBox()
        self.invoice_num_spinbox.setRange(1, 200)
        self.invoice_num_spinbox.setValue(2)  # Default στήλη B
        self.invoice_num_spinbox.setToolTip("Η στήλη με τους αριθμούς τιμολογίων (π.χ., 2 για B)")
        self.columns_layout_internal.addWidget(self.invoice_num_spinbox, 4, 1) 
        
        self.options_group = QGroupBox("Επιλογές Επεξεργασίας")
        self.settings_layout.addWidget(self.options_group)
        self.options_layout = QVBoxLayout(self.options_group)
        self.create_backup_check = QCheckBox("Δημιουργία αντιγράφου ασφαλείας (.backup)")
        self.create_backup_check.setChecked(True)
        self.create_backup_check.setToolTip("Αν επιλεγεί, δημιουργείται αντίγραφο του αρχικού αρχείου πριν τροποποιηθεί")
        self.options_layout.addWidget(self.create_backup_check)
        self.overwrite_check = QCheckBox("Αντικατάσταση αρχείου εξόδου αν υπάρχει ήδη")
        self.overwrite_check.setChecked(True)
        self.overwrite_check.setToolTip("Αν επιλεγεί, τυχόν υπάρχον αρχείο εξόδου θα αντικατασταθεί.\nΑλλιώς, η επεξεργασία για αυτό το αρχείο θα παραλειφθεί.")
        self.options_layout.addWidget(self.overwrite_check)
        
        self.log_group = QGroupBox("Ιστορικό Ενεργειών και Σφαλμάτων")
        self.log_layout.addWidget(self.log_group)
        self.log_group_layout = QVBoxLayout(self.log_group)
        self.log_editor = QTextEdit() 
        self.log_editor.setReadOnly(True)
        self.log_editor.setFont(QFont("Courier New", 9))
        self.log_editor.setToolTip("Εμφανίζει τα μηνύματα κατά την εκτέλεση της εφαρμογής")
        self.log_group_layout.addWidget(self.log_editor)
        self.log_buttons_layout = QHBoxLayout()
        self.log_group_layout.addLayout(self.log_buttons_layout)
        self.log_buttons_layout.addStretch()
        self.clear_log_btn = QPushButton(self.style().standardIcon(QStyle.SP_DialogResetButton), " Καθαρισμός Οθόνης")
        self.clear_log_btn.setIconSize(QSize(16, 16))
        self.clear_log_btn.clicked.connect(self.clear_log)
        self.clear_log_btn.setToolTip("Καθαρίζει τα μηνύματα από αυτή την οθόνη")
        self.log_buttons_layout.addWidget(self.clear_log_btn)
        self.open_log_file_btn = QPushButton(self.style().standardIcon(QStyle.SP_FileLinkIcon), " Άνοιγμα Αρχείου Log")
        self.open_log_file_btn.setIconSize(QSize(16, 16))
        self.open_log_file_btn.clicked.connect(self.open_log_file)
        self.log_buttons_layout.addWidget(self.open_log_file_btn)

        self.split_layout.addStretch()
        self.settings_layout.addStretch()
        
        try:
            log_file_path = self.logger.get_log_file()
            self.open_log_file_btn.setToolTip(f"Ανοίγει το αρχείο καταγραφής ({log_file_path})")
        except Exception:
            self.open_log_file_btn.setToolTip("Άνοιγμα του τρέχοντος αρχείου καταγραφής")

    

    def select_file(self):
        """Επιλογή ενός αρχείου Excel"""
        last_dir = self.output_dir if os.path.isdir(self.output_dir) else ""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Επιλογή Αρχείου Excel", last_dir, "Αρχεία Excel (*.xls *.xlsx *.xlsm);;Όλα τα αρχεία (*.*)"
        )
        if file_path:
            if not FileManager.validate_excel_file(file_path):
                QMessageBox.warning(self, "Μη έγκυρο αρχείο", f"Το αρχείο '{os.path.basename(file_path)}' δεν φαίνεται να είναι υποστηριζόμενο αρχείο Excel.")
                self.logger.warning(f"Απορρίφθηκε μη έγκυρο αρχείο: {file_path}")
                return
            items = [self.file_list.item(i).text() for i in range(self.file_list.count())]
            if file_path not in items:
                self.file_list.addItem(file_path)
                self.logger.info(f"Προστέθηκε αρχείο: {file_path}")
            else:
                 QMessageBox.information(self, "Αρχείο υπάρχει ήδη", f"Το αρχείο '{os.path.basename(file_path)}' υπάρχει ήδη στη λίστα.")
                 self.logger.info(f"Το αρχείο {file_path} υπάρχει ήδη στη λίστα, δεν προστέθηκε ξανά.")

    def select_multiple_files(self):
        """Επιλογή πολλαπλών αρχείων Excel"""
        last_dir = self.output_dir if os.path.isdir(self.output_dir) else ""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Επιλογή Πολλαπλών Αρχείων Excel", last_dir, "Αρχεία Excel (*.xls *.xlsx *.xlsm);;Όλα τα αρχεία (*.*)"
        )
        if file_paths:
            added_count = 0
            skipped_count = 0
            invalid_count = 0
            existing_items = [self.file_list.item(i).text() for i in range(self.file_list.count())]
            for file_path in file_paths:
                if not FileManager.validate_excel_file(file_path):
                    self.logger.warning(f"Παράλειψη μη έγκυρου αρχείου: {file_path}")
                    invalid_count += 1
                    continue
                if file_path in existing_items:
                    skipped_count += 1
                    continue
                self.file_list.addItem(file_path)
                existing_items.append(file_path)
                added_count += 1
            if added_count > 0: self.logger.info(f"Προστέθηκαν {added_count} νέα αρχεία στη λίστα.")
            if skipped_count > 0: self.logger.info(f"Παραλείφθηκαν {skipped_count} αρχεία που υπήρχαν ήδη στη λίστα.")
            if invalid_count > 0:
                 QMessageBox.warning(self,"Μη έγκυρα αρχεία", f"Παραλείφθηκαν {invalid_count} μη υποστηριζόμενα αρχεία.")

    def clear_files(self):
        """Καθαρισμός της λίστας αρχείων"""
        if self.file_list.count() > 0:
            reply = QMessageBox.question(self, 'Επιβεβαίωση Καθαρισμού',
                                         "Να αφαιρεθούν όλα τα αρχεία από τη λίστα;",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.file_list.clear()
                self.logger.info("Καθαρίστηκε η λίστα αρχείων.")
        else:
             self.logger.info("Προσπάθεια καθαρισμού άδειας λίστας αρχείων.")

    def select_output_dir(self):
        """Επιλογή φακέλου αποθήκευσης"""
        output_dir = QFileDialog.getExistingDirectory(self, "Επιλογή Φακέλου Αποθήκευσης", self.output_dir)
        if output_dir:
            if not os.access(output_dir, os.W_OK):
                 QMessageBox.warning(self, "Σφάλμα Δικαιωμάτων", f"Δεν υπάρχουν δικαιώματα εγγραφής στον φάκελο:\n{output_dir}")
                 self.logger.error(f"Αποτυχία επιλογής φακέλου εξόδου (δικαιώματα): {output_dir}")
                 return
            self.output_dir = output_dir
            self.output_path_edit.setText(output_dir)
            self.logger.info(f"Επιλέχθηκε νέος φάκελος αποθήκευσης: {output_dir}")
     
    def handle_log_message_for_ui(self, formatted_message):
        """ Φιλτράρει τα μηνύματα για το UI Log. """
        
        show_keywords = [
            
            
            "--- Έναρξη Επεξεργασίας Αρχείων ---",
            "--- Η επεξεργασία ολοκληρώθηκε ---",
            "--- Αποτελέσματα Επεξεργασίας ---",
            "Σύνολο αρχείων προς επεξεργασία:", 
            "Αρχεία που επεξεργάστηκαν:",   
            "Αρχεία που παραλείφθηκαν :",  
            "Σύνολο γραμμών που διασπάστηκαν:", 
            "Σύνολο αποτυχημένων διασπάσεων:", 
            "Αποτυχημένες διασπάσεις:",      
            
            "Ξεκινά η διαδικασία επεξεργασίας", 
            "Φάκελος Εξόδου:",               
            "Όριο Διάσπασης:",               
            "Μέγιστη Τιμή Μετά:",            
            "Στήλη Αξίας:",                  
            "Στήλες Αναλ/κές:",              
            "Backup:",                       
            "Overwrite:",                    
            "Split Mode:",                   # Προσθήκη για να φαίνεται το mode
            "Οι ρυθμίσεις φορτώθηκαν",
            "Οι ρυθμίσεις αποθηκεύτηκαν",
            "Εκκίνηση εφαρμογής",
            "Κλείσιμο εφαρμογής",
                        
        ]
        
        try:
            if any(keyword in formatted_message for keyword in show_keywords):
                self.log_editor.append(formatted_message)
            
            elif "[ERROR] Αποτυχία generate_n_splits" in formatted_message:
                pass 
            elif "[DEBUG]" in formatted_message:
                pass 
        except Exception as e:
            print(f"ERROR in handle_log_message_for_ui: {e}. Message was: {formatted_message}")

    
    def clear_log(self):
        """Καθαρισμός του πεδίου καταγραφής στην οθόνη"""
        self.log_editor.clear()

    def open_log_file(self):
        """Άνοιγμα του αρχείου καταγραφής"""
        try:
            log_file = self.logger.get_log_file()
            if log_file and os.path.exists(log_file):
                self.logger.info(f"Προσπάθεια ανοίγματος αρχείου log: {log_file}")
                if sys.platform == 'win32': os.startfile(log_file)
                elif sys.platform == 'darwin': subprocess.call(('open', log_file))
                else: subprocess.call(('xdg-open', log_file))
            else:
                QMessageBox.warning(self, "Σφάλμα", f"Το αρχείο καταγραφής δεν βρέθηκε: {log_file}")
                self.logger.error(f"Αποτυχία ανοίγματος αρχείου log - Δεν βρέθηκε: {log_file}")
        except AttributeError:
             QMessageBox.critical(self, "Σφάλμα Logger", "Αδυναμία λήψης διαδρομής log.")
             print("Σφάλμα: Δεν βρέθηκε η μέθοδος get_log_file() στον logger.")
        except Exception as e:
             QMessageBox.critical(self, "Σφάλμα Ανοίγματος", f"Σφάλμα κατά το άνοιγμα του log:\n{e}")
             self.logger.error(f"Σφάλμα κατά το άνοιγμα του αρχείου log: {str(e)}")

    
    def start_processing(self):
        """Έναρξη της διαδικασίας επεξεργασίας των αρχείων"""

        
        if self.worker is not None and self.worker.isRunning():
             QMessageBox.information(self,"Επεξεργασία σε Εξέλιξη", "Μια διαδικασία επεξεργασίας είναι ήδη σε εξέλιξη. Παρακαλώ περιμένετε να ολοκληρωθεί.")
             return

        
        if self.file_list.count() == 0:
            QMessageBox.warning(self, "Δεν Επιλέχθηκαν Αρχεία", "Παρακαλώ επιλέξτε τουλάχιστον ένα αρχείο Excel από την καρτέλα 'Διάσπαση Αρχείων'.")
            self.tabs.setCurrentWidget(self.split_tab) 
            return

        if not self.output_dir or not os.path.isdir(self.output_dir):
            QMessageBox.warning(self, "Μη Έγκυρος Φάκελος Εξόδου", f"Ο φάκελος αποθήκευσης '{self.output_dir}' δεν είναι έγκυρος. Παρακαλώ επιλέξτε έναν έγκυρο φάκελο.")
            self.tabs.setCurrentWidget(self.split_tab) 
            self.select_output_dir() 
            return

        if not os.access(self.output_dir, os.W_OK):
             QMessageBox.warning(self, "Σφάλμα Δικαιωμάτων Φακέλου Εξόδου", f"Δεν υπάρχουν δικαιώματα εγγραφής στον φάκελο εξόδου:\n{self.output_dir}\n\nΠαρακαλώ επιλέξτε έναν άλλο φάκελο.")
             self.tabs.setCurrentWidget(self.split_tab)
             return

        files_to_process = [self.file_list.item(i).text() for i in range(self.file_list.count())]

        
        threshold = self.threshold_spinbox.value()
        value_col = self.value_col_spinbox.value()
        prop_col1 = self.prop_col1_spinbox.value()
        prop_col2 = self.prop_col2_spinbox.value()
        prop_cols = list(dict.fromkeys([prop_col1, prop_col2])) 
        max_split_value = self.max_split_value_spinbox.value()
        if max_split_value < 0.01:
            QMessageBox.warning(self, "Μη Έγκυρη Ρύθμιση", "Η 'Μέγιστη τιμή μετά τη διάσπαση' πρέπει να είναι τουλάχιστον 0.01€.")
            return
        create_backup = self.create_backup_check.isChecked()
        overwrite = self.overwrite_check.isChecked()
        
        use_integer_split = self.integer_split_check.isChecked()
        split_mode = 'integer_5' if use_integer_split else 'decimal'
        
        auto_numbering = self.auto_numbering_check.isChecked()
        invoice_num_col = self.invoice_num_spinbox.value()
        
        self.logger.info("="*40)
        self.logger.info(f"Ξεκινά η διαδικασία επεξεργασίας για {len(files_to_process)} αρχεία.")
        self.logger.info(f"  Φάκελος Εξόδου: {self.output_dir}")
        self.logger.info(f"  Όριο Διάσπασης: {threshold:.2f} €") 
        self.logger.info(f"  Μέγιστη Τιμή Μετά: {max_split_value:.2f} €") 
        self.logger.info(f"  Στήλη Αξίας: {value_col}") 
        self.logger.info(f"  Στήλες Αναλ/κές: {prop_cols}") 
        self.logger.info(f"  Backup: {'Ναι' if create_backup else 'Όχι'}") 
        self.logger.info(f"  Overwrite: {'Ναι' if overwrite else 'Όχι'}") 
        self.logger.info(f"  Split Mode: {'Ακέραια (x5)' if split_mode == 'integer_5' else 'Δεκαδικά'}")
        self.logger.info(f"  Αυτόματη Αρίθμηση: {'Ναι' if auto_numbering else 'Όχι'}")
        if auto_numbering:
            self.logger.info(f"  Στήλη Αριθμού: {invoice_num_col}")
        self.logger.info("="*40)

        self.tabs.setCurrentWidget(self.log_tab) 

        
        if create_backup:
            self.logger.info("Έναρξη δημιουργίας αντιγράφων ασφαλείας...")
            self.progress_label.setText("Δημιουργία αντιγράφων ασφαλείας...")
            QApplication.processEvents() 
            backups_ok = True
            for idx, file_path in enumerate(files_to_process):
                self.progress_bar.setValue(int(((idx + 1) / len(files_to_process)) * 100))
                try:
                    backup_path = FileManager.create_backup(file_path)
                    self.logger.info(f"OK Backup: {os.path.basename(backup_path)}")
                except Exception as e:
                    self.logger.error(f"Κρίσιμο σφάλμα κατά τη δημιουργία backup για το '{os.path.basename(file_path)}': {str(e)}")
                    QMessageBox.critical(self, "Σφάλμα Backup", f"Αδυναμία δημιουργίας αντιγράφου ασφαλείας για το αρχείο:\n{file_path}\n\nΣφάλμα: {str(e)}\n\nΗ επεξεργασία θα ακυρωθεί.")
                    backups_ok = False
                    break
            self.progress_bar.setValue(0) 
            if not backups_ok:
                 self.logger.error("Η επεξεργασία ακυρώθηκε λόγω αποτυχίας δημιουργίας backup.")
                 self.progress_label.setText("Η επεξεργασία ακυρώθηκε (σφάλμα backup).")
                 return 
            self.logger.info("Ολοκληρώθηκε η δημιουργία αντιγράφων ασφαλείας.")

        
        self.set_ui_enabled(False)
        self.progress_bar.setValue(0)
        self.progress_label.setText("Προετοιμασία επεξεργασίας...")
        self.progress_label.setStyleSheet("") 
        self.progress_bar.setFormat("%p%") 
        self.process_btn.setText(" Επεξεργασία σε Εξέλιξη...")

        
        self.worker = WorkerThread(
            self.processor, files_to_process, self.output_dir,
            threshold, value_col, prop_cols, overwrite,
            max_split_value, split_mode, auto_numbering, invoice_num_col
        )

        
        
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.file_signal.connect(self.update_file_label)
        self.worker.finished_signal.connect(self.processing_finished)
        self.worker.finished.connect(self.worker.deleteLater) 

        self.worker.start() 
    

    def update_progress(self, current_step, total_steps):
        """Ενημερώνει ΜΟΝΟ την μπάρα προόδου (τιμή και κείμενο)."""
        if total_steps > 0:
            
            
            percentage = int(((current_step + 1) / total_steps) * 100)
            self.progress_bar.setValue(percentage)
            
            self.progress_bar.setFormat(f"{current_step + 1}/{total_steps} (%p%)")
        else:
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("%p%")

    
    def update_file_label(self, file_name):
        """Ενημερώνει την ετικέτα κειμένου με το τρέχον αρχείο."""
        
        status_text = f"Επεξεργασία: {file_name}"
        self.progress_label.setStyleSheet("") 
        self.progress_label.setText(status_text)

    
    
    def processing_finished(self, results):
        """Καλείται όταν το νήμα επεξεργασίας ολοκληρώσει"""
        self.logger.info("--- Η επεξεργασία ολοκληρώθηκε ---") 

        
        processed_ok = results.get('processed_files', 0)
        skipped = results.get('skipped_files', 0)
        errors = results.get('errors', 0)
        split_rows = results.get('total_rows_split', 0)
        skipped_impossible = results.get('skipped_impossible_splits', 0)
        multi_splits = results.get('multi_splits_performed', 0)

        
        self.logger.info("\n--- Αποτελέσματα Επεξεργασίας ---")
        self.logger.info(f"Σύνολο αρχείων προς επεξεργασία: {results.get('total_files', 0)}")
        self.logger.info(f"Αρχεία που επεξεργάστηκαν: {processed_ok}")
        self.logger.info(f"Αρχεία που παραλείφθηκαν : {skipped}")
        self.logger.info(f"Σύνολο γραμμών που διασπάστηκαν: {split_rows}")
        if multi_splits > 0:
             self.logger.info(f"  (Εκ των οποίων {multi_splits} διασπάστηκαν σε >2 μέρη)")
        self.logger.info(f"Σύνολο αποτυχημένων διασπάσεων: {skipped_impossible}")

        failed_splits_details = []
        for file_name, file_res in results.get('file_results', {}).items():
            if file_res and 'skipped_details' in file_res:
                 failed_splits_details.extend(file_res['skipped_details'])

        if failed_splits_details:
            self.logger.info("Αποτυχημένες διασπάσεις:") 
            unique_failures = set()
            for failure in failed_splits_details:
                 fail_str = f"- Αρχείο: {failure['file']}, Φύλλο: {failure['sheet']}, Γραμμή: {failure['row']}, Ποσό: {failure['value']}"
                 if fail_str not in unique_failures:
                      self.logger.info(fail_str) 
                      unique_failures.add(fail_str)
        elif errors > 0 : 
             self.logger.info(f"Σύνολο σφαλμάτων κατά την επεξεργασία: {errors}")
        self.logger.info("---------------------------------")
        

        
        final_message = f"Ολοκληρώθηκε ({processed_ok} επεξεργ., {skipped} παραλ., {errors} σφάλμ., {skipped_impossible} αδύν. διασπ.)."
        style_sheet = "" 
        if errors > 0:
             style_sheet = "color: red; font-weight: bold;"; final_message = f"Ολοκληρώθηκε με {errors} ΣΦΑΛΜΑΤΑ."
        elif skipped > 0 and processed_ok == 0 and skipped_impossible == 0:
             style_sheet = "color: orange;"; final_message = f"Ολοκληρώθηκε. Όλα τα αρχεία ({skipped}) παραλείφθηκαν."
        elif skipped > 0 or skipped_impossible > 0:
             style_sheet = "color: orange;"
        else: style_sheet = "color: green; font-weight: bold;" 

        
        QApplication.processEvents()

        
        self.set_ui_enabled(True)
        self.worker = None 

        
        msg_title = "Ολοκλήρωση Επεξεργασίας"
        msg_details = (f"Επεξεργάστηκαν: {processed_ok}\n"
                       f"Παραλείφθηκαν: {skipped}\n"
                       f"Αδύνατες Διασπάσεις: {skipped_impossible}\n"
                       f"Σφάλματα: {errors}\n\n"
                       f"Τα νέα αρχεία βρίσκονται:\n{self.output_dir}")
        if errors > 0: QMessageBox.warning(self, msg_title, f"Ολοκλήρωση με {errors} σφάλματα.\n{msg_details}\n\nΕλέγξτε την καρτέλα 'Καταγραφή'.")
        elif skipped > 0 and processed_ok == 0 and skipped_impossible == 0: QMessageBox.information(self, msg_title, f"Όλα τα αρχεία παραλείφθηκαν.\n{msg_details}")
        else: QMessageBox.information(self, msg_title, f"Η επεξεργασία ολοκληρώθηκε.\n{msg_details}")

        
        self.set_ui_enabled(True)
        self.worker = None

        
        self.progress_label.setStyleSheet("")
        self.progress_label.setText("Έτοιμο για νέα επεξεργασία.")
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        
        self.process_btn.setText(" Έναρξη Επεξεργασίας")
        self.process_btn.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay)) 

    def set_ui_enabled(self, enabled):
        """Ενεργοποιεί/απενεργοποιεί στοιχεία του UI"""
        self.split_tab.setEnabled(enabled)
        self.settings_tab.setEnabled(enabled)
    
    

     
    def load_settings(self):
        settings = QSettings("MyCompanyName", "InvoiceSplitter") 
        if self.logger: self.logger.debug("--- Φόρτωση Ρυθμίσεων ---")
        default_output = os.path.join(os.path.expanduser("~"), "Documents", "Díaspasména_Timológia")
        try:
            threshold = settings.value(SETTING_THRESHOLD, 500.0, type=float)
            max_split = settings.value(SETTING_MAX_SPLIT, threshold, type=float)
            val_col = settings.value(SETTING_VALUE_COL, 6, type=int)
            prop1 = settings.value(SETTING_PROP_COL1, 8, type=int)
            prop2 = settings.value(SETTING_PROP_COL2, 19, type=int)
            backup = settings.value(SETTING_CREATE_BACKUP, True, type=bool)
            overwrite = settings.value(SETTING_OVERWRITE, False, type=bool)
            output_dir = settings.value(SETTING_OUTPUT_DIR, default_output, type=str)
            integer_split = settings.value(SETTING_INTEGER_SPLIT, False, type=bool)
            auto_numbering = settings.value(SETTING_AUTO_NUMBERING, False, type=bool)
            invoice_num_col = settings.value(SETTING_INVOICE_NUM_COL, 2, type=int)


            if not os.path.isdir(output_dir): output_dir = default_output
            
            self.integer_split_check.setChecked(integer_split)
            self.auto_numbering_check.setChecked(auto_numbering)
            self.invoice_num_spinbox.setValue(invoice_num_col)
            self.threshold_spinbox.setValue(threshold)
            self.max_split_value_spinbox.setValue(max_split)
            self.value_col_spinbox.setValue(val_col)
            self.prop_col1_spinbox.setValue(prop1)
            self.prop_col2_spinbox.setValue(prop2)
            self.create_backup_check.setChecked(backup)
            self.overwrite_check.setChecked(overwrite)
            self.output_dir = output_dir
            self.output_path_edit.setText(output_dir)

            if self.logger: self.logger.info("Οι ρυθμίσεις φορτώθηκαν.")
        except Exception as e:
            if self.logger: self.logger.error(f"Σφάλμα φόρτωσης ρυθμίσεων: {e}", exc_info=True)
            
            self.threshold_spinbox.setValue(500.0); self.max_split_value_spinbox.setValue(500.0)
            self.value_col_spinbox.setValue(6); self.prop_col1_spinbox.setValue(8); self.prop_col2_spinbox.setValue(19)
            self.create_backup_check.setChecked(True); self.overwrite_check.setChecked(False)
            self.output_dir = default_output; self.output_path_edit.setText(default_output)


    
    def save_settings(self):
        settings = QSettings("MyCompanyName", "InvoiceSplitter")
        if self.logger: self.logger.debug("--- Αποθήκευση Ρυθμίσεων ---")
        try:
            settings.setValue(SETTING_INTEGER_SPLIT, self.integer_split_check.isChecked())
            settings.setValue(SETTING_AUTO_NUMBERING, self.auto_numbering_check.isChecked())
            settings.setValue(SETTING_INVOICE_NUM_COL, self.invoice_num_spinbox.value())
            
            settings.setValue(SETTING_THRESHOLD, self.threshold_spinbox.value())
            if self.logger: self.logger.debug(f"  Saving Threshold: {self.threshold_spinbox.value()}")

            settings.setValue(SETTING_MAX_SPLIT, self.max_split_value_spinbox.value())
            if self.logger: self.logger.debug(f"  Saving Max Split: {self.max_split_value_spinbox.value()}")

            settings.setValue(SETTING_VALUE_COL, self.value_col_spinbox.value())
            if self.logger: self.logger.debug(f"  Saving Value Col: {self.value_col_spinbox.value()}")

            settings.setValue(SETTING_PROP_COL1, self.prop_col1_spinbox.value())
            if self.logger: self.logger.debug(f"  Saving Prop Col 1: {self.prop_col1_spinbox.value()}")

            settings.setValue(SETTING_PROP_COL2, self.prop_col2_spinbox.value())
            if self.logger: self.logger.debug(f"  Saving Prop Col 2: {self.prop_col2_spinbox.value()}")

            settings.setValue(SETTING_CREATE_BACKUP, self.create_backup_check.isChecked())
            if self.logger: self.logger.debug(f"  Saving Create Backup: {self.create_backup_check.isChecked()}")

            settings.setValue(SETTING_OVERWRITE, self.overwrite_check.isChecked())
            if self.logger: self.logger.debug(f"  Saving Overwrite: {self.overwrite_check.isChecked()}")

            settings.setValue(SETTING_OUTPUT_DIR, self.output_dir)
            if self.logger: self.logger.debug(f"  Saving Output Dir: {self.output_dir}")

            if self.logger: self.logger.info("Οι ρυθμίσεις αποθηκεύτηκαν.")
        except Exception as e:
             if self.logger: self.logger.error(f"Σφάλμα κατά την αποθήκευση ρυθμίσεων: {e}", exc_info=True)

    
    def closeEvent(self, event):
        if self.worker is not None and self.worker.isRunning():
            reply = QMessageBox.question(self, 'Επεξεργασία σε Εξέλιξη',
                                         "Η επεξεργασία τρέχει. Να διακοπεί και να κλείσει η εφαρμογή;",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.logger.warning("Κλείσιμο από χρήστη κατά την επεξεργασία.")
                self.worker.requestInterruption()
                
                self.save_settings()
                event.accept()
            else:
                event.ignore()
        else:
            self.logger.info("Κλείσιμο εφαρμογής.")
            
            self.save_settings()
            event.accept()