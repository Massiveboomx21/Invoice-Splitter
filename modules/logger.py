
import os
import logging
from datetime import datetime
from PyQt5.QtCore import QObject, pyqtSignal

class Logger(QObject):
    log_signal = pyqtSignal(str)

    def __init__(self, log_dir=None, log_level=logging.INFO):
        super().__init__()

        if log_dir is None:
            try:
                script_dir = os.path.dirname(__file__)
                parent_dir = os.path.dirname(script_dir)
                log_dir = os.path.join(parent_dir, 'logs')
            except NameError: 
                 log_dir = os.path.join(os.getcwd(), 'logs')


        try:
            os.makedirs(log_dir, exist_ok=True)
        except OSError as e:
             print(f"CRITICAL: Could not create log directory '{log_dir}'. Error: {e}")
             
             log_dir = os.getcwd()


        self.logger = logging.getLogger('invoice_splitter_app')
        self.logger.setLevel(log_level)
        self.log_file = None 

        
        if not self.logger.handlers:
            try:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                self.log_file = os.path.join(log_dir, f'invoice_splitter_{timestamp}.log')

                
                file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
                file_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
                file_handler.setFormatter(file_format)
                self.logger.addHandler(file_handler)

                
                console_handler = logging.StreamHandler()
                console_format = logging.Formatter('%(levelname)s: %(message)s')
                console_handler.setFormatter(console_format)
                console_handler.setLevel(logging.INFO) 
                self.logger.addHandler(console_handler)

                
                

            except Exception as e:
                 print(f"CRITICAL: Failed to configure logging handlers. Error: {e}")

        else: 
             for handler in self.logger.handlers:
                 if isinstance(handler, logging.FileHandler):
                     self.log_file = handler.baseFilename
                     break

    def _emit_signal(self, level, message):
        """Εκπέμπει το σήμα για το UI log."""
        formatted_message = f"[{level}] {message}"
        try:
            self.log_signal.emit(formatted_message)
        except Exception as e:
            print(f"Error emitting log signal: {e}. Message: {formatted_message}")

    
    def info(self, message, exc_info=False):
        self.logger.info(message, exc_info=exc_info)
        
        
        self._emit_signal("INFO", message)

    def warning(self, message, exc_info=False):
        self.logger.warning(message, exc_info=exc_info)
        self._emit_signal("WARNING", message) 

    def error(self, message, exc_info=False):
        
        self.logger.error(message, exc_info=exc_info)
        
        self._emit_signal("ERROR", message)

    def debug(self, message, exc_info=False):
        self.logger.debug(message, exc_info=exc_info)
        
        

    def get_log_file(self):
        """Επιστρέφει τη διαδρομή του τρέχοντος αρχείου log."""
        return self.log_file if hasattr(self, 'log_file') else None