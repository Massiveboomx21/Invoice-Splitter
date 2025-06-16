import os
import shutil

class FileManager:
    @staticmethod
    def create_backup(file_path):
        """
        Δημιουργία αντιγράφου ασφαλείας του αρχείου.
        
        Args:
            file_path (str): Η διαδρομή του αρχείου
            
        Returns:
            str: Η διαδρομή του αντιγράφου ασφαλείας
        """
        backup_path = f"{file_path}.backup"
        shutil.copy2(file_path, backup_path)
        return backup_path
    
    @staticmethod
    def validate_excel_file(file_path):
        """
        Έλεγχος εγκυρότητας αρχείου Excel.
        
        Args:
            file_path (str): Η διαδρομή του αρχείου
            
        Returns:
            bool: True αν είναι έγκυρο αρχείο Excel
        """
        if not os.path.exists(file_path):
            return False
        
        ext = os.path.splitext(file_path)[1].lower()
        return ext in ['.xls', '.xlsx', '.xlsm']
    
    @staticmethod
    def get_output_path(input_path, output_dir=None, suffix='_διασπασμένο'):
        """
        Δημιουργία διαδρομής εξόδου για το αρχείο.
        
        Args:
            input_path (str): Η διαδρομή του αρχείου εισόδου
            output_dir (str): Ο φάκελος εξόδου (προεπιλογή: ίδιος με της εισόδου)
            suffix (str): Το επίθεμα για το νέο όνομα αρχείου
        
        Returns:
            str: Η διαδρομή του αρχείου εξόδου
        """
        file_name = os.path.basename(input_path)
        name, ext = os.path.splitext(file_name)
        
        if output_dir is None:
            output_dir = os.path.dirname(input_path)
        
        return os.path.join(output_dir, f"{name}{suffix}{ext}")