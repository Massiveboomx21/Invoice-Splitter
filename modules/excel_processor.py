import os
import decimal
import random
import win32com.client as win32
import pythoncom
import time

class ExcelProcessor:
    
    def __init__(self, logger=None):
        self.logger = logger
        
    def _generate_n_splits_integer_multiple_of_5(self, value_decimal, N, max_split_decimal):
        """
        Διασπά μια τιμή (πολλαπλάσιο του 5) σε Ν ακέραια κομμάτια, πολλαπλάσια του 5.
        Κάθε κομμάτι πρέπει να είναι < max_split_decimal.
        """
        if self.logger: self.logger.info(f"Integer (x5) split: V={value_decimal}, N={N}, Max={max_split_decimal}")

        
        if value_decimal % 5 != 0:
            if self.logger: self.logger.error(f"Integer (x5) split: Η αρχική τιμή {value_decimal} δεν είναι πολλαπλάσιο του 5. Αδύνατη η διάσπαση.")
            return None

        
        value_units = int(value_decimal / 5)
        
        max_split_units = int((max_split_decimal - 1) / 5)

        if self.logger: self.logger.debug(f"Integer (x5) split -> Units: Value={value_units}, N={N}, Max={max_split_units}")

        if N <= 0 or value_units < N:
            return None
        if max_split_units <= 0:
            if self.logger: self.logger.error(f"Integer (x5) split: Το μέγιστο όριο ({max_split_decimal}) είναι πολύ μικρό.")
            return None

        
        base_part_units = value_units // N
        remainder_units = value_units % N

        if base_part_units > max_split_units:
             if self.logger: self.logger.warning(f"Integer (x5) split: Το βασικό κομμάτι ({base_part_units*5}€) υπερβαίνει το όριο.")
             return None
        if base_part_units + 1 > max_split_units and remainder_units > 0:
            if self.logger: self.logger.warning(f"Integer (x5) split: Η προσθήκη υπολοίπου υπερβαίνει το όριο.")
            return None

        parts_units = [base_part_units] * N
        for i in range(remainder_units):
            parts_units[i] += 1
        
        
        if any(p > max_split_units for p in parts_units):
            if self.logger: self.logger.error(f"Integer (x5) split: Αποτυχία. Ένα κομμάτι υπερέβη το όριο μετά τη διανομή υπολοίπου.")
            return None

        
        final_parts_decimal = [decimal.Decimal(p * 5) for p in parts_units]
        
        if self.logger: self.logger.info(f"Integer (x5) split success. Parts: {final_parts_decimal}")
        return final_parts_decimal

    def generate_n_splits_normalized(self, value_decimal, N, max_split_decimal, epsilon, max_retries=100):
        rounding_precision = decimal.Decimal('0.01')
        current_context = decimal.getcontext()
        original_prec = current_context.prec
        current_context.prec = max(original_prec, 34)
        try:
            if N <= 0 or value_decimal < N * epsilon:
                if self.logger: self.logger.error(f"generate_n_splits_normalized: Αδύνατη η διάσπαση του {value_decimal} σε {N} κομμάτια >= {epsilon}.")
                return None
            for attempt in range(max_retries):
                rand_nums = [decimal.Decimal(str(random.uniform(0.01, 1.0))) for _ in range(N)]
                total_rand = sum(rand_nums)
                if total_rand == 0: continue
                parts_initial = [(r / total_rand) * value_decimal for r in rand_nums]
                parts_quantized = [p.quantize(rounding_precision, rounding=decimal.ROUND_HALF_UP) for p in parts_initial]
                all_valid_pre_adjust = all(pq >= epsilon and pq < max_split_decimal for pq in parts_quantized)
                if not all_valid_pre_adjust: continue
                current_sum = sum(parts_quantized)
                diff = value_decimal - current_sum
                idx = 0
                max_adjust_loops = N * 5
                adjustment_step = rounding_precision if diff > 0 else -rounding_precision
                while abs(diff) >= rounding_precision / 2 and idx < max_adjust_loops:
                    part_idx = random.randrange(N)
                    adjusted_part = parts_quantized[part_idx] + adjustment_step
                    if adjusted_part >= epsilon and adjusted_part < max_split_decimal:
                        parts_quantized[part_idx] = adjusted_part
                        diff -= adjustment_step
                    idx += 1
                final_sum = sum(parts_quantized)
                if abs(value_decimal - final_sum) < rounding_precision / 2:
                    all_valid_final = all(pq >= epsilon and pq < max_split_decimal for pq in parts_quantized)
                    if all_valid_final:
                        if self.logger: self.logger.debug(f"generate_n_splits_normalized success after {attempt+1} tries.")
                        return parts_quantized
            if self.logger: self.logger.error(f"Αποτυχία generate_n_splits_normalized για V={value_decimal}, N={N}, max_split={max_split_decimal} μετά από {max_retries}.")
            return None
        finally:
             current_context.prec = original_prec

    def _generate_n_splits_deterministic(self, value_decimal, N, max_split_decimal, epsilon):
        rounding_precision = decimal.Decimal('0.01')
        with decimal.localcontext() as ctx:
             ctx.prec = max(ctx.prec, 28)
             if N <= 0 or value_decimal < N * epsilon:
                 if self.logger: self.logger.debug(f"_generate_n_splits_deterministic: Αδύνατη είσοδος N={N}, V={value_decimal}")
                 return None
             base_part = (value_decimal / N).quantize(rounding_precision, rounding=decimal.ROUND_FLOOR)
             if base_part < epsilon:
                 if self.logger: self.logger.debug(f"_deterministic fallback: Base part {base_part} < epsilon for N={N}.")
                 return None
             remainder = value_decimal - (base_part * N)
             num_parts_to_increment = int(remainder / epsilon)
             if base_part + epsilon >= max_split_decimal and num_parts_to_increment > 0:
                 if self.logger: self.logger.debug(f"_deterministic fallback: Incrementing base part {base_part} would exceed max_split {max_split_decimal}.")
                 return None
             parts = [base_part] * N
             for i in range(num_parts_to_increment):
                 if i < N: parts[i] += epsilon
             final_sum = sum(parts)
             if abs(final_sum - value_decimal) >= rounding_precision / 2:
                 if self.logger: self.logger.error(f"_deterministic fallback: Σφάλμα αθροίσματος! Αναμενόταν {value_decimal}, βρέθηκε {final_sum}. Parts: {parts}")
                 return None
             all_valid = all(p >= epsilon and p < max_split_decimal for p in parts)
             if all_valid:
                 if self.logger: self.logger.info(f"Η ντετερμινιστική διάσπαση (fallback) παρήγαγε {N} έγκυρα κομμάτια.")
                 return parts
             else:
                 if self.logger: self.logger.warning(f"Η ντετερμινιστική διάσπαση (fallback) απέτυχε ελέγχους ορίων για V={value_decimal}, N={N}, max_split={max_split_decimal}.")
                 return None

    
    def process_file(self, input_path, output_path, threshold=500, value_col=6, prop_cols=None, overwrite=False, max_split_value=None, split_mode='decimal', auto_numbering=False, invoice_num_col=2):
        if prop_cols is None: prop_cols = [8, 19]
        if max_split_value is None: max_split_value = threshold

        rounding_precision = decimal.Decimal('0.01')
        epsilon = decimal.Decimal('0.01')
        try:
             with decimal.localcontext() as ctx:
                 ctx.prec = 34
                 max_split_decimal = decimal.Decimal(str(max_split_value)).quantize(rounding_precision)
                 if max_split_decimal < epsilon:
                     if self.logger: self.logger.warning(f"max_split_value ({max_split_value}) πολύ μικρό, χρησιμοποιείται {epsilon}.")
                     max_split_decimal = epsilon
        except decimal.InvalidOperation:
             if self.logger: self.logger.error(f"Μη έγκυρη τιμή max_split_value: {max_split_value}. Χρησιμοποιείται threshold ({threshold}).")
             max_split_decimal = decimal.Decimal(str(threshold)).quantize(rounding_precision)

        file_basename = os.path.basename(input_path)
        output_basename = os.path.basename(output_path)

        if self.logger: self.logger.info(f"--- Έναρξη επεξεργασίας: {file_basename} ---")

        if not overwrite and os.path.exists(output_path):
            if self.logger: self.logger.warning(f"Το αρχείο '{output_basename}' υπάρχει ήδη. Παράλειψη (overwrite=False).")
            return {'processed_rows': 0, 'split_rows': 0, 'errors': 0, 'skipped': True, 'message': f"File '{output_basename}' already exists."}

        results = {
            'processed_rows': 0, 'split_rows': 0, 'errors': 0,
            'skipped_impossible_splits': 0, 'multi_splits_performed': 0,
            'skipped': False, 'message': ''
        }
        excel = None
        workbook = None
        original_calculation_mode = None 

        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error as e:
             if self.logger: self.logger.warning(f"Pythoncom.CoInitialize com_error: {e} (Maybe ignorable)")

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            
            excel.Visible = False
            excel.DisplayAlerts = False
            

            try:
                 workbook = excel.Workbooks.Open(os.path.abspath(input_path))
                 if self.logger: self.logger.debug(f"Workbook '{file_basename}' opened successfully.")
                 
                 try:
                      original_calculation_mode = excel.Calculation 
                      excel.Calculation = win32.constants.xlCalculationManual
                      excel.ScreenUpdating = False
                      excel.EnableEvents = False
                      if self.logger: self.logger.debug("Excel optimization settings applied.")
                 except pythoncom.com_error as opt_err:
                      if self.logger: self.logger.warning(f"Σφάλμα COM κατά την εφαρμογή ρυθμίσεων βελτιστοποίησης: {opt_err}. Η επεξεργασία συνεχίζεται...")
                      
                 except Exception as gen_opt_err: 
                       if self.logger: self.logger.warning(f"Γενικό σφάλμα κατά την εφαρμογή ρυθμίσεων βελτιστοποίησης: {gen_opt_err}. Η επεξεργασία συνεχίζεται...")


            except pythoncom.com_error as open_error:
                 if self.logger: self.logger.error(f"Σφάλμα COM ανοίγματος workbook '{file_basename}': {open_error}")
                 results['errors'] += 1; results['message'] = f"COM Error opening workbook: {open_error}"
                 if excel: excel.Quit()
                 pythoncom.CoUninitialize()
                 return results

            
            for worksheet_idx in range(1, workbook.Worksheets.Count + 1):
                 
                 
                worksheet = workbook.Worksheets(worksheet_idx)
                if self.logger: self.logger.info(f"Επεξεργασία φύλλου: '{worksheet.Name}'")

                try: 
                    last_row = worksheet.UsedRange.Rows.Count
                    if last_row <= 1: 
                         cell_val = None
                         try: cell_val = worksheet.Cells(1,1).Value
                         except: pass
                         if cell_val is None or str(cell_val).strip() == "":
                             if self.logger: self.logger.info(f"Παράλειψη (πιθανώς) κενού φύλλου: '{worksheet.Name}'")
                             continue

                    rows_to_split = []
                    sheet_processed_rows = 0
                    
                    for row in range(last_row, 1, -1):
                        sheet_processed_rows += 1
                        try:
                             cell_a_value = worksheet.Cells(row, 1).Value
                             if cell_a_value and "σύνολα" in str(cell_a_value).lower(): continue
                             value = worksheet.Cells(row, value_col).Value
                             if not (value and isinstance(value, (int, float)) and value >= threshold): continue
                             cell_b_value = worksheet.Cells(row, 2).Value
                             if cell_a_value is not None and str(cell_a_value).strip() != "" and \
                                cell_b_value is not None and str(cell_b_value).strip() != "":
                                 rows_to_split.append(row)
                        except pythoncom.com_error as cell_read_err:
                              if self.logger: self.logger.warning(f"Σφάλμα COM ανάγνωσης κελιού γραμμής {row}, φύλλο '{worksheet.Name}': {cell_read_err}. Παράλειψη.")
                        except Exception as gen_read_err:
                              if self.logger: self.logger.warning(f"Σφάλμα ανάγνωσης δεδομένων γραμμής {row}, φύλλο '{worksheet.Name}': {gen_read_err}. Παράλειψη.")

                except Exception as sheet_prep_err:
                     if self.logger: self.logger.error(f"Σφάλμα κατά την προετοιμασία φύλλου '{worksheet.Name}': {sheet_prep_err}")
                     results['errors'] += 1
                     continue 

                
                sheet_split_count = 0
                for row in rows_to_split:
                    try: 
                        value = worksheet.Cells(row, value_col).Value
                        if not (value and isinstance(value, (int, float))): continue

                        with decimal.localcontext() as ctx:
                            ctx.prec = 34
                            value_decimal = decimal.Decimal(str(value))
                            N = 1
                            split_values_decimal = None
                            split_method_used = "None"

                            
                            if value_decimal >= threshold:
                                
                                if split_mode == 'integer_5':
                                    
                                    if value_decimal % 5 != 0:
                                        if self.logger: self.logger.warning(f"Η λειτουργία 'Ακέραια Διάσπαση' απαιτεί πολλαπλάσια του 5. Παράλειψη για {value_decimal:.2f} στη Γρ.{row}.")
                                        if 'skipped_details' not in results: results['skipped_details'] = []
                                        results['skipped_details'].append({'file': file_basename, 'sheet': worksheet.Name, 'row': row, 'value': f"{value_decimal:.2f}"})
                                        split_values_decimal = None
                                        N = 1
                                    else:
                                        N_calc = int((value_decimal / max_split_decimal).to_integral_value(rounding=decimal.ROUND_CEILING))
                                        if N_calc < 2: N_calc = 2
                                        N = N_calc
                                        split_values_decimal = self._generate_n_splits_integer_multiple_of_5(value_decimal, N, max_split_decimal)
                                        if split_values_decimal:
                                            split_method_used = f"Integer x5 ({N}-way)"
                                        else:
                                            if 'skipped_details' not in results: results['skipped_details'] = []
                                            results['skipped_details'].append({'file': file_basename, 'sheet': worksheet.Name, 'row': row, 'value': f"{value_decimal:.2f}"})
                                            N = 1

                                else:
                                    if value_decimal < max_split_decimal:
                                        if self.logger: self.logger.info(f"Η τιμή {value_decimal:.2f} (Γρ.{row}) >= όριο αλλά < μέγιστο. Παραμένει.")
                                        N = 1; split_values_decimal = None; split_method_used = "None (Below Max)"
                                    elif value_decimal < 2 * max_split_decimal:
                                        N = 2
                                        lower_bound = max(epsilon, value_decimal - max_split_decimal + epsilon)
                                        upper_bound = min(value_decimal - epsilon, max_split_decimal - epsilon)
                                        lower_bound = lower_bound.quantize(rounding_precision, rounding=decimal.ROUND_CEILING)
                                        upper_bound = upper_bound.quantize(rounding_precision, rounding=decimal.ROUND_FLOOR)
                                        split_possible_2way = False; temp_split_values = None
                                        if lower_bound <= upper_bound:
                                            try:
                                                if lower_bound == upper_bound: split1_decimal = lower_bound
                                                else:
                                                    rand_float = random.uniform(float(lower_bound), float(upper_bound))
                                                    split1_decimal = decimal.Decimal(str(rand_float)).quantize(rounding_precision, rounding=decimal.ROUND_HALF_UP)
                                                    if split1_decimal < lower_bound: split1_decimal = lower_bound
                                                    if split1_decimal > upper_bound: split1_decimal = upper_bound
                                                split2_decimal = (value_decimal - split1_decimal).quantize(rounding_precision, rounding=decimal.ROUND_HALF_UP)
                                                sum_check_tolerance = epsilon / 10
                                                if (split1_decimal >= epsilon and split1_decimal < max_split_decimal and
                                                    split2_decimal >= epsilon and split2_decimal < max_split_decimal and
                                                    abs(split1_decimal + split2_decimal - value_decimal) < sum_check_tolerance):
                                                    temp_split_values = [split1_decimal, split2_decimal]; split_possible_2way = True; split_method_used = "Random (2-way)"
                                                else:
                                                    s1_half = (value_decimal / 2).quantize(rounding_precision, rounding=decimal.ROUND_HALF_UP)
                                                    s2_half = value_decimal - s1_half
                                                    if (s1_half >= epsilon and s1_half < max_split_decimal and
                                                        s2_half >= epsilon and s2_half < max_split_decimal and
                                                        abs(s1_half + s2_half - value_decimal) < sum_check_tolerance):
                                                        temp_split_values = [s1_half, s2_half]; split_possible_2way = True; split_method_used = "Half (2-way)"
                                            except Exception as e_2way:
                                                if self.logger: self.logger.error(f"Σφάλμα υπολ. 2-way split για γραμμή {row}: {e_2way}")
                                        if split_possible_2way: split_values_decimal = temp_split_values
                                        else:
                                            if 'skipped_details' not in results: results['skipped_details'] = []
                                            results['skipped_details'].append({'file': file_basename, 'sheet': worksheet.Name, 'row': row, 'value': f"{value_decimal:.2f}"})
                                    else: 
                                        try:
                                            N_calc = int((value_decimal / max_split_decimal).to_integral_value(rounding=decimal.ROUND_CEILING))
                                            if N_calc <= 2: N_calc = 3
                                        except Exception as calc_n_err:
                                            N_calc = 1; split_values_decimal = None
                                            if self.logger: self.logger.error(f"Σφάλμα υπολ. N για γραμμή {row}: {calc_n_err}", exc_info=True)
                                        if N_calc > 1:
                                            N = N_calc; temp_split_values_n = None
                                            if self.logger: self.logger.info(f"Τιμή {value_decimal:.2f} (Γρ.{row}) -> N={N} κομμάτια (< {max_split_value:.2f}). Προσπάθεια random...")
                                            temp_split_values_n = self.generate_n_splits_normalized(value_decimal, N, max_split_decimal, epsilon)
                                            if temp_split_values_n is None:
                                                if self.logger: self.logger.warning(f"Random N-way split απέτυχε για Γρ.{row}. Δοκιμή deterministic fallback...")
                                                temp_split_values_n = self._generate_n_splits_deterministic(value_decimal, N, max_split_decimal, epsilon)
                                                if temp_split_values_n is not None: split_method_used = "Deterministic Fallback"
                                                else:
                                                    if 'skipped_details' not in results: results['skipped_details'] = []
                                                    results['skipped_details'].append({'file': file_basename, 'sheet': worksheet.Name, 'row': row, 'value': f"{value_decimal:.2f}"})
                                            else: split_method_used = f"Random ({N}-way)"
                                            if temp_split_values_n is not None: split_values_decimal = temp_split_values_n
                                        else: N = 1; split_values_decimal = None; split_method_used = "None (N Calc Error?)"
                            

                            
                            if N <= 1 or split_values_decimal is None:
                                continue 

                            
                            if self.logger: self.logger.info(f"Διάσπαση γραμμής {row} σε {N} κομμάτια ({split_method_used}): {', '.join(f'{s:.2f}' for s in split_values_decimal)}")
                            split_values_float = [float(s) for s in split_values_decimal]
                            ratios = [s / value_decimal if value_decimal else decimal.Decimal(1/N) for s in split_values_decimal]

                            
                            original_row_data = {}
                            first_col=1; last_col_to_copy=value_col-1; other_static_cols=[]
                            columns_to_copy_indices = list(range(first_col, last_col_to_copy + 1)) + other_static_cols
                            columns_to_copy_indices = [c for c in columns_to_copy_indices if c not in prop_cols and c != value_col]
                            if self.logger: self.logger.debug(f"Θα αντιγραφούν δεδομένα από στήλες: {columns_to_copy_indices}")
                            for col_idx in columns_to_copy_indices:
                                try: original_row_data[col_idx] = worksheet.Cells(row, col_idx).Value
                                except: original_row_data[col_idx] = None
                            original_prop_values = {}
                            for prop_col in prop_cols:
                                try: original_prop_values[prop_col] = worksheet.Cells(row, prop_col).Value
                                except: original_prop_values[prop_col] = None

                            if N > 1:
                                 try:
                                      start_cell = worksheet.Cells(row + 1, 1); end_cell = worksheet.Cells(row + N - 1, 1)
                                      if self.logger: self.logger.debug(f"Προσπάθεια εισαγωγής {N-1} γραμμών από {row + 1}...")
                                      worksheet.Range(start_cell, end_cell).EntireRow.Insert(Shift=win32.constants.xlShiftDown)
                                      if self.logger: self.logger.debug(f"Επιτυχής εισαγωγή {N-1} γραμμών. Παύση...")
                                      time.sleep(0.2) 
                                 except Exception as insert_err:
                                      if self.logger: self.logger.error(f"Σφάλμα εισαγωγής {N-1} γραμμών στη γραμμή {row+1}: {insert_err}")
                                      results['errors'] += 1; continue

                            prop_sums_calculated = {pc: decimal.Decimal(0) for pc in prop_cols}
                            for i in range(N):
                                current_row_index = row + i
                                if i > 0: 
                                     for col_idx, value_to_copy in original_row_data.items():
                                         try: worksheet.Cells(current_row_index, col_idx).Value = value_to_copy
                                         except Exception as write_err:
                                              if self.logger: self.logger.warning(f"Σφάλμα εγγραφής αντιγραφής στο ({current_row_index},{col_idx}): {write_err}")
                                
                                try:
                                     worksheet.Cells(current_row_index, value_col).Value = split_values_float[i]
                                except Exception as write_err:
                                    
                                    results['errors'] += 1
                                    if self.logger:
                                        self.logger.error(f"Σφάλμα εγγραφής βασικής τιμής στο ({current_row_index},{value_col}): {write_err}")
                                
                                for prop_col in prop_cols:
                                    original_value = original_prop_values.get(prop_col)
                                    if original_value is not None and isinstance(original_value, (int, float)):
                                        prop_decimal = decimal.Decimal(str(original_value))
                                        try:
                                            if i < N - 1: prop_i_decimal = (prop_decimal * ratios[i]).quantize(rounding_precision, rounding=decimal.ROUND_HALF_UP)
                                            else:
                                                prop_i_decimal = prop_decimal - prop_sums_calculated[prop_col]
                                                if prop_i_decimal < 0 and prop_decimal >= 0: prop_i_decimal = decimal.Decimal(0)
                                            worksheet.Cells(current_row_index, prop_col).Value = float(prop_i_decimal)
                                            prop_sums_calculated[prop_col] += prop_i_decimal
                                        except Exception as prop_calc_e:
                                             if self.logger: self.logger.warning(f"Σφάλμα υπολ./εγγραφής αναλ. στήλης {prop_col} γραμμής {current_row_index}: {prop_calc_e}")
                                    elif original_value is not None and i > 0:
                                           try: worksheet.Cells(current_row_index, prop_col).Value = original_value
                                           except: pass

                            
                            sheet_split_count += 1
                            results['split_rows'] += 1
                            if N > 2: results['multi_splits_performed'] = results.get('multi_splits_performed', 0) + 1
                        
                    except pythoncom.com_error as split_com_err:
                         results['errors'] += 1
                         if self.logger: self.logger.error(f"Σφάλμα COM κατά τη διάσπαση γραμμής {row}, φύλλο '{worksheet.Name}': {split_com_err}")
                    except Exception as e:
                         results['errors'] += 1
                         if self.logger: self.logger.error(f"Γενικό σφάλμα κατά τη διάσπαση γραμμής {row}, φύλλο '{worksheet.Name}': {str(e)}", exc_info=True)
                

                if self.logger: self.logger.info(f"Ολοκληρώθηκε το φύλλο '{worksheet.Name}'. Διασπάστηκαν {sheet_split_count} γραμμές.")
                results['processed_rows'] += sheet_processed_rows
            
            
            try:

                try:
                    excel.ScreenUpdating = True
                    excel.EnableEvents = True
                    
                    if original_calculation_mode is not None:
                        excel.Calculation = original_calculation_mode
                    else: 
                        excel.Calculation = win32.constants.xlCalculationAutomatic
                    if self.logger: self.logger.debug("Excel optimization settings restored before save.")
                except Exception as restore_err:
                     if self.logger: self.logger.warning(f"Σφάλμα επαναφοράς ρυθμίσεων Excel πριν την αποθήκευση: {restore_err}")

                workbook.SaveAs(os.path.abspath(output_path))
                if self.logger: self.logger.info(f"Το επεξεργασμένο αρχείο αποθηκεύτηκε ως: {output_basename}")
                results['message'] = f"Successfully processed and saved to {output_basename}"
            except pythoncom.com_error as save_error:
                 if self.logger: self.logger.error(f"Σφάλμα COM κατά την αποθήκευση '{output_basename}': {save_error}")
                 results['errors'] += 1; results['message'] = f"COM Error saving file: {save_error}"

            workbook.Close(SaveChanges=False)
            workbook = None

        except pythoncom.com_error as main_com_error:
            results['errors'] += 1; results['message'] = f"Main processing COM Error: {main_com_error}"
            if self.logger: self.logger.error(f"Κύριο σφάλμα COM επεξεργασίας '{file_basename}': {main_com_error}")
        except Exception as general_error:
            results['errors'] += 1; results['message'] = f"General processing error: {general_error}"
            if self.logger: self.logger.error(f"Γενικό σφάλμα επεξεργασίας '{file_basename}': {str(general_error)}", exc_info=True)
        finally:
            
            if workbook is not None:
                try: workbook.Close(SaveChanges=False)
                except: pass
            if excel is not None:
                try:
                     
                     excel.ScreenUpdating = True; excel.EnableEvents = True
                     if original_calculation_mode is not None:
                          excel.Calculation = original_calculation_mode
                     else:
                          excel.Calculation = win32.constants.xlCalculationAutomatic
                     excel.DisplayAlerts = True; excel.Quit()
                except: pass
            pythoncom.CoUninitialize()
            if self.logger: self.logger.info(f"--- Ολοκλήρωση επεξεργασίας: {file_basename} (Errors: {results['errors']}) ---")

        return results

    
    def process_multiple_files(self, input_files, output_dir, threshold=500, value_col=6, prop_cols=None, overwrite=False, max_split_value=None):
        """
        Επεξεργασία πολλαπλών αρχείων Excel καλώντας την process_file για το καθένα.
        """
        if self.logger: self.logger.info(f"Ξεκινά η μαζική επεξεργασία {len(input_files)} αρχείων...")
        overall_results = {
            'total_files': len(input_files), 'processed_files': 0, 'skipped_files': 0,
            'total_rows_processed': 0, 'total_rows_split': 0,
            'skipped_impossible_splits': 0, 'multi_splits_performed': 0,
            'errors': 0, 'file_results': {}
        }
        for input_file in input_files:
            file_name = os.path.basename(input_file)
            try:
                file_name_base, file_ext = os.path.splitext(file_name)
                output_name = f"{file_name_base}_διασπασμένο{file_ext}"
                output_path = os.path.join(output_dir, output_name)
                results = self.process_file(
                    input_file, output_path, threshold, value_col, prop_cols,
                    overwrite, max_split_value
                )
                if results.get('skipped'): overall_results['skipped_files'] += 1
                else:
                    
                    if results.get('errors', 0) == 0:
                         overall_results['processed_files'] += 1
                    overall_results['total_rows_processed'] += results.get('processed_rows', 0)
                    overall_results['total_rows_split'] += results.get('split_rows', 0)
                    overall_results['skipped_impossible_splits'] += results.get('skipped_impossible_splits', 0)
                    overall_results['multi_splits_performed'] += results.get('multi_splits_performed', 0)
                overall_results['errors'] += results.get('errors', 0)
                overall_results['file_results'][file_name] = results
            except Exception as e:
                overall_results['errors'] += 1
                error_msg = f"Κρίσιμο σφάλμα διαχείρισης {file_name}: {str(e)}"
                if self.logger: self.logger.error(error_msg, exc_info=True)
                overall_results['file_results'][file_name] = {'error': True, 'message': error_msg, 'skipped': False}
        if self.logger:
            self.logger.info(f"Ολοκληρώθηκε η μαζική επεξεργασία.")
            self.logger.info(f"Σύνοψη: Επεξεργάστηκαν={overall_results['processed_files']}, Παραλείφθηκαν={overall_results['skipped_files']}, Σφάλματα={overall_results['errors']}, Διασπάσεις={overall_results['total_rows_split']}")
        return overall_results