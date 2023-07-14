import threading
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import tkinter.ttk as ttk
from tkinter import messagebox, Toplevel

import time

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Manufacturer specific DTCs")
        self.configure(background="#f2f2f2")
        self.create_widgets()

    def create_widgets(self):
        self.create_summary_table_widgets()
        self.create_column_indices_widgets()
        self.create_sae_j2012da_widgets()
        self.create_manufacturer_dtc_widgets()
        self.create_fault_code_manufacturer_dtc_widgets()
        self.create_control_unit_widgets()
        self.create_process_button()
        self.create_footer()
        
    def create_summary_table_widgets(self):
        summary_table_label = tk.Label(self, text="Selection File:", bg="#f2f2f2", font=("Arial", 12))
        summary_table_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.summary_table_entry = tk.Entry(self, width=50, font=("Arial", 12))
        self.summary_table_entry.grid(row=0, column=1, padx=10)
        browse_summary_table_button = tk.Button(self, text="Browse", command=self.browse_summary_table, font=("Arial", 12))
        browse_summary_table_button.grid(row=0, column=2, padx=10)

    def create_column_indices_widgets(self):
        fault_code_column_label = tk.Label(self, text="Fault Code index:", bg="#f2f2f2", font=("Arial", 12))
        fault_code_column_label.grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.fault_code_column_spinbox = tk.Spinbox(self, from_=0, to=100, width=5, font=("Arial", 12))
        self.fault_code_column_spinbox.grid(row=1, column=1, padx=10, pady=5)

        start_row_label = tk.Label(self, text="Starting Row:", bg="#f2f2f2", font=("Arial", 12))
        start_row_label.grid(row=2, column=0, sticky="w", padx=10, pady=10)
        self.start_row_spinbox = tk.Spinbox(self, from_=1, to=100, width=5, font=("Arial", 12))
        self.start_row_spinbox.grid(row=2, column=1, padx=10)

    def create_sae_j2012da_widgets(self):
        sae_j2012da_label = tk.Label(self, text="SAE J2012DA:", bg="#f2f2f2", font=("Arial", 12))
        sae_j2012da_label.grid(row=3, column=0, sticky="w", padx=10, pady=10)
        self.sae_j2012da_entry = tk.Entry(self, width=50, font=("Arial", 12))
        self.sae_j2012da_entry.grid(row=3, column=1, padx=10)
        browse_sae_j2012da_button = tk.Button(self, text="Browse", command=self.browse_sae_j2012da, font=("Arial", 12))
        browse_sae_j2012da_button.grid(row=3, column=2, padx=10)

    def create_manufacturer_dtc_widgets(self):
        manufacturer_dtc_label = tk.Label(self, text="Manufacturer DTC:", bg="#f2f2f2", font=("Arial", 12))
        manufacturer_dtc_label.grid(row=4, column=0, sticky="w", padx=10, pady=10)
        self.manufacturer_dtc_entry = tk.Entry(self, width=50, font=("Arial", 12))
        self.manufacturer_dtc_entry.grid(row=4, column=1, padx=10)
        browse_manufacturer_dtc_button = tk.Button(self, text="Browse", command=self.browse_manufacturer_dtc, font=("Arial", 12))
        browse_manufacturer_dtc_button.grid(row=4, column=2, padx=10)

    def create_fault_code_manufacturer_dtc_widgets(self):
        fault_code_column_manufacturer_dtc_label = tk.Label(self, text="Fault Code index (Manufacturer DTC):", bg="#f2f2f2", font=("Arial", 12))
        fault_code_column_manufacturer_dtc_label.grid(row=5, column=0, sticky="w", padx=10, pady=10)
        self.fault_code_column_manufacturer_dtc_spinbox = tk.Spinbox(self, from_=1, to=100, width=5, font=("Arial", 12))
        self.fault_code_column_manufacturer_dtc_spinbox.grid(row=5, column=1, padx=10)

    def create_control_unit_widgets(self):
        control_unit_label = tk.Label(self, text="Control Unit:", bg="#f2f2f2", font=("Arial", 12))
        control_unit_label.grid(row=6, column=0, sticky="w", padx=10, pady=10)
        self.control_unit_entry = tk.Entry(self, width=50, font=("Arial", 12))
        self.control_unit_entry.grid(row=6, column=1, padx=10)

    def create_process_button(self):
        process_button = tk.Button(self, text="Start", command=self.process_summary_table, font=("Arial", 14, "bold"), bg="#336699", fg="white")
        process_button.grid(row=8, column=0, columnspan=3, pady=20)

    def create_footer(self):
        footer_label = tk.Label(self, text="Author: Kuyubasioglu Ilhami Capgemini-engineering, OBD-Team, 13.06.2023", bg="#f2f2f2", fg="gray", font=("Arial", 10))
        footer_label.grid(row=9, column=0, columnspan=3, pady=10)

    def create_style(self):
        style = ttk.Style()
        style.configure("TCombobox", fieldbackground="#ffffff")

    def browse_summary_table(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        self.summary_table_entry.delete(0, tk.END)
        self.summary_table_entry.insert(0, file_path)

    def browse_sae_j2012da(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        self.sae_j2012da_entry.delete(0, tk.END)
        self.sae_j2012da_entry.insert(0, file_path)

    def browse_manufacturer_dtc(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        self.manufacturer_dtc_entry.delete(0, tk.END)
        self.manufacturer_dtc_entry.insert(0, file_path)
        

    
    def fault_code_in_normal_area(area, fault_code):
        for a, f in zip(area, fault_code):
            if a.isdigit() and f.isdigit():
                if int(f) < int(a):
                    return False  # Fault code is below the valid range
            elif a == 'X':
                if not f.isdigit() or int(f, 16) > 15:
                    return False  # Fault code is outside the X-value range
            elif a != f:
                return False  # Mismatched characters

        return True
    
    
    def fault_code_in_area(area, fault_code):
        if '-' in area:
            start, end = area.split('-')
            if len(fault_code) != len(start) or len(fault_code) != len(end):
                return False

            for a_start, a_end, f in zip(start, end, fault_code):
                if a_start == "X" or a_end == "X":
                    continue

                if f < a_start or f > a_end:
                    return False

            return True
        
        




    
    
    
    
    
    
    
    
    
    
    
    
    

    def process_row(self, row):
        prefix_mapping = {
            "P": "P0000-P3FFF",
            "B": "B0000-B3FFF",
            "C": "C0000-C3FFF",
            "U": "U0000-U3FFF"
        }
        
        # get the fault code 
        fault_code_summary = row[self.fault_code_column]
        
        # get sae fault code and column_description
        sae_j2012da_sheet = prefix_mapping.get(fault_code_summary[0], "Invalid Prefix")
        sae_j2012da_df = self.sae_j2012da.parse(sae_j2012da_sheet)
        fault_code_sae = sae_j2012da_df.iloc[:, 2].values.tolist()
        check_column = sae_j2012da_df.iloc[:, 3].values.tolist()
  
        # check if fault code in sae file
        if fault_code_summary in fault_code_sae:
            if check_column[fault_code_sae.index(fault_code_summary)] == "ISO/SAE Reserved":
                self.summary_table.at[self.index, "Kommentar Capgemini"] = "ISO/SAE Reserved: akutellen Kondiax-Auszug prüfen"
                return
            
            elif check_column[fault_code_sae.index(fault_code_summary)] == "Manufacturer Controlled DTC":
                # check if fault code in manufacturer dtc sheets
                manufacturer_dtc_sheets = self.manufacturer_dtc.sheet_names 
                for sheet in manufacturer_dtc_sheets:
                    manufacturer_dtc_df = self.manufacturer_dtc.parse(sheet)
                    fault_code_manufacturer_dtc = manufacturer_dtc_df.iloc[:, self.fault_code_column_manufacturer_dtc].values.tolist()
                    
                    if fault_code_summary in fault_code_manufacturer_dtc:
                        self.summary_table.at[self.index, "Kommentar Capgemini"] = "In Manufacturer DTCs gefunden"
                        return
                    
                    else:
                        self.summary_table.at[self.index, "Kommentar Capgemini"] = "DTC nicht beantragt: bitte PAG benachrichtigen"
                        return
                    
            else:
                self.summary_table.at[self.index, "Kommentar Capgemini"] = "In SAE J2012DA gefunden"
                return
            
        else:
            for column in check_column:
                if "Table" not in column:
                    continue
                
                index = check_column.index(column)
                if not self.fault_code_in_area(fault_code_sae[index], fault_code_summary):
                    self.summary_table.at[self.index, "Kommentar Capgemini"] = "Fault Code in SAE J2012DA nicht gefunden (nicht direkt und in keinem Bereich): bitte Fault Code auf Tippfehler überprüfen"
                    return
                
                if "ISO/SAE Reserved" in check_column[index]:
                    self.summary_table.at[self.index, "Kommentar Capgemini"] = "ISO/SAE Reserved: akutellen Kondiax-Auszug prüfen"
                    return
                    
                elif "Manufacturer Controlled DTC" in check_column[index]:
                    # check if fault code in manufacturer dtc sheets
                    manufacturer_dtc_sheets = self.manufacturer_dtc.sheet_names 
                    for sheet in manufacturer_dtc_sheets:
                        manufacturer_dtc_df = self.manufacturer_dtc.parse(sheet)
                        fault_code_manufacturer_dtc = manufacturer_dtc_df.iloc[:, self.fault_code_column_manufacturer_dtc].values.tolist()
                        
                        if fault_code_summary in fault_code_manufacturer_dtc:
                            self.summary_table.at[self.index, "Kommentar Capgemini"] = "In Manufacturer DTCs gefunden"
                            return
                        
                        else:
                            self.summary_table.at[self.index, "Kommentar Capgemini"] = "DTC nicht beantragt: bitte PAG benachrichtigen"
                            return
        
        # # checking if fault code not in sae file
        # if fault_code_summary not in fault_code_sae:
        #     self.summary_table.at[self.index, "Kommentar Capgemini"] = "Nicht vorhanden"
        #     self.summary_table.at[self.index, "Finding"] = "x"
        #     return
        
        # # checking if column_description is ISO/SAE Reserved
        # index = fault_code_sae.index(fault_code_summary)
        # if "ISO/SAE Reserved" in check_column[index]:
        #     self.summary_table.at[self.index, "Kommentar Capgemini"] = "DTC ISO/SAE reserved: Aktuellen Kondiax Auszug prüfen."
        #     return

        # self.summary_table.at[self.index, "Kommentar Capgemini"] = "Manufacturer-specific DTC. Aktuellen Kondiax Auszug prüfen."

    def get_all_inputs(self):
        try:
            self.summary_table_path = "C:\\Users\\mkanoua\\Downloads\\Excel\\Abgleich_herstellerspezifische_DTCs.xlsx"
            self.fault_code_column = int(self.fault_code_column_spinbox.get())
            self.start_row = int(self.start_row_spinbox.get()) - 1
            self.sae_j2012da_path = "C:\\Users\\mkanoua\\Downloads\\Excel\\J2012DA_201812.xlsx"
            self.manufacturer_dtc_path = "C:\\Users\\mkanoua\\Downloads\\Excel\\20220830_Manufacturer_specific_DTCs_PAG.xlsx"
            self.fault_code_column_manufacturer_dtc = int(self.fault_code_column_manufacturer_dtc_spinbox.get())
            self.control_unit = self.control_unit_entry.get()
            
            if not self.summary_table_path or not self.sae_j2012da_path or not self.manufacturer_dtc_path:
                return messagebox.showerror("Error", "Please provide paths for all required files.")
            
            if not self.control_unit:
                continue_process = messagebox.askokcancel("Warning", "Control Unit was not provided. Do you want to proceed without a Control Unit name?", default='cancel')
                if not continue_process:
                    return
        
        except FileNotFoundError as e:
            return messagebox.showerror("Error", f"Path not found: {e.filename}")
        except Exception as e:
            return messagebox.showerror("Error", "An error occurred while getting the inputs. Please try again.")

    def process_summary_table(self): 
        start_time = time.time()
        try:
            self.get_all_inputs()
            
            #####
            summary_table_dir = os.path.dirname(self.summary_table_path)
            #####
            
            self.summary_table = pd.read_excel(self.summary_table_path)
            self.sae_j2012da = pd.ExcelFile(self.sae_j2012da_path)
            self.manufacturer_dtc = pd.ExcelFile(self.manufacturer_dtc_path)

            progress_window = Toplevel(self)
            progress_window.title("Processing...")
            progress_window.configure(background="#f2f2f2")
            progress_window.geometry("300x50")
            progress_window.resizable(False, False)

            progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
            progress_bar.pack(pady=10)
            progress_bar.start(10)

            def process_complete():
                progress_bar.stop()
                progress_window.destroy()
                messagebox.showinfo("Process Complete", f"The summary table has been processed. Output saved to 'Abgleich_herstellerspezifische_DTCs_{self.control_unit}.xlsx'.")
                print("--- %s seconds ---" % (time.time() - start_time))

            def process_error(e):
                progress_bar.stop()
                progress_window.destroy()
                messagebox.showerror("Error", f"An error occurred during the process: {e}")

            def process_summary_table_thread():
                try:
                    for self.index, row in self.summary_table.iterrows():
                        if self.index < self.start_row:
                            continue
                        self.process_row(row)

                    output_path = os.path.join(summary_table_dir, f"Abgleich_herstellerspezifische_DTCs_{self.control_unit}.xlsx")
                    self.summary_table.to_excel(output_path, index=False)
                    self.after(0, process_complete)
                except Exception as e:
                    self.after(0, process_error(e))

            processing_thread = threading.Thread(target=process_summary_table_thread)
            processing_thread.daemon = True
            processing_thread.start()
            
        except Exception as e:
            messagebox.showerror("Error", "An error occurred during the process: ")

if __name__ == "__main__":
    app = Application()
    app.mainloop()











###################################################################################################################



def is_fault_code_in_area(area, fault_code):
    if len(area) != len(fault_code):
        return False

    for a, f in zip(area, fault_code):
        print(f"{a}  -   {f}")
        if a != 'X' and a != f:
            return False  # Mismatched characters

    return True

# Example usage
area = "P1XXX"
fault_code = "B12FD"

if is_fault_code_in_area(area, fault_code):
    print("Fault code falls into the area.")
else:
    print("Fault code does not fall into the area.")
