import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import threading
import base64
from csv_to_json import csv_to_json_threads

class MaximoSenderUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Maximo Data Sender")
        self.root.geometry("800x600")
        
        # Variables
        self.data_file_path = tk.StringVar()
        self.config_file_path = tk.StringVar()
        self.maxauth_token = tk.StringVar()
        self.base_url = tk.StringVar()
        self.obj_structure = tk.StringVar()
        self.obj_search_attr = tk.StringVar()
        self.obj_id_attr_name = tk.StringVar()
        self.oslc_where = tk.StringVar()
        self.oslc_select = tk.StringVar()
        self.request_type = tk.StringVar(value="-c")
        self.progress_var = tk.DoubleVar()
        self.current_entry = tk.StringVar(value="0/0")
        self.failed_entries = tk.StringVar(value="Failed: 0")
        
        self.setup_ui()
        
    def setup_ui(self):
        # Create main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(file_frame, text="Data File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.data_file_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_data_file).grid(row=0, column=2)
        
        # Request Type Section
        request_frame = ttk.LabelFrame(main_frame, text="Request Type", padding="5")
        request_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        request_types = [
            ("Bulk Create", "-bc"),
            ("Create", "-c"),
            ("Update", "-u"),
            ("Merge Update", "-mu"),
            ("Delete", "-d")
        ]
        
        for i, (label, value) in enumerate(request_types):
            ttk.Radiobutton(request_frame, text=label, value=value, 
                          variable=self.request_type).grid(row=i//2, column=i%2, sticky=tk.W, padx=5)
        
        # Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="5")
        config_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Maximo Authentication
        auth_frame = ttk.Frame(config_frame)
        auth_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(auth_frame, text="Username:").grid(row=0, column=0, sticky=tk.W)
        self.username_entry = ttk.Entry(auth_frame)
        self.username_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(auth_frame, text="Password:").grid(row=0, column=2, sticky=tk.W)
        self.password_entry = ttk.Entry(auth_frame, show="*")
        self.password_entry.grid(row=0, column=3, padx=5)
        
        ttk.Button(auth_frame, text="Generate Token", 
                  command=self.generate_token).grid(row=0, column=4, padx=5)
        
        # Base URL and Object Structure
        ttk.Label(config_frame, text="Base URL:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(config_frame, textvariable=self.base_url, width=50).grid(row=1, column=1, columnspan=2, padx=5)
        
        ttk.Label(config_frame, text="Object Structure:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(config_frame, textvariable=self.obj_structure).grid(row=2, column=1, padx=5)
        
        # Search and ID Attributes
        ttk.Label(config_frame, text="Search Attribute:").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(config_frame, textvariable=self.obj_search_attr).grid(row=3, column=1, padx=5)
        
        ttk.Label(config_frame, text="ID Attribute:").grid(row=4, column=0, sticky=tk.W)
        ttk.Entry(config_frame, textvariable=self.obj_id_attr_name).grid(row=4, column=1, padx=5)
        
        # OSLC Configuration
        ttk.Label(config_frame, text="OSLC Where:").grid(row=5, column=0, sticky=tk.W)
        ttk.Entry(config_frame, textvariable=self.oslc_where).grid(row=5, column=1, padx=5)
        
        ttk.Label(config_frame, text="OSLC Select:").grid(row=6, column=0, sticky=tk.W)
        ttk.Entry(config_frame, textvariable=self.oslc_select).grid(row=6, column=1, padx=5)
        
        # Progress Section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="5")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(progress_frame, textvariable=self.current_entry).grid(row=1, column=0, sticky=tk.W)
        ttk.Label(progress_frame, textvariable=self.failed_entries).grid(row=1, column=1, sticky=tk.E)
        
        # Summary Section
        summary_frame = ttk.LabelFrame(main_frame, text="Summary", padding="5")
        summary_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.summary_text = tk.Text(summary_frame, height=5, width=70)
        self.summary_text.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        # Control Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Start Processing", 
                  command=self.start_processing).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Clear", 
                  command=self.clear_all).grid(row=0, column=1, padx=5)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
    def browse_data_file(self):
        try:
            file_path = filedialog.askopenfilename(
                title="Select Data File",
                filetypes=[
                    ("CSV files", "*.csv"),
                    ("JSON files", "*.json")
                ]
            )
            
            if not file_path:
                return
                
            if not (file_path.lower().endswith('.csv') or file_path.lower().endswith('.json')):
                messagebox.showwarning("Warning", "Please select a CSV or JSON file")
                return
                
            self.data_file_path.set(file_path)
            if file_path.lower().endswith('.csv'):
                self.show_csv_conversion_dialog(file_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting file: {str(e)}")
            print(f"File dialog error: {str(e)}")  # For debugging
    
    def show_csv_conversion_dialog(self, csv_path):
        dialog = tk.Toplevel(self.root)
        dialog.title("CSV to JSON Conversion")
        dialog.geometry("400x300")
        
        # Variables for CSV conversion
        parse_dates = tk.BooleanVar(value=True)
        person_transform = tk.StringVar()
        
        ttk.Label(dialog, text="CSV Conversion Options").pack(pady=10)
        
        ttk.Checkbutton(dialog, text="Parse Dates", 
                       variable=parse_dates).pack(pady=5)
        
        ttk.Label(dialog, text="Person Transform Columns (comma-separated):").pack(pady=5)
        ttk.Entry(dialog, textvariable=person_transform).pack(pady=5)
        
        def convert_csv():
            try:
                output_path = csv_path.rsplit('.', 1)[0] + '.json'
                person_transform_cols = [col.strip() for col in person_transform.get().split(',')] if person_transform.get() else None
                
                csv_to_json_threads(
                    input_file=csv_path,
                    output_file=output_path,
                    parse_dates=parse_dates.get(),
                    person_transform_columns=person_transform_cols
                )
                
                self.data_file_path.set(output_path)
                dialog.destroy()
                messagebox.showinfo("Success", "CSV converted to JSON successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to convert CSV: {str(e)}")
        
        ttk.Button(dialog, text="Convert", command=convert_csv).pack(pady=10)
    
    def generate_token(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        
        if not username or not password:
            messagebox.showerror("Error", "Please enter both username and password")
            return
        
        token = base64.b64encode(f"{username}:{password}".encode()).decode()
        self.maxauth_token.set(token)
    
    def start_processing(self):
        if not self.validate_inputs():
            return
        
        # Create config dictionary
        config = {
            "base_url": self.base_url.get(),
            "obj_structure": self.obj_structure.get(),
            "obj_search_attr": self.obj_search_attr.get(),
            "obj_id_attr_name": self.obj_id_attr_name.get(),
            "oslc.where": self.oslc_where.get(),
            "oslc.select": self.oslc_select.get()
        }
        
        # Save config to temporary file
        config_path = "temp_config.json"
        with open(config_path, "w") as f:
            json.dump(config, f)
        
        # Start processing in a separate thread
        thread = threading.Thread(
            target=self.process_data,
            args=(config_path, self.data_file_path.get(), self.request_type.get())
        )
        thread.daemon = True
        thread.start()
    
    def validate_inputs(self):
        if not self.data_file_path.get():
            messagebox.showerror("Error", "Please select a data file")
            return False
        
        if not self.base_url.get():
            messagebox.showerror("Error", "Please enter the base URL")
            return False
        
        if not self.obj_structure.get():
            messagebox.showerror("Error", "Please enter the object structure")
            return False
        
        if not self.maxauth_token.get():
            messagebox.showerror("Error", "Please generate the Maximo authentication token")
            return False
        
        return True
    
    def process_data(self, config_path, data_path, action):
        try:
            # Import here to avoid circular imports
            from maximo_sender import process_one_record, process_in_bulk
            
            # Load data
            with open(data_path, "r") as f:
                data = json.load(f)
            
            if isinstance(data, list):
                data_array = data
                records_to_process = None
            else:
                data_array = data.get("data", [])
                records_to_process = data.get("records_to_process")
            
            # Load config
            with open(config_path, "r") as f:
                config = json.load(f)
            
            # Set up session
            import requests
            session = requests.Session()
            timeout_seconds = 30
            
            # Process based on action
            if action == "-bc":
                process_in_bulk(records_to_process, data_array, 0, f"{config['base_url']}/{config['obj_structure']}?lean=1")
            else:
                # Process records one by one
                all_pairs = []
                if records_to_process:
                    for i in records_to_process:
                        if 0 <= i < len(data_array):
                            all_pairs.append((i, data_array[i]))
                else:
                    for i in range(len(data_array)):
                        all_pairs.append((i, data_array[i]))
                
                total_records = len(all_pairs)
                processed = 0
                failed = 0
                
                for idx, rec in all_pairs:
                    success = process_one_record(
                        idx, rec, session, config, action,
                        f"{config['base_url']}/{config['obj_structure']}?lean=1",
                        timeout_seconds
                    )
                    
                    processed += 1
                    if not success:
                        failed += 1
                    
                    # Update progress
                    progress = (processed / total_records) * 100
                    self.progress_var.set(progress)
                    self.current_entry.set(f"{processed}/{total_records}")
                    self.failed_entries.set(f"Failed: {failed}")
                    self.root.update_idletasks()
                
                # Show summary
                summary = f"Processing completed!\n"
                summary += f"Total records: {total_records}\n"
                summary += f"Successful: {total_records - failed}\n"
                summary += f"Failed: {failed}\n"
                summary += f"Success rate: {((total_records - failed) / total_records * 100):.2f}%"
                
                self.summary_text.delete(1.0, tk.END)
                self.summary_text.insert(tk.END, summary)
                
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed: {str(e)}")
        finally:
            # Clean up temporary config file
            try:
                os.remove(config_path)
            except:
                pass
    
    def clear_all(self):
        self.data_file_path.set("")
        self.base_url.set("")
        self.obj_structure.set("")
        self.obj_search_attr.set("")
        self.obj_id_attr_name.set("")
        self.oslc_where.set("")
        self.oslc_select.set("")
        self.request_type.set("-c")
        self.progress_var.set(0)
        self.current_entry.set("0/0")
        self.failed_entries.set("Failed: 0")
        self.summary_text.delete(1.0, tk.END)
        self.username_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)
        self.maxauth_token.set("")

def main():
    root = tk.Tk()
    app = MaximoSenderUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 