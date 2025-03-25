import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import threading
import base64
from csv_to_json import csv_to_json_threads
from PIL import Image, ImageTk

class MaximoSenderUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Softwrench's Maximo Data Importer")
        self.root.geometry("800x600")
        
        # Variables
        self.data_file_path = tk.StringVar()
        self.config_file_path = tk.StringVar()
        self.maxauth_token = tk.StringVar()
        self.base_url = tk.StringVar()
        self.maximo_instance = tk.StringVar()
        self.obj_structure = tk.StringVar()
        self.obj_search_attr = tk.StringVar()
        self.obj_id_attr_name = tk.StringVar()
        self.oslc_where = tk.StringVar()
        self.oslc_select = tk.StringVar()
        self.request_type = tk.StringVar(value="-c")
        self.progress_var = tk.DoubleVar()
        self.current_entry = tk.StringVar(value="0/0")
        self.failed_entries = tk.StringVar(value="Failed: 0")
        
        # Load logo
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            logo_image = Image.open(logo_path)
            # Calculate new size maintaining aspect ratio
            target_size = 200
            ratio = min(target_size / logo_image.width, target_size / logo_image.height)
            new_size = (int(logo_image.width * ratio), int(logo_image.height * ratio))
            logo_image = logo_image.resize(new_size, Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_image)
        except Exception as e:
            print(f"Error loading logo: {e}")
            self.logo_photo = None
        
        # Bind request type changes to update UI
        self.request_type.trace_add("write", self.on_request_type_change)
        
        # Bind mouse wheel for scrolling
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Add a queue for UI updates
        self.update_queue = []
        # Start the UI update checker
        self.check_updates()
        
        self.setup_ui()
        self.update_search_fields_visibility()
        
    def setup_ui(self):
        # Create main container with padding and scrolling
        self.canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Configure the canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create window in canvas
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=self.canvas.winfo_reqwidth())
        
        # Pack the scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Create main container with padding
        main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure the scrollable frame to expand
        self.scrollable_frame.grid_rowconfigure(0, weight=1)
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        
        # Bind canvas resize to update frame width and scroll region
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.scrollable_frame.bind('<Configure>', self._on_frame_configure)
        
        # File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(file_frame, text="Data File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.data_file_path, width=50).grid(row=0, column=1, sticky=tk.W)
        ttk.Button(file_frame, text="Browse", command=self.browse_data_file).grid(row=0, column=2, padx=5)
        
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
        
        # Maximo Instance Section
        instance_frame = ttk.Frame(config_frame)
        instance_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(instance_frame, text="Maximo Instance Name:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(instance_frame, textvariable=self.maximo_instance, width=20).grid(row=0, column=1, sticky=tk.W)
        ttk.Label(instance_frame, text=".softwrench2.com/maximo/oslc/os").grid(row=0, column=2, sticky=tk.W)
        
        # Object Structure Section
        structure_frame = ttk.Frame(config_frame)
        structure_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(structure_frame, text="Object Structure:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(structure_frame, textvariable=self.obj_structure).grid(row=0, column=1, sticky=tk.W)
        
        # Maximo Authentication
        auth_frame = ttk.Frame(config_frame)
        auth_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(auth_frame, text="Username:").grid(row=0, column=0, sticky=tk.W)
        self.username_entry = ttk.Entry(auth_frame)
        self.username_entry.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(auth_frame, text="Password:").grid(row=0, column=2, sticky=tk.W)
        self.password_entry = ttk.Entry(auth_frame, show="*")
        self.password_entry.grid(row=0, column=3, sticky=tk.W)
        
        # Additional Configuration Section (initially hidden)
        self.additional_config_frame = ttk.LabelFrame(main_frame, text="Additional Configuration", padding="5")
        self.additional_config_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Search and ID Attributes
        self.search_attr_label = ttk.Label(self.additional_config_frame, text="Search Attribute:")
        self.search_attr_label.grid(row=0, column=0, sticky=tk.W)
        self.search_attr_entry = ttk.Entry(self.additional_config_frame, textvariable=self.obj_search_attr)
        self.search_attr_entry.grid(row=0, column=1, sticky=tk.W)
        
        self.id_attr_label = ttk.Label(self.additional_config_frame, text="ID Attribute:")
        self.id_attr_label.grid(row=1, column=0, sticky=tk.W)
        self.id_attr_entry = ttk.Entry(self.additional_config_frame, textvariable=self.obj_id_attr_name)
        self.id_attr_entry.grid(row=1, column=1, sticky=tk.W)
        
        # OSLC Configuration
        self.oslc_where_label = ttk.Label(self.additional_config_frame, text="OSLC Where:")
        self.oslc_where_label.grid(row=2, column=0, sticky=tk.W)
        self.oslc_where_entry = ttk.Entry(self.additional_config_frame, textvariable=self.oslc_where)
        self.oslc_where_entry.grid(row=2, column=1, sticky=tk.W)
        
        self.oslc_select_label = ttk.Label(self.additional_config_frame, text="OSLC Select:")
        self.oslc_select_label.grid(row=3, column=0, sticky=tk.W)
        self.oslc_select_entry = ttk.Entry(self.additional_config_frame, textvariable=self.oslc_select)
        self.oslc_select_entry.grid(row=3, column=1, sticky=tk.W)
        
        # Progress Section
        self.progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="5")
        self.progress_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Configure the progress frame columns to expand
        self.progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar with mode='determinate' and full width
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            mode='determinate',
            length=0  # Setting length to 0 allows it to expand with its container
        )
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        # Progress labels
        progress_info_frame = ttk.Frame(self.progress_frame)
        progress_info_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5)
        progress_info_frame.columnconfigure(1, weight=1)  # Make the middle space expand
        
        ttk.Label(progress_info_frame, textvariable=self.current_entry).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(progress_info_frame, textvariable=self.failed_entries).grid(row=0, column=2, sticky=tk.E)
        
        # Summary Section
        self.summary_frame = ttk.LabelFrame(main_frame, text="Summary", padding="5")
        self.summary_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.summary_text = tk.Text(self.summary_frame, height=5, width=70)
        self.summary_text.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        # Control Buttons and Logo
        self.bottom_frame = ttk.Frame(main_frame)
        self.bottom_frame.grid(row=6, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        button_frame = ttk.Frame(self.bottom_frame)
        button_frame.pack(side=tk.LEFT)
        
        ttk.Button(button_frame, text="Start Processing", 
                  command=self.start_processing).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear", 
                  command=self.clear_all).pack(side=tk.LEFT, padx=5)
        
        # Add logo to bottom right
        if self.logo_photo:
            logo_label = ttk.Label(self.bottom_frame, image=self.logo_photo)
            logo_label.pack(side=tk.RIGHT, padx=10)
        
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
    
    def on_request_type_change(self, *args):
        self.update_search_fields_visibility()
    
    def update_search_fields_visibility(self):
        action = self.request_type.get()
        show_search_fields = action in ["-u", "-mu", "-d"]
        
        # Update visibility of additional configuration section
        self.additional_config_frame.grid_remove() if not show_search_fields else self.additional_config_frame.grid()
        
        # Update grid positions of subsequent sections
        if show_search_fields:
            self.progress_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
            self.summary_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
            self.bottom_frame.grid(row=6, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        else:
            self.progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
            self.summary_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
            self.bottom_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
    
    def start_processing(self):
        if not self.validate_inputs():
            return
        
        if not self.maxauth_token.get():
            self.generate_token()
        
        config = {
            "base_url": f"https://{self.maximo_instance.get()}.softwrench2.com/maximo/oslc/os",
            "obj_structure": self.obj_structure.get(),
        }
        
        action = self.request_type.get()
        if action in ["-u", "-mu", "-d"]:
            config.update({
                "obj_search_attr": self.obj_search_attr.get(),
                "obj_id_attr_name": self.obj_id_attr_name.get(),
                "oslc.where": self.oslc_where.get(),
                "oslc.select": self.oslc_select.get()
            })
        
        config_path = "temp_config.json"
        with open(config_path, "w") as f:
            json.dump(config, f)
        
        thread = threading.Thread(
            target=self.process_data,
            args=(config_path, self.data_file_path.get(), action)
        )
        thread.daemon = True
        thread.start()
    
    def validate_inputs(self):
        if not self.data_file_path.get():
            messagebox.showerror("Error", "Please select a data file")
            return False
        
        if not self.maximo_instance.get():
            messagebox.showerror("Error", "Please enter your Maximo instance name")
            return False
        
        if not self.obj_structure.get():
            messagebox.showerror("Error", "Please enter the object structure")
            return False
        
        # Validate username and password
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not username:
            messagebox.showerror("Error", "Username is required")
            self.username_entry.focus()
            return False
            
        if not password:
            messagebox.showerror("Error", "Password is required")
            self.password_entry.focus()
            return False
        
        if not self.maxauth_token.get():
            self.generate_token()
        
        action = self.request_type.get()
        if action in ["-u", "-mu", "-d"]:
            if not self.obj_search_attr.get():
                messagebox.showerror("Error", "Please enter the search attribute")
                return False
            if not self.obj_id_attr_name.get():
                messagebox.showerror("Error", "Please enter the ID attribute")
                return False
            if not self.oslc_where.get():
                messagebox.showerror("Error", "Please enter the OSLC Where clause")
                return False
            if not self.oslc_select.get():
                messagebox.showerror("Error", "Please enter the OSLC Select fields")
                return False
        
        return True
    
    def check_updates(self):
        """Process any pending UI updates"""
        while self.update_queue:
            update_func = self.update_queue.pop(0)
            try:
                update_func()
            except Exception as e:
                print(f"Error in UI update: {e}")
        # Schedule the next check
        self.root.after(100, self.check_updates)

    def process_data(self, config_path, data_path, action):
        try:
            print("Starting data processing...")  # Debug log
            
            # Import here to avoid circular imports
            from maximo_sender import process_one_record, process_in_bulk
            
            # Override the MAXAUTH_TOKEN in maximo_sender module
            import maximo_sender
            maximo_sender.MAXAUTH_TOKEN = self.maxauth_token.get()
            
            # Handle CSV file conversion if needed
            temp_json_path = None
            if data_path.lower().endswith('.csv'):
                try:
                    print(f"Converting CSV file: {data_path}")  # Debug log
                    # Create a temporary JSON file for the converted data
                    temp_json_path = data_path.rsplit('.', 1)[0] + '_temp.json'
                    
                    # Convert CSV to JSON with default settings
                    csv_to_json_threads(
                        input_file=data_path,
                        output_file=temp_json_path,
                        parse_dates=True  # Enable date parsing by default
                    )
                    
                    # Wait for the file to exist
                    import time
                    max_retries = 10
                    retries = 0
                    while not os.path.exists(temp_json_path) and retries < max_retries:
                        time.sleep(0.5)
                        retries += 1
                    
                    if not os.path.exists(temp_json_path):
                        raise Exception("Timeout waiting for CSV conversion to complete")
                    
                    print(f"CSV conversion completed. Using: {temp_json_path}")  # Debug log
                    # Use the converted JSON file for processing
                    data_path = temp_json_path
                except Exception as e:
                    print(f"CSV conversion error: {str(e)}")  # Debug log
                    self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to convert CSV file: {str(e)}"))
                    return
            
            print(f"Loading data from: {data_path}")  # Debug log
            # Load data
            try:
                # Ensure the file exists before trying to read it
                if not os.path.exists(data_path):
                    raise FileNotFoundError(f"Data file not found: {data_path}")
                
                with open(data_path, "r") as f:
                    data = json.load(f)
                print(f"Data loaded successfully. Type: {type(data)}")  # Debug log
            except Exception as e:
                print(f"Error loading data: {str(e)}")  # Debug log
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to load data file: {str(e)}"))
                return
            
            if isinstance(data, list):
                data_array = data
                records_to_process = None
            else:
                data_array = data.get("data", [])
                records_to_process = data.get("records_to_process")
            
            print(f"Data array length: {len(data_array)}")  # Debug log
            
            # Load config
            try:
                with open(config_path, "r") as f:
                    config = json.load(f)
                print("Config loaded successfully")  # Debug log
            except Exception as e:
                print(f"Error loading config: {str(e)}")  # Debug log
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to load configuration: {str(e)}"))
                return
            
            # Set up session
            import requests
            session = requests.Session()
            timeout_seconds = 30
            
            # Process based on action
            if action == "-bc":
                print("Starting bulk creation process...")  # Debug log
                print(config['base_url'])
                print(config['obj_structure'])
                process_in_bulk(records_to_process, data_array, 0, f"{config['base_url']}/{config['obj_structure']}?lean=1")
            else:
                print(f"Starting {action} process...")  # Debug log
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

                def queue_progress_update():
                    """Queue a progress update"""
                    nonlocal processed, failed, total_records
                    progress = (processed / total_records) * 100
                    
                    def update():
                        self.progress_var.set(progress)
                        self.current_entry.set(f"{processed}/{total_records}")
                        self.failed_entries.set(f"Failed: {failed}")
                        self.progress_frame.update()
                    
                    self.update_queue.append(update)

                for idx, rec in all_pairs:
                    try:
                        success = process_one_record(
                            idx, rec, session, config, action,
                            f"{config['base_url']}/{config['obj_structure']}?lean=1",
                            timeout_seconds
                        )
                        
                        processed += 1
                        if not success:
                            failed += 1
                        
                        queue_progress_update()
                        print(f"Processed record {idx + 1}/{total_records}")  # Debug log
                    except Exception as e:
                        print(f"Error processing record {idx}: {str(e)}")  # Debug log
                        failed += 1
                        processed += 1
                        queue_progress_update()

                # Queue final summary update
                def queue_summary_update():
                    def update():
                        try:
                            if total_records > 0:
                                success_rate = ((total_records - failed) / total_records * 100)
                            else:
                                success_rate = 0
                                
                            summary = (
                                f"Processing completed!\n"
                                f"Total records: {total_records}\n"
                                f"Successful: {total_records - failed}\n"
                                f"Failed: {failed}\n"
                                f"Success rate: {success_rate:.2f}%"
                            )
                            
                            self.summary_text.delete("1.0", tk.END)
                            self.summary_text.insert("1.0", summary)
                            self.summary_frame.update()
                        except Exception as e:
                            print(f"Error updating summary: {str(e)}")
                    
                    self.update_queue.append(update)

                queue_summary_update()
                
        except Exception as e:
            print(f"Main process error: {str(e)}")  # Debug log
            self.root.after(0, lambda: messagebox.showerror("Error", f"Processing failed: {str(e)}"))
        finally:
            print("Cleaning up temporary files...")  # Debug log
            # Clean up temporary files
            try:
                print(f"Cleaning up files: {config_path}, {temp_json_path}")  # Debug log
                
                if os.path.exists(config_path):
                    os.remove(config_path)
                if temp_json_path and os.path.exists(temp_json_path):
                    os.remove(temp_json_path)
            except Exception as e:
                print(f"Error cleaning up files: {str(e)}")  # Debug log

    def clear_all(self):
        self.data_file_path.set("")
        self.maximo_instance.set("")
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
        self.update_search_fields_visibility()

    def generate_token(self):
        """Generate authentication token from username and password."""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not username or not password:
            return False
            
        # Create the token by encoding username:password in base64
        auth_string = f"{username}:{password}"
        auth_bytes = auth_string.encode('ascii')
        base64_bytes = base64.b64encode(auth_bytes)
        self.maxauth_token.set(base64_bytes.decode('ascii'))
        return True

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_canvas_configure(self, event):
        # Update the frame width to match the canvas width
        self.canvas.itemconfig(self.canvas.find_withtag("all")[0], width=event.width)
        
    def _on_frame_configure(self, event):
        # Update the scroll region to encompass the inner frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Show/hide scrollbar based on content size
        if self.scrollable_frame.winfo_reqheight() > self.canvas.winfo_height():
            self.canvas.yview_moveto(0)
        else:
            self.canvas.yview_moveto(0)

def main():
    root = tk.Tk()
    app = MaximoSenderUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 