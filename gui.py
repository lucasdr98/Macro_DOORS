import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import threading
import importlib.util
import sys
from datetime import datetime
import shutil
import time


class MacroGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOORS Macro - VF Downloader")
        self.root.geometry("800x900")  # Tamanho inicial menor
        self.root.minsize(600, 900)  # Tamanho mínimo para garantir visibilidade dos botões
        self.root.resizable(True, True)
        
        # Set icon if available
        try:
            self.root.iconbitmap("images/icon.ico")
        except:
            pass
        
        # Main frame
        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=False, padx=5, pady=5)
        
        # Configure settings
        self.setup_settings(settings_frame)
        
        # Status frame - tamanho aumentado
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Status box - altura aumentada
        self.status_text = tk.Text(status_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(status_frame, orient=tk.HORIZONTAL, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=5, pady=5)
        
        # Buttons frame - diretamente no main_frame para garantir visibilidade
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Start button
        self.start_button = ttk.Button(buttons_frame, text="Start Macro", command=self.run_macro)
        self.start_button.pack(side=tk.RIGHT, padx=5)
        
        # Exit button
        self.exit_button = ttk.Button(buttons_frame, text="Exit", command=root.destroy)
        self.exit_button.pack(side=tk.RIGHT, padx=5)
        
        # Initialize variables
        self.projects_from_excel = False
        self.excel_path = None
        self.region = None
        self.output_dir = os.path.join(os.getcwd(), "output")
        
        # Create output directory if it doesn't exist
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        
        # Set the initial output directory
        self.output_dir_var.set(self.output_dir)
        
        # Log to the status box
        self.log("GUI initialized. Please configure settings and click 'Start Macro' to begin.")
        
        # Countdown timer flag
        self.countdown_active = False
    
    def setup_settings(self, parent_frame):
        # Use a single column layout for better organization
        parent_frame.columnconfigure(0, weight=1)
        
        row = 0
        
        # Project selection frame
        project_frame = ttk.LabelFrame(parent_frame, text="Project Selection", padding=(15, 10))
        project_frame.grid(row=row, column=0, sticky="ew", padx=5, pady=10)
        row += 1
        
        # Configure project selection frame
        self.project_method = tk.StringVar(value="manual")
        
        # Radio buttons for project input method
        radio_frame = ttk.Frame(project_frame)
        radio_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Radiobutton(
            radio_frame, 
            text="Manual Input", 
            variable=self.project_method, 
            value="manual",
            command=self.toggle_project_input
        ).pack(side=tk.LEFT, padx=(0, 15), pady=2)
        
        ttk.Radiobutton(
            radio_frame, 
            text="Load from Excel", 
            variable=self.project_method, 
            value="excel",
            command=self.toggle_project_input
        ).pack(side=tk.LEFT, padx=5, pady=2)
        
        # Manual project input
        self.manual_frame = ttk.Frame(project_frame)
        self.manual_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.projects_var = tk.StringVar()
        ttk.Entry(self.manual_frame, textvariable=self.projects_var).pack(fill=tk.X, padx=5, pady=5)
        
        # Excel project input
        self.excel_frame = ttk.Frame(project_frame)
        self.excel_frame.pack(fill=tk.X, padx=5, pady=5)
        self.excel_frame.pack_forget()  # Hide initially
        
        excel_file_frame = ttk.Frame(self.excel_frame)
        excel_file_frame.pack(fill=tk.X, expand=True, pady=5)
        
        self.excel_path_var = tk.StringVar()
        ttk.Entry(excel_file_frame, textvariable=self.excel_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(excel_file_frame, text="Browse...", command=self.browse_excel).pack(side=tk.RIGHT, padx=5)
        
        region_frame = ttk.Frame(self.excel_frame)
        region_frame.pack(fill=tk.X, expand=True, pady=5)
        
        ttk.Label(region_frame, text="Region:").pack(side=tk.LEFT, padx=5, pady=2)
        self.region_var = tk.StringVar()
        region_combobox = ttk.Combobox(region_frame, textvariable=self.region_var, width=15, state="readonly")
        region_combobox['values'] = ('EMEA', 'NAFTA', 'LATAM')
        region_combobox.pack(side=tk.LEFT, padx=5, pady=2)
        
        # Domains frame
        domains_frame = ttk.LabelFrame(parent_frame, text="Domains", padding=(15, 10))
        domains_frame.grid(row=row, column=0, sticky="ew", padx=5, pady=10)
        row += 1
        
        # Configure domains frame
        self.domains_var = tk.StringVar()
        ttk.Entry(domains_frame, textvariable=self.domains_var).pack(fill=tk.X, padx=5, pady=5)
        
        # Use Cases frame
        usecases_frame = ttk.LabelFrame(parent_frame, text="Use Cases", padding=(15, 10))
        usecases_frame.grid(row=row, column=0, sticky="ew", padx=5, pady=10)
        row += 1
        
        # Configure use cases frame
        self.usecases_var = tk.StringVar()
        ttk.Entry(usecases_frame, textvariable=self.usecases_var).pack(fill=tk.X, padx=5, pady=5)
        
        # VFs frame
        vfs_frame = ttk.LabelFrame(parent_frame, text="VFs", padding=(15, 10))
        vfs_frame.grid(row=row, column=0, sticky="ew", padx=5, pady=10)
        row += 1
        
        # Configure VFs frame
        self.vfs_var = tk.StringVar()
        ttk.Entry(vfs_frame, textvariable=self.vfs_var).pack(fill=tk.X, padx=5, pady=5)
        
        # Output directory frame
        output_frame = ttk.LabelFrame(parent_frame, text="Output Settings", padding=(15, 10))
        output_frame.grid(row=row, column=0, sticky="ew", padx=5, pady=10)
        row += 1
        
        # Configure output settings
        output_dir_frame = ttk.Frame(output_frame)
        output_dir_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        
        self.output_dir_var = tk.StringVar()
        ttk.Entry(output_dir_frame, textvariable=self.output_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_dir_frame, text="Browse...", command=self.browse_output_dir).pack(side=tk.RIGHT, padx=5)
        
        # Debug checkbox
        debug_frame = ttk.Frame(output_frame)
        debug_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        
        self.debug_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(debug_frame, text="Enable Debug Mode", variable=self.debug_var).pack(anchor=tk.W, padx=5)
    
    def toggle_project_input(self):
        if self.project_method.get() == "manual":
            self.manual_frame.pack(fill=tk.X, padx=5, pady=5)
            self.excel_frame.pack_forget()
            self.projects_from_excel = False
        else:
            self.manual_frame.pack_forget()
            self.excel_frame.pack(fill=tk.X, padx=5, pady=5)
            self.projects_from_excel = True
    
    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self.excel_path = file_path
    
    def browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="Select Output Directory")
        if dir_path:
            self.output_dir_var.set(dir_path)
            self.output_dir = dir_path
    
    def log(self, message):
        self.status_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def validate_inputs(self):
        errors = []
        
        if self.projects_from_excel:
            if not self.excel_path_var.get():
                errors.append("Please select an Excel file")
            
            if not self.region_var.get():
                errors.append("Please select a region")
        else:
            if not self.projects_var.get():
                errors.append("Please enter at least one project")
        
        if not self.domains_var.get():
            errors.append("Please enter at least one domain")
        
        if not self.vfs_var.get():
            errors.append("Please enter at least one VF")
        
        if errors:
            messagebox.showerror("Validation Error", "\n".join(errors))
            return False
        
        return True
    
    def create_output_dirs(self):
        # Create debug and logs directories inside output directory
        debug_dir = os.path.join(self.output_dir, "debug")
        logs_dir = os.path.join(self.output_dir, "logs")
        
        if not os.path.exists(debug_dir):
            os.makedirs(debug_dir)
        
        if not os.path.exists(logs_dir):
            os.makedirs(logs_dir)
        
        return debug_dir, logs_dir
    
    def countdown(self, seconds):
        self.countdown_active = True
        for i in range(seconds, 0, -1):
            if not self.countdown_active:
                break
            self.log(f"Starting macro in {i} seconds... Prepare your screen.")
            time.sleep(1)
        self.countdown_active = False
    
    def run_macro(self):
        if not self.validate_inputs():
            return
        
        # Disable buttons during execution
        self.start_button.config(state=tk.DISABLED)
        self.exit_button.config(state=tk.DISABLED)
        
        # Start progress bar
        self.progress.start()
        
        # Create output directories
        debug_dir, logs_dir = self.create_output_dirs()
        
        # Start countdown in a separate thread
        countdown_thread = threading.Thread(target=self.countdown, args=(5,))
        countdown_thread.daemon = True
        countdown_thread.start()
        
        # Create a thread to run the macro after countdown
        thread = threading.Thread(target=self.execute_macro)
        thread.daemon = True
        
        # Start the macro execution thread after countdown completes
        self.root.after(5000, thread.start)
    
    def execute_macro(self):
        try:
            self.log("Starting macro execution...")
            
            # Load macro.py as a module
            
            spec = importlib.util.spec_from_file_location("macro", "macro.py")
            macro = importlib.util.module_from_spec(spec)
            
            # Set global variables in the macro module before executing
            spec.loader.exec_module(macro)
            
            # Update debug directory
            macro.debug_dir = os.path.join(self.output_dir, "debug")
            
            # Update logs directory
            macro.logs_dir = os.path.join(self.output_dir, "logs")
            
            # Set debug mode
            macro.debug = self.debug_var.get()
            
            # Override the messagebox function to use our logger
            original_messagebox_showerror = messagebox.showerror
            original_messagebox_showinfo = messagebox.showinfo
            
            def custom_showerror(title, message):
                self.log(f"ERROR: {message}")
                return original_messagebox_showerror(title, message)
            
            def custom_showinfo(title, message):
                self.log(f"INFO: {message}")
                return original_messagebox_showinfo(title, message)
            
            macro.messagebox.showerror = custom_showerror
            macro.messagebox.showinfo = custom_showinfo
            
            # Prepare arguments for main_logic function
            projects_list = []
            if self.projects_from_excel:
                self.log(f"Loading projects from Excel file: {self.excel_path_var.get()}")
                self.log(f"Using region: {self.region_var.get()}")
                
                # Call macro.filtrar_codigos_por_regiao to get projects
                try:
                    # Get projects from Excel file
                    raw_projects_list = macro.filtrar_codigos_por_regiao(
                        self.excel_path_var.get(),
                        self.region_var.get()
                    )
                    
                    # Convert all items to strings to avoid type errors
                    projects_list = [str(p).strip() for p in raw_projects_list if p is not None]
                    
                    if not projects_list:
                        self.log(f"No projects found for region {self.region_var.get()} in the Excel file.")
                        messagebox.showerror("Error", f"No projects found for region {self.region_var.get()} in the Excel file.")
                        self.cleanup()
                        return
                    
                    self.log(f"Found projects: {', '.join(projects_list)}")
                    
                except Exception as e:
                    self.log(f"Error loading projects from Excel: {str(e)}")
                    messagebox.showerror("Error", f"Error loading projects from Excel: {str(e)}")
                    self.cleanup()
                    return
            else:
                # Parse manual project input
                projects_text = self.projects_var.get().strip()
                projects_list = [p.strip() for p in projects_text.split(',') if p.strip()]
                self.log(f"Using projects: {', '.join(projects_list)}")
            
            # Get domains
            domains_text = self.domains_var.get().strip()
            domains_list = [d.strip() for d in domains_text.split(',') if d.strip()]
            self.log(f"Using domains: {', '.join(domains_list)}")
            
            # Get use cases
            usecases_text = self.usecases_var.get().strip()
            usecases_list = [u.strip() for u in usecases_text.split(',') if u.strip()]
            self.log(f"Using use cases: {', '.join(usecases_list)}")
            
            # Get VFs
            vfs_text = self.vfs_var.get().strip()
            vfs_list = [v.strip() for v in vfs_text.split(',') if v.strip()]
            self.log(f"Using VFs: {', '.join(vfs_list)}")
            
            # Initialize paths for log files
            timestamp_execucao = datetime.now().strftime("%Y%m%d_%H%M%S")
            macro.timestamp_execucao = timestamp_execucao
            macro.nome_arquivo_log = f"{macro.logs_dir}/log_{timestamp_execucao}.txt"
            macro.nome_arquivo_caminhos = f"{macro.logs_dir}/caminhos_{timestamp_execucao}.txt"
            macro.caminhos_registrados = set()
            
            # Clean up old log files
            macro.limpar_arquivos_antigos(macro.logs_dir, "log_", 10)
            macro.limpar_arquivos_antigos(macro.logs_dir, "caminhos_", 10)
            
            self.log("Starting macro operations. Please do not interfere with the mouse or keyboard.")
            self.log("This may take several minutes...")
            
            # Call the main_logic function with our parameters
            try:
                # Execute the main logic function
                macro.main_logic(projects_list, domains_list, usecases_list, vfs_list, self.output_dir)
                self.log("Macro execution completed successfully!")
                
            except Exception as e:
                self.log(f"Error during macro execution: {str(e)}")
                import traceback
                self.log(f"Traceback: {traceback.format_exc()}")
                macro.messagebox.showerror("Error", f"Error during macro execution: {str(e)}")
            
        except Exception as e:
            self.log(f"Error setting up macro execution: {str(e)}")
            import traceback
            self.log(f"Traceback: {traceback.format_exc()}")
            messagebox.showerror("Error", f"Error setting up macro execution: {str(e)}")
        
        finally:
            self.cleanup()
    
    def cleanup(self):
        # Stop progress bar
        self.progress.stop()
        
        # Cancel any active countdown
        self.countdown_active = False
        
        # Re-enable buttons
        self.start_button.config(state=tk.NORMAL)
        self.exit_button.config(state=tk.NORMAL)
        
        self.log("Ready for next operation.")

def main():
    root = tk.Tk()
    app = MacroGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 