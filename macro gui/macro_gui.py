import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os
import sys
from macro import main_macro

class MacroGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Macro DOORS")
        self.root.geometry("600x500")
        
        # Variáveis para armazenar os inputs com valores padrão
        self.projects_var = tk.StringVar(value="226MCA")
        self.domains_var = tk.StringVar(value="defroster")
        self.use_cases_var = tk.StringVar(value="")
        self.vfs_var = tk.StringVar(value="291")
        self.doors_var = tk.StringVar(value="EMEA")
        
        # Variável para controlar o estado do botão
        self.running = False
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        ttk.Label(main_frame, text="Macro DOORS", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
        
        # Projetos
        ttk.Label(main_frame, text="Projetos (separados por vírgula):").grid(row=1, column=0, sticky=tk.W, pady=5)
        projects_entry = ttk.Entry(main_frame, textvariable=self.projects_var, width=50)
        projects_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Domínios
        ttk.Label(main_frame, text="Domínios (separados por vírgula):").grid(row=2, column=0, sticky=tk.W, pady=5)
        domains_entry = ttk.Entry(main_frame, textvariable=self.domains_var, width=50)
        domains_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Use Cases
        ttk.Label(main_frame, text="Use Cases (separados por vírgula):").grid(row=3, column=0, sticky=tk.W, pady=5)
        use_cases_entry = ttk.Entry(main_frame, textvariable=self.use_cases_var, width=50)
        use_cases_entry.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # VFs
        ttk.Label(main_frame, text="VFs (separadas por vírgula):").grid(row=4, column=0, sticky=tk.W, pady=5)
        vfs_entry = ttk.Entry(main_frame, textvariable=self.vfs_var, width=50)
        vfs_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # DOORS
        ttk.Label(main_frame, text="DOORS:").grid(row=5, column=0, sticky=tk.W, pady=5)
        doors_frame = ttk.Frame(main_frame)
        doors_frame.grid(row=5, column=1, sticky=tk.W, pady=5)
        
        ttk.Radiobutton(doors_frame, text="LATAM", variable=self.doors_var, value="LATAM").pack(side=tk.LEFT)
        ttk.Radiobutton(doors_frame, text="EMEA", variable=self.doors_var, value="EMEA").pack(side=tk.LEFT)
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="Iniciar", command=self.start_macro)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Parar", command=self.stop_macro, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar", font=("Arial", 10))
        self.status_label.grid(row=7, column=0, columnspan=2, pady=10)
        
    def start_macro(self):
        # Obtém os valores dos campos e remove espaços em branco
        projects_text = self.projects_var.get().strip()
        domains_text = self.domains_var.get().strip()
        use_cases_text = self.use_cases_var.get().strip()
        vfs_text = self.vfs_var.get().strip()
        doors = self.doors_var.get()
        
        # Lista para armazenar mensagens de erro
        errors = []
        
        # Validação detalhada
        if not projects_text:
            errors.append("- Campo 'Projetos' está vazio")
        if not domains_text:
            errors.append("- Campo 'Domínios' está vazio")
        if not vfs_text:
            errors.append("- Campo 'VFs' está vazio")
            
        # Se houver erros, mostra todos de uma vez
        if errors:
            messagebox.showerror("Erro de Validação", "Por favor, corrija os seguintes erros:\n\n" + "\n".join(errors))
            return
            
        # Converte os textos em listas, removendo espaços em branco extras
        projects = [p.strip() for p in projects_text.split(',') if p.strip()]
        domains = [d.strip() for d in domains_text.split(',') if d.strip()]
        use_cases = [uc.strip() for uc in use_cases_text.split(',') if uc.strip()]
        vfs = [vf.strip() for vf in vfs_text.split(',') if vf.strip()]
        
        # Debug - mostra os valores que serão usados
        print("Valores que serão usados na macro:")
        print(f"Projetos: {projects}")
        print(f"Domínios: {domains}")
        print(f"Use Cases: {use_cases}")
        print(f"VFs: {vfs}")
        print(f"DOORS: {doors}")
        
        # Atualiza a interface
        self.running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_label.config(text="Executando...")
        
        # Inicia a macro em uma thread separada
        thread = threading.Thread(target=self.run_macro, args=(projects, domains, use_cases, vfs, doors))
        thread.daemon = True
        thread.start()
        
    def stop_macro(self):
        self.running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="Parado pelo usuário")
        
    def run_macro(self, projects, domains, use_cases, vfs, doors):
        try:
            main_macro(projects, domains, use_cases, vfs, doors)
            if self.running:  # Só atualiza se não foi parado pelo usuário
                self.root.after(0, self.macro_completed)
        except Exception as error:
            if self.running:  # Só atualiza se não foi parado pelo usuário
                error_msg = str(error)
                self.root.after(0, lambda msg=error_msg: self.macro_error(msg))
                
    def macro_completed(self):
        self.running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="Concluído com sucesso!")
        messagebox.showinfo("Concluído", "Macro executada com sucesso!")
        
    def macro_error(self, error_message):
        self.running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="Erro durante a execução")
        messagebox.showerror("Erro", f"Ocorreu um erro durante a execução:\n{error_message}")
        
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MacroGUI()
    app.run() 