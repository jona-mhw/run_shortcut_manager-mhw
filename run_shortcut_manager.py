import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import winreg
import os

class RunShortcutManager:
    def __init__(self, master):
        self.master = master
        master.title("Administrador de Accesos Directos de Ejecutar / Run Shortcut Manager")
        master.geometry("800x600")

        # Main Frame
        main_frame = tk.Frame(master)
        main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Existing Shortcuts Section
        tk.Label(main_frame, text="Accesos Directos Existentes / Existing Shortcuts", font=("Arial", 14)).pack()

        # Treeview to display existing shortcuts
        self.tree = ttk.Treeview(main_frame, columns=('Nombre', 'Ruta', 'Tipo'), show='headings')
        self.tree.heading('Nombre', text='Nombre / Name')
        self.tree.heading('Ruta', text='Ruta / Path')
        self.tree.heading('Tipo', text='Tipo / Type')
        self.tree.pack(fill=tk.BOTH, expand=True, pady=10)

        # Scrollbar for Treeview
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Add Shortcut Section
        tk.Label(main_frame, text="Agregar Nuevo Acceso Directo / Add New Shortcut", font=("Arial", 12)).pack()

        # Shortcut Name
        tk.Label(main_frame, text="Nombre del Acceso Directo / Shortcut Name:").pack()
        self.shortcut_name_var = tk.StringVar()
        self.shortcut_name_entry = tk.Entry(main_frame, textvariable=self.shortcut_name_var, width=30)
        self.shortcut_name_entry.pack()

        # Shortcut Path
        tk.Label(main_frame, text="Ruta del Ejecutable / Executable Path:").pack()
        self.shortcut_path_var = tk.StringVar()
        self.shortcut_path_entry = tk.Entry(main_frame, textvariable=self.shortcut_path_var, width=50)
        self.shortcut_path_entry.pack()

        # Browse Button
        tk.Button(main_frame, text="Buscar / Browse", command=self.browse_executable).pack()

        # Add Shortcut Button
        tk.Button(main_frame, text="Agregar Acceso Directo / Add Shortcut", command=self.add_shortcut).pack(pady=10)

        # Buttons for Actions
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="Actualizar Lista / Refresh List", command=self.load_existing_shortcuts).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Eliminar Seleccionado / Remove Selected", command=self.remove_shortcut).pack(side=tk.LEFT, padx=5)

        # Initial load of existing shortcuts
        self.load_existing_shortcuts()

    def browse_executable(self):
        """Permite al usuario buscar un archivo ejecutable / Allows user to browse for an executable"""
        filename = filedialog.askopenfilename(
            title="Seleccionar Ejecutable / Select Executable",
            filetypes=[("Ejecutables", "*.exe"), ("Accesos Directos", "*.lnk"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.shortcut_path_var.set(filename)

    def load_existing_shortcuts(self):
        """Cargar accesos directos existentes / Load existing shortcuts"""
        # Clear existing items
        for i in self.tree.get_children():
            self.tree.delete(i)

        try:
            # Open the App Paths key
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\App Paths")
            
            # Enumerate subkeys
            index = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, index)
                    
                    # Open the subkey
                    subkey = winreg.OpenKey(key, subkey_name)
                    
                    # Try to get the default value (executable path)
                    try:
                        path, _ = winreg.QueryValueEx(subkey, "")
                        
                        # Determine shortcut type
                        shortcut_type = "Ejecutable" if path.lower().endswith('.exe') else "Acceso Directo"
                        
                        # Insert into treeview
                        self.tree.insert('', 'end', values=(
                            subkey_name.replace('.exe', ''), 
                            path, 
                            shortcut_type
                        ))
                    except FileNotFoundError:
                        pass
                    
                    winreg.CloseKey(subkey)
                    index += 1
                except OSError:
                    break
            
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los accesos directos:\n{str(e)}")

    def add_shortcut(self):
        """Agregar un nuevo acceso directo / Add a new shortcut"""
        shortcut_name = self.shortcut_name_var.get().strip()
        shortcut_path = self.shortcut_path_var.get().strip()

        # Validaciones básicas / Basic validations
        if not shortcut_name:
            messagebox.showerror("Error", "Ingrese un nombre de acceso directo / Enter a shortcut name")
            return
        
        if not shortcut_path or not os.path.exists(shortcut_path):
            messagebox.showerror("Error", "Seleccione un archivo válido / Select a valid file")
            return

        try:
            # Crear entrada de registro / Create registry entry
            key_path = rf"Software\Microsoft\Windows\CurrentVersion\App Paths\{shortcut_name}.exe"
            
            # Abrir la clave de registro / Open registry key
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            
            # Establecer valores / Set values
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, shortcut_path)
            winreg.SetValueEx(key, "Path", 0, winreg.REG_SZ, os.path.dirname(shortcut_path))
            
            # Cerrar la clave / Close key
            winreg.CloseKey(key)

            # Actualizar lista / Refresh list
            self.load_existing_shortcuts()

            # Limpiar campos / Clear fields
            self.shortcut_name_var.set("")
            self.shortcut_path_var.set("")

            # Mensaje de éxito / Success message
            messagebox.showinfo("Éxito / Success", 
                f"Acceso directo '{shortcut_name}' creado exitosamente.\n"
                f"Shortcut '{shortcut_name}' created successfully.")

        except Exception as e:
            messagebox.showerror("Error", 
                f"No se pudo crear el acceso directo:\n"
                f"Could not create shortcut:\n{str(e)}")

    def remove_shortcut(self):
        """Eliminar el acceso directo seleccionado / Remove selected shortcut"""
        # Obtener el elemento seleccionado / Get selected item
        selected_item = self.tree.selection()
        
        if not selected_item:
            messagebox.showwarning("Advertencia / Warning", 
                "Seleccione un acceso directo para eliminar.\n"
                "Select a shortcut to remove.")
            return

        # Obtener el nombre del acceso directo / Get shortcut name
        shortcut_name = self.tree.item(selected_item[0])['values'][0]

        try:
            # Eliminar la clave de registro / Delete registry key
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, 
                rf"Software\Microsoft\Windows\CurrentVersion\App Paths\{shortcut_name}.exe")

            # Actualizar lista / Refresh list
            self.load_existing_shortcuts()

            messagebox.showinfo("Éxito / Success", 
                f"Acceso directo '{shortcut_name}' eliminado.\n"
                f"Shortcut '{shortcut_name}' removed.")

        except Exception as e:
            messagebox.showerror("Error", 
                f"No se pudo eliminar el acceso directo:\n"
                f"Could not remove shortcut:\n{str(e)}")

def main():
    root = tk.Tk()
    app = RunShortcutManager(root)
    root.mainloop()

if __name__ == "__main__":
    main()
