import os
import pandas as pd
from tkinter import Tk, Label, Button, filedialog, messagebox, Frame

class ExcelUnitiApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Uniti - Realizzato da ELpythonEMI")
        self.root.geometry("400x300")
        self.root.configure(bg="#f0f8ff")  

        self.folder_path = ""
        self.output_path = ""

       
        header_frame = Frame(root, bg="#4682b4")
        header_frame.pack(fill="x")
        Label(header_frame, text="📊 Excel Uniti", font=("Helvetica", 16, "bold"), bg="#4682b4", fg="white").pack(pady=10)

        
        Label(root, text="Unisci i dati di più file Excel in un unico file!", bg="#f0f8ff", font=("Helvetica", 12)).pack(pady=10)
        Button(root, text="📂 Seleziona cartella con file Excel", command=self.select_folder, bg="#add8e6", font=("Helvetica", 10, "bold")).pack(pady=5)
        Button(root, text="💾 Scegli dove salvare il file unito", command=self.select_output_file, bg="#add8e6", font=("Helvetica", 10, "bold")).pack(pady=5)
        Button(root, text="🔗 Unisci i file", command=self.merge_files, bg="#32cd32", fg="white", font=("Helvetica", 10, "bold")).pack(pady=15)

       
        Label(root, text="Realizzato da ELpythonEMI", font=("Helvetica", 10, "italic"), bg="#f0f8ff", fg="#696969").pack(side="bottom", pady=10)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            messagebox.showinfo("Cartella selezionata", f"Cartella: {self.folder_path}")

    def select_output_file(self):
        self.output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                        filetypes=[("Excel Files", "*.xlsx")])
        if self.output_path:
            messagebox.showinfo("File di uscita selezionato", f"Salverai il file unito qui: {self.output_path}")

    def merge_files(self):
        if not self.folder_path or not self.output_path:
            messagebox.showerror("Errore", "Devi selezionare sia la cartella sia il file di uscita.")
            return

        try:
            all_data = []
            for file_name in os.listdir(self.folder_path):
                if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
                    file_path = os.path.join(self.folder_path, file_name)
                    df = pd.read_excel(file_path)
                    all_data.append(df)

            if not all_data:
                messagebox.showerror("Errore", "Nessun file Excel trovato nella cartella selezionata.")
                return

           
            combined_df = pd.concat(all_data, ignore_index=True)
            combined_df.to_excel(self.output_path, index=False)

            messagebox.showinfo("Successo", "🎉 File Excel unito salvato con successo!")
        except Exception as e:
            messagebox.showerror("Errore", f"Si è verificato un errore: {e}")


if __name__ == "__main__":
    root = Tk()
    app = ExcelUnitiApp(root)
    root.mainloop()
