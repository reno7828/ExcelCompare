import pandas as pd
from tkinter import Tk, filedialog, messagebox, Toplevel, Text, Scrollbar, END

def display_differences(differences):
    # Créer une nouvelle fenêtre pour afficher les différences
    diff_window = Toplevel()
    diff_window.title("Différences entre les fichiers Excel")

    # Créer une zone de texte pour afficher les différences
    text_widget = Text(diff_window, wrap="word", height=20, width=80)
    text_widget.pack(padx=20, pady=20)

    # Ajouter les différences à la zone de texte
    text_widget.insert(END, differences)

    # Désactiver l'édition dans la zone de texte
    text_widget.configure(state='disabled')

    # Afficher la fenêtre
    diff_window.mainloop()


def compare_excel():
    # Demander à l'utilisateur de choisir les fichiers Excel
    root = Tk()
    root.withdraw()  # Cacher la fenêtre principale
    file_path1 = filedialog.askopenfilename(title="Sélectionner le premier fichier Excel")
    file_path2 = filedialog.askopenfilename(title="Sélectionner le deuxième fichier Excel")

    # Vérifier si les fichiers ont été sélectionnés
    if not file_path1 or not file_path2:
        messagebox.showwarning("Avertissement", "Veuillez sélectionner deux fichiers Excel.")
        return

    # Lire les fichiers Excel
    df1 = pd.read_excel(file_path1)
    df2 = pd.read_excel(file_path2)

    # Comparer les données
    differences = []

    for row in range(max(len(df1), len(df2))):
        for col in range(max(len(df1.columns), len(df2.columns))):
            # Obtenir les valeurs des cellules pour chaque fichier
            cell1 = df1.iat[row, col] if row < len(df1) and col < len(df1.columns) else None
            cell2 = df2.iat[row, col] if row < len(df2) and col < len(df2.columns) else None

            if cell1 != cell2:
                differences.append(f"Différence à la ligne {row + 2}, colonne {col + 1}: {cell1} vs {cell2}")

    # Afficher les différences
    if differences:
        differences_text = "\n".join(differences)
        print(differences_text)
        display_differences(differences_text)
    else:
        messagebox.showinfo("Comparaison des fichiers Excel", "Les fichiers sont identiques.")

# Appeler la fonction pour comparer les fichiers
compare_excel()
