import win32com.client

# Chemin vers ton fichier Excel"
fichier_excel = "./Tests.xlsm"


# Nom de la macro (doit inclure le nom du module si nécessaire, ex : Module1.MaMacro)
nom_macro = "ID"

# Ouvrir Excel via COM
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  # Mettre sur True si tu veux voir Excel se lancer

try:
    # Ouvrir le fichier Excel
    wb = excel.Workbooks.Open(fichier_excel)
    
    # Lancer la macro
    excel.Application.Run(nom_macro)
    
    # Fermer le fichier sans enregistrer
    wb.Close(SaveChanges=True)
except Exception as e:
    print(f"Erreur : {e}")
finally:
    # Quitter Excel
    excel.Quit()

print("Macro exécutée avec succès.")