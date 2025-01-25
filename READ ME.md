A noter

Sauvegarde complète : Même si tu modifies une seule cellule, pandas réécrit toute la feuille Excel. Pour modifier une cellule sans affecter le reste, il faut utiliser directement openpyxl (voir ci-dessous).

from openpyxl import load_workbook

# Charger le fichier Excel
chemin_fichier = 'chemin/vers/ton_fichier.xlsx'
wb = load_workbook(chemin_fichier)
sheet = wb['Sheet1']  # Charger la feuille "Sheet1"