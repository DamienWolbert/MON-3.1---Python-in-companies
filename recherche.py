import pandas as pd
from openpyxl import load_workbook
import time



path_excel = r"C:\Users\damie\Documents\GitHub\MON-3.1---Python-in-companies\Tests 5.xlsm"
nom_donnee = "Données 1"
nom_resultat = "Sommaire"

def derniere_ligne_remplie(sheet):
    """
    Retourne le numéro de la dernière ligne remplie dans la feuille.
    """
    derniere_ligne = 0
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        for cell in row:
            if cell.value is not None:
                derniere_ligne = max(derniere_ligne, cell.row)
    return derniere_ligne

def derniere_colonne_remplie(sheet):
    """
    Retourne le numéro de la dernière colonne remplie dans la feuille.
    """
    derniere_colonne = 0
    for col in sheet.iter_cols(min_col=1, max_col=sheet.max_column):
        for cell in col:
            if cell.value is not None:
                derniere_colonne = max(derniere_colonne, cell.column)
    return derniere_colonne

print("Ouverture")

#Ouverture de la feuille
fichier = load_workbook(path_excel,keep_vba=True)
feuille = fichier[nom_donnee]
data = fichier[nom_resultat]
df = pd.DataFrame(data, columns=columns)

print("Dimensions")

temps_1 = time.time()

derniere_ligne = derniere_ligne_remplie(data)
derniere_colonne = derniere_colonne_remplie(data)

temps_2 = time.time()

compteur = 0
somme = 0

print("Parcours")

# Complétion des données
compteur = 0
for i in range(1,10):#last_row+1):
    for j in range(1,10):#last_column):
        if type(feuille.cell(i,j).value)!= str:
            compteur += 1
            somme += feuille.cell(i,j).value

temps_3 = time.time()

print("Reporting")

data.cell(11,3).value = compteur
data.cell(12,3).value = somme
data.cell(13,3).value = temps_2 - temps_1
data.cell(14,3).value = temps_3 - temps_2
data.cell(15,3).value = temps_3 - temps_1

#Sauvegrade des modificaitons
fichier.save(path_excel)

print("FIN")
