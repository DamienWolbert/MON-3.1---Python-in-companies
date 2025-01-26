import pandas as pd
import matplotlib.pyplot as plt

# Charger le fichier Excel
file_path = "Tests 6.xlsm"

# Lire toutes les feuilles pour déterminer où se trouvent les données
xls = pd.ExcelFile(file_path)

# Affiche les noms de feuilles pour s'assurer que nous lisons la bonne
print("Feuilles disponibles :", xls.sheet_names)

# Charger les données d'une feuille spécifique (à adapter si besoin)
data = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

# Afficher un aperçu des données pour vérifier leur structure
print(data.head())

# Identifier la ligne 'Poids' et utiliser ses valeurs comme axe des abscisses
poids = data.iloc[0, 1:].values  # Supposons que la ligne Poids est la première

# Boucler sur les autres lignes et créer un graphique pour chaque
for i in range(1, len(data)):
    ligne = data.iloc[i, 1:].values  # Toutes les colonnes sauf la première
    label = data.iloc[i, 0]  # Utiliser la première colonne comme label

    # Créer un graphique
    plt.figure()
    plt.plot(poids, ligne, marker='o', label=label)
    plt.title(f"Graphique pour {label}")
    plt.xlabel("Poids")
    plt.ylabel("Valeur")
    plt.legend()
    plt.grid()

    # Sauvegarder ou afficher le graphique
    plt.savefig(f"graphique_{label}.png")
    plt.close()

print("Graphiques générés !")

