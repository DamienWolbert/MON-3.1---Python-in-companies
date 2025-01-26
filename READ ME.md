A noter

Sauvegarde complète : Même si tu modifies une seule cellule, pandas réécrit toute la feuille Excel. Pour modifier une cellule sans affecter le reste, il faut utiliser directement openpyxl (voir ci-dessous).

from openpyxl import load_workbook

# Charger le fichier Excel
chemin_fichier = 'chemin/vers/ton_fichier.xlsx'
wb = load_workbook(chemin_fichier)
sheet = wb['Sheet1']  # Charger la feuille "Sheet1"


On ne peux pas lancer une macro qui modifie le fichier excel sur lequel il y la macro

Si vous utilisez openpyxl.load_workbook() pour charger un fichier Excel .xlsm, le code que j'ai proposé avec Pandas ne fonctionnera pas directement sans quelques ajustements. Voici pourquoi et ce que vous pouvez faire pour obtenir les informations sur la dernière ligne et colonne remplies avec openpyxl.

Pourquoi ça ne fonctionne pas directement :
Pandas fonctionne très bien avec les fichiers .xlsx (formats Excel standard), mais quand vous utilisez openpyxl avec des fichiers .xlsm, vous travaillez directement avec le fichier au niveau des feuilles Excel, ce qui nécessite d'interagir avec la structure des cellules et des feuilles.

openpyxl permet de travailler avec les fichiers .xlsm (et d'autres types de fichiers Excel), mais pour obtenir la dernière ligne ou colonne remplie, il faut traiter les données au niveau de la feuille de calcul de manière spécifique.

Ne pas intérompre le programme avant d'avoir enregistré => Fichier endommagé irrécupérable
