def lancement_macro_sans_argument(path_fichier, nom_macro):
    print("Lancement de la macro "+ nom_macro + " contenue dans le ficher" + path_fichier)
    
    import win32com.client

    fichier_excel = path_fichier

    # Ouverture d'Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True # Permet de voir Excel se lancer ou non

    # Ouvrir le fichier Excel
    wb = excel.Workbooks.Open(fichier_excel)
    
    # Lancer la macro
    excel.Application.Run(nom_macro)
    
    # Fermeture du fichier excel
    wb.Close(SaveChanges=True)

    # Dermeture Excel
    excel.Quit()

    print("Macro exécutée")

def lancement_macro_avec_argument(path_fichier, nom_macro, liste_arguments):
    print("Lancement de la macro "+ nom_macro + " contenue dans le ficher" + path_fichier)
    
    import win32com.client

    fichier_excel = path_fichier

    # Ouverture d'Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True # Permet de voir Excel se lancer ou non

    # Ouvrir le fichier Excel
    wb = excel.Workbooks.Open(fichier_excel)
    
    # Lancer la macro
    excel.Application.Run(nom_macro,*liste_arguments)
    
    # Fermeture du fichier excel
    wb.Close(SaveChanges=True)

    # Dermeture Excel
    excel.Quit()

    print("Macro exécutée")