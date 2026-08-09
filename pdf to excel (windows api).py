import win32com.client as win32

# Le chemin du pdf 
chemin_pdf = r"C:\Users\hsour\OneDrive\Desktop\Data\Bkam\note_d__information_EMPLOI.pdf"

excel = win32.Dispatch("Excel.Application")
excel.Visible = True 
wb = excel.Workbooks.Add()


page_index = 1


while True:
    
    nom_page = f"Page{page_index:03d}" 
    nom_requete = f"Extraction_{nom_page}"
    
    formule_m = f'let Source = Pdf.Tables(File.Contents("{chemin_pdf}"), [Implementation="1.3"]), Tableau_Cible = Source{{[Id="{nom_page}"]}}[Data] in Tableau_Cible'
    
    try:
        # 1. Création de l'onglet Excel 
        if page_index == 1:
            ws = wb.Worksheets(1) 
        else:
            # Ajout d'une nouvelle feuille 
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            
        ws.Name = nom_page # Renomme l'onglet 

        wb.Queries.Add(Name=nom_requete, Formula=formule_m)
        chaine_connexion = f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={nom_requete};Extended Properties=""'
        
        wb.Connections.Add2(
            Name=f"Query - {nom_requete}", 
            Description=f"Connexion à la {nom_page}", 
            ConnectionString=chaine_connexion, 
            CommandText=f'SELECT * FROM [{nom_requete}]', 
            lCmdtype=2
        )

        # 3. Création du tableau visuel
        list_object = ws.ListObjects.Add(0, [chaine_connexion], True, 1, ws.Range("A1"))
        list_object.QueryTable.CommandType = 2 
        list_object.QueryTable.CommandText = [f"SELECT * FROM [{nom_requete}]"]

        # 4. Actualisation 
        list_object.QueryTable.Refresh(BackgroundQuery=False)
        
        print(f"{nom_page} extraite dans l'onglet '{nom_page}' !")
        page_index += 1

    except Exception as e:
        # Si on arrive ici, c'est que la page demandée n'existe pas 
        print(f"\nFin du document atteinte après {page_index - 1} page(s).")
        
        # Nettoyage : On supprime l'onglet vide qui vient d'être créé en trop
        excel.DisplayAlerts = False 
        if page_index > 1:
            ws.Delete()
        excel.DisplayAlerts = True
        
        break 

print("\nOpération terminée !")