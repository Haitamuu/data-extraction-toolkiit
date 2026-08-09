import win32com.client as win32

 
chemin_pdf = r"C:\Users\Admin\Desktop\File\doc.pdf"

excel = win32.Dispatch("Excel.Application")
excel.Visible = True 
wb = excel.Workbooks.Add()


page_index = 1


while True:
    
    nom_page = f"Page{page_index:03d}" 
    nom_requete = f"Extraction_{nom_page}"
    
    formule_m = f'let Source = Pdf.Tables(File.Contents("{chemin_pdf}"), [Implementation="1.3"]), Tableau_Cible = Source{{[Id="{nom_page}"]}}[Data] in Tableau_Cible'
    
    try:
 
        if page_index == 1:
            ws = wb.Worksheets(1) 
        else:
             
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            
        ws.Name = nom_page 

        wb.Queries.Add(Name=nom_requete, Formula=formule_m)
        chaine_connexion = f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={nom_requete};Extended Properties=""'
        
        wb.Connections.Add2(
            Name=f"Query - {nom_requete}", 
            Description=f"Connexion à la {nom_page}", 
            ConnectionString=chaine_connexion, 
            CommandText=f'SELECT * FROM [{nom_requete}]', 
            lCmdtype=2
        )

        
        list_object = ws.ListObjects.Add(0, [chaine_connexion], True, 1, ws.Range("A1"))
        list_object.QueryTable.CommandType = 2 
        list_object.QueryTable.CommandText = [f"SELECT * FROM [{nom_requete}]"]

         
        list_object.QueryTable.Refresh(BackgroundQuery=False)
        
        print(f"{nom_page} extraite dans l'onglet '{nom_page}' !")
        page_index += 1

    except Exception as e:
     
        print(f"\nFin du document atteinte après {page_index - 1} page(s).")
        
       
        excel.DisplayAlerts = False 
        if page_index > 1:
            ws.Delete()
        excel.DisplayAlerts = True
        
        break 

print("\nOpération terminée !")
