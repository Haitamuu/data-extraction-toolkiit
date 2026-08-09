import tabula
import pandas as pd
import warnings

#warnings.filterwarnings('ignore')

doc = "doc.pdf"
conversion = "doc.xlsx"


liste_dfs = tabula.read_pdf(
    doc, 
    pages='8', 
    multiple_tables=True, 
    stream=False,
    guess = False 
)


if liste_dfs:
    with pd.ExcelWriter(conversion, engine='openpyxl') as writer:
        for i, df in enumerate(liste_dfs):
            
            for col in df.columns:
                
                
                colonne_originale = df[col]
                
                colonne_nombres = pd.to_numeric(df[col].astype(str).str.replace('  ', '', regex=False), errors='coerce')
                
                df[col] = colonne_nombres.fillna(colonne_originale)
            
            nom_onglet = f"Tableau_{i+1}"
            df.to_excel(writer, sheet_name=nom_onglet, index=False)


print(liste_dfs)



