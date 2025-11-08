import openpyxl

# Ouvrir le fichier Excel
wb = openpyxl.load_workbook('PLANNING_MUSEE_TEST.xlsm', keep_vba=True)

# Afficher tous les noms d'onglets
print("=== NOMS DES ONGLETS DANS LE FICHIER ===")
for i, sheet in enumerate(wb.sheetnames, 1):
    print(f"{i}. [{sheet}]")
    
wb.close()
