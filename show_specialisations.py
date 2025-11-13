#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rend visible la feuille Spécialisations
"""

from openpyxl import load_workbook

# Ouvrir le fichier
wb = load_workbook('PLANNING.xlsm', keep_vba=True)

# Rendre visible la feuille Spécialisations
ws = wb['Spécialisations']
ws.sheet_state = 'visible'

# Sauvegarder
wb.save('PLANNING.xlsm')
wb.close()

print("✅ Feuille 'Spécialisations' maintenant VISIBLE !")
print()
print("La cliente pourra maintenant la voir dans les onglets Excel.")
