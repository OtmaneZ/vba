#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configure Spécialisations : veryHidden par défaut, l'admin l'affichera via VBA
"""

from openpyxl import load_workbook

# Ouvrir le fichier
wb = load_workbook('PLANNING.xlsm', keep_vba=True)

# Mettre Spécialisations en veryHidden (cachée, seul VBA peut l'afficher)
ws = wb['Spécialisations']
ws.sheet_state = 'veryHidden'

# Sauvegarder
wb.save('PLANNING.xlsm')
wb.close()

print("✅ Feuille 'Spécialisations' configurée :")
print("   - Par défaut : CACHÉE (veryHidden)")
print("   - L'admin pourra la voir après connexion (VBA l'affiche)")
print()
print("📝 Prochaine étape : Recopier Module_Authentification.bas dans Excel")
