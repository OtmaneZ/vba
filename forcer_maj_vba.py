#!/usr/bin/env python3
"""
FORCER la mise à jour du VBA dans PLANNING.xlsm avec win32com
"""

import sys
from pathlib import Path

try:
    import win32com.client
    print("✅ win32com disponible")
except ImportError:
    print("❌ win32com non disponible sur macOS")
    print("\n🔧 SOLUTION MANUELLE OBLIGATOIRE:")
    print("="*80)
    print("1. Ouvrez PLANNING.xlsm dans Excel")
    print("2. Alt+F11 pour ouvrir VBA")
    print("3. Dans Module_Authentification, fonction AfficherToutesFeuillesAdmin")
    print("4. Cherchez la ligne:")
    print('   ThisWorkbook.Sheets("Spécialisations").Visible = xlSheetVisible')
    print("5. Remplacez par:")
    print('   ThisWorkbook.Sheets(FEUILLE_SPECIALISATIONS).Visible = xlSheetVisible')
    print("6. Ctrl+S pour sauvegarder")
    print("7. Fermez VBA et Excel")
    print("8. Rouvrez et testez")
    print("="*80)

    # Montrer le code exact
    print("\n📝 CODE EXACT À COPIER:")
    print("="*80)

    code = """Private Sub AfficherToutesFeuillesAdmin()
    On Error Resume Next

    ' Afficher toutes les feuilles pour l'admin
    ThisWorkbook.Sheets(FEUILLE_GUIDES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_VISITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CALCULS).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CONTRATS).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CONFIG).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_SPECIALISATIONS).Visible = xlSheetVisible

    On Error GoTo 0
End Sub"""

    print(code)
    print("="*80)

    sys.exit(1)
