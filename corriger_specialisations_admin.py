#!/usr/bin/env python3
"""
Correction: Afficher la feuille Spécialisations pour l'admin
Réimporte les modules VBA corrigés dans PLANNING.xlsm
"""

import sys
from pathlib import Path
import oletools.olevba as olevba

# Chemin du fichier Excel
fichier_planning = Path(__file__).parent / "PLANNING.xlsm"

if not fichier_planning.exists():
    print(f"❌ ERREUR: Fichier {fichier_planning} introuvable")
    sys.exit(1)

print(f"📂 Ouverture du fichier: {fichier_planning}")

# Modules VBA à réimporter
modules_a_reimporter = {
    "Module_Config": "vba-modules/Module_Config.bas",
    "Module_Authentification": "vba-modules/Module_Authentification.bas"
}

try:
    # Charger le fichier XLSM
    vba = olevba.VBA_Parser(str(fichier_planning))

    print("\n📝 Modules VBA actuels dans le fichier:")
    for (filename, stream_path, vba_filename, vba_code) in vba.extract_all_macros():
        if vba_filename:
            print(f"  - {vba_filename}")

    print("\n🔄 Réimportation des modules corrigés...")

    for module_name, module_path in modules_a_reimporter.items():
        module_file = Path(__file__).parent / module_path

        if not module_file.exists():
            print(f"  ⚠️  Fichier {module_path} introuvable, ignoré")
            continue

        print(f"  ✅ Lecture de {module_path}")

        with open(module_file, 'r', encoding='utf-8') as f:
            code = f.read()

        # Note: oletools.olevba ne permet pas facilement de modifier le VBA
        # Il faut utiliser win32com ou openpyxl avec xlwings
        print(f"  ℹ️  Module {module_name} lu ({len(code)} caractères)")

    vba.close()

    print("\n" + "="*70)
    print("⚠️  IMPORTANT: La modification automatique du VBA nécessite:")
    print("   1. Soit win32com (Windows seulement)")
    print("   2. Soit une réimportation manuelle dans Excel VBA")
    print("="*70)

    print("\n📋 INSTRUCTIONS MANUELLES:")
    print("   1. Ouvrez PLANNING.xlsm dans Excel")
    print("   2. Appuyez sur Alt+F11 pour ouvrir l'éditeur VBA")
    print("   3. Double-cliquez sur 'Module_Config' dans la liste")
    print("   4. Vérifiez que la ligne suivante existe après FEUILLE_CONFIG:")
    print('      Public Const FEUILLE_SPECIALISATIONS As String = "Spécialisations"')
    print("   5. Double-cliquez sur 'Module_Authentification'")
    print("   6. Cherchez la fonction 'AfficherToutesFeuillesAdmin'")
    print("   7. Vérifiez que la ligne suivante existe:")
    print('      ThisWorkbook.Sheets(FEUILLE_SPECIALISATIONS).Visible = xlSheetVisible')
    print("   8. Enregistrez (Ctrl+S) et fermez l'éditeur VBA")
    print("   9. Reconnectez-vous en tant qu'admin")
    print("\n✅ Les fichiers .bas sont déjà corrigés dans le dossier vba-modules/")

except Exception as e:
    print(f"\n❌ ERREUR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
