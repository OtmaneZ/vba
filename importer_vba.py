#!/usr/bin/env python3
"""
Script pour importer automatiquement le code VBA dans le fichier Excel enrichi
Utilise win32com sur Windows ou applescript sur macOS
"""

import os
import sys
import platform
from pathlib import Path

# Chemins
EXCEL_FILE = "PLANNING_MUSEE_ENRICHI.xlsx"
VBA_MODULES_DIR = "vba-modules"

# Modules VBA à importer
MODULES_BAS = [
    "Module_Accueil.bas",
    "Module_Authentification.bas",
    "Module_Calculs.bas",
    "Module_Config.bas",
    "Module_Contrats.bas",
    "Module_Disponibilites.bas",
    "Module_Emails.bas",
    "Module_Planning.bas"
]

# Classes VBA (code à insérer dans les objets feuilles)
CLASSES_CLS = {
    "ThisWorkbook.cls": "ThisWorkbook",
    "Feuille_Accueil.cls": "Feuille1",  # Accueil
    "Feuille_Visites.cls": "Feuille4"   # Visites
}


def importer_vba_macos():
    """
    Import VBA sur macOS en utilisant AppleScript
    """
    print("🍎 Détection macOS - Utilisation d'AppleScript")

    excel_path = os.path.abspath(EXCEL_FILE)
    xlsm_path = excel_path.replace('.xlsx', '.xlsm')

    print(f"\n📂 Fichier Excel : {xlsm_path}")
    print(f"📁 Dossier VBA : {VBA_MODULES_DIR}")

    # Générer le code pour créer les modules
    modules_code = []
    for module_name in MODULES_BAS:
        module_path = os.path.join(VBA_MODULES_DIR, module_name)
        if not os.path.exists(module_path):
            print(f"❌ Module introuvable : {module_path}")
            continue

        with open(module_path, 'r', encoding='utf-8') as f:
            vba_code = f.read()

        # Nettoyer le code (enlever l'en-tête Attribute VB_Name si présent)
        lines = vba_code.split('\n')
        clean_lines = [line for line in lines if not line.startswith('Attribute VB_Name')]
        clean_code = '\n'.join(clean_lines)

        module_name_only = module_name.replace('.bas', '')
        modules_code.append({
            'name': module_name_only,
            'code': clean_code
        })

    print(f"\n✅ {len(modules_code)} modules prêts à importer")

    # Afficher les instructions manuelles
    print("\n" + "="*60)
    print("📋 INSTRUCTIONS POUR IMPORT MANUEL (Excel macOS)")
    print("="*60)
    print("\n1️⃣  Ouvrir Excel VBA Editor : Alt+F11 (ou Tools > Macro > Visual Basic Editor)")
    print("\n2️⃣  Pour chaque MODULE (.bas) :")
    print("    → Insert > Module")
    print("    → Copier-coller le contenu du fichier .bas")
    print("\n3️⃣  Pour les CLASSES (.cls) :")
    print("    → Double-cliquer sur l'objet de feuille correspondant :")
    print("      • ThisWorkbook → Coller le contenu de ThisWorkbook.cls")
    print("      • Feuille1 (Accueil) → Coller le contenu de Feuille_Accueil.cls")
    print("      • Feuille4 (Visites) → Coller le contenu de Feuille_Visites.cls")
    print("\n4️⃣  Sauvegarder : Ctrl+S ou File > Save")
    print("="*60)

    # Créer un fichier résumé avec tout le code
    output_file = "VBA_CODE_COMPLET.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("CODE VBA COMPLET À COPIER-COLLER DANS EXCEL\n")
        f.write("="*80 + "\n\n")

        f.write("MODULES STANDARDS (.bas)\n")
        f.write("-"*80 + "\n\n")

        for module_info in modules_code:
            f.write(f"\n{'='*80}\n")
            f.write(f"MODULE : {module_info['name']}\n")
            f.write(f"{'='*80}\n\n")
            f.write(module_info['code'])
            f.write("\n\n")

        f.write("\n" + "="*80 + "\n")
        f.write("CLASSES / OBJETS FEUILLES (.cls)\n")
        f.write("-"*80 + "\n\n")

        for cls_file, object_name in CLASSES_CLS.items():
            cls_path = os.path.join(VBA_MODULES_DIR, cls_file)
            if os.path.exists(cls_path):
                with open(cls_path, 'r', encoding='utf-8') as cf:
                    cls_code = cf.read()

                f.write(f"\n{'='*80}\n")
                f.write(f"CLASSE : {cls_file} → À copier dans {object_name}\n")
                f.write(f"{'='*80}\n\n")
                f.write(cls_code)
                f.write("\n\n")

    print(f"\n✅ Fichier créé : {output_file}")
    print("   Ce fichier contient TOUT le code VBA pour copier-coller facilement")

    return True


def importer_vba_windows():
    """
    Import VBA sur Windows en utilisant win32com
    """
    print("🪟 Détection Windows - Utilisation de win32com")

    try:
        import win32com.client
    except ImportError:
        print("❌ Module pywin32 non installé")
        print("   Installation : pip install pywin32")
        return False

    excel_path = os.path.abspath(EXCEL_FILE)
    xlsm_path = excel_path.replace('.xlsx', '.xlsm')

    if not os.path.exists(xlsm_path):
        print(f"❌ Fichier introuvable : {xlsm_path}")
        print("   Assurez-vous que le fichier est en .xlsm (pas .xlsx)")
        return False

    print(f"📂 Ouverture de {xlsm_path}...")

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(xlsm_path)
        vbproj = wb.VBProject

        # Importer les modules .bas
        print("\n📥 Import des modules...")
        for module_name in MODULES_BAS:
            module_path = os.path.abspath(os.path.join(VBA_MODULES_DIR, module_name))
            if os.path.exists(module_path):
                vbproj.VBComponents.Import(module_path)
                print(f"   ✅ {module_name}")
            else:
                print(f"   ❌ {module_name} introuvable")

        # Pour les classes, il faut les insérer dans les objets existants
        print("\n📝 Insertion du code dans les objets feuilles...")
        for cls_file, object_name in CLASSES_CLS.items():
            cls_path = os.path.join(VBA_MODULES_DIR, cls_file)
            if os.path.exists(cls_path):
                with open(cls_path, 'r', encoding='utf-8') as f:
                    cls_code = f.read()

                # Trouver l'objet correspondant
                try:
                    component = vbproj.VBComponents(object_name)
                    code_module = component.CodeModule

                    # Effacer le code existant
                    if code_module.CountOfLines > 0:
                        code_module.DeleteLines(1, code_module.CountOfLines)

                    # Insérer le nouveau code
                    code_module.AddFromString(cls_code)
                    print(f"   ✅ {cls_file} → {object_name}")
                except Exception as e:
                    print(f"   ❌ {cls_file} : {e}")

        print("\n💾 Sauvegarde...")
        wb.Save()

        print("\n✅ Import VBA terminé avec succès !")
        print(f"📁 Fichier : {xlsm_path}")

        return True

    except Exception as e:
        print(f"❌ Erreur : {e}")
        return False


def main():
    print("="*60)
    print("🔧 IMPORT AUTOMATIQUE DU CODE VBA")
    print("="*60)

    # Vérifier que les fichiers existent
    if not os.path.exists(VBA_MODULES_DIR):
        print(f"❌ Dossier {VBA_MODULES_DIR} introuvable")
        return 1

    # Convertir .xlsx en .xlsm si nécessaire
    excel_path = os.path.abspath(EXCEL_FILE)
    xlsm_path = excel_path.replace('.xlsx', '.xlsm')

    if os.path.exists(excel_path) and not os.path.exists(xlsm_path):
        print(f"\n📝 Renommage : {EXCEL_FILE} → {os.path.basename(xlsm_path)}")
        os.rename(excel_path, xlsm_path)
        print("   ✅ Fichier renommé en .xlsm (format avec macros)")

    # Détecter le système d'exploitation
    system = platform.system()

    if system == "Darwin":  # macOS
        success = importer_vba_macos()
    elif system == "Windows":
        success = importer_vba_windows()
    else:
        print(f"❌ Système non supporté : {system}")
        print("   Ce script fonctionne sur macOS et Windows uniquement")
        return 1

    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
