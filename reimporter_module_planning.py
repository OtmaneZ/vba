#!/usr/bin/env python3
"""
Réimporter Module_Planning.bas corrigé dans PLANNING.xlsm
"""
import zipfile
import shutil
from pathlib import Path
import tempfile

fichier_excel = Path("PLANNING.xlsm")
module_corrige = Path("vba-modules/Module_Planning.bas")

print("="*80)
print("RÉIMPORT DE MODULE_PLANNING.BAS DANS EXCEL")
print("="*80)

# Backup
backup = Path("PLANNING_backup_avant_fix_planning.xlsm")
shutil.copy2(fichier_excel, backup)
print(f"\n✅ Backup: {backup.name}")

# Lire le module corrigé
print(f"\n📂 Lecture de {module_corrige}...")
with open(module_corrige, 'r', encoding='utf-8') as f:
    code_corrige = f.read()

print(f"   ✅ {len(code_corrige)} caractères lus")

# Extraire le XLSM
print(f"\n📦 Extraction de {fichier_excel.name}...")
with tempfile.TemporaryDirectory() as tmpdir:
    tmpdir = Path(tmpdir)
    
    # Extraire tout
    with zipfile.ZipFile(fichier_excel, 'r') as zin:
        zin.extractall(tmpdir)
    
    print("   ✅ Extrait")
    
    # Trouver le fichier VBA
    vba_bin = tmpdir / "xl" / "vbaProject.bin"
    
    if vba_bin.exists():
        print(f"\n⚠️  VBA binaire trouvé: {vba_bin.name}")
        print("   ⚠️  Modification directe du VBA impossible via Python")
        print("\n📝 SOLUTION MANUELLE REQUISE:")
        print("="*80)
        print("1. Ouvre PLANNING.xlsm dans Excel")
        print("2. Alt+F11 (ou Cmd+F11 sur Mac) pour ouvrir VBA")
        print("3. Double-clique sur 'Module_Planning' dans la liste")
        print("4. Sélectionne TOUT le code (Cmd+A)")
        print("5. Copie le code depuis vba-modules/Module_Planning.bas")
        print("6. Colle dans Excel VBA")
        print("7. Sauvegarde (Cmd+S) et ferme VBA")
        print("8. Ferme et rouvre PLANNING.xlsm")
        print("9. Lance la génération du planning")
        print("="*80)
        print("\n🔧 CORRECTIONS APPLIQUÉES:")
        print("   • ObtenirGuidesDisponibles: Lit Date en col 1, Dispo en col 2")
        print("   • Format heures: Format(time, 'hh:mm') au lieu de nombres")
        print("   • Colonnes Visites corrigées: Durée=col4, TypeVisite=col5, Musée=col6")
        print("="*80)
    else:
        print(f"\n❌ VBA binaire non trouvé!")

print("\n✅ Terminé")
