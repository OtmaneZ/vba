#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VALIDATION FINALE - Vérifier que tout est prêt
"""

import os
from pathlib import Path

def valider_livraison():
    """
    Vérifie que tous les fichiers nécessaires sont présents et valides
    """

    base_dir = "/Users/otmaneboulahia/Documents/Excel-Auto"

    print("=" * 80)
    print("✅ VALIDATION FINALE - CORRECTIONS PLANNING GUIDES")
    print("=" * 80)

    checks = []

    # 1. Fichier principal
    print("\n📋 1. FICHIER PRINCIPAL")
    fichier_planning = os.path.join(base_dir, "PLANNING.xlsm")
    if os.path.exists(fichier_planning):
        size = os.path.getsize(fichier_planning)
        print(f"  ✅ PLANNING.xlsm ({size:,} bytes)")
        checks.append(True)
    else:
        print(f"  ❌ PLANNING.xlsm manquant")
        checks.append(False)

    # 2. Modules VBA corrigés
    print("\n🔧 2. MODULES VBA CORRIGÉS")
    modules = [
        "vba-modules/Module_Planning_CORRECTED.bas",
        "vba-modules/Module_Specialisations_CORRECTED.bas"
    ]

    for module in modules:
        module_path = os.path.join(base_dir, module)
        if os.path.exists(module_path):
            size = os.path.getsize(module_path)
            with open(module_path, 'r', encoding='utf-8') as f:
                lines = len(f.readlines())
            print(f"  ✅ {module.split('/')[-1]} ({lines} lignes, {size:,} bytes)")
            checks.append(True)
        else:
            print(f"  ❌ {module} manquant")
            checks.append(False)

    # 3. Scripts Python
    print("\n🐍 3. SCRIPTS PYTHON D'ANALYSE")
    scripts = [
        "analyser_planning_structure.py",
        "corriger_structure_disponibilites.py",
        "corriger_modules_vba_complet.py",
        "simuler_resultat_planning.py"
    ]

    for script in scripts:
        script_path = os.path.join(base_dir, script)
        if os.path.exists(script_path):
            size = os.path.getsize(script_path)
            print(f"  ✅ {script} ({size:,} bytes)")
            checks.append(True)
        else:
            print(f"  ❌ {script} manquant")
            checks.append(False)

    # 4. Documentation
    print("\n📚 4. DOCUMENTATION")
    docs = [
        "GUIDE_CORRECTION_COMPLET.md",
        "RESUME_EXECUTIF.md",
        "CORRECTIONS_VBA_A_APPLIQUER.md",
        "email.md"
    ]

    for doc in docs:
        doc_path = os.path.join(base_dir, doc)
        if os.path.exists(doc_path):
            size = os.path.getsize(doc_path)
            with open(doc_path, 'r', encoding='utf-8') as f:
                lines = len(f.readlines())
            print(f"  ✅ {doc} ({lines} lignes, {size:,} bytes)")
            checks.append(True)
        else:
            print(f"  ❌ {doc} manquant")
            checks.append(False)

    # 5. Backups
    print("\n💾 5. BACKUPS DE SÉCURITÉ")
    backup_dir = base_dir
    backups = [f for f in os.listdir(backup_dir) if f.startswith("PLANNING_backup_") and f.endswith(".xlsm")]

    if len(backups) >= 2:
        print(f"  ✅ {len(backups)} backups créés")
        for backup in sorted(backups)[-2:]:  # Afficher les 2 derniers
            backup_path = os.path.join(backup_dir, backup)
            size = os.path.getsize(backup_path)
            print(f"     - {backup} ({size:,} bytes)")
        checks.append(True)
    else:
        print(f"  ⚠️ Seulement {len(backups)} backup(s) trouvé(s)")
        checks.append(False)

    # 6. Validation du contenu VBA
    print("\n🔍 6. VALIDATION CONTENU VBA")

    module_planning = os.path.join(base_dir, "vba-modules/Module_Planning_CORRECTED.bas")
    if os.path.exists(module_planning):
        with open(module_planning, 'r', encoding='utf-8') as f:
            content = f.read()

        validations = [
            ('Format(heureDebut, "hh:mm")', "Format heure corrigé"),
            ('wsVisites.Cells(i, 3).Value', "Lecture col 3 (Heure)"),
            ('wsVisites.Cells(i, 6).Value', "Lecture col 6 (Type)"),
            ('wsDispo.Cells(i, 1).Value', "Lecture col 1 Dispo (Date)"),
            ('wsDispo.Cells(i, 2).Value', "Lecture col 2 Dispo (OUI/NON)"),
            ('listeGuidesDispos', "Liste guides disponibles")
        ]

        for code_snippet, description in validations:
            if code_snippet in content:
                print(f"  ✅ {description}")
                checks.append(True)
            else:
                print(f"  ❌ {description} - CODE MANQUANT")
                checks.append(False)
    else:
        print("  ❌ Impossible de valider le contenu")
        checks.append(False)

    # 7. Validation du contenu Spécialisations
    print("\n⭐ 7. VALIDATION MODULE SPÉCIALISATIONS")

    module_spec = os.path.join(base_dir, "vba-modules/Module_Specialisations_CORRECTED.bas")
    if os.path.exists(module_spec):
        with open(module_spec, 'r', encoding='utf-8') as f:
            content = f.read()

        validations = [
            ('ws.Cells(i, 2).Value', "Lecture col 2 (Nom_Guide)"),
            ('ws.Cells(i, 4).Value', "Lecture col 4 (Type_Prestation)"),
            ('ws.Cells(i, 5).Value', "Lecture col 5 (Autorise)"),
            ('If autorise = "OUI" Then', "Logique OUI/NON")
        ]

        for code_snippet, description in validations:
            if code_snippet in content:
                print(f"  ✅ {description}")
                checks.append(True)
            else:
                print(f"  ❌ {description} - CODE MANQUANT")
                checks.append(False)
    else:
        print("  ❌ Impossible de valider le contenu")
        checks.append(False)

    # RÉSUMÉ FINAL
    print("\n" + "=" * 80)
    print("📊 RÉSUMÉ DE LA VALIDATION")
    print("=" * 80)

    total_checks = len(checks)
    passed_checks = sum(checks)
    success_rate = (passed_checks / total_checks) * 100

    print(f"\n✅ Checks réussis : {passed_checks}/{total_checks} ({success_rate:.1f}%)")

    if success_rate == 100:
        print("\n🎉 VALIDATION COMPLÈTE ! TOUT EST PRÊT !")
        print("\n📋 PROCHAINES ÉTAPES :")
        print("  1. Ouvrir PLANNING.xlsm")
        print("  2. Alt+F11 (ouvrir VBA)")
        print("  3. Importer Module_Planning_CORRECTED.bas")
        print("  4. Importer Module_Specialisations_CORRECTED.bas")
        print("  5. Exécuter GenererPlanningAutomatique")
        print("\n📚 Consulter GUIDE_CORRECTION_COMPLET.md pour les détails")
    elif success_rate >= 90:
        print("\n⚠️ VALIDATION PARTIELLE - Quelques éléments manquent")
        print("Vérifiez les éléments marqués ❌ ci-dessus")
    else:
        print("\n❌ VALIDATION ÉCHOUÉE - Problèmes majeurs détectés")
        print("Relancez les scripts de correction")

    print("\n" + "=" * 80)

    return success_rate == 100

if __name__ == "__main__":
    valider_livraison()
