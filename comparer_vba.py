#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compare le VBA dans PLANNING.xlsm avec les fichiers dans vba-modules/
"""

import os

# Fonctions/procédures importantes de Phase 2 à vérifier
ELEMENTS_PHASE2 = {
    "Module_Calculs.bas": [
        "CalculerTarifVisite",
        "CalculerDureeVisite",
        "Nom_Structure",
        "Type_Prestation",
        "Duree_Heures",
        "Nb_Participants"
    ],
    "Module_Planning.bas": [
        "AttribuerGuideAutomatiquement",
        "AttribuerGuideParSpecialisation",
        "VerifierDisponibiliteGuide",
        "Type_Prestation"
    ],
    "Module_Emails.bas": [
        "EnvoyerEmailOutlook",
        "EnvoyerNotificationGuide",
        "FormatEmailNotification"
    ],
    "Module_Emails_SMTP.bas": [
        "EnvoyerEmailSMTP",
        "ConfigurerSMTP"
    ],
    "Module_Disponibilites.bas": [
        "AjouterDisponibilite",
        "ModifierDisponibilite",
        "SupprimerDisponibilite"
    ],
    "Module_Specialisations.bas": [
        "ChargerSpecialisationsGuide",
        "SauvegarderSpecialisations",
        "VerifierSpecialisation"
    ]
}

def lire_vba_extracted():
    """Lit le VBA extrait"""
    with open("vbaProject_extracted.bin", "rb") as f:
        return f.read()

def lire_fichier_vba(filepath):
    """Lit un fichier VBA"""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.read()
    except:
        with open(filepath, "r", encoding="latin-1") as f:
            return f.read()

def verifier_elements():
    """Vérifie quels éléments sont présents dans le VBA extrait"""

    print("=" * 100)
    print("🔍 COMPARAISON VBA : PLANNING.xlsm vs vba-modules/")
    print("=" * 100)
    print()

    # Lire le VBA binaire extrait
    vba_content = lire_vba_extracted()

    manquants = []
    presents = []

    for module, elements in ELEMENTS_PHASE2.items():
        print(f"📋 {module}")
        print("-" * 100)

        # Vérifier si le fichier existe dans vba-modules/
        filepath = os.path.join("vba-modules", module)
        if not os.path.exists(filepath):
            print(f"  ⚠️  Fichier {module} n'existe pas dans vba-modules/")
            continue

        # Lire le contenu attendu
        file_content = lire_fichier_vba(filepath)

        for element in elements:
            # Chercher dans le VBA binaire
            element_bytes = element.encode('utf-8')

            if element_bytes in vba_content:
                print(f"  ✅ {element}")
                presents.append(f"{module}::{element}")
            else:
                print(f"  ❌ {element} - MANQUANT")
                manquants.append(f"{module}::{element}")

        print()

    # Résumé
    print("=" * 100)
    print("📊 RÉSUMÉ")
    print("=" * 100)
    print()

    total = len(presents) + len(manquants)
    pourcentage = (len(presents) / total * 100) if total > 0 else 0

    print(f"✅ Éléments présents : {len(presents)}/{total} ({pourcentage:.0f}%)")
    print(f"❌ Éléments manquants : {len(manquants)}/{total}")
    print()

    if manquants:
        print("🚨 ÉLÉMENTS MANQUANTS À COPIER :")
        modules_a_copier = set()
        for item in manquants:
            module = item.split("::")[0]
            modules_a_copier.add(module)
            print(f"  - {item}")

        print()
        print("📝 MODULES À RECOPIER COMPLÈTEMENT :")
        for module in sorted(modules_a_copier):
            print(f"  ➤ {module}")
    else:
        print("✅ Tous les éléments Phase 2 sont présents !")

    print()
    return len(manquants) == 0

if __name__ == "__main__":
    if not os.path.exists("vbaProject_extracted.bin"):
        print("❌ Erreur : vbaProject_extracted.bin n'existe pas")
        print("   Exécute d'abord : python3 extraire_vba_planning.py")
    else:
        verifier_elements()
