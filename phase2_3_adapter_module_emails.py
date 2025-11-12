#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PHASE 2.3 - ADAPTER MODULE_EMAILS.BAS

Adapte les références de colonnes et enrichit les templates emails.

MAPPING COLONNES (AVANT → APRÈS):
  Col 3: Heure → Col 3 Heure_Debut ✅ (concaténé avec col 4)
  Col 4: Lieu/Musée → Col 7 Nom_Structure
  Col 5: guideID → Col 12 Guide_Attribue
  Col 6: guideNom → Col 12 Guide_Attribue (même col)

ENRICHISSEMENT EMAILS:
  Ajouter dans templates:
  - Type_Prestation (col 6)
  - Nom_Structure (col 7)
  - Niveau (col 8)
  - Theme (col 9)
"""

import re
import sys


def lire_fichier(chemin):
    """Lit le fichier VBA."""
    with open(chemin, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()


def ecrire_fichier(chemin, contenu):
    """Écrit le fichier VBA."""
    with open(chemin, 'w', encoding='utf-8', newline='\r\n') as f:
        f.write(contenu)


def adapter_module_emails():
    """Adapte Module_Emails.bas avec les nouvelles colonnes."""

    chemin = 'vba-modules/Module_Emails.bas'
    print("=" * 100)
    print("🔧 PHASE 2.3 - ADAPTATION MODULE_EMAILS.BAS")
    print("=" * 100)

    # Lire le fichier
    print("\n📖 Lecture du fichier...")
    contenu = lire_fichier(chemin)
    lignes = contenu.split('\n')

    print(f"   ✅ {len(lignes)} lignes lues")

    modifications = 0

    # =========================================================================
    # MODIFICATION 1: Toutes les lectures de guideID depuis wsPlanning
    # AVANT: guideID = wsPlanning.Cells(i, 5).Value
    # APRÈS: guideID = wsPlanning.Cells(i, 12).Value ' Guide_Attribue
    # =========================================================================
    print("\n🔧 Modification 1: guideID lectures (multiples lignes)")

    compteur_guideid = 0
    for idx, ligne in enumerate(lignes):
        if 'guideID = wsPlanning.Cells(i, 5).Value' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 5)',
                'wsPlanning.Cells(i, 12) \' Guide_Attribue'
            )
            compteur_guideid += 1

    print(f"   ✅ {compteur_guideid} lignes: guideID → colonne 12 (Guide_Attribue)")
    modifications += compteur_guideid

    # =========================================================================
    # MODIFICATION 2: Toutes les lectures de Lieu/Musée (col 4 → col 7)
    # AVANT: wsPlanning.Cells(i, 4).Value
    # APRÈS: wsPlanning.Cells(i, 7).Value ' Nom_Structure
    # =========================================================================
    print("\n🔧 Modification 2: Lieu/Musée (multiples lignes)")

    compteur_lieu = 0
    for idx, ligne in enumerate(lignes):
        # Rechercher les références à col 4 qui sont pour le lieu
        if 'wsPlanning.Cells(i, 4).Value' in ligne and 'Lieu' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 4)',
                'wsPlanning.Cells(i, 7) \' Nom_Structure'
            )
            compteur_lieu += 1

    print(f"   ✅ {compteur_lieu} lignes: Lieu → colonne 7 (Nom_Structure)")
    modifications += compteur_lieu

    # =========================================================================
    # MODIFICATION 3: Lecture guideNom (col 6 → col 12)
    # AVANT: guideNom = wsPlanning.Cells(i, 6).Value
    # APRÈS: guideNom = wsPlanning.Cells(i, 12).Value ' Guide_Attribue
    # =========================================================================
    print("\n🔧 Modification 3: guideNom lectures (multiples lignes)")

    compteur_guidenom = 0
    for idx, ligne in enumerate(lignes):
        if 'guideNom = wsPlanning.Cells(i, 6).Value' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 6)',
                'wsPlanning.Cells(i, 12) \' Guide_Attribue'
            )
            compteur_guidenom += 1

    print(f"   ✅ {compteur_guidenom} lignes: guideNom → colonne 12 (Guide_Attribue)")
    modifications += compteur_guidenom

    # =========================================================================
    # ENRICHISSEMENT: Ajouter nouvelles colonnes dans templates emails
    # =========================================================================
    print("\n🆕 Enrichissement: Templates emails avec nouvelles colonnes")

    compteur_enrichi = 0

    # Trouver les lignes infoVisite et enrichir
    for idx, ligne in enumerate(lignes):
        # Chercher les constructions infoVisite
        if 'infoVisite = "Date :' in ligne and 'Format(dateVisite' in ligne:
            # Vérifier les 3 lignes suivantes
            if idx + 2 < len(lignes):
                ligne_heure = lignes[idx + 1]
                ligne_lieu = lignes[idx + 2]

                # Vérifier qu'on a bien la structure attendue
                if '"Heure :' in ligne_heure and '"Lieu :' in ligne_lieu:
                    # Insérer après ligne_lieu les nouvelles infos
                    nouvelles_lignes = [
                        '                                "Type : " & wsPlanning.Cells(i, 6).Value & vbCrLf & _ \' Type_Prestation',
                        '                                "Niveau : " & wsPlanning.Cells(i, 8).Value & vbCrLf & _ \' Niveau',
                        '                                "Thème : " & wsPlanning.Cells(i, 9).Value \' Theme'
                    ]

                    # Modifier la dernière ligne existante pour ajouter & vbCrLf & _
                    if 'Nom_Structure' in lignes[idx + 2]:
                        lignes[idx + 2] = lignes[idx + 2].rstrip() + ' & vbCrLf & _'

                    # Insérer les nouvelles lignes
                    for offset, nouvelle_ligne in enumerate(nouvelles_lignes, start=1):
                        lignes.insert(idx + 2 + offset, nouvelle_ligne)

                    compteur_enrichi += 1

    print(f"   ✅ {compteur_enrichi} templates enrichis avec Type, Niveau, Thème")
    modifications += compteur_enrichi

    # =========================================================================
    # ENRICHISSEMENT 2: Template récapitulatif mensuel
    # =========================================================================
    print("\n🆕 Enrichissement 2: Template récapitulatif mensuel")

    compteur_recap = 0
    for idx, ligne in enumerate(lignes):
        # Chercher la construction du récapitulatif mensuel (ligne ~70)
        if 'infoVisite = Format(dateVisite, "dd/mm/yyyy") & " | "' in ligne:
            # Vérifier les 2 lignes suivantes
            if idx + 2 < len(lignes):
                if 'wsPlanning.Cells(i, 3).Value' in lignes[idx + 1]:
                    # Remplacer la construction complète
                    lignes[idx] = '                    infoVisite = Format(dateVisite, "dd/mm/yyyy") & " | " & _'
                    lignes[idx + 1] = '                                wsPlanning.Cells(i, 3).Value & " | " & _'
                    lignes[idx + 2] = '                                wsPlanning.Cells(i, 7).Value & " | " & _ \' Nom_Structure'

                    # Insérer nouvelles infos
                    lignes.insert(idx + 3, '                                "Type: " & wsPlanning.Cells(i, 6).Value \' Type_Prestation')

                    compteur_recap += 1

    print(f"   ✅ {compteur_recap} récapitulatifs enrichis")
    modifications += compteur_recap

    # Reconstituer le contenu
    contenu_modifie = '\n'.join(lignes)

    # Écrire le fichier
    print(f"\n💾 Écriture des modifications...")
    ecrire_fichier(chemin, contenu_modifie)

    print(f"   ✅ Fichier sauvegardé: {chemin}")
    print(f"\n📊 RÉSUMÉ:")
    print(f"   • {modifications} modifications effectuées")
    print(f"   • Colonnes adaptées: 7 (Nom_Structure), 12 (Guide_Attribue)")
    print(f"   • Templates enrichis avec Type, Niveau, Thème")

    print("\n" + "=" * 100)
    print("✅ MODULE_EMAILS.BAS ADAPTÉ AVEC SUCCÈS !")
    print("=" * 100)

    return True


if __name__ == '__main__':
    try:
        succes = adapter_module_emails()
        sys.exit(0 if succes else 1)
    except Exception as e:
        print(f"\n❌ ERREUR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
