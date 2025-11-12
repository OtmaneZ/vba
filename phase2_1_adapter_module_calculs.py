#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PHASE 2.1 - ADAPTER MODULE_CALCULS.BAS

Adapte toutes les références de colonnes après restructuration Phase 1.

MAPPING COLONNES (AVANT → APRÈS):
  Col 2:  Date              → Col 2  Date ✅ INCHANGÉ
  Col 3:  Heure (doublon)   → SUPPRIMÉ
  Col 4:  Musée             → Col 7  Nom_Structure
  Col 5:  Type_Visite       → Col 6  Type_Prestation
  Col 6:  Durée_Heures      → Col 14 Duree_Heures (calculée)
  Col 7:  Nombre_Visiteurs  → Col 5  Nb_Participants
  Col 8:  Statut            → Col 11 Statut
  Col 11: Heure_Debut       → Col 3  Heure_Debut
  Col 12: Heure_Fin         → Col 4  Heure_Fin
  Col 13: Langue            → Col 15 Langue
  Col 15: Tarif             → Col 13 Tarif
  Col 16: Guide             → Col 12 Guide_Attribue
  Col 17: Notes             → Col 10 Commentaires
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


def adapter_module_calculs():
    """Adapte Module_Calculs.bas avec les nouvelles colonnes."""

    chemin = 'vba-modules/Module_Calculs.bas'
    print("=" * 100)
    print("🔧 PHASE 2.1 - ADAPTATION MODULE_CALCULS.BAS")
    print("=" * 100)

    # Lire le fichier
    print("\n📖 Lecture du fichier...")
    contenu = lire_fichier(chemin)
    lignes = contenu.split('\n')

    print(f"   ✅ {len(lignes)} lignes lues")

    # Compteur de modifications
    modifications = 0

    # =========================================================================
    # MODIFICATION 1: Ligne 63 - guideID
    # AVANT: guideID = Trim(wsPlanning.Cells(i, 5).Value)
    # APRÈS: guideID = Trim(wsPlanning.Cells(i, 12).Value) ' Guide_Attribue
    # =========================================================================
    print("\n🔧 Modification 1/4: guideID (ligne ~63)")

    for idx, ligne in enumerate(lignes):
        if 'guideID = Trim(wsPlanning.Cells(i, 5).Value)' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 5)',
                'wsPlanning.Cells(i, 12) \' Guide_Attribue'
            )
            print(f"   ✅ Ligne {idx+1}: guideID → colonne 12 (Guide_Attribue)")
            modifications += 1
            break

    # =========================================================================
    # MODIFICATION 2: Ligne 277 - IdentifierTypeVisite
    # AVANT: typeVisite = UCase(Trim(wsVisites.Cells(i, 5).Value))
    # APRÈS: typeVisite = UCase(Trim(wsVisites.Cells(i, 6).Value)) ' Type_Prestation
    # =========================================================================
    print("\n🔧 Modification 2/4: IdentifierTypeVisite (ligne ~277)")

    for idx, ligne in enumerate(lignes):
        if 'typeVisite = UCase(Trim(wsVisites.Cells(i, 5).Value))' in ligne:
            lignes[idx] = ligne.replace(
                'wsVisites.Cells(i, 5)',
                'wsVisites.Cells(i, 6) \' Type_Prestation'
            )
            print(f"   ✅ Ligne {idx+1}: typeVisite → colonne 6 (Type_Prestation)")
            modifications += 1
            break

    # =========================================================================
    # MODIFICATION 3: Ligne 375-376 - ObtenirDureeVisite
    # AVANT: heureDebut = CDate(wsVisites.Cells(i, 3).Value)
    #        heureFin = CDate(wsVisites.Cells(i, 4).Value)
    # DÉJÀ CORRECT ! Colonnes 3 et 4 = Heure_Debut et Heure_Fin
    # =========================================================================
    print("\n✅ Vérification 3/4: ObtenirDureeVisite (lignes ~375-376)")

    for idx, ligne in enumerate(lignes):
        if 'heureDebut = CDate(wsVisites.Cells(i, 3).Value)' in ligne:
            print(f"   ✅ Ligne {idx+1}: heureDebut → colonne 3 ✅ DÉJÀ CORRECT")
            break

    for idx, ligne in enumerate(lignes):
        if 'heureFin = CDate(wsVisites.Cells(i, 4).Value)' in ligne:
            print(f"   ✅ Ligne {idx+1}: heureFin → colonne 4 ✅ DÉJÀ CORRECT")
            break

    # =========================================================================
    # MODIFICATION 4: Ligne 467 - GenererFichesPaie (guideID dans wsPlanning)
    # AVANT: If Trim(wsPlanning.Cells(i, 5).Value) = Trim(guideID) Then
    # APRÈS: If Trim(wsPlanning.Cells(i, 12).Value) = Trim(guideID) Then ' Guide_Attribue
    # =========================================================================
    print("\n🔧 Modification 4/4: GenererFichesPaie (ligne ~467)")

    for idx, ligne in enumerate(lignes):
        if 'If Trim(wsPlanning.Cells(i, 5).Value) = Trim(guideID) Then' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 5)',
                'wsPlanning.Cells(i, 12) \' Guide_Attribue'
            )
            print(f"   ✅ Ligne {idx+1}: guideID filter → colonne 12 (Guide_Attribue)")
            modifications += 1
            break

    # =========================================================================
    # RÉÉCRITURE COMPLÈTE de IdentifierTypeVisite
    # =========================================================================
    print("\n🔄 Réécriture complète: IdentifierTypeVisite")

    nouvelle_fonction = '''Private Function IdentifierTypeVisite(idVisite As String) As String
    '===============================================================================
    ' FONCTION: IdentifierTypeVisite
    ' DESCRIPTION: Identifie le type de visite depuis colonne Type_Prestation
    ' PARAMETRE: idVisite - ID de la visite (ex: V001)
    ' RETOUR: Type de prestation (VISITE CONTEE BRANLY, MARINE, HORS LES MURS, etc.)
    '===============================================================================
    Dim wsVisites As Worksheet
    Dim i As Long
    Dim typePrestation As String

    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    IdentifierTypeVisite = "AUTRE" ' Par defaut

    ' Chercher la visite dans la feuille Visites
    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
        If wsVisites.Cells(i, 1).Value = idVisite Then
            ' Lire la colonne Type_Prestation (colonne 6)
            typePrestation = UCase(Trim(wsVisites.Cells(i, 6).Value)) ' Type_Prestation

            ' Normaliser les types pour correspondre aux tarifs
            Select Case typePrestation
                Case "VISITE CONTEE BRANLY", "VISITE BRANLY"
                    IdentifierTypeVisite = "BRANLY"
                Case "VISITE CONTEE MARINE", "VISITE MARINE"
                    IdentifierTypeVisite = "MARINE"
                Case "HORS LES MURS", "HORS-LES-MURS", "HORSLEMURS"
                    IdentifierTypeVisite = "HORSLEMURS"
                Case "VISIO"
                    IdentifierTypeVisite = "VISIO"
                Case "EVENEMENT BRANLY", "EVENEMENT"
                    IdentifierTypeVisite = "EVENEMENT"
                Case Else
                    ' Retourner le type tel quel si non reconnu
                    If typePrestation <> "" Then
                        IdentifierTypeVisite = typePrestation
                    End If
            End Select

            Exit Function
        End If
    Next i
End Function'''

    # Trouver et remplacer la fonction IdentifierTypeVisite
    debut_fonction = -1
    fin_fonction = -1

    for idx, ligne in enumerate(lignes):
        if 'Private Function IdentifierTypeVisite(' in ligne:
            debut_fonction = idx
        elif debut_fonction != -1 and ligne.strip().startswith('End Function') and 'IdentifierTypeVisite' not in lignes[idx-5:idx][-1]:
            fin_fonction = idx
            break

    if debut_fonction != -1 and fin_fonction != -1:
        # Remplacer la fonction
        lignes_avant = lignes[:debut_fonction]
        lignes_apres = lignes[fin_fonction+1:]

        lignes = lignes_avant + nouvelle_fonction.split('\n') + lignes_apres

        print(f"   ✅ Fonction IdentifierTypeVisite réécrite (lignes {debut_fonction+1}-{fin_fonction+1})")
        print(f"      • Lecture directe de colonne 6 (Type_Prestation)")
        print(f"      • Normalisation des 5 types de prestations")
        modifications += 1
    else:
        print(f"   ⚠️  Fonction IdentifierTypeVisite non trouvée pour remplacement")

    # Reconstituer le contenu
    contenu_modifie = '\n'.join(lignes)

    # Écrire le fichier
    print(f"\n💾 Écriture des modifications...")
    ecrire_fichier(chemin, contenu_modifie)

    print(f"   ✅ Fichier sauvegardé: {chemin}")
    print(f"\n📊 RÉSUMÉ:")
    print(f"   • {modifications} modifications effectuées")
    print(f"   • Colonnes adaptées à la nouvelle structure Phase 1")
    print(f"   • Fonction IdentifierTypeVisite réécrite")

    print("\n" + "=" * 100)
    print("✅ MODULE_CALCULS.BAS ADAPTÉ AVEC SUCCÈS !")
    print("=" * 100)

    return True


if __name__ == '__main__':
    try:
        succes = adapter_module_calculs()
        sys.exit(0 if succes else 1)
    except Exception as e:
        print(f"\n❌ ERREUR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
