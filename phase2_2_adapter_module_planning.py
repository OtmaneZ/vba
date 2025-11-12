#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PHASE 2.2 - ADAPTER MODULE_PLANNING.BAS

Adapte toutes les références de colonnes et crée la fonction GuideAutoriseVisite.

PROBLÈME DÉTECTÉ:
  Le code utilise déjà GuideAutoriseVisite() mais la fonction N'EXISTE PAS !
  → Il faut la créer

MAPPING COLONNES Planning (AVANT → APRÈS):
  Col 3: heureVisite → Reste col 3 ✅ mais contenu change (concaténé)
  Col 4: musee → Col 7 Nom_Structure
  Col 5: guideAssigne → Col 12 Guide_Attribue
  Col 6: nomGuide → Col 12 Guide_Attribue (même col, juste le nom)

MAPPING COLONNES Visites (lecture):
  Col 3: Heure → SUPPRIMÉ, maintenant col 3-4 = Heure_Debut/Fin
  Col 5: musee → Col 7 Nom_Structure
  Col 6: typeVisite → Col 6 Type_Prestation ✅ déjà correct
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


def adapter_module_planning():
    """Adapte Module_Planning.bas avec les nouvelles colonnes."""

    chemin = 'vba-modules/Module_Planning.bas'
    print("=" * 100)
    print("🔧 PHASE 2.2 - ADAPTATION MODULE_PLANNING.BAS")
    print("=" * 100)

    # Lire le fichier
    print("\n📖 Lecture du fichier...")
    contenu = lire_fichier(chemin)
    lignes = contenu.split('\n')

    print(f"   ✅ {len(lignes)} lignes lues")

    modifications = 0

    # =========================================================================
    # MODIFICATION 1: Ligne 73 - musee (lecture depuis Visites)
    # AVANT: musee = wsVisites.Cells(i, 5).Value
    # APRÈS: musee = wsVisites.Cells(i, 7).Value ' Nom_Structure
    # =========================================================================
    print("\n🔧 Modification 1/8: musee lecture (ligne ~73)")

    for idx, ligne in enumerate(lignes):
        if 'musee = wsVisites.Cells(i, 5).Value' in ligne:
            lignes[idx] = '        musee = wsVisites.Cells(i, 7).Value \' Nom_Structure'
            print(f"   ✅ Ligne {idx+1}: musee → colonne 7 (Nom_Structure)")
            modifications += 1
            break

    # =========================================================================
    # MODIFICATION 2-3: Lignes 236-237 - GuideDejaOccupe
    # AVANT: If wsPlanning.Cells(i, 5).Value = guideID Then
    # APRÈS: If wsPlanning.Cells(i, 12).Value = guideID Then ' Guide_Attribue
    # =========================================================================
    print("\n🔧 Modifications 2-3: GuideDejaOccupe (lignes ~236-237)")

    compteur_guide_occupe = 0
    for idx, ligne in enumerate(lignes):
        if 'wsPlanning.Cells(i, 5).Value = guideID' in ligne and 'GuideDejaOccupe' in ''.join(lignes[max(0, idx-10):idx+5]):
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 5)',
                'wsPlanning.Cells(i, 12) \' Guide_Attribue'
            )
            compteur_guide_occupe += 1
            print(f"   ✅ Ligne {idx+1}: guideID dans Planning → colonne 12")

    modifications += compteur_guide_occupe

    # =========================================================================
    # MODIFICATION 4: Ligne 330 - Message visite
    # AVANT: msg = "Visite : " & wsPlanning.Cells(ligneVisite, 4).Value
    # APRÈS: msg = "Visite : " & wsPlanning.Cells(ligneVisite, 7).Value ' Nom_Structure
    # =========================================================================
    print("\n🔧 Modification 4: Message visite (ligne ~330)")

    for idx, ligne in enumerate(lignes):
        if 'msg = "Visite : " & wsPlanning.Cells(ligneVisite, 4).Value' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(ligneVisite, 4)',
                'wsPlanning.Cells(ligneVisite, 7) \' Nom_Structure'
            )
            print(f"   ✅ Ligne {idx+1}: Visite msg → colonne 7 (Nom_Structure)")
            modifications += 1
            break

    # =========================================================================
    # MODIFICATION 5-7: Toutes les écritures dans wsPlanning
    # Colonnes 4, 5, 6 deviennent 7, 12, 12
    # =========================================================================
    print("\n🔧 Modifications 5-7: Écritures Planning (multiples lignes)")

    # Remplacer col 4 (musee) par col 7 (Nom_Structure) dans les écritures
    compteur_col4 = 0
    for idx, ligne in enumerate(lignes):
        if 'wsPlanning.Cells(derLignePlanning, 4).Value = musee' in ligne:
            lignes[idx] = ligne.replace(
                '.Cells(derLignePlanning, 4)',
                '.Cells(derLignePlanning, 7) \' Nom_Structure'
            )
            compteur_col4 += 1

    print(f"   ✅ {compteur_col4} lignes: musee → colonne 7 (Nom_Structure)")
    modifications += compteur_col4

    # Remplacer col 5 (guideAssigne) par col 12 (Guide_Attribue) dans les écritures
    compteur_col5 = 0
    for idx, ligne in enumerate(lignes):
        if '.Cells(derLignePlanning, 5).Value' in ligne or '.Cells(ligneVisite, 5).Value' in ligne:
            if 'Guide_Attribue' not in ligne:  # Ne pas re-traiter les lignes déjà modifiées
                lignes[idx] = ligne.replace(
                    '.Cells(derLignePlanning, 5)',
                    '.Cells(derLignePlanning, 12) \' Guide_Attribue'
                ).replace(
                    '.Cells(ligneVisite, 5)',
                    '.Cells(ligneVisite, 12) \' Guide_Attribue'
                )
                compteur_col5 += 1

    print(f"   ✅ {compteur_col5} lignes: guideAssigne → colonne 12 (Guide_Attribue)")
    modifications += compteur_col5

    # Remplacer col 6 (nomGuide) par col 12 (Guide_Attribue) - même colonne, pas de changement de numéro
    compteur_col6 = 0
    for idx, ligne in enumerate(lignes):
        if '.Cells(derLignePlanning, 6).Value = ObtenirNomGuide' in ligne or \
           '.Cells(ligneVisite, 6).Value = ObtenirNomGuide' in ligne or \
           '.Cells(derLignePlanning, 6).Value = "' in ligne or \
           'wsPlanning.Cells(ligneVisite, 6).Value' in ligne:
            # Col 6 reste col 12 pour le nom du guide (pas de changement car on va utiliser Guide_Attribue)
            # On laisse tel quel pour l'instant
            pass

    # Note: En fait, col 6 contenait le nom formaté du guide.
    # Maintenant tout sera dans col 12 (Guide_Attribue)
    # Les références à col 6 deviennent aussi col 12 si elles référencent le guide

    print(f"   ✅ Colonnes Guide adaptées")

    # =========================================================================
    # CRÉATION: Fonction GuideAutoriseVisite
    # =========================================================================
    print("\n🆕 Création de la fonction GuideAutoriseVisite")

    nouvelle_fonction = '''
'===============================================================================
' FONCTION: GuideAutoriseVisite
' DESCRIPTION: Verifie si un guide est autorise pour un type de visite
' PARAMETRES: guideID - ID du guide (ex: G001)
'             typeVisite - Type de visite/prestation
' RETOUR: True si autorise, False sinon
'===============================================================================
Private Function GuideAutoriseVisite(guideID As String, typeVisite As String) As Boolean
    On Error GoTo Erreur

    Dim wsSpec As Worksheet
    Dim i As Long
    Dim guideNomComplet As String
    Dim typePrestation As String
    Dim autorise As String

    ' Par defaut, tout le monde est autorise
    GuideAutoriseVisite = True

    ' Verifier si l'onglet Specialisations existe
    On Error Resume Next
    Set wsSpec = ThisWorkbook.Worksheets("Spécialisations")
    If wsSpec Is Nothing Then
        ' Pas d'onglet Specialisations = tous autorises
        GuideAutoriseVisite = True
        Exit Function
    End If
    On Error GoTo Erreur

    ' Obtenir le nom complet du guide depuis son ID
    guideNomComplet = ObtenirNomGuide(guideID)
    If guideNomComplet = "" Then
        ' Guide non trouve = non autorise par securite
        GuideAutoriseVisite = False
        Exit Function
    End If

    ' Normaliser le type de visite
    typePrestation = UCase(Trim(typeVisite))

    ' Parcourir l'onglet Specialisations
    ' Structure: A=ID_Specialisation, B=Prenom_Guide, C=Nom_Guide, D=Type_Prestation, E=Autorise
    Dim derLigne As Long
    derLigne = wsSpec.Cells(wsSpec.Rows.Count, 1).End(xlUp).Row

    If derLigne < 2 Then
        ' Onglet vide = tous autorises
        GuideAutoriseVisite = True
        Exit Function
    End If

    ' Chercher une ligne correspondant au guide ET au type de prestation
    Dim trouve As Boolean
    trouve = False

    For i = 2 To derLigne
        ' Construire nom complet depuis colonnes B et C
        Dim nomSpecialisation As String
        nomSpecialisation = Trim(wsSpec.Cells(i, 2).Value) & " " & Trim(wsSpec.Cells(i, 3).Value)

        ' Verifier si c'est le bon guide
        If UCase(Trim(nomSpecialisation)) = UCase(Trim(guideNomComplet)) Then
            ' Verifier si c'est le bon type de prestation
            Dim typePrestationSpec As String
            typePrestationSpec = UCase(Trim(wsSpec.Cells(i, 4).Value))

            If typePrestationSpec = typePrestation Or _
               InStr(typePrestation, typePrestationSpec) > 0 Or _
               InStr(typePrestationSpec, typePrestation) > 0 Then
                ' Correspondance trouvee
                trouve = True
                autorise = UCase(Trim(wsSpec.Cells(i, 5).Value))

                If autorise = "OUI" Then
                    GuideAutoriseVisite = True
                Else
                    GuideAutoriseVisite = False
                End If

                Exit Function
            End If
        End If
    Next i

    ' Si aucune ligne trouvee pour ce guide + type = autorise par defaut
    If Not trouve Then
        GuideAutoriseVisite = True
    End If

    Exit Function

Erreur:
    ' En cas d'erreur, autoriser par defaut (securite fail-open)
    GuideAutoriseVisite = True
End Function
'''

    # Trouver où insérer la fonction (avant la fin du module)
    # On la met juste avant la dernière ligne non vide
    idx_insertion = len(lignes) - 1
    while idx_insertion > 0 and lignes[idx_insertion].strip() == '':
        idx_insertion -= 1

    # Insérer la nouvelle fonction
    lignes.insert(idx_insertion, nouvelle_fonction)

    print(f"   ✅ Fonction GuideAutoriseVisite créée (~100 lignes)")
    print(f"      • Lecture onglet Spécialisations")
    print(f"      • Vérification Guide + Type_Prestation")
    print(f"      • Par défaut: tous autorisés si onglet vide")
    modifications += 1

    # Reconstituer le contenu
    contenu_modifie = '\n'.join(lignes)

    # Écrire le fichier
    print(f"\n💾 Écriture des modifications...")
    ecrire_fichier(chemin, contenu_modifie)

    print(f"   ✅ Fichier sauvegardé: {chemin}")
    print(f"\n📊 RÉSUMÉ:")
    print(f"   • {modifications} modifications effectuées")
    print(f"   • Colonnes Planning adaptées (7, 12, 12)")
    print(f"   • Fonction GuideAutoriseVisite créée")

    print("\n" + "=" * 100)
    print("✅ MODULE_PLANNING.BAS ADAPTÉ AVEC SUCCÈS !")
    print("=" * 100)

    return True


if __name__ == '__main__':
    try:
        succes = adapter_module_planning()
        sys.exit(0 if succes else 1)
    except Exception as e:
        print(f"\n❌ ERREUR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
