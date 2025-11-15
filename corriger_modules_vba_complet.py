#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CORRECTION COMPLÈTE DES MODULES VBA
Basé sur l'analyse réelle de la structure du fichier PLANNING.xlsm

PROBLÈMES IDENTIFIÉS:
1. Feuille DISPONIBILITES: structure incorrecte (Col1=Guide au lieu de Date)
2. Module VBA lit les mauvaises colonnes
3. Spécialisations: Nom_Guide au lieu de Prénom/Nom séparés
4. Heures mal formatées dans Planning
"""

import openpyxl
from openpyxl import load_workbook
import shutil
from datetime import datetime

# ============================================================================
# NOUVEAU CODE VBA - MODULE_PLANNING (CORRIGÉ)
# ============================================================================

MODULE_PLANNING_CORRECTED = """Attribute VB_Name = "Module_Planning"
Option Explicit

' ===== CONSTANTES =====
Private Const FEUILLE_DISPONIBILITES As String = "Disponibilites"
Private Const FEUILLE_VISITES As String = "Visites"
Private Const FEUILLE_PLANNING As String = "Planning"
Private Const FEUILLE_SPECIALISATIONS As String = "Spécialisations"

' ===== GÉNÉRATION AUTOMATIQUE DU PLANNING =====
Public Sub GenererPlanningAutomatique()
    On Error GoTo GestionErreur

    Dim wsVisites As Worksheet
    Dim wsPlanning As Worksheet
    Dim i As Long, ligneP As Long
    Dim dateVisite As Date
    Dim heureDebut As Date
    Dim heureFin As String
    Dim typeVisite As String
    Dim nomStructure As String
    Dim nbParticipants As String
    Dim niveau As String
    Dim theme As String
    Dim guidesDispos As Collection
    Dim guideAttribue As String
    Dim listeGuidesDispos As String

    Application.ScreenUpdating = False

    ' Récupération des feuilles
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)

    ' Vider le planning existant (garder les en-têtes)
    If wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row > 1 Then
        wsPlanning.Range("A2:K" & wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row).ClearContents
    End If

    ligneP = 2 ' Commencer à la ligne 2 (après en-têtes)

    ' Parcourir toutes les visites
    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row

        ' LECTURE DES COLONNES CORRECTES (selon analyse)
        ' Col 1: ID_Visite
        ' Col 2: Date
        ' Col 3: Heure_Debut
        ' Col 4: Heure_Fin
        ' Col 5: Nb_Participants
        ' Col 6: Type_Prestation
        ' Col 7: Nom_Structure
        ' Col 8: Niveau
        ' Col 9: Theme

        dateVisite = wsVisites.Cells(i, 2).Value
        heureDebut = wsVisites.Cells(i, 3).Value
        heureFin = wsVisites.Cells(i, 4).Value
        nbParticipants = wsVisites.Cells(i, 5).Value
        typeVisite = wsVisites.Cells(i, 6).Value
        nomStructure = wsVisites.Cells(i, 7).Value
        niveau = wsVisites.Cells(i, 8).Value
        theme = wsVisites.Cells(i, 9).Value

        ' Obtenir les guides disponibles pour cette date
        Set guidesDispos = ObtenirGuidesDisponibles(dateVisite)

        ' Filtrer par spécialisation
        Set guidesDispos = FiltrerParSpecialisation(guidesDispos, typeVisite)

        ' Attribuer un guide
        If guidesDispos.Count > 0 Then
            guideAttribue = guidesDispos(1)
        Else
            guideAttribue = "AUCUN GUIDE DISPONIBLE"
        End If

        ' Construire liste guides disponibles
        listeGuidesDispos = ConstruireListeGuides(guidesDispos)

        ' ÉCRIRE DANS PLANNING
        wsPlanning.Cells(ligneP, 1).Value = wsVisites.Cells(i, 1).Value ' ID_Visite
        wsPlanning.Cells(ligneP, 2).Value = dateVisite
        wsPlanning.Cells(ligneP, 3).Value = Format(heureDebut, "hh:mm") ' ✅ FORMAT HEURE CORRIGÉ
        wsPlanning.Cells(ligneP, 4).Value = nomStructure
        wsPlanning.Cells(ligneP, 5).Value = typeVisite
        wsPlanning.Cells(ligneP, 6).Value = heureFin
        wsPlanning.Cells(ligneP, 7).Value = guideAttribue
        wsPlanning.Cells(ligneP, 8).Value = theme
        wsPlanning.Cells(ligneP, 9).Value = niveau
        wsPlanning.Cells(ligneP, 10).Value = listeGuidesDispos ' ✅ LISTE GUIDES
        wsPlanning.Cells(ligneP, 11).Value = "À confirmer"

        ligneP = ligneP + 1
    Next i

    Application.ScreenUpdating = True

    MsgBox "Planning généré avec succès !" & vbCrLf & _
           (ligneP - 2) & " visites traitées.", vbInformation

    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la génération du planning : " & Err.Description, vbCritical
End Sub

' ===== OBTENIR GUIDES DISPONIBLES (CORRIGÉ) =====
Private Function ObtenirGuidesDisponibles(dateVisite As Date) As Collection
    On Error Resume Next

    Dim wsDispo As Worksheet
    Dim col As New Collection
    Dim i As Long
    Dim guideID As String
    Dim dateGuide As Date
    Dim disponible As String
    Dim nomGuide As String

    Set wsDispo = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)

    ' STRUCTURE RÉELLE (selon analyse):
    ' Col 1: Guide (DATE au format texte bizarre)
    ' Col 2: Date (contient "OUI" ou vide)
    ' Col 3: Disponible (vide)
    ' Col 4: Commentaire (contient PRÉNOM)
    ' Col 5: Prenom (contient NOM)
    ' Col 6: Nom (vide)

    ' ⚠️ STRUCTURE INCORRECTE DANS EXCEL !
    ' Il faudra corriger l'import des données
    ' Pour l'instant, on adapte le code VBA

    For i = 2 To wsDispo.Cells(wsDispo.Rows.Count, 1).End(xlUp).Row
        On Error Resume Next

        ' Col 1 contient la date
        dateGuide = CDate(wsDispo.Cells(i, 1).Value)

        ' Col 2 contient OUI/NON
        disponible = UCase(Trim(wsDispo.Cells(i, 2).Value))

        ' Col 4 = Prénom, Col 5 = Nom
        nomGuide = Trim(wsDispo.Cells(i, 4).Value) & " " & Trim(wsDispo.Cells(i, 5).Value)

        If dateGuide = dateVisite And disponible = "OUI" Then
            ' Éviter doublons
            Dim existe As Boolean
            existe = False
            Dim j As Integer
            For j = 1 To col.Count
                If col(j) = nomGuide Then
                    existe = True
                    Exit For
                End If
            Next j

            If Not existe And nomGuide <> " " Then
                col.Add nomGuide
            End If
        End If

        On Error GoTo 0
    Next i

    Set ObtenirGuidesDisponibles = col
End Function

' ===== FILTRER PAR SPÉCIALISATION (CORRIGÉ) =====
Private Function FiltrerParSpecialisation(guidesDispos As Collection, typeVisite As String) As Collection
    Dim col As New Collection
    Dim guide As Variant
    Dim i As Integer

    If guidesDispos.Count = 0 Then
        Set FiltrerParSpecialisation = col
        Exit Function
    End If

    For Each guide In guidesDispos
        If GuideAutoriseVisite(CStr(guide), typeVisite) Then
            col.Add guide
        End If
    Next guide

    Set FiltrerParSpecialisation = col
End Function

' ===== CONSTRUIRE LISTE GUIDES =====
Private Function ConstruireListeGuides(guidesCol As Collection) As String
    Dim resultat As String
    Dim guide As Variant

    resultat = ""
    For Each guide In guidesCol
        If resultat = "" Then
            resultat = guide
        Else
            resultat = resultat & ", " & guide
        End If
    Next guide

    If resultat = "" Then
        resultat = "Aucun"
    End If

    ConstruireListeGuides = resultat
End Function

' ===== VÉRIFICATION SPÉCIALISATION =====
Private Function GuideAutoriseVisite(nomGuide As String, typeVisite As String) As Boolean
    ' Appel vers Module_Specialisations
    GuideAutoriseVisite = Module_Specialisations.GuideAutoriseVisite(nomGuide, typeVisite)
End Function
"""

# ============================================================================
# NOUVEAU CODE VBA - MODULE_SPECIALISATIONS (CORRIGÉ)
# ============================================================================

MODULE_SPECIALISATIONS_CORRECTED = """Attribute VB_Name = "Module_Specialisations"
Option Explicit

' ===== VÉRIFIER AUTORISATION GUIDE POUR TYPE VISITE =====
Public Function GuideAutoriseVisite(nomGuide As String, typeVisite As String) As Boolean
    On Error Resume Next

    Dim ws As Worksheet
    Dim derLigne As Long
    Dim i As Long
    Dim nomGuideSpec As String
    Dim typeVisiteSpec As String
    Dim autorise As String
    Dim trouve As Boolean

    ' Par défaut : autorisé (si pas de règle spécifique)
    GuideAutoriseVisite = True
    trouve = False

    ' Récupérer feuille Spécialisations
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("Spécialisations")

    If ws Is Nothing Then
        Exit Function
    End If

    derLigne = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' STRUCTURE RÉELLE (selon analyse):
    ' Col 1: ID_Specialisation
    ' Col 2: Nom_Guide (NOM uniquement)
    ' Col 3: Email_Guide
    ' Col 4: Type_Prestation
    ' Col 5: Autorise (OUI/NON)

    For i = 2 To derLigne
        nomGuideSpec = UCase(Trim(ws.Cells(i, 2).Value))
        typeVisiteSpec = UCase(Trim(ws.Cells(i, 4).Value))
        autorise = UCase(Trim(ws.Cells(i, 5).Value))

        ' Vérifier correspondance (nom peut être partiel)
        If (InStr(1, UCase(nomGuide), nomGuideSpec, vbTextCompare) > 0 Or _
            InStr(1, nomGuideSpec, UCase(nomGuide), vbTextCompare) > 0) And _
           (InStr(1, UCase(typeVisite), typeVisiteSpec, vbTextCompare) > 0 Or _
            InStr(1, typeVisiteSpec, UCase(typeVisite), vbTextCompare) > 0) Then

            trouve = True

            If autorise = "OUI" Then
                GuideAutoriseVisite = True
            Else
                GuideAutoriseVisite = False
            End If

            Exit Function
        End If
    Next i

    ' Si aucune règle trouvée, autoriser par défaut
    If Not trouve Then
        GuideAutoriseVisite = True
    End If

    On Error GoTo 0
End Function

' ===== OBTENIR SPÉCIALISATIONS D'UN GUIDE =====
Public Function ObtenirSpecialisationsGuide(nomGuide As String) As Collection
    On Error Resume Next

    Dim ws As Worksheet
    Dim col As New Collection
    Dim i As Long
    Dim nomGuideSpec As String
    Dim typeVisite As String
    Dim autorise As String

    Set ws = ThisWorkbook.Worksheets("Spécialisations")

    If ws Is Nothing Then
        Set ObtenirSpecialisationsGuide = col
        Exit Function
    End If

    For i = 2 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        nomGuideSpec = UCase(Trim(ws.Cells(i, 2).Value))
        typeVisite = Trim(ws.Cells(i, 4).Value)
        autorise = UCase(Trim(ws.Cells(i, 5).Value))

        If InStr(1, UCase(nomGuide), nomGuideSpec, vbTextCompare) > 0 And autorise = "OUI" Then
            col.Add typeVisite
        End If
    Next i

    Set ObtenirSpecialisationsGuide = col

    On Error GoTo 0
End Function
"""

# ============================================================================
# FONCTION PRINCIPALE
# ============================================================================

def corriger_modules_vba():
    """
    Importe les modules VBA corrigés dans PLANNING.xlsm
    """

    fichier_planning = "/Users/otmaneboulahia/Documents/Excel-Auto/PLANNING.xlsm"

    print("=" * 80)
    print("🔧 CORRECTION DES MODULES VBA")
    print("=" * 80)

    # Backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f"/Users/otmaneboulahia/Documents/Excel-Auto/PLANNING_backup_{timestamp}.xlsm"
    shutil.copy2(fichier_planning, backup_file)
    print(f"\n✅ Backup créé : {backup_file}")

    # Sauvegarder les modules corrigés
    module_planning_path = "/Users/otmaneboulahia/Documents/Excel-Auto/vba-modules/Module_Planning_CORRECTED.bas"
    module_spec_path = "/Users/otmaneboulahia/Documents/Excel-Auto/vba-modules/Module_Specialisations_CORRECTED.bas"

    with open(module_planning_path, 'w', encoding='utf-8') as f:
        f.write(MODULE_PLANNING_CORRECTED)
    print(f"\n✅ Module Planning corrigé sauvegardé : {module_planning_path}")

    with open(module_spec_path, 'w', encoding='utf-8') as f:
        f.write(MODULE_SPECIALISATIONS_CORRECTED)
    print(f"✅ Module Spécialisations corrigé sauvegardé : {module_spec_path}")

    print("\n" + "=" * 80)
    print("📋 RÉSUMÉ DES CORRECTIONS APPLIQUÉES")
    print("=" * 80)

    print("""
✅ MODULE_PLANNING :
   - Format heure corrigé : Format(heureDebut, "hh:mm")
   - Lecture colonnes Visites corrigée (Col 3=Heure, Col 6=Type, Col 7=Structure)
   - Lecture Disponibilites adaptée à la structure actuelle
   - Liste guides disponibles ajoutée

✅ MODULE_SPECIALISATIONS :
   - Lecture Col 2=Nom_Guide, Col 4=Type_Prestation, Col 5=Autorise
   - Logique OUI/NON correcte
   - Comparaison insensible à la casse

⚠️ ATTENTION : STRUCTURE DISPONIBILITES INCORRECTE !
   Actuellement :
     Col 1: Date (format texte)
     Col 2: "OUI" ou vide
     Col 4: Prénom
     Col 5: Nom

   Le code VBA a été adapté, mais il faudrait corriger l'import des données.
""")

    print("\n" + "=" * 80)
    print("📝 PROCHAINES ÉTAPES")
    print("=" * 80)
    print("""
1. Importer les modules VBA dans Excel :
   - Ouvrir PLANNING.xlsm
   - Alt+F11 (ouvrir VBA)
   - Clic droit sur VBAProject > Importer un fichier
   - Sélectionner Module_Planning_CORRECTED.bas
   - Répéter pour Module_Specialisations_CORRECTED.bas

2. Tester le planning :
   - Alt+F8 > GenererPlanningAutomatique > Exécuter
   - Vérifier colonnes HEURE et GUIDES_DISPONIBLES

3. Corriger l'import des disponibilités (optionnel mais recommandé)
""")

if __name__ == "__main__":
    corriger_modules_vba()
