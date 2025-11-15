Attribute VB_Name = "Module_Planning"
Option Explicit

' ===== CONSTANTES =====
Private Const FEUILLE_DISPONIBILITES As String = "Disponibilites"
Private Const FEUILLE_VISITES As String = "Visites"
Private Const FEUILLE_PLANNING As String = "Planning"
Private Const FEUILLE_SPECIALISATIONS As String = "Specialisations"

' ===== GENERATION AUTOMATIQUE DU PLANNING =====
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

    ' Recuperation des feuilles
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)

    ' Vider le planning existant (garder les en-tetes)
    If wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row > 1 Then
        wsPlanning.Range("A2:K" & wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row).ClearContents
    End If

    ligneP = 2 ' Commencer a la ligne 2 (apres en-tetes)

    ' Parcourir toutes les visites
    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row

        ' LECTURE DES COLONNES REELLES (structure actuelle du fichier)
        ' ATTENTION: Les donnees sont decalees dans votre fichier Excel !
        ' Col 1: ID_Visite
        ' Col 2: Date
        ' Col 3: Heure_Debut
        ' Col 4: Heure_Fin (contient duree "1h")
        ' Col 5: Nb_Participants (contient TYPE_VISITE !)
        ' Col 6: Type_Prestation (contient NOM_STRUCTURE !)
        ' Col 7: Nom_Structure (contient NB_PARTICIPANTS !)
        ' Col 8: Niveau
        ' Col 9: Theme

        dateVisite = wsVisites.Cells(i, 2).Value
        heureDebut = wsVisites.Cells(i, 3).Value
        heureFin = wsVisites.Cells(i, 4).Value

        ' CORRECTION: Lire les colonnes selon la structure REELLE
        nbParticipants = wsVisites.Cells(i, 5).Value  ' Nb_Participants en col 5
        typeVisite = wsVisites.Cells(i, 6).Value      ' Type_Prestation en col 6
        nomStructure = wsVisites.Cells(i, 7).Value    ' Nom_Structure en col 7
        niveau = wsVisites.Cells(i, 8).Value
        theme = wsVisites.Cells(i, 9).Value

        ' Obtenir les guides disponibles pour cette date
        Set guidesDispos = ObtenirGuidesDisponibles(dateVisite)

        ' Filtrer par specialisation
        Set guidesDispos = FiltrerParSpecialisation(guidesDispos, typeVisite)

        ' Attribuer un guide
        If guidesDispos.Count > 0 Then
            guideAttribue = guidesDispos(1)
        Else
            guideAttribue = "AUCUN GUIDE DISPONIBLE"
        End If

        ' Construire liste guides disponibles
        listeGuidesDispos = ConstruireListeGuides(guidesDispos)

        ' ECRIRE DANS PLANNING
        wsPlanning.Cells(ligneP, 1).Value = wsVisites.Cells(i, 1).Value ' ID_Visite
        wsPlanning.Cells(ligneP, 2).Value = Format(dateVisite, "dd/mm/yyyy") ' Date format uniforme
        wsPlanning.Cells(ligneP, 3).Value = Format(heureDebut, "hh:mm") ' Heure
        wsPlanning.Cells(ligneP, 4).Value = typeVisite ' Type_Visite
        wsPlanning.Cells(ligneP, 5).Value = nbParticipants ' Nb_Participants
        wsPlanning.Cells(ligneP, 6).Value = heureFin ' Duree
        wsPlanning.Cells(ligneP, 7).Value = guideAttribue ' Guide_Attribue
        wsPlanning.Cells(ligneP, 8).Value = theme ' Theme
        wsPlanning.Cells(ligneP, 9).Value = niveau ' Niveau
        wsPlanning.Cells(ligneP, 10).Value = listeGuidesDispos ' Guides_Disponibles
        wsPlanning.Cells(ligneP, 11).Value = "A confirmer" ' Statut_Confirmation

        ligneP = ligneP + 1
    Next i

    Application.ScreenUpdating = True

    MsgBox "Planning genere avec succes !" & vbCrLf & _
           (ligneP - 2) & " visites traitees.", vbInformation

    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la generation du planning : " & Err.Description, vbCritical
End Sub

' ===== OBTENIR GUIDES DISPONIBLES (CORRIGE) =====
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

    ' STRUCTURE REELLE (selon analyse):
    ' Col 1: Guide (DATE au format texte bizarre)
    ' Col 2: Date (contient "OUI" ou vide)
    ' Col 3: Disponible (vide)
    ' Col 4: Commentaire (contient PRENOM)
    ' Col 5: Prenom (contient NOM)
    ' Col 6: Nom (vide)

    ' ⚠️ STRUCTURE INCORRECTE DANS EXCEL !
    ' Il faudra corriger l'import des donnees
    ' Pour l'instant, on adapte le code VBA

    For i = 2 To wsDispo.Cells(wsDispo.Rows.Count, 1).End(xlUp).Row
        On Error Resume Next

        ' Col 1 contient la date
        dateGuide = CDate(wsDispo.Cells(i, 1).Value)

        ' Col 2 contient OUI/NON
        disponible = UCase(Trim(wsDispo.Cells(i, 2).Value))

        ' Col 4 = Prenom, Col 5 = Nom
        nomGuide = Trim(wsDispo.Cells(i, 4).Value) & " " & Trim(wsDispo.Cells(i, 5).Value)

        If dateGuide = dateVisite And disponible = "OUI" Then
            ' Eviter doublons
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

' ===== FILTRER PAR SPECIALISATION (CORRIGE) =====
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

' ===== VERIFICATION SPECIALISATION =====
Private Function GuideAutoriseVisite(nomGuide As String, typeVisite As String) As Boolean
    ' Appel vers Module_Specialisations
    GuideAutoriseVisite = Module_Specialisations.GuideAutoriseVisite(nomGuide, typeVisite)
End Function
