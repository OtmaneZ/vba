Attribute VB_Name = "Module_Planning"
'===============================================================================
' MODULE: Gestion du Planning
' DESCRIPTION: Croisement disponibilites/visites et attribution automatique
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

'===============================================================================
' FONCTION: GenererPlanningAutomatique
' DESCRIPTION: Attribue automatiquement les guides disponibles aux visites
'===============================================================================
Public Sub GenererPlanningAutomatique()
    Dim wsVisites As Worksheet
    Dim wsPlanning As Worksheet
    Dim wsDispo As Worksheet
    Dim wsGuides As Worksheet
    Dim i As Long
    Dim derLigneVisites As Long
    Dim derLignePlanning As Long
    Dim idVisite As String
    Dim dateVisite As Date
    Dim heureVisite As String
    Dim musee As String
    Dim guideAssigne As String
    Dim guidesDispos As Collection
    Dim compteurAttribue As Long
    Dim compteurNonAttribue As Long

    On Error GoTo Erreur

    Application.ScreenUpdating = False

    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsDispo = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)

    ' Verifier qu'il y a des visites
    derLigneVisites = wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
    If derLigneVisites < 2 Then
        MsgBox "Aucune visite a planifier.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Effacer l'ancien planning (conserver les en-tetes)
    derLignePlanning = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
    If derLignePlanning > 1 Then
        wsPlanning.Range("A2:F" & derLignePlanning).ClearContents
        wsPlanning.Range("A2:F" & derLignePlanning).Interior.ColorIndex = xlNone
    End If

    compteurAttribue = 0
    compteurNonAttribue = 0
    derLignePlanning = 2

    ' Parcourir chaque visite
    For i = 2 To derLigneVisites
        idVisite = wsVisites.Cells(i, 1).Value

        On Error Resume Next
        dateVisite = CDate(wsVisites.Cells(i, 2).Value)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo VisiteSuivante
        End If
        On Error GoTo Erreur

        heureVisite = wsVisites.Cells(i, 3).Value & " - " & wsVisites.Cells(i, 4).Value
        musee = wsVisites.Cells(i, 7).Value ' Nom_Structure

        ' Recuperer le type de visite (pas de colonne Categorie dans FEUILLE_VISITES)
        Dim typeVisite As String
        typeVisite = wsVisites.Cells(i, 6).Value ' Colonne F = "Type_Visite"

        ' Chercher un guide disponible ET autorise pour ce type de visite
        Set guidesDispos = ObtenirGuidesDisponibles(dateVisite)

        ' Filtrer les guides selon les specialisations
        Dim guidesAutorises As New Collection
        Dim k As Long
        For k = 1 To guidesDispos.Count
            If GuideAutoriseVisite(guidesDispos(k), typeVisite) Then
                guidesAutorises.Add guidesDispos(k)
            End If
        Next k

        If guidesAutorises.Count > 0 Then
            ' Selectionner le premier guide disponible ET autorise
            guideAssigne = guidesAutorises(1)

            ' Verifier que le guide n'a pas deja une visite ce jour-la
            If Not GuideDejaOccupe(guideAssigne, dateVisite, derLignePlanning - 1) Then
                ' Ajouter au planning
                wsPlanning.Cells(derLignePlanning, 1).Value = idVisite
                wsPlanning.Cells(derLignePlanning, 2).Value = dateVisite
                wsPlanning.Cells(derLignePlanning, 3).Value = heureVisite
                wsPlanning.Cells(derLignePlanning, 4).Value = musee ' Musée (col 4)
                wsPlanning.Cells(derLignePlanning, 5).Value = typeVisite ' Type_Visite (col 5)
                wsPlanning.Cells(derLignePlanning, 6).Value = wsVisites.Cells(i, 14).Value ' Durée (col 6)
                wsPlanning.Cells(derLignePlanning, 7).Value = guideAssigne ' Guide_Attribué (col 7)
                wsPlanning.Cells(derLignePlanning, 8).Value = wsVisites.Cells(i, 9).Value ' Thème (col H/8) - INVERSÉ
                wsPlanning.Cells(derLignePlanning, 9).Value = wsVisites.Cells(i, 8).Value ' Niveau (col I/9) - INVERSÉ

                ' Appliquer le code couleur selon le type de visite
                AppliquerCodeCouleurLigne wsPlanning, derLignePlanning, typeVisite

                derLignePlanning = derLignePlanning + 1
                compteurAttribue = compteurAttribue + 1
            Else
                ' Guide deja occupe, chercher le suivant (parmi les autorises)
                Dim trouve As Boolean
                trouve = False
                Dim j As Integer

                For j = 2 To guidesAutorises.Count
                    If Not GuideDejaOccupe(guidesAutorises(j), dateVisite, derLignePlanning - 1) Then
                        guideAssigne = guidesAutorises(j)
                        trouve = True
                        Exit For
                    End If
                Next j

                If trouve Then
                    wsPlanning.Cells(derLignePlanning, 1).Value = idVisite
                    wsPlanning.Cells(derLignePlanning, 2).Value = dateVisite
                    wsPlanning.Cells(derLignePlanning, 3).Value = heureVisite
                    wsPlanning.Cells(derLignePlanning, 4).Value = musee ' Musée (col 4)
                    wsPlanning.Cells(derLignePlanning, 5).Value = typeVisite ' Type_Visite (col 5)
                    wsPlanning.Cells(derLignePlanning, 6).Value = wsVisites.Cells(i, 14).Value ' Durée (col 6)
                    wsPlanning.Cells(derLignePlanning, 7).Value = guideAssigne ' Guide_Attribué (col 7)
                    wsPlanning.Cells(derLignePlanning, 8).Value = wsVisites.Cells(i, 9).Value ' Thème (col H/8)
                    wsPlanning.Cells(derLignePlanning, 9).Value = wsVisites.Cells(i, 8).Value ' Niveau (col I/9)
                    AppliquerCodeCouleurLigne wsPlanning, derLignePlanning, typeVisite
                    derLignePlanning = derLignePlanning + 1
                    compteurAttribue = compteurAttribue + 1
                Else
                    ' Aucun guide autorise disponible
                    wsPlanning.Cells(derLignePlanning, 1).Value = idVisite
                    wsPlanning.Cells(derLignePlanning, 2).Value = dateVisite
                    wsPlanning.Cells(derLignePlanning, 3).Value = heureVisite
                    wsPlanning.Cells(derLignePlanning, 4).Value = musee ' Musée (col 4)
                    wsPlanning.Cells(derLignePlanning, 5).Value = typeVisite ' Type_Visite (col 5)
                    wsPlanning.Cells(derLignePlanning, 6).Value = wsVisites.Cells(i, 14).Value ' Durée (col 6)
                    wsPlanning.Cells(derLignePlanning, 7).Value = "NON ATTRIBUE" ' Guide_Attribué (col 7)
                    wsPlanning.Cells(derLignePlanning, 8).Value = wsVisites.Cells(i, 9).Value ' Thème (col H/8)
                    wsPlanning.Cells(derLignePlanning, 9).Value = wsVisites.Cells(i, 8).Value ' Niveau (col I/9)
                    wsPlanning.Cells(derLignePlanning, 11).Value = "Aucun guide autorise disponible" ' Statut (col 11)
                    wsPlanning.Rows(derLignePlanning).Interior.Color = COULEUR_OCCUPE
                    derLignePlanning = derLignePlanning + 1
                    compteurNonAttribue = compteurNonAttribue + 1
                End If
            End If
        Else
            ' Aucun guide autorise disponible pour cette visite
            wsPlanning.Cells(derLignePlanning, 1).Value = idVisite
            wsPlanning.Cells(derLignePlanning, 2).Value = dateVisite
            wsPlanning.Cells(derLignePlanning, 3).Value = heureVisite
            wsPlanning.Cells(derLignePlanning, 4).Value = musee ' Musée (col 4)
            wsPlanning.Cells(derLignePlanning, 5).Value = typeVisite ' Type_Visite (col 5)
            wsPlanning.Cells(derLignePlanning, 6).Value = wsVisites.Cells(i, 14).Value ' Durée (col 6)
            wsPlanning.Cells(derLignePlanning, 7).Value = "NON ATTRIBUE" ' Guide_Attribué (col 7)
            wsPlanning.Cells(derLignePlanning, 8).Value = wsVisites.Cells(i, 9).Value ' Thème (col H/8)
            wsPlanning.Cells(derLignePlanning, 9).Value = wsVisites.Cells(i, 8).Value ' Niveau (col I/9)
            wsPlanning.Cells(derLignePlanning, 11).Value = "Aucun guide autorise pour ce type de visite" ' Statut (col 11)

            ' Colorer en rouge
            wsPlanning.Rows(derLignePlanning).Interior.Color = COULEUR_OCCUPE

            derLignePlanning = derLignePlanning + 1
            compteurNonAttribue = compteurNonAttribue + 1
        End If

VisiteSuivante:
    Next i

    wsPlanning.Columns.AutoFit
    Application.ScreenUpdating = True

    ' Message de resume
    MsgBox "Planning genere !" & vbCrLf & vbCrLf & _
           "[OK] Visites attribuees : " & compteurAttribue & vbCrLf & _
           "[X] Visites non attribuees : " & compteurNonAttribue, _
           vbInformation, "Generation du Planning"

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la generation du planning : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: ObtenirGuidesDisponibles
' DESCRIPTION: Retourne la liste des guides disponibles pour une date
'===============================================================================
Private Function ObtenirGuidesDisponibles(dateVisite As Date) As Collection
    Dim wsDispo As Worksheet
    Dim col As New Collection
    Dim i As Long
    Dim guideID As String

    Set wsDispo = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)

    ' Parcourir les disponibilites
    For i = 2 To wsDispo.Cells(wsDispo.Rows.Count, 1).End(xlUp).Row
        On Error Resume Next
        If CDate(wsDispo.Cells(i, 2).Value) = dateVisite Then
            If UCase(wsDispo.Cells(i, 3).Value) = "OUI" Then
                guideID = wsDispo.Cells(i, 1).Value

                ' Ajouter seulement si pas deja dans la collection
                Dim existe As Boolean
                existe = False
                Dim j As Integer
                For j = 1 To col.Count
                    If col(j) = guideID Then
                        existe = True
                        Exit For
                    End If
                Next j

                If Not existe Then
                    col.Add guideID
                End If
            End If
        End If
        On Error GoTo 0
    Next i

    Set ObtenirGuidesDisponibles = col
End Function

'===============================================================================
' FONCTION: GuideDejaOccupe
' DESCRIPTION: Verifie si un guide a deja une visite assignee ce jour-la
'===============================================================================
Private Function GuideDejaOccupe(guideID As String, dateVisite As Date, derniereLigne As Long) As Boolean
    Dim wsPlanning As Worksheet
    Dim i As Long

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    GuideDejaOccupe = False

    For i = 2 To derniereLigne
        On Error Resume Next
        If wsPlanning.Cells(i, 7).Value = guideID Then ' Guide_Attribué (col 7)
            If CDate(wsPlanning.Cells(i, 2).Value) = dateVisite Then
                GuideDejaOccupe = True
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next i
End Function

'===============================================================================
' FONCTION: ObtenirNomGuide
' DESCRIPTION: Retourne le nom complet d'un guide
'===============================================================================
Private Function ObtenirNomGuide(guideID As String) As String
    Dim wsGuides As Worksheet
    Dim i As Long

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    ObtenirNomGuide = ""

    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            ObtenirNomGuide = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value ' Prenom + Nom
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' FONCTION: AfficherGuidesDisponiblesPourVisite
' DESCRIPTION: Affiche les guides disponibles pour une visite specifique
'===============================================================================
Public Sub AfficherGuidesDisponiblesPourVisite()
    Dim dateVisite As Date
    Dim guidesDispos As Collection
    Dim msg As String
    Dim guide As Variant

    On Error Resume Next
    dateVisite = CDate(InputBox("Date de la visite (jj/mm/aaaa):", "Rechercher guides disponibles"))
    If dateVisite = 0 Then Exit Sub
    On Error GoTo 0

    Set guidesDispos = ObtenirGuidesDisponibles(dateVisite)

    If guidesDispos.Count = 0 Then
        MsgBox "Aucun guide disponible pour le " & Format(dateVisite, "dd/mm/yyyy"), vbInformation
    Else
        msg = "Guides disponibles le " & Format(dateVisite, "dd/mm/yyyy") & " :" & vbCrLf & vbCrLf

        For Each guide In guidesDispos
            msg = msg & "- " & ObtenirNomGuide(CStr(guide)) & " (ID: " & guide & ")" & vbCrLf
        Next guide

        MsgBox msg, vbInformation, "Guides disponibles"
    End If
End Sub

'===============================================================================
' FONCTION: ModifierAttribution
' DESCRIPTION: Permet de modifier manuellement l'attribution d'un guide
'===============================================================================
Public Sub ModifierAttribution()
    Dim wsPlanning As Worksheet
    Dim idVisite As String
    Dim nouveauGuide As String
    Dim ligneVisite As Long
    Dim trouve As Boolean

    On Error GoTo Erreur

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)

    ' Demander l'ID de la visite
    idVisite = InputBox("Entrez l'ID de la visite a modifier:", "Modification")
    If idVisite = "" Then Exit Sub

    ' Chercher la visite dans le planning
    trouve = False
    For ligneVisite = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        If wsPlanning.Cells(ligneVisite, 1).Value = idVisite Then
            trouve = True
            Exit For
        End If
    Next ligneVisite

    If Not trouve Then
        MsgBox "Visite non trouvee dans le planning.", vbExclamation
        Exit Sub
    End If

    ' Afficher les infos de la visite
    Dim msg As String
    msg = "Visite : " & wsPlanning.Cells(ligneVisite, 4).Value & vbCrLf ' Musée (col 4)
    msg = msg & "Date : " & Format(wsPlanning.Cells(ligneVisite, 2).Value, "dd/mm/yyyy") & vbCrLf
    msg = msg & "Heure : " & wsPlanning.Cells(ligneVisite, 3).Value & vbCrLf
    msg = msg & "Type : " & wsPlanning.Cells(ligneVisite, 5).Value & vbCrLf & vbCrLf ' Type_Visite (col 5)
    msg = msg & "Guide actuellement assigne : " & wsPlanning.Cells(ligneVisite, 7).Value ' Guide_Attribué (col 7)

    MsgBox msg, vbInformation, "Informations visite"

    ' Demander le nouveau guide
    nouveauGuide = InputBox("Entrez l'ID du nouveau guide:", "Nouveau guide")
    If nouveauGuide = "" Then Exit Sub

    ' Verifier que le guide existe
    If ObtenirNomGuide(nouveauGuide) = "" Then
        MsgBox "Guide non trouve.", vbExclamation
        Exit Sub
    End If

    ' Mettre a jour
    wsPlanning.Cells(ligneVisite, 7).Value = nouveauGuide ' Guide_Attribué (col 7)
    wsPlanning.Cells(ligneVisite, 11).Value = "Modifié manuellement le " & Format(Now, "dd/mm/yyyy hh:nn") ' Statut (col 11)
    wsPlanning.Rows(ligneVisite).Interior.Color = COULEUR_ASSIGNE

    MsgBox "Attribution modifiee avec succes !", vbInformation

    Exit Sub

Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: ExporterPlanning
' DESCRIPTION: Export le planning dans un fichier separe
'===============================================================================
Public Sub ExporterPlanning()
    Dim wsPlanning As Worksheet
    Dim wbExport As Workbook
    Dim fichier As String

    On Error GoTo Erreur

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)

    Application.ScreenUpdating = False

    ' Creer un nouveau classeur
    Set wbExport = Workbooks.Add

    ' Copier le planning
    wsPlanning.UsedRange.Copy
    wbExport.Worksheets(1).Range("A1").PasteSpecial xlPasteAll
    wbExport.Worksheets(1).Name = "Planning"
    wbExport.Worksheets(1).Columns.AutoFit

    Application.CutCopyMode = False

    ' Proposer de sauvegarder
    fichier = Application.GetSaveAsFilename("Planning_Guides_" & Format(Date, "yyyymmdd") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")
    If fichier <> "False" Then
        wbExport.SaveAs fichier
        MsgBox "Planning exporte avec succes !" & vbCrLf & fichier, vbInformation
    End If

    wbExport.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'export : " & Err.Description, vbCritical

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

End Sub
