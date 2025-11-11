Attribute VB_Name = "Module_Contrats"
'===============================================================================
' MODULE: Gestion des Contrats
' DESCRIPTION: Generation automatique des contrats avec dates/horaires
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

'===============================================================================
' NOTE: Anciennes fonctions GenererContratGuide() et RemplirModeleContrat()
' supprimees car obsoletes (systeme tarif horaire remplace par systeme cachets)
' Utiliser desormais GenererContratDebutMois() et GenererContratFinMois()
'===============================================================================

'===============================================================================
' FONCTION: ObtenirDureeVisiteContrat
' DESCRIPTION: Calcule la duree d'une visite (fonction auxiliaire)
'===============================================================================
Private Function ObtenirDureeVisiteContrat(idVisite As String) As Double
    Dim wsVisites As Worksheet
    Dim i As Long
    Dim heureDebut As Date
    Dim heureFin As Date

    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    ObtenirDureeVisiteContrat = 2 ' Duree par defaut

    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
        If wsVisites.Cells(i, 1).Value = idVisite Then
            On Error Resume Next
            heureDebut = CDate(wsVisites.Cells(i, 3).Value)
            heureFin = CDate(wsVisites.Cells(i, 4).Value)

            If Err.Number = 0 Then
                ObtenirDureeVisiteContrat = (heureFin - heureDebut) * 24
            End If
            On Error GoTo 0
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' FONCTION: GenererContratsEnMasse
' DESCRIPTION: Genere les contrats pour tous les guides d'un mois
'===============================================================================
Public Sub GenererContratsEnMasse()
    Dim wsPlanning As Worksheet
    Dim wsGuides As Worksheet
    Dim moisFiltre As String
    Dim moisCible As Integer
    Dim anneeCible As Integer
    Dim dictGuides As Object
    Dim i As Long
    Dim guideID As String
    Dim dateVisite As Date
    Dim dossier As String
    Dim compteur As Integer

    On Error GoTo Erreur

    ' Demander le mois
    moisFiltre = InputBox("Mois des contrats (MM/AAAA):", "Generation en masse", Format(Date, "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    moisCible = CInt(Left(moisFiltre, 2))
    anneeCible = CInt(Right(moisFiltre, 4))

    ' Demander le dossier de destination
    dossier = BrowseForFolder("Selectionnez le dossier pour enregistrer les contrats")
    If dossier = "" Then Exit Sub

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set dictGuides = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Identifier tous les guides ayant des visites ce mois
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = wsPlanning.Cells(i, 5).Value

        If guideID <> "NON ATTRIBUE" And guideID <> "" Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 Then
                If Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                    If Not dictGuides.exists(guideID) Then
                        dictGuides.Add guideID, True
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    ' Generer un contrat pour chaque guide
    compteur = 0
    Dim key As Variant

    For Each key In dictGuides.Keys
        guideID = CStr(key)

        ' Generer le contrat (version simplifiee sans interaction)
        Call GenererContratSilencieux(guideID, moisCible, anneeCible, dossier)
        compteur = compteur + 1
    Next key

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Generation terminee !" & vbCrLf & vbCrLf & _
           "Nombre de contrats generes : " & compteur & vbCrLf & _
           "Dossier : " & dossier, vbInformation

    Exit Sub

Erreur:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la generation en masse : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: GenererContratSilencieux
' DESCRIPTION: Genere un contrat sans interaction utilisateur (pour batch)
'===============================================================================
Private Sub GenererContratSilencieux(guideID As String, mois As Integer, annee As Integer, dossier As String)
    Dim guideNom As String
    Dim wsPlanning As Worksheet, wsGuides As Worksheet
    Dim wbContrat As Workbook, wsContrat As Worksheet
    Dim i As Long, nbVisites As Integer
    Dim dateVisite As Date
    Dim fichier As String

    On Error GoTo Erreur

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)

    ' Recuperer nom du guide
    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            guideNom = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value
            Exit For
        End If
    Next i

    If guideNom = "" Then Exit Sub

    ' Compter visites
    nbVisites = 0
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        If wsPlanning.Cells(i, 5).Value = guideID Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)
            If Err.Number = 0 And Month(dateVisite) = mois And Year(dateVisite) = annee Then
                nbVisites = nbVisites + 1
            End If
            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    If nbVisites = 0 Then Exit Sub

    ' Creer contrat simple
    Set wbContrat = Workbooks.Add
    Set wsContrat = wbContrat.Worksheets(1)

    wsContrat.Cells(1, 1).Value = "CONTRAT - " & guideNom
    wsContrat.Cells(2, 1).Value = "Mois : " & Format(DateSerial(annee, mois, 1), "MMMM YYYY")
    wsContrat.Cells(3, 1).Value = "Visites prevues : " & nbVisites

    ' Sauvegarder
    fichier = dossier & "\Contrat_" & Replace(guideNom, " ", "_") & "_" & _
              Format(DateSerial(annee, mois, 1), "yyyymm") & ".xlsx"

    wbContrat.SaveAs fichier
    wbContrat.Close SaveChanges:=False

    Exit Sub

Erreur:
    If Not wbContrat Is Nothing Then wbContrat.Close SaveChanges:=False
End Sub

'===============================================================================
' FONCTION: BrowseForFolder
' DESCRIPTION: Ouvre une boite de dialogue pour selectionner un dossier
'===============================================================================
Private Function BrowseForFolder(Optional titre As String = "Selectionner un dossier") As String
    Dim shellApp As Object
    Dim dossier As Object

    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    Set dossier = shellApp.BrowseForFolder(0, titre, 0, 0)

    If Not dossier Is Nothing Then
        BrowseForFolder = dossier.Self.Path
    Else
        BrowseForFolder = ""
    End If

    On Error GoTo 0
End Function

'===============================================================================
' FONCTION: AfficherContratsGeneres
' DESCRIPTION: Affiche la liste des contrats generes
'===============================================================================
Public Sub AfficherContratsGeneres()
    Dim wsContrats As Worksheet
    Dim msg As String
    Dim i As Long
    Dim derLigne As Long

    Set wsContrats = ThisWorkbook.Worksheets(FEUILLE_CONTRATS)
    derLigne = wsContrats.Cells(wsContrats.Rows.Count, 1).End(xlUp).Row

    If derLigne < 2 Then
        MsgBox "Aucun contrat genere pour le moment.", vbInformation
        Exit Sub
    End If

    msg = "CONTRATS GENERES" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf

    For i = 2 To derLigne
        msg = msg & wsContrats.Cells(i, 2).Value & " - " & _
                    wsContrats.Cells(i, 3).Value & " - " & _
                    wsContrats.Cells(i, 6).Value & vbCrLf
    Next i

    MsgBox msg, vbInformation, "Historique des contrats"
End Sub

'===============================================================================
' FONCTION: GenererContratDebutMois
' DESCRIPTION: Genere contrat debut de mois avec pre-planning et tarif minimum
'===============================================================================
Public Sub GenererContratDebutMois()
    Dim guideID As String
    Dim guideNom As String, emailGuide As String, telGuide As String
    Dim wsPlanning As Worksheet, wsGuides As Worksheet, wsConfig As Worksheet
    Dim wbContrat As Workbook, wsContrat As Worksheet
    Dim moisCible As Integer, anneeCible As Integer, moisFiltre As String
    Dim i As Long, ligne As Long
    Dim dateVisite As Date, nbJoursPrevus As Integer
    Dim tarifMinimum As Double
    Dim fichier As String
    Dim listeVisites As String

    On Error GoTo Erreur

    ' Demander le guide
    guideID = InputBox("Entrez l'ID du guide:", "Contrat Debut de Mois")
    If guideID = "" Then Exit Sub

    ' Demander le mois
    moisFiltre = InputBox("Mois du contrat (MM/AAAA):", "Periode", Format(DateAdd("m", 1, Date), "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    moisCible = CInt(Left(moisFiltre, 2))
    anneeCible = CInt(Right(moisFiltre, 4))

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsConfig = ThisWorkbook.Worksheets("Configuration")

    ' Recuperer infos guide
    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            guideNom = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value ' Prenom + Nom
            emailGuide = wsGuides.Cells(i, 3).Value ' Email
            telGuide = wsGuides.Cells(i, 4).Value ' Telephone
            Exit For
        End If
    Next i

    If guideNom = "" Then
        MsgBox "Guide non trouve.", vbExclamation
        Exit Sub
    End If

    ' Recuperer tarif minimum (80 par defaut)
    tarifMinimum = 80
    For i = 1 To wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
        If UCase(Trim(wsConfig.Cells(i, 1).Value)) = "TARIF_MINIMUM" Then
            tarifMinimum = wsConfig.Cells(i, 2).Value
            Exit For
        End If
    Next i

    Application.ScreenUpdating = False

    ' Compter les jours prevus
    nbJoursPrevus = 0
    listeVisites = ""

    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        If wsPlanning.Cells(i, 5).Value = guideID Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 And Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                If listeVisites <> "" Then listeVisites = listeVisites & vbCrLf
                listeVisites = listeVisites & Format(dateVisite, "dd/mm/yyyy")
                nbJoursPrevus = nbJoursPrevus + 1
            End If

            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    If nbJoursPrevus = 0 Then
        MsgBox "Aucune visite prevue pour ce guide ce mois-ci.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Creer le contrat
    Set wbContrat = Workbooks.Add
    Set wsContrat = wbContrat.Worksheets(1)
    wsContrat.Name = "Contrat_Provisoire"

    ' Remplir le contrat
    ligne = 1
    wsContrat.Cells(ligne, 1).Value = "CONTRAT DE VACATION - VERSION PROVISOIRE"
    wsContrat.Cells(ligne, 1).Font.Size = 16
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Range("A1:D1").Merge
    wsContrat.Range("A1").HorizontalAlignment = xlCenter

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "Guide :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 2).Value = guideNom

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Email :"
    wsContrat.Cells(ligne, 2).Value = emailGuide

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Telephone :"
    wsContrat.Cells(ligne, 2).Value = telGuide

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "Periode :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 2).Value = Format(DateSerial(anneeCible, moisCible, 1), "MMMM YYYY")

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "DATES PREVUES (PRE-PLANNING) :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 1).Font.Underline = True

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = listeVisites

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "REMUNERATION PREVUE :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 1).Font.Underline = True

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Nombre de jours prevus :"
    wsContrat.Cells(ligne, 2).Value = nbJoursPrevus & " jours"

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Tarif minimum par cachet :"
    wsContrat.Cells(ligne, 2).Value = Format(tarifMinimum, "#,##0.00") & " "

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "MONTANT MINIMUM ESTIME :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 2).Value = Format(nbJoursPrevus * tarifMinimum, "#,##0.00") & " "
    wsContrat.Cells(ligne, 2).Font.Bold = True
    wsContrat.Cells(ligne, 2).Font.Size = 12

    ligne = ligne + 3
    wsContrat.Cells(ligne, 1).Value = "Note : Ce contrat sera mis a jour en fin de mois avec les dates et montants exacts."
    wsContrat.Cells(ligne, 1).Font.Italic = True
    wsContrat.Cells(ligne, 1).Font.Color = RGB(255, 0, 0)
    wsContrat.Range("A" & ligne & ":D" & ligne).Merge

    wsContrat.Columns.AutoFit
    Application.ScreenUpdating = True

    ' Sauvegarder
    fichier = Application.GetSaveAsFilename("Contrat_Provisoire_" & Replace(guideNom, " ", "_") & "_" & _
                                            Format(DateSerial(anneeCible, moisCible, 1), "yyyymm") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")

    If fichier <> "False" Then
        wbContrat.SaveAs fichier
        MsgBox "Contrat provisoire genere !" & vbCrLf & vbCrLf & _
               "Guide : " & guideNom & vbCrLf & _
               "Mois : " & Format(DateSerial(anneeCible, moisCible, 1), "MMMM YYYY") & vbCrLf & _
               "Jours prevus : " & nbJoursPrevus & vbCrLf & _
               "Montant minimum : " & Format(nbJoursPrevus * tarifMinimum, "#,##0.00") & " ", _
               vbInformation
    End If

    wbContrat.Close SaveChanges:=False

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbContrat Is Nothing Then wbContrat.Close SaveChanges:=False
    MsgBox "Erreur lors de la generation du contrat : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: GenererContratFinMois
' DESCRIPTION: Genere contrat fin de mois avec dates reelles et cachets calcules
'===============================================================================
Public Sub GenererContratFinMois()
    Dim guideID As String
    Dim guideNom As String, emailGuide As String, telGuide As String
    Dim wsPlanning As Worksheet, wsGuides As Worksheet, wsCalculs As Worksheet
    Dim wbContrat As Workbook, wsContrat As Worksheet
    Dim moisCible As Integer, anneeCible As Integer, moisFiltre As String
    Dim i As Long, ligne As Long
    Dim dateVisite As Date, heureVisite As String
    Dim nbJoursReel As Integer, montantParCachet As Double, montantTotal As Double
    Dim fichier As String
    Dim dictJours As Object
    Dim listeVisitesDetail As String

    On Error GoTo Erreur

    ' Demander le guide
    guideID = InputBox("Entrez l'ID du guide:", "Contrat Fin de Mois")
    If guideID = "" Then Exit Sub

    ' Demander le mois
    moisFiltre = InputBox("Mois du contrat (MM/AAAA):", "Periode", Format(Date, "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    moisCible = CInt(Left(moisFiltre, 2))
    anneeCible = CInt(Right(moisFiltre, 4))

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsCalculs = ThisWorkbook.Worksheets(FEUILLE_CALCULS)
    Set dictJours = CreateObject("Scripting.Dictionary")

    ' Recuperer infos guide
    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            guideNom = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value ' Prenom + Nom
            emailGuide = wsGuides.Cells(i, 3).Value ' Email
            telGuide = wsGuides.Cells(i, 4).Value ' Telephone
            Exit For
        End If
    Next i

    If guideNom = "" Then
        MsgBox "Guide non trouve.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Collecter les visites reelles du mois
    listeVisitesDetail = ""

    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        If wsPlanning.Cells(i, 5).Value = guideID Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)
            heureVisite = wsPlanning.Cells(i, 3).Value

            If Err.Number = 0 And Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                Dim cleJour As String
                cleJour = Format(dateVisite, "yyyy-mm-dd")

                If Not dictJours.exists(cleJour) Then
                    dictJours.Add cleJour, dateVisite
                End If

                If listeVisitesDetail <> "" Then listeVisitesDetail = listeVisitesDetail & vbCrLf
                listeVisitesDetail = listeVisitesDetail & Format(dateVisite, "dd/mm/yyyy") & " a " & heureVisite
            End If

            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    nbJoursReel = dictJours.Count

    If nbJoursReel = 0 Then
        MsgBox "Aucune visite realisee pour ce guide ce mois-ci.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Chercher les calculs de paie pour ce guide
    montantParCachet = 0
    montantTotal = 0

    For i = 2 To wsCalculs.Cells(wsCalculs.Rows.Count, 1).End(xlUp).Row
        If wsCalculs.Cells(i, 1).Value = guideID Then
            montantParCachet = wsCalculs.Cells(i, 6).Value ' Colonne F
            montantTotal = wsCalculs.Cells(i, 7).Value     ' Colonne G
            Exit For
        End If
    Next i

    If montantParCachet = 0 Then
        MsgBox "Veuillez d'abord calculer les salaires (Module_Calculs.CalculerVisitesEtSalaires)", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Creer le contrat
    Set wbContrat = Workbooks.Add
    Set wsContrat = wbContrat.Worksheets(1)
    wsContrat.Name = "Contrat_Final"

    ' Remplir le contrat
    ligne = 1
    wsContrat.Cells(ligne, 1).Value = "CONTRAT DE VACATION - VERSION FINALE"
    wsContrat.Cells(ligne, 1).Font.Size = 16
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Range("A1:D1").Merge
    wsContrat.Range("A1").HorizontalAlignment = xlCenter
    wsContrat.Cells(ligne, 1).Interior.Color = RGB(146, 208, 80)

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "Guide :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 2).Value = guideNom

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Email :"
    wsContrat.Cells(ligne, 2).Value = emailGuide

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Telephone :"
    wsContrat.Cells(ligne, 2).Value = telGuide

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "Periode :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 2).Value = Format(DateSerial(anneeCible, moisCible, 1), "MMMM YYYY")

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "DATES ET HORAIRES REELS :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 1).Font.Underline = True

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = listeVisitesDetail
    wsContrat.Range("A" & ligne & ":D" & ligne).Merge
    wsContrat.Cells(ligne, 1).WrapText = True

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "REMUNERATION FINALE :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 1).Font.Underline = True
    wsContrat.Cells(ligne, 1).Interior.Color = RGB(255, 242, 204)

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Nombre de jours travailles (cachets) :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 2).Value = nbJoursReel & " cachets"
    wsContrat.Cells(ligne, 2).Font.Bold = True

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "Montant par cachet :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 2).Value = Format(montantParCachet, "#,##0.00") & " "
    wsContrat.Cells(ligne, 2).Font.Bold = True

    ligne = ligne + 1
    wsContrat.Cells(ligne, 1).Value = "MONTANT TOTAL DU :"
    wsContrat.Cells(ligne, 1).Font.Bold = True
    wsContrat.Cells(ligne, 1).Font.Size = 12
    wsContrat.Cells(ligne, 2).Value = Format(montantTotal, "#,##0.00") & " "
    wsContrat.Cells(ligne, 2).Font.Bold = True
    wsContrat.Cells(ligne, 2).Font.Size = 14
    wsContrat.Cells(ligne, 2).Font.Color = RGB(0, 128, 0)

    ligne = ligne + 2
    wsContrat.Cells(ligne, 1).Value = "Calcul : " & nbJoursReel & " cachets  " & Format(montantParCachet, "#,##0.00") & "  = " & Format(montantTotal, "#,##0.00") & " "
    wsContrat.Cells(ligne, 1).Font.Italic = True
    wsContrat.Range("A" & ligne & ":D" & ligne).Merge

    ligne = ligne + 3
    wsContrat.Cells(ligne, 1).Value = "Signature du guide :"
    wsContrat.Cells(ligne, 3).Value = "Signature de l'association :"

    wsContrat.Columns.AutoFit
    Application.ScreenUpdating = True

    ' Sauvegarder
    fichier = Application.GetSaveAsFilename("Contrat_Final_" & Replace(guideNom, " ", "_") & "_" & _
                                            Format(DateSerial(anneeCible, moisCible, 1), "yyyymm") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")

    If fichier <> "False" Then
        wbContrat.SaveAs fichier
        MsgBox "Contrat final genere !" & vbCrLf & vbCrLf & _
               "Guide : " & guideNom & vbCrLf & _
               "Mois : " & Format(DateSerial(anneeCible, moisCible, 1), "MMMM YYYY") & vbCrLf & _
               "Jours travailles : " & nbJoursReel & " cachets" & vbCrLf & _
               "Montant par cachet : " & Format(montantParCachet, "#,##0.00") & " " & vbCrLf & _
               "TOTAL : " & Format(montantTotal, "#,##0.00") & " ", _
               vbInformation
    End If

    wbContrat.Close SaveChanges:=False

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbContrat Is Nothing Then wbContrat.Close SaveChanges:=False
    MsgBox "Erreur lors de la generation du contrat : " & Err.Description, vbCritical
End Sub
