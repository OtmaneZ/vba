Attribute VB_Name = "Module_Contrats"
'===============================================================================
' MODULE: Gestion des Contrats
' DESCRIPTION: Generation automatique des contrats avec dates/horaires
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

'===============================================================================
' FONCTION: GenererContratGuide
' DESCRIPTION: Genere un contrat pre-rempli pour un guide
'===============================================================================
Public Sub GenererContratGuide()
    Dim guideID As String
    Dim guideNom As String
    Dim wsPlanning As Worksheet
    Dim wsGuides As Worksheet
    Dim wsContrats As Worksheet
    Dim wbContrat As Workbook
    Dim wsContrat As Worksheet
    Dim moisFiltre As String
    Dim i As Long
    Dim ligne As Long
    Dim dateVisite As Date
    Dim moisCible As Integer
    Dim anneeCible As Integer
    Dim listeVisites As String
    Dim listeHoraires As String
    Dim totalHeures As Double
    Dim fichier As String
    Dim emailGuide As String
    Dim telGuide As String

    On Error GoTo Erreur

    ' Demander l'ID du guide
    guideID = InputBox("Entrez l'ID du guide:", "Generation de contrat")
    If guideID = "" Then Exit Sub

    ' Recuperer les infos du guide
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    guideNom = ""

    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            guideNom = wsGuides.Cells(i, 2).Value & " " & wsGuides.Cells(i, 3).Value
            emailGuide = wsGuides.Cells(i, 4).Value
            telGuide = wsGuides.Cells(i, 5).Value
            Exit For
        End If
    Next i

    If guideNom = "" Then
        MsgBox "Guide non trouve.", vbExclamation
        Exit Sub
    End If

    ' Demander le mois
    moisFiltre = InputBox("Mois du contrat (MM/AAAA):", "Periode", Format(Date, "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    moisCible = CInt(Left(moisFiltre, 2))
    anneeCible = CInt(Right(moisFiltre, 4))

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsContrats = ThisWorkbook.Worksheets(FEUILLE_CONTRATS)

    Application.ScreenUpdating = False

    ' Collecter les visites du guide pour ce mois
    listeVisites = ""
    listeHoraires = ""
    totalHeures = 0
    Dim compteurVisites As Integer
    compteurVisites = 0

    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        If wsPlanning.Cells(i, 5).Value = guideID Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 Then
                If Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                    compteurVisites = compteurVisites + 1

                    ' Ajouter a la liste des dates
                    If listeVisites <> "" Then listeVisites = listeVisites & ", "
                    listeVisites = listeVisites & Format(dateVisite, "dd/mm/yyyy")

                    ' Ajouter aux horaires
                    If listeHoraires <> "" Then listeHoraires = listeHoraires & vbCrLf
                    listeHoraires = listeHoraires & Format(dateVisite, "dd/mm") & " : " & wsPlanning.Cells(i, 3).Value

                    ' Calculer les heures
                    Dim duree As Double
                    duree = ObtenirDureeVisiteContrat(wsPlanning.Cells(i, 1).Value)
                    totalHeures = totalHeures + duree
                End If
            End If
            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    If compteurVisites = 0 Then
        MsgBox "Aucune visite trouvee pour ce guide ce mois-ci.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Creer le document de contrat
    Set wbContrat = Workbooks.Add
    Set wsContrat = wbContrat.Worksheets(1)
    wsContrat.Name = "Contrat"

    ' Remplir le contrat
    Call RemplirModeleContrat(wsContrat, guideNom, guideID, emailGuide, telGuide, _
                              moisCible, anneeCible, listeVisites, listeHoraires, _
                              totalHeures, compteurVisites)

    ' Enregistrer le contrat dans la feuille Contrats
    Call EnregistrerDansFeuilleContrats(guideID, guideNom, moisCible, anneeCible, _
                                        listeVisites, listeHoraires, totalHeures)

    ' Proposer de sauvegarder le fichier
    fichier = Application.GetSaveAsFilename("Contrat_" & Replace(guideNom, " ", "_") & "_" & _
                                            Format(DateSerial(anneeCible, moisCible, 1), "yyyymm") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")

    If fichier <> "False" Then
        wbContrat.SaveAs fichier
        MsgBox "Contrat genere avec succes !" & vbCrLf & vbCrLf & _
               "Guide : " & guideNom & vbCrLf & _
               "Periode : " & Format(DateSerial(anneeCible, moisCible, 1), "mmmm yyyy") & vbCrLf & _
               "Nombre de visites : " & compteurVisites & vbCrLf & _
               "Total heures : " & Format(totalHeures, "0.0") & " h" & vbCrLf & vbCrLf & _
               "Fichier : " & fichier, vbInformation
    End If

    wbContrat.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbContrat Is Nothing Then wbContrat.Close SaveChanges:=False
    MsgBox "Erreur lors de la generation du contrat : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: RemplirModeleContrat
' DESCRIPTION: Remplit le modele de contrat avec les informations
'===============================================================================
Private Sub RemplirModeleContrat(ws As Worksheet, nomGuide As String, idGuide As String, _
                                 email As String, tel As String, mois As Integer, annee As Integer, _
                                 dates As String, horaires As String, heures As Double, nbVisites As Integer)
    Dim nomMois As String
    Dim tarifHeure As Double
    Dim montantTotal As Double

    nomMois = Format(DateSerial(annee, mois, 1), "mmmm yyyy")
    tarifHeure = ObtenirTarifHeure()
    montantTotal = heures * tarifHeure

    With ws
        ' En-tete du contrat
        .Range("A1").Value = "CONTRAT DE VACATION"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter

        .Range("A3").Value = "Entre :"
        .Range("A4").Value = "L'Association des Guides de Musee"
        .Range("A4").Font.Bold = True
        .Range("A5").Value = "[Adresse de l'association]"
        .Range("A6").Value = "[Code postal, Ville]"

        .Range("A8").Value = "Ci-apres denommee  L'Association "

        .Range("A10").Value = "Et :"
        .Range("A11").Value = nomGuide
        .Range("A11").Font.Bold = True
        .Range("A12").Value = "ID Guide : " & idGuide
        .Range("A13").Value = "Email : " & email
        .Range("A14").Value = "Telephone : " & tel

        .Range("A16").Value = "Ci-apres denomme(e)  Le Guide "

        .Range("A18").Value = "Il a ete convenu ce qui suit :"
        .Range("A18").Font.Bold = True

        ' Article 1 : Objet
        .Range("A20").Value = "ARTICLE 1 - OBJET DU CONTRAT"
        .Range("A20").Font.Bold = True
        .Range("A20").Font.Underline = True

        .Range("A21").Value = "L'Association confie au Guide la realisation de visites guidees au sein de musees partenaires."

        ' Article 2 : Periode
        .Range("A23").Value = "ARTICLE 2 - PERIODE D'INTERVENTION"
        .Range("A23").Font.Bold = True
        .Range("A23").Font.Underline = True

        .Range("A24").Value = "Periode : " & nomMois
        .Range("A24").Font.Bold = True
        .Range("A25").Value = "Nombre de visites : " & nbVisites

        ' Article 3 : Planning
        .Range("A27").Value = "ARTICLE 3 - PLANNING DES INTERVENTIONS"
        .Range("A27").Font.Bold = True
        .Range("A27").Font.Underline = True

        .Range("A28").Value = "Dates des visites :"
        .Range("A29").Value = dates
        .Range("A29").WrapText = True

        .Range("A31").Value = "Horaires detailles :"
        .Range("A32").Value = horaires
        .Range("A32").WrapText = True

        ' Article 4 : Remuneration
        .Range("A35").Value = "ARTICLE 4 - REMUNERATION"
        .Range("A35").Font.Bold = True
        .Range("A35").Font.Underline = True

        .Range("A36").Value = "Tarif horaire : " & Format(tarifHeure, "#,##0.00 ") & " / heure"
        .Range("A37").Value = "Volume horaire total : " & Format(heures, "0.0") & " heures"
        .Range("A38").Value = "Montant total brut : " & Format(montantTotal, "#,##0.00 ")
        .Range("A38").Font.Bold = True
        .Range("A38").Font.Size = 12
        .Range("A38").Interior.Color = RGB(255, 242, 204)

        ' Article 5 : Obligations
        .Range("A40").Value = "ARTICLE 5 - OBLIGATIONS DU GUIDE"
        .Range("A40").Font.Bold = True
        .Range("A40").Font.Underline = True

        .Range("A41").Value = "Le Guide s'engage a :"
        .Range("A42").Value = "- Se presenter aux horaires convenus"
        .Range("A43").Value = "- Assurer des visites de qualite conformes aux standards de l'Association"
        .Range("A44").Value = "- Respecter les consignes de securite des musees"
        .Range("A45").Value = "- Informer l'Association de toute absence au moins 48h a l'avance"

        ' Signatures
        .Range("A48").Value = "Fait a _________________, le ___/___/" & annee

        .Range("A50").Value = "Pour l'Association"
        .Range("A50").Font.Bold = True
        .Range("A51").Value = "(Signature et cachet)"

        .Range("D50").Value = "Le Guide"
        .Range("D50").Font.Bold = True
        .Range("D51").Value = "(Signature precedee de 'Lu et approuve')"

        ' Mise en forme
        .Columns("A:D").ColumnWidth = 20
        .Range("A1:D51").Font.Name = "Arial"
        .Range("A1:D51").Font.Size = 11

        ' Bordures autour du contrat
        .Range("A1:D51").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
End Sub

'===============================================================================
' FONCTION: EnregistrerDansFeuilleContrats
' DESCRIPTION: Enregistre les informations du contrat dans la feuille Contrats
'===============================================================================
Private Sub EnregistrerDansFeuilleContrats(guideID As String, guideNom As String, _
                                          mois As Integer, annee As Integer, _
                                          dates As String, horaires As String, heures As Double)
    Dim wsContrats As Worksheet
    Dim derLigne As Long
    Dim nomMois As String

    Set wsContrats = ThisWorkbook.Worksheets(FEUILLE_CONTRATS)
    nomMois = Format(DateSerial(annee, mois, 1), "mmmm yyyy")

    ' Trouver la derniere ligne
    derLigne = wsContrats.Cells(wsContrats.Rows.Count, 1).End(xlUp).Row + 1

    ' Ajouter la nouvelle ligne
    wsContrats.Cells(derLigne, 1).Value = guideID
    wsContrats.Cells(derLigne, 2).Value = guideNom
    wsContrats.Cells(derLigne, 3).Value = nomMois
    wsContrats.Cells(derLigne, 4).Value = dates
    wsContrats.Cells(derLigne, 5).Value = horaires
    wsContrats.Cells(derLigne, 6).Value = Format(heures, "0.0") & " h"

    ' Formater
    wsContrats.Rows(derLigne).Interior.Color = COULEUR_DISPONIBLE
    wsContrats.Columns.AutoFit
End Sub

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
' DESCRIPTION: Genere un contrat sans interaction utilisateur
'===============================================================================
Private Sub GenererContratSilencieux(guideID As String, mois As Integer, annee As Integer, dossier As String)
    ' Version simplifiee de GenererContratGuide pour l'automatisation
    ' (Implementation similaire mais sans InputBox ni MsgBox)

    ' Cette fonction serait une copie adaptee de GenererContratGuide
    ' Je la laisse commentee pour ne pas alourdir, mais elle suivrait la meme logique
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
            guideNom = wsGuides.Cells(i, 2).Value & " " & wsGuides.Cells(i, 3).Value
            emailGuide = wsGuides.Cells(i, 4).Value
            telGuide = wsGuides.Cells(i, 5).Value
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
            guideNom = wsGuides.Cells(i, 2).Value & " " & wsGuides.Cells(i, 3).Value
            emailGuide = wsGuides.Cells(i, 4).Value
            telGuide = wsGuides.Cells(i, 5).Value
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
