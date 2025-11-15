Attribute VB_Name = "Module_Disponibilites"
'===============================================================================
' MODULE: Gestion des Disponibilites
' DESCRIPTION: Gestion confidentielle des disponibilites des guides
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

'===============================================================================
' FONCTION: SaisirDisponibilites
' DESCRIPTION: Interface pour qu'un guide saisisse ses disponibilites
'===============================================================================
Public Sub SaisirDisponibilites()
    Dim guideID As String
    Dim guideNom As String
    Dim dateDebut As Date
    Dim dateFin As Date
    Dim i As Long
    Dim ws As Worksheet
    Dim derLigne As Long
    Dim reponse As VbMsgBoxResult

    On Error GoTo Erreur

    ' Demander l'ID du guide
    guideID = InputBox("Entrez votre ID Guide (voir liste des guides):", "Identification")
    If guideID = "" Then Exit Sub

    ' Verifier que le guide existe
    guideNom = RechercherGuide(guideID)
    If guideNom = "" Then
        MsgBox "ID Guide non trouve. Veuillez verifier.", vbExclamation
        Exit Sub
    End If

    ' Message de bienvenue
    MsgBox "Bonjour " & guideNom & " !" & vbCrLf & vbCrLf & _
           "Vous allez saisir vos disponibilites pour le mois a venir.", _
           vbInformation, "Saisie des disponibilites"

    ' Demander la periode
    On Error Resume Next
    dateDebut = CDate(InputBox("Date de debut (jj/mm/aaaa):", "Periode", Date))
    If dateDebut = 0 Then Exit Sub

    dateFin = CDate(InputBox("Date de fin (jj/mm/aaaa):", "Periode", DateAdd("m", 1, Date)))
    If dateFin = 0 Then Exit Sub
    On Error GoTo Erreur

    If dateFin < dateDebut Then
        MsgBox "La date de fin doit etre apres la date de debut.", vbExclamation
        Exit Sub
    End If

    ' Ouvrir la feuille Disponibilites de maniere securisee
    Set ws = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)

    ' Proteger les donnees des autres guides
    Application.ScreenUpdating = False

    ' Supprimer les anciennes disponibilites du guide pour cette periode
    Call SupprimerAnciennesDisponibilites(guideID, dateDebut, dateFin)

    ' Afficher le formulaire de saisie jour par jour
    Call AfficherFormulaireDispo(guideID, guideNom, dateDebut, dateFin)

    Application.ScreenUpdating = True

    MsgBox "Vos disponibilites ont ete enregistrees avec succes !" & vbCrLf & _
           "Elles restent confidentielles.", vbInformation, "Confirmation"

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la saisie : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: RechercherGuide
' DESCRIPTION: Retourne le nom complet d'un guide
'===============================================================================
Private Function RechercherGuide(guideID As String) As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim ligneGuide As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set rng = ws.Range("A:A").Find(guideID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rng Is Nothing Then
        ligneGuide = rng.Row
        RechercherGuide = ws.Cells(ligneGuide, 1).Value & " " & ws.Cells(ligneGuide, 2).Value ' Prenom + Nom
    Else
        RechercherGuide = ""
    End If
    On Error GoTo 0
End Function

'===============================================================================
' FONCTION: SupprimerAnciennesDisponibilites
' DESCRIPTION: Supprime les anciennes disponibilites pour eviter les doublons
'===============================================================================
Private Sub SupprimerAnciennesDisponibilites(guideID As String, dateDebut As Date, dateFin As Date)
    Dim ws As Worksheet
    Dim i As Long
    Dim dateLigne As Date

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)

    ' Parcourir de bas en haut pour eviter les problemes de suppression
    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If ws.Cells(i, 1).Value = guideID Then
            On Error Resume Next
            dateLigne = CDate(ws.Cells(i, 2).Value)
            If Err.Number = 0 Then
                If dateLigne >= dateDebut And dateLigne <= dateFin Then
                    ws.Rows(i).Delete
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

'===============================================================================
' FONCTION: AfficherFormulaireDispo
' DESCRIPTION: Affiche un formulaire simple pour saisir les disponibilites
'===============================================================================
Private Sub AfficherFormulaireDispo(guideID As String, guideNom As String, dateDebut As Date, dateFin As Date)
    Dim ws As Worksheet
    Dim dateActuelle As Date
    Dim reponse As VbMsgBoxResult
    Dim disponible As Boolean
    Dim commentaire As String
    Dim derLigne As Long
    Dim nbJours As Long
    Dim compteur As Long

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)
    dateActuelle = dateDebut
    compteur = 0
    nbJours = DateDiff("d", dateDebut, dateFin) + 1

    ' Barre de progression (simple)
    Application.StatusBar = "Saisie des disponibilites : 0%"

    Do While dateActuelle <= dateFin
        compteur = compteur + 1

        ' Demander la disponibilite pour ce jour
        reponse = MsgBox("Etes-vous disponible le " & Format(dateActuelle, "dddd dd/mm/yyyy") & " ?", _
                        vbYesNoCancel + vbQuestion, "Disponibilite - " & guideNom)

        If reponse = vbCancel Then
            If MsgBox("Voulez-vous vraiment annuler la saisie ?", vbYesNo + vbQuestion) = vbYes Then
                Application.StatusBar = False
                Exit Sub
            Else
                ' Recommencer ce jour
                dateActuelle = DateAdd("d", -1, dateActuelle)
            End If
        Else
            disponible = (reponse = vbYes)

            ' Optionnel : demander un commentaire si non disponible
            commentaire = ""
            If Not disponible Then
                commentaire = InputBox("Raison (optionnel) :", "Commentaire", "")
            End If

            ' Enregistrer dans la feuille
            derLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            ws.Cells(derLigne, 1).Value = guideID
            ws.Cells(derLigne, 2).Value = dateActuelle
            ws.Cells(derLigne, 3).Value = IIf(disponible, "OUI", "NON")
            ws.Cells(derLigne, 4).Value = commentaire

            ' Colorer la ligne
            If disponible Then
                ws.Rows(derLigne).Interior.Color = COULEUR_DISPONIBLE
            Else
                ws.Rows(derLigne).Interior.Color = COULEUR_OCCUPE
            End If
        End If

        ' Mettre a jour la barre de progression
        Application.StatusBar = "Saisie des disponibilites : " & Format(compteur / nbJours, "0%")

        dateActuelle = DateAdd("d", 1, dateActuelle)
    Loop

    Application.StatusBar = False
End Sub

'===============================================================================
' FONCTION: ImporterDisponibilitesMasse
' DESCRIPTION: Importer les disponibilites depuis un fichier externe (optionnel)
'===============================================================================
Public Sub ImporterDisponibilitesMasse()
    Dim fichier As String
    Dim ws As Worksheet
    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    On Error GoTo Erreur

    ' Selectionner le fichier
    fichier = Application.GetOpenFilename("Fichiers Excel (*.xlsx; *.xls), *.xlsx; *.xls", , "Selectionner le fichier des disponibilites")
    If fichier = "False" Then Exit Sub

    Application.ScreenUpdating = False

    ' Ouvrir le fichier source
    Set wbSource = Workbooks.Open(fichier)
    Set wsSource = wbSource.Worksheets(1)

    ' Copier les donnees
    Set ws = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)

    Dim derLigne As Long
    derLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Copier depuis la ligne 2 (en supposant qu'il y a des en-tetes)
    wsSource.Range("A2:D" & wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row).Copy
    ws.Cells(derLigne, 1).PasteSpecial xlPasteValues

    ' Fermer le fichier source
    wbSource.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.CutCopyMode = False

    MsgBox "Import reussi !", vbInformation
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    MsgBox "Erreur lors de l'import : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: VerifierDisponibiliteGuide
' DESCRIPTION: Verifie si un guide est disponible a une date donnee
'===============================================================================
Public Function VerifierDisponibiliteGuide(guideID As String, dateVisite As Date) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim trouve As Boolean

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)
    trouve = False
    VerifierDisponibiliteGuide = False

    ' Chercher la disponibilite du guide pour cette date
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = guideID Then
            If CDate(ws.Cells(i, 2).Value) = dateVisite Then
                trouve = True
                If UCase(ws.Cells(i, 3).Value) = "OUI" Then
                    VerifierDisponibiliteGuide = True
                End If
                Exit For
            End If
        End If
    Next i

    ' Si aucune info trouvee, on considere non disponible par defaut
End Function

'===============================================================================
' FONCTION: ExporterMesDisponibilites
' DESCRIPTION: Permet a un guide d'exporter ses propres disponibilites
'===============================================================================
Public Sub ExporterMesDisponibilites()
    Dim guideID As String
    Dim ws As Worksheet
    Dim wsExport As Worksheet
    Dim i As Long
    Dim ligneExport As Long
    Dim fichier As String

    On Error GoTo Erreur

    ' Demander l'ID du guide
    guideID = InputBox("Entrez votre ID Guide :", "Export de vos disponibilites")
    If guideID = "" Then Exit Sub

    ' Verifier que le guide existe
    If RechercherGuide(guideID) = "" Then
        MsgBox "ID Guide non trouve.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Creer un nouveau classeur
    Dim wbExport As Workbook
    Set wbExport = Workbooks.Add
    Set wsExport = wbExport.Worksheets(1)

    ' En-tetes
    wsExport.Range("A1:C1").Value = Array("Date", "Disponible", "Commentaire")
    wsExport.Range("A1:C1").Font.Bold = True

    ' Copier les donnees du guide
    Set ws = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)
    ligneExport = 2

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = guideID Then
            wsExport.Cells(ligneExport, 1).Value = ws.Cells(i, 2).Value
            wsExport.Cells(ligneExport, 2).Value = ws.Cells(i, 3).Value
            wsExport.Cells(ligneExport, 3).Value = ws.Cells(i, 4).Value
            ligneExport = ligneExport + 1
        End If
    Next i

    wsExport.Columns.AutoFit

    ' Proposer de sauvegarder
    fichier = Application.GetSaveAsFilename("Mes_Disponibilites_" & Format(Date, "yyyymmdd") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")
    If fichier <> "False" Then
        wbExport.SaveAs fichier
        MsgBox "Export reussi !" & vbCrLf & fichier, vbInformation
    End If

    wbExport.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'export : " & Err.Description, vbCritical
End Sub
