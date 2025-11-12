Attribute VB_Name = "Module_Import_Visites"
'===============================================================================
' MODULE: Import de Visites depuis fichier Excel externe
' DESCRIPTION: Permet d'importer des visites depuis un fichier Excel de la cliente
' AUTEUR: Système de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

'===============================================================================
' FONCTION: ImporterVisitesDepuisFichier
' DESCRIPTION: Import automatique de visites depuis un fichier Excel externe
'===============================================================================
Public Sub ImporterVisitesDepuisFichier()
    Dim fichierSource As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsVisites As Worksheet
    Dim derniereLigne As Long
    Dim dernierID As Long
    Dim ligneSource As Long
    Dim ligneDestination As Long
    Dim nbImportees As Long
    Dim nbErreurs As Long

    ' Colonnes dans le fichier source (à adapter)
    Dim colDate As Integer
    Dim colHeure As Integer
    Dim colMusee As Integer
    Dim colType As Integer
    Dim colDuree As Integer
    Dim colVisiteurs As Integer

    On Error GoTo Erreur

    Application.ScreenUpdating = False

    ' Demander le fichier source
    fichierSource = Application.GetOpenFilename( _
        "Fichiers Excel (*.xlsx; *.xls), *.xlsx; *.xls", _
        , "Sélectionnez le fichier Excel contenant vos visites")

    If fichierSource = "Faux" Then
        MsgBox "Import annulé.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Message d'avertissement
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("ATTENTION :" & vbCrLf & vbCrLf & _
                     "- Les visites du fichier seront AJOUTÉES à l'onglet Visites" & vbCrLf & _
                     "- Les ID seront générés automatiquement (V001, V002, etc.)" & vbCrLf & _
                     "- Une sauvegarde est recommandée avant import" & vbCrLf & vbCrLf & _
                     "Continuer l'import ?", _
                     vbYesNo + vbQuestion, "Confirmer l'import")

    If reponse <> vbYes Then
        MsgBox "Import annulé.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Ouvrir le fichier source
    Set wbSource = Workbooks.Open(fichierSource, ReadOnly:=True)
    Set wsSource = wbSource.Worksheets(1) ' Premier onglet par défaut

    ' Demander les colonnes (ou détecter automatiquement)
    Dim detectionAuto As VbMsgBoxResult
    detectionAuto = MsgBox("Voulez-vous que je détecte automatiquement les colonnes ?" & vbCrLf & vbCrLf & _
                           "OUI = Détection automatique" & vbCrLf & _
                           "NON = Vous indiquez les colonnes manuellement", _
                           vbYesNo + vbQuestion, "Détection des colonnes")

    If detectionAuto = vbYes Then
        ' Détection automatique
        Call DetecterColonnes(wsSource, colDate, colHeure, colMusee, colType, colDuree, colVisiteurs)
    Else
        ' Demander manuellement
        colDate = Val(InputBox("Numéro de colonne pour la DATE (ex: 1 pour colonne A):", "Colonne Date", "1"))
        colHeure = Val(InputBox("Numéro de colonne pour l'HEURE (ex: 2 pour colonne B):", "Colonne Heure", "2"))
        colMusee = Val(InputBox("Numéro de colonne pour le MUSÉE (ex: 3 pour colonne C):", "Colonne Musée", "3"))
        colType = Val(InputBox("Numéro de colonne pour le TYPE DE VISITE (ex: 4 pour colonne D):", "Colonne Type", "4"))
        colDuree = Val(InputBox("Numéro de colonne pour la DURÉE (ex: 5 pour colonne E):", "Colonne Durée", "5"))
        colVisiteurs = Val(InputBox("Numéro de colonne pour le NOMBRE DE VISITEURS (ex: 6 pour colonne F):", "Colonne Visiteurs", "6"))
    End If

    ' Vérifier les colonnes
    If colType = 0 Then
        MsgBox "La colonne TYPE DE VISITE est obligatoire !" & vbCrLf & "Import annulé.", vbExclamation
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Préparer l'onglet Visites
    Set wsVisites = ThisWorkbook.Worksheets("Visites")

    ' Trouver la dernière ligne et le dernier ID
    derniereLigne = wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
    dernierID = TrouverDernierID(wsVisites)

    ' Commencer l'import
    ligneDestination = derniereLigne + 1
    nbImportees = 0
    nbErreurs = 0

    ' Boucle sur les lignes du fichier source (en commençant à la ligne 2 si en-têtes)
    Dim premiereLigneData As Long
    premiereLigneData = Val(InputBox("À quelle ligne commencent les données ?" & vbCrLf & _
                                     "(2 si vous avez des en-têtes, 1 sinon):", _
                                     "Première ligne de données", "2"))

    For ligneSource = premiereLigneData To wsSource.Cells(wsSource.Rows.Count, colType).End(xlUp).Row
        On Error Resume Next

        ' Lire les données source
        Dim dateVal As Variant
        Dim heureVal As Variant
        Dim museeVal As String
        Dim typeVal As String
        Dim dureeVal As Variant
        Dim visiteursVal As Variant

        If colDate > 0 Then dateVal = wsSource.Cells(ligneSource, colDate).Value
        If colHeure > 0 Then heureVal = wsSource.Cells(ligneSource, colHeure).Value
        If colMusee > 0 Then museeVal = wsSource.Cells(ligneSource, colMusee).Value
        typeVal = wsSource.Cells(ligneSource, colType).Value
        If colDuree > 0 Then dureeVal = wsSource.Cells(ligneSource, colDuree).Value
        If colVisiteurs > 0 Then visiteursVal = wsSource.Cells(ligneSource, colVisiteurs).Value

        ' Vérifier que le type de visite existe
        If Len(Trim(typeVal)) = 0 Then
            nbErreurs = nbErreurs + 1
            GoTo LigneSuivante
        End If

        ' Générer le nouvel ID
        dernierID = dernierID + 1
        Dim nouvelID As String
        nouvelID = "V" & Format(dernierID, "000")

        ' Écrire dans l'onglet Visites
        wsVisites.Cells(ligneDestination, 1).Value = nouvelID ' ID_Visite
        wsVisites.Cells(ligneDestination, 2).Value = dateVal ' Date
        wsVisites.Cells(ligneDestination, 3).Value = heureVal ' Heure
        wsVisites.Cells(ligneDestination, 4).Value = IIf(Len(museeVal) > 0, museeVal, "Musée du Quai Branly") ' Musée
        wsVisites.Cells(ligneDestination, 5).Value = typeVal ' Type_Visite
        wsVisites.Cells(ligneDestination, 6).Value = dureeVal ' Durée_Heures
        wsVisites.Cells(ligneDestination, 7).Value = visiteursVal ' Nombre_Visiteurs

        ' Statut
        If IsDate(dateVal) Then
            wsVisites.Cells(ligneDestination, 8).Value = "Planifié"
        Else
            wsVisites.Cells(ligneDestination, 8).Value = "À planifier"
        End If

        ligneDestination = ligneDestination + 1
        nbImportees = nbImportees + 1

        On Error GoTo Erreur

LigneSuivante:
    Next ligneSource

    ' Fermer le fichier source
    wbSource.Close SaveChanges:=False

    ' Message de confirmation
    Application.ScreenUpdating = True

    MsgBox "IMPORT TERMINÉ !" & vbCrLf & vbCrLf & _
           "✅ Visites importées : " & nbImportees & vbCrLf & _
           "⚠️ Erreurs (lignes ignorées) : " & nbErreurs & vbCrLf & vbCrLf & _
           "Prochaine étape :" & vbCrLf & _
           "→ Lancez la macro 'GenererPlanningAutomatique' pour attribuer les guides", _
           vbInformation, "Import réussi"

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'import :" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Ligne source : " & ligneSource, _
           vbExclamation, "Erreur"

    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If
End Sub

'===============================================================================
' FONCTION: TrouverDernierID
' DESCRIPTION: Trouve le dernier ID utilisé dans l'onglet Visites
'===============================================================================
Private Function TrouverDernierID(ws As Worksheet) As Long
    Dim derniereLigne As Long
    Dim i As Long
    Dim idVal As String
    Dim maxID As Long

    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    maxID = 0

    For i = 2 To derniereLigne
        idVal = ws.Cells(i, 1).Value

        If Left(idVal, 1) = "V" Then
            Dim numID As Long
            numID = Val(Mid(idVal, 2))
            If numID > maxID Then maxID = numID
        End If
    Next i

    TrouverDernierID = maxID
End Function

'===============================================================================
' FONCTION: DetecterColonnes
' DESCRIPTION: Détecte automatiquement les colonnes basées sur les en-têtes
'===============================================================================
Private Sub DetecterColonnes(ws As Worksheet, _
                            ByRef colDate As Integer, _
                            ByRef colHeure As Integer, _
                            ByRef colMusee As Integer, _
                            ByRef colType As Integer, _
                            ByRef colDuree As Integer, _
                            ByRef colVisiteurs As Integer)

    Dim col As Integer
    Dim header As String

    ' Initialiser à 0 (non trouvé)
    colDate = 0
    colHeure = 0
    colMusee = 0
    colType = 0
    colDuree = 0
    colVisiteurs = 0

    ' Parcourir les 20 premières colonnes
    For col = 1 To 20
        header = LCase(Trim(ws.Cells(1, col).Value))

        ' Détecter Date
        If colDate = 0 And (InStr(header, "date") > 0 Or InStr(header, "jour") > 0) Then
            colDate = col
        End If

        ' Détecter Heure
        If colHeure = 0 And (InStr(header, "heure") > 0 Or InStr(header, "horaire") > 0 Or InStr(header, "time") > 0) Then
            colHeure = col
        End If

        ' Détecter Musée
        If colMusee = 0 And (InStr(header, "musee") > 0 Or InStr(header, "musée") > 0 Or InStr(header, "lieu") > 0) Then
            colMusee = col
        End If

        ' Détecter Type
        If colType = 0 And (InStr(header, "type") > 0 Or InStr(header, "titre") > 0 Or InStr(header, "visite") > 0 Or InStr(header, "nom") > 0) Then
            colType = col
        End If

        ' Détecter Durée
        If colDuree = 0 And (InStr(header, "duree") > 0 Or InStr(header, "durée") > 0 Or InStr(header, "duration") > 0) Then
            colDuree = col
        End If

        ' Détecter Visiteurs
        If colVisiteurs = 0 And (InStr(header, "visiteur") > 0 Or InStr(header, "personne") > 0 Or InStr(header, "nombre") > 0 Or InStr(header, "participant") > 0) Then
            colVisiteurs = col
        End If
    Next col

    ' Afficher les résultats de la détection
    Dim message As String
    message = "Détection automatique des colonnes :" & vbCrLf & vbCrLf

    If colDate > 0 Then message = message & "✅ Date : Colonne " & colDate & vbCrLf
    If colHeure > 0 Then message = message & "✅ Heure : Colonne " & colHeure & vbCrLf
    If colMusee > 0 Then message = message & "✅ Musée : Colonne " & colMusee & vbCrLf
    If colType > 0 Then message = message & "✅ Type : Colonne " & colType & vbCrLf
    If colDuree > 0 Then message = message & "✅ Durée : Colonne " & colDuree & vbCrLf
    If colVisiteurs > 0 Then message = message & "✅ Visiteurs : Colonne " & colVisiteurs & vbCrLf

    message = message & vbCrLf & "Continuer avec ces colonnes ?"

    Dim reponse As VbMsgBoxResult
    reponse = MsgBox(message, vbYesNo + vbQuestion, "Colonnes détectées")

    If reponse <> vbYes Then
        ' L'utilisateur veut spécifier manuellement
        colDate = 0
        colHeure = 0
        colMusee = 0
        colType = 0
        colDuree = 0
        colVisiteurs = 0
    End If
End Sub
