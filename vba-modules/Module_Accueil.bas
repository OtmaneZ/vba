Attribute VB_Name = "Module_Accueil"
' ============================================
' MODULE ACCUEIL
' Gestion de l'ecran d'accueil et evenements
' ============================================

Option Explicit

' ============================================
' Creer la feuille d'accueil
' ============================================
Sub CreerFeuilleAccueil()
    Dim wsAccueil As Worksheet
    Dim ligneActuelle As Long

    ' Supprimer l'ancienne feuille d'accueil si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Accueil").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Creer la nouvelle feuille d'accueil
    Set wsAccueil = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    wsAccueil.Name = "Accueil"

    With wsAccueil
        ' Configuration de la page
        .Tab.Color = RGB(70, 173, 71)
        .Cells.Font.Name = "Arial"
        .Cells.Font.Size = 11

        ' Largeur des colonnes
        .Columns("A:A").ColumnWidth = 5
        .Columns("B:E").ColumnWidth = 15
        .Columns("F:F").ColumnWidth = 5

        ' ===== TITRE =====
        ligneActuelle = 3
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "*** SYSTEME DE PLANNING ***"
        With .Range("B" & ligneActuelle)
            .Font.Bold = True
            .Font.Size = 20
            .Font.Color = RGB(70, 173, 71)
            .HorizontalAlignment = xlCenter
        End With

        ligneActuelle = ligneActuelle + 1
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "Gestion des Guides de Musee"
        With .Range("B" & ligneActuelle)
            .Font.Size = 14
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
        End With

        ' ===== SEPARATEUR =====
        ligneActuelle = ligneActuelle + 2
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("B" & ligneActuelle).Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B" & ligneActuelle).Borders(xlEdgeBottom).Color = RGB(70, 173, 71)

        ' ===== BLOC GUIDE =====
        ligneActuelle = ligneActuelle + 3
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "[GUIDE] JE SUIS UN GUIDE"
        With .Range("B" & ligneActuelle)
            .Font.Bold = True
            .Font.Size = 16
            .Interior.Color = RGB(70, 173, 71)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        .Range("B" & ligneActuelle).RowHeight = 40

        ligneActuelle = ligneActuelle + 1
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "Consulter mon planning personnel"
        With .Range("B" & ligneActuelle)
            .Font.Size = 11
            .Font.Italic = True
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(242, 249, 242)
            .Borders.LineStyle = xlContinuous
        End With
        .Range("B" & ligneActuelle).RowHeight = 30

        ligneActuelle = ligneActuelle + 1
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "- Voir mes visites a venir" & vbLf & _
                                             "- Confirmer ou refuser des missions" & vbLf & _
                                             "- Exporter mon planning"
        With .Range("B" & ligneActuelle)
            .Font.Size = 10
            .Font.Color = RGB(50, 50, 50)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .Interior.Color = RGB(242, 249, 242)
            .Borders.LineStyle = xlContinuous
            .WrapText = True
        End With
        .Range("B" & ligneActuelle).RowHeight = 60

        ' Stocker la ligne du bouton Guide pour l'evenement clic
        .Range("Z1").Value = ligneActuelle - 2 ' Ligne du titre "JE SUIS UN GUIDE"

        ' ===== SEPARATEUR =====
        ligneActuelle = ligneActuelle + 2

        ' ===== BLOC ADMIN =====
        ligneActuelle = ligneActuelle + 1
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "[ADMIN] JE SUIS L'ADMINISTRATEUR"
        With .Range("B" & ligneActuelle)
            .Font.Bold = True
            .Font.Size = 16
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        .Range("B" & ligneActuelle).RowHeight = 40

        ligneActuelle = ligneActuelle + 1
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "Acces complet au systeme"
        With .Range("B" & ligneActuelle)
            .Font.Size = 11
            .Font.Italic = True
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(237, 244, 252)
            .Borders.LineStyle = xlContinuous
        End With
        .Range("B" & ligneActuelle).RowHeight = 30

        ligneActuelle = ligneActuelle + 1
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "- Gerer tous les plannings" & vbLf & _
                                             "- Attribuer les visites automatiquement" & vbLf & _
                                             "- Envoyer des e-mails" & vbLf & _
                                             "- Calculer les salaires"
        With .Range("B" & ligneActuelle)
            .Font.Size = 10
            .Font.Color = RGB(50, 50, 50)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .Interior.Color = RGB(237, 244, 252)
            .Borders.LineStyle = xlContinuous
            .WrapText = True
        End With
        .Range("B" & ligneActuelle).RowHeight = 70

        ' Stocker la ligne du bouton Admin
        .Range("Z2").Value = ligneActuelle - 2 ' Ligne du titre "JE SUIS L'ADMINISTRATEUR"

        ' ===== PIED DE PAGE =====
        ligneActuelle = ligneActuelle + 4
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "[i] Cliquez sur le bloc qui correspond a votre profil"
        With .Range("B" & ligneActuelle)
            .Font.Size = 10
            .Font.Italic = True
            .Font.Color = RGB(150, 150, 150)
            .HorizontalAlignment = xlCenter
        End With

        ligneActuelle = ligneActuelle + 2
        .Range("B" & ligneActuelle & ":E" & ligneActuelle).Merge
        .Range("B" & ligneActuelle).Value = "Version 1.0 - Systeme de Planning Automatise - " & Format(Date, "dd/mm/yyyy")
        With .Range("B" & ligneActuelle)
            .Font.Size = 8
            .Font.Color = RGB(180, 180, 180)
            .HorizontalAlignment = xlCenter
        End With

        ' Masquer les quadrillages
        ActiveWindow.DisplayGridlines = False

        ' Proteger la feuille (navigation uniquement, clics autorises)
        .Protect Password:="protection", UserInterfaceOnly:=True, _
                 AllowFiltering:=True, AllowSorting:=True

        ' Activer la feuille
        .Activate
        .Range("B3").Select
    End With

    MsgBox "[OK] Feuille d'accueil creee avec succes !" & vbCrLf & vbCrLf & _
           "Les utilisateurs peuvent cliquer sur les blocs pour se connecter.", _
           vbInformation, "Accueil configure"
End Sub

' ============================================
' Gerer le clic sur la feuille d'accueil
' ============================================
' NOTE: Cette fonction doit etre appelee depuis l'evenement Worksheet_SelectionChange
' de la feuille Accueil
Public Sub GererClicAccueil(Target As Range, ws As Worksheet)
    Dim ligneGuide As Long
    Dim ligneAdmin As Long

    ' Recuperer les lignes des boutons
    ligneGuide = ws.Range("Z1").Value
    ligneAdmin = ws.Range("Z2").Value

    ' Verifier si le clic est sur le bloc Guide
    If Target.Row = ligneGuide And Target.Column >= 2 And Target.Column <= 5 Then
        ' Lancer la connexion Guide
        Call Module_Authentification.SeConnecter
        Exit Sub
    End If

    ' Verifier si le clic est sur le bloc Admin
    If Target.Row = ligneAdmin And Target.Column >= 2 And Target.Column <= 5 Then
        ' Lancer la connexion Admin
        Call Module_Authentification.SeConnecter
        Exit Sub
    End If
End Sub
