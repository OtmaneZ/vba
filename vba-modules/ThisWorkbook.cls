Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next

    ' Masquer TOUTES les feuilles sauf Accueil par defaut
    Call MasquerToutesFeuillesParDefaut

    ' Activer la page d'accueil
    ThisWorkbook.Sheets("Accueil").Activate

    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If niveauAcces <> "" Then
        utilisateurConnecte = ""
        niveauAcces = ""
        emailUtilisateur = ""
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If niveauAcces = "GUIDE" Then
        ThisWorkbook.Sheets("Calculs_Paie").Visible = xlSheetVeryHidden
        ThisWorkbook.Sheets("Configuration").Visible = xlSheetVeryHidden
    ElseIf niveauAcces = "ADMIN" Then
        ThisWorkbook.Sheets("Calculs_Paie").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Configuration").Visible = xlSheetVisible
    End If
    On Error GoTo 0
End Sub

' Masquer toutes les feuilles sauf Accueil au demarrage
Private Sub MasquerToutesFeuillesParDefaut()
    On Error Resume Next

    Dim ws As Worksheet

    ' Boucle sur TOUTES les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        ' Masquer toutes sauf Accueil
        If ws.Name <> "Accueil" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' S'assurer que Accueil est visible
    ThisWorkbook.Sheets("Accueil").Visible = xlSheetVisible

    On Error GoTo 0
End Sub
