Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
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
