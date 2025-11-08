Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next

    Dim ligneGuide As Long, ligneAdmin As Long

    ligneGuide = Me.Range("Z1").Value
    ligneAdmin = Me.Range("Z2").Value

    ' Clic sur le bloc GUIDE
    If Target.Row >= ligneGuide And Target.Row <= ligneGuide + 2 Then
        If Target.Column >= 2 And Target.Column <= 5 Then
            Call SeConnecter
        End If
    End If

    ' Clic sur le bloc ADMIN
    If ligneAdmin > 0 Then
        If Target.Row >= ligneAdmin And Target.Row <= ligneAdmin + 3 Then
            If Target.Column >= 2 And Target.Column <= 5 Then
                Call SeConnecter
            End If
        End If
    End If

    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next

    If utilisateurConnecte <> "" Then
        Me.Range("B25").Value = ">>> Connecte en tant que : " & utilisateurConnecte & " (" & niveauAcces & ")"
        Me.Range("B25").Font.Bold = True
        Me.Range("B25").Font.Color = RGB(0, 128, 0)
    Else
        Me.Range("B25").Value = ""
    End If

    On Error GoTo 0
End Sub
