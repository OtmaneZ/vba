' ============================================
' EVENEMENTS DE LA FEUILLE ACCUEIL
' A copier dans l'objet feuille "Accueil" (pas un module)
' ATTENTION : NE PAS copier la ligne "Attribute VB_Name"
' ============================================

Option Explicit

' ============================================
' Evenement : Clic sur la feuille
' ============================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Detection du clic sur les blocs d'authentification

    On Error Resume Next

    Dim ligneGuide As Long
    Dim ligneAdmin As Long

    ' Recuperer les lignes des boutons (stockees en Z1 et Z2)
    ligneGuide = Me.Range("Z1").Value
    ligneAdmin = Me.Range("Z2").Value

    ' Si clic sur le bloc GUIDE (titre ou description)
    If Target.Row >= ligneGuide And Target.Row <= ligneGuide + 2 Then
        If Target.Column >= 2 And Target.Column <= 5 Then
            ' Lancer l'authentification guide
            Call SeConnecter
        End If
    End If

    ' Si clic sur le bloc ADMIN
    If ligneAdmin > 0 Then
        If Target.Row >= ligneAdmin And Target.Row <= ligneAdmin + 3 Then
            If Target.Column >= 2 And Target.Column <= 5 Then
                ' Lancer l'authentification admin
                Call SeConnecter
            End If
        End If
    End If

    On Error GoTo 0
End Sub

' ============================================
' Evenement : Double-clic sur la feuille
' ============================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' Alternative : utiliser le double-clic au lieu du simple clic

    On Error Resume Next

    Dim ligneGuide As Long
    Dim ligneAdmin As Long

    ligneGuide = Me.Range("Z1").Value
    ligneAdmin = Me.Range("Z2").Value

    ' Si double-clic sur le bloc GUIDE
    If Target.Row >= ligneGuide And Target.Row <= ligneGuide + 2 Then
        If Target.Column >= 2 And Target.Column <= 5 Then
            Cancel = True ' Empecher l'edition de la cellule
            Call SeConnecter
        End If
    End If

    ' Si double-clic sur le bloc ADMIN
    If ligneAdmin > 0 Then
        If Target.Row >= ligneAdmin And Target.Row <= ligneAdmin + 3 Then
            If Target.Column >= 2 And Target.Column <= 5 Then
                Cancel = True
                Call SeConnecter
            End If
        End If
    End If

    On Error GoTo 0
End Sub

' ============================================
' Evenement : Activation de la feuille
' ============================================
Private Sub Worksheet_Activate()
    ' Optionnel : rafraichir la feuille d'accueil a chaque activation
    On Error Resume Next

    ' Afficher un message si l'utilisateur est deja connecte
    If utilisateurConnecte <> "" Then
        Me.Range("B25").Value = ">>> Connecte en tant que : " & utilisateurConnecte & " (" & niveauAcces & ")"
        Me.Range("B25").Font.Bold = True
        Me.Range("B25").Font.Color = RGB(0, 128, 0)
    Else
        Me.Range("B25").Value = ""
    End If

    On Error GoTo 0
End Sub

' ============================================
' Evenement : Desactivation de la feuille
' ============================================
Private Sub Worksheet_Deactivate()
    ' Rien de special pour le moment
End Sub
