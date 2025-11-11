VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuille_Mon_Planning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ============================================
' EVENEMENTS DE LA FEUILLE Mon_Planning
' Gestion des clics pour confirmer/refuser visites
' ============================================

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next

    ' Verifier que c'est un clic sur la colonne Action (H)
    If Target.Column = 8 And Target.Row > 1 Then
        ' Verifier le statut actuel
        Dim statutActuel As String
        statutActuel = Target.Offset(0, -1).Value ' Colonne G (Statut)

        ' Ne permettre action que si en attente
        If statutActuel = "En attente" Or statutActuel = "" Then
            ' Appeler la fonction de confirmation/refus
            Call Module_Authentification.ConfirmerOuRefuserVisite
        ElseIf statutActuel = "Confirme" Then
            MsgBox "[i] Cette visite est deja confirmee.", vbInformation, "Deja confirme"
        ElseIf statutActuel = "Refuse" Then
            MsgBox "[i] Cette visite a deja ete refusee.", vbInformation, "Deja refuse"
        End If
    End If

    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next

    ' Message d'aide au survol
    If niveauAcces = "GUIDE" Then
        Me.Range("H1").AddComment
        Me.Range("H1").Comment.Text "Cliquez sur les cellules de cette colonne pour confirmer ou refuser vos visites"
        Me.Range("H1").Comment.Visible = False
    End If

    On Error GoTo 0
End Sub
