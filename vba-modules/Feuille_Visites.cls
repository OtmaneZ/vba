Option Explicit

' ============================================
' EVENEMENTS DE LA FEUILLE VISITES
' A copier dans l'objet feuille "Visites"
' ============================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next

    ' Ne rien faire si pas admin connecte
    If niveauAcces <> "ADMIN" Then Exit Sub

    ' Detecter ajout dans la colonne A (ID_Visite) ou B (Date)
    If Not Intersect(Target, Me.Range("A:B")) Is Nothing Then
        ' Eviter boucle infinie
        Application.EnableEvents = False

        ' Lancer attribution automatique
        Call GenererPlanningAutomatique

        ' Reactiver evenements
        Application.EnableEvents = True

        MsgBox "[OK] Planning mis a jour automatiquement !" & vbCrLf & _
               "Les visites ont ete attribuees aux guides disponibles.", _
               vbInformation, "Attribution automatique"
    End If

    On Error GoTo 0
End Sub

' ============================================
' Evenement : Activation de la feuille
' ============================================
Private Sub Worksheet_Activate()
    On Error Resume Next

    ' Message informatif pour l'admin
    If niveauAcces = "ADMIN" Then
        Me.Range("A1").AddComment
        Me.Range("A1").Comment.Text "Attribution automatique activee" & Chr(10) & _
                                     "Ajoutez une visite, le guide sera assigne automatiquement !"
        Me.Range("A1").Comment.Visible = False
    End If

    On Error GoTo 0
End Sub
