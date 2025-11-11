Option Explicit

' ============================================
' EVENEMENTS DE LA FEUILLE VISITES
' A copier dans l'objet feuille "Visites"
' ============================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Erreur

    ' Ne rien faire si pas admin connecte
    If niveauAcces <> "ADMIN" Then Exit Sub

    ' Detecter ajout d'une nouvelle visite (colonne A = ID_Visite non vide)
    ' Declenchement uniquement si nouvelle ligne ajoutee (pas modification existante)
    If Not Intersect(Target, Me.Range("A:A")) Is Nothing Then
        If Target.Row > 1 And Target.Value <> "" Then
            ' Eviter boucle infinie
            Application.EnableEvents = False

            ' Lancer attribution automatique
            Call GenererPlanningAutomatique

            ' Reactiver evenements
            Application.EnableEvents = True

            ' Note: MsgBox supprime pour eviter interruptions lors de saisie multiple
        End If
    End If

    Exit Sub

Erreur:
    Application.EnableEvents = True  ' CRITIQUE : Reactiver meme si erreur
    MsgBox "ERREUR lors de l'attribution automatique :" & vbCrLf & _
           Err.Description, vbCritical, "Erreur"
End Sub

' ============================================
' Evenement : Activation de la feuille
' ============================================
Private Sub Worksheet_Activate()
    On Error Resume Next

    ' Message informatif pour l'admin
    If niveauAcces = "ADMIN" Then
        ' Supprimer ancien commentaire si existe (securise)
        On Error Resume Next
        Me.Range("A1").Comment.Delete
        On Error GoTo 0

        ' Ajouter nouveau commentaire
        Me.Range("A1").AddComment
        Me.Range("A1").Comment.Text "Attribution automatique activee" & Chr(10) & _
                                     "Ajoutez une visite (ID en colonne A), le guide sera assigne automatiquement !"
        Me.Range("A1").Comment.Visible = True  ' Visible pour etre utile
    End If

    On Error GoTo 0
End Sub
