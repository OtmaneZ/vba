# 🎯 INSTRUCTIONS RAPIDES - Configuration du système

## Problème actuel
- La feuille Accueil est recréée à chaque exécution de `CreerFeuilleAccueil`
- Le code événementiel disparaît à chaque fois
- Vous avez maintenant Feuil9 (Accueil)

## ✅ Solution simple (5 minutes)

### Étape 1 : Dans Excel VBA (Alt+F11)
Double-cliquez sur **Feuil9 (Accueil)** dans l'arborescence

### Étape 2 : Copiez ce code dans l'éditeur

```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next

    Dim ligneGuide As Long
    Dim ligneAdmin As Long

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
```

### Étape 3 : Testez !
1. Fermez VBA
2. Retournez dans Excel sur la feuille Accueil
3. **Cliquez sur le bloc vert [GUIDE]** → Ça devrait lancer la connexion !

## 🚀 Si ça marche

Vous verrez une fenêtre de connexion s'ouvrir. Par défaut :
- **Guide** : Choisissez un nom dans la liste
- **Admin** : Mot de passe = `admin123`

## ❌ Si ça ne marche toujours pas

Vérifiez que :
1. Les modules sont bien importés (dossier Modules dans VBA doit contenir 8 fichiers)
2. La cellule Z1 contient le numéro de ligne du bloc GUIDE (devrait être 8)
3. La cellule Z2 contient le numéro de ligne du bloc ADMIN (devrait être 14)
