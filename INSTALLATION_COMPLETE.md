# 🚀 INSTALLATION COMPLÈTE - Depuis Zéro

## ✅ Prérequis
- Fichier Excel vierge ou `PLANNING_MUSEE_TEST.xlsm`
- Dossier `vba-modules/` avec les 10 fichiers VBA

---

## 📋 ÉTAPE 1 : Importer les modules (.bas)

1. **Ouvrir VBA** : `Alt + F11` (ou Outils → Macros → Visual Basic)
2. **Fichier** → **Importer un fichier...**
3. Sélectionner **TOUS les 8 fichiers .bas** :
   - Module_Config.bas
   - Module_Authentification.bas
   - Module_Disponibilites.bas
   - Module_Planning.bas
   - Module_Emails.bas
   - Module_Calculs.bas
   - Module_Contrats.bas
   - Module_Accueil.bas
4. Cliquer **Ouvrir**

---

## 📋 ÉTAPE 2 : Copier le code de ThisWorkbook

1. Dans VBA, **double-cliquer** sur **ThisWorkbook** (dans "Microsoft Excel Objets")
2. **Copier-coller** ce code :

```vb
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
```

---

## 📋 ÉTAPE 3 : Créer la page d'accueil

1. Dans VBA, aller dans **Outils** → **Macros** (ou `Alt + F8`)
2. Sélectionner la macro **`CreerFeuilleAccueil`**
3. Cliquer **Exécuter**
4. Message "Feuille creee avec succes !" → OK

---

## 📋 ÉTAPE 4 : Copier le code de la feuille Accueil

1. Dans VBA, **double-cliquer** sur la feuille **Accueil** (dans "Microsoft Excel Objets")
2. **Copier-coller** ce code :

```vb
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
```

---

## 📋 ÉTAPE 5 : Initialiser le système

1. Dans VBA, aller dans **Outils** → **Macros** (ou `Alt + F8`)
2. Sélectionner la macro **`InitialiserApplication`**
3. Cliquer **Exécuter**
4. Vérifie que les feuilles ont été créées : Guides, Disponibilites, Visites, Planning, Calculs_Paie, Contrats, Configuration

---

## ✅ ÉTAPE 6 : Tester !

1. **Fermer VBA** et retourner dans Excel
2. **Sauvegarder** le fichier : `Cmd + S`
3. **Fermer Excel** complètement
4. **Ré-ouvrir** le fichier `PLANNING_MUSEE_TEST.xlsm`
5. **Activer les macros** si demandé
6. **Cliquer** sur le bloc vert **[GUIDE] JE SUIS UN GUIDE**
7. Fenêtre de connexion devrait s'ouvrir ! 🎉

---

## 🔐 Connexion par défaut

**Admin :**
- Mot de passe : `admin123`

**Guide :**
- Pas de mot de passe
- Ajouter des guides dans la feuille "Guides" avec colonnes :
  - Nom
  - Email
  - Telephone

---

## ⚠️ Important

- **NE PLUS exécuter** `CreerFeuilleAccueil` après l'étape 4
- Si tu veux recommencer : supprimer manuellement la feuille Accueil avant de la recréer
- Les fichiers .cls ne s'importent PAS, on copie juste leur contenu
