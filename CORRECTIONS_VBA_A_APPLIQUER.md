# CORRECTIONS VBA À APPLIQUER - PLANNING GUIDES

## CONTEXTE

La cliente signale 3 problèmes majeurs :
1. La colonne **HEURE** affiche des nombres (0.4375) au lieu de "10:30"
2. La colonne **GUIDES_DISPONIBLES** reste vide
3. La feuille **SPÉCIALISATIONS** disparaît ou ne fonctionne pas

## STRUCTURE DES DONNÉES EXCEL

### Feuille DISPONIBILITES
```
Colonne 1 : Date
Colonne 2 : Disponibilite (OUI/NON)
Colonne 3 : Commentaire
Colonne 4 : Prenom_Guide
Colonne 5 : Nom_Guide
```

### Feuille VISITES
```
Colonne 1 : ID_Visite
Colonne 2 : Date
Colonne 3 : Heure
Colonne 4 : Duree
Colonne 5 : Type_Visite
Colonne 6 : Musee
Colonne 7 : Nb_Personnes
Colonne 8 : Niveau
Colonne 9 : Theme
Colonne 10 : Commentaire
```

### Feuille SPECIALISATIONS
```
Ligne 1 : En-têtes
À partir de ligne 2 : Données
Colonne 1 : Prenom
Colonne 2 : Nom
Colonne 3 : Type_Visite
Colonne 4 : Autorise (OUI/NON)
```

---

## 🔴 CORRECTION 1 : Format des heures (Module_Planning.bas)

### Problème
Le code concatène l'heure comme un nombre, donnant 0.4375 au lieu de "10:30"

### Localisation
**Fichier :** Module_Planning.bas  
**Ligne :** ~72 (dans la fonction GenererPlanningAutomatique)

### Code actuel (INCORRECT)
```vba
heureVisite = wsVisites.Cells(i, 3).Value & " - " & wsVisites.Cells(i, 4).Value
```

### Code corrigé
```vba
heureVisite = Format(wsVisites.Cells(i, 3).Value, "hh:mm")
```

---

## 🔴 CORRECTION 2 : Lecture des disponibilités

### Problème
La fonction ObtenirGuidesDisponibles lit les mauvaises colonnes

**Fichier :** Module_Planning.bas  
**Fonction :** ObtenirGuidesDisponibles (lignes ~200-235)

### Code actuel lit :
- Date en colonne 2 (devrait être colonne 1)
- Disponibilité en colonne 3 (devrait être colonne 2)
- Guide en colonne 1 (devrait être Prénom col 4 + Nom col 5)

### Code corrigé à copier-coller :
```vba
Private Function ObtenirGuidesDisponibles(dateVisite As Date) As Collection
    Dim wsDispo As Worksheet
    Dim col As New Collection
    Dim i As Long
    Dim guideID As String
    Dim prenomGuide As String
    Dim nomGuide As String

    Set wsDispo = ThisWorkbook.Worksheets(FEUILLE_DISPONIBILITES)

    For i = 2 To wsDispo.Cells(wsDispo.Rows.Count, 1).End(xlUp).Row
        On Error Resume Next
        If CDate(wsDispo.Cells(i, 1).Value) = dateVisite Then
            If UCase(wsDispo.Cells(i, 2).Value) = "OUI" Then
                prenomGuide = Trim(wsDispo.Cells(i, 4).Value)
                nomGuide = Trim(wsDispo.Cells(i, 5).Value)
                guideID = prenomGuide & " " & nomGuide

                Dim existe As Boolean
                existe = False
                Dim j As Integer
                For j = 1 To col.Count
                    If col(j) = guideID Then
                        existe = True
                        Exit For
                    End If
                Next j

                If Not existe Then
                    col.Add guideID
                End If
            End If
        End If
        On Error GoTo 0
    Next i

    Set ObtenirGuidesDisponibles = col
End Function
```

---

## 🔴 CORRECTION 3 : Colonnes Visites

**Fichier :** Module_Planning.bas (lignes ~73-75)

### Changer :
```vba
musee = wsVisites.Cells(i, 7).Value
typeVisite = wsVisites.Cells(i, 6).Value
```

### Par :
```vba
duree = wsVisites.Cells(i, 4).Value
typeVisite = wsVisites.Cells(i, 5).Value
musee = wsVisites.Cells(i, 6).Value
```

---

## 🔴 CORRECTION 4 : Spécialisations

**Fichier :** Module_Specialisations.bas  
**Fonction :** GuideAutoriseVisite

### Problème
Lit ligne 4, colonnes 1-2 avec logique "Tous sauf"/"UNIQUEMENT"  
Devrait lire ligne 2, colonnes 1-4 avec simple OUI/NON

### Code corrigé complet :
```vba
Function GuideAutoriseVisite(nomGuide As String, typeVisite As String) As Boolean
    On Error Resume Next

    Dim ws As Worksheet
    Dim derLigne As Long
    Dim i As Long
    Dim guideSpec As String
    Dim visiteSpec As String
    Dim trouve As Boolean
    Dim prenomGuide As String
    Dim nomGuideSpec As String
    Dim autorise As String

    GuideAutoriseVisite = True
    trouve = False

    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("Specialisations")

    If ws Is Nothing Then
        Exit Function
    End If

    derLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To derLigne
        prenomGuide = Trim(ws.Cells(i, 1).Value)
        nomGuideSpec = Trim(ws.Cells(i, 2).Value)
        guideSpec = prenomGuide & " " & nomGuideSpec
        visiteSpec = Trim(ws.Cells(i, 3).Value)
        autorise = UCase(Trim(ws.Cells(i, 4).Value))

        If (InStr(1, guideSpec, nomGuide, vbTextCompare) > 0 Or _
            InStr(1, nomGuide, guideSpec, vbTextCompare) > 0) And _
           (InStr(1, visiteSpec, typeVisite, vbTextCompare) > 0 Or _
            InStr(1, typeVisite, visiteSpec, vbTextCompare) > 0) Then

            trouve = True

            If autorise = "OUI" Then
                GuideAutoriseVisite = True
            Else
                GuideAutoriseVisite = False
            End If

            Exit Function
        End If
    Next i

    If trouve Then
        GuideAutoriseVisite = False
    End If

    On Error GoTo 0
End Function
```

---

## PROCÉDURE

1. Ouvrir PLANNING.xlsm
2. Alt+F11 (éditeur VBA)
3. Module_Planning → appliquer corrections 1, 2, 3
4. Module_Specialisations → appliquer correction 4
5. Ctrl+S pour sauvegarder
6. Fermer VBA et Excel
7. Rouvrir et tester

---

## RÉSULTATS ATTENDUS

✅ Heure affiche "10:30" au lieu de 0.4375
✅ Guides_Disponibles remplie avec noms des guides
✅ Spécialisations respectées (OUI/NON)


---

## �� PROBLÈME BONUS : Bouton "Générer Planning" manquant

### Situation
Le code VBA contient la fonction `GenererPlanningAutomatique` mais il n'y a **pas de bouton** dans l'interface Excel pour l'exécuter.

### Solutions possibles

#### SOLUTION 1 : Exécuter via le menu Macros (RAPIDE)
1. Dans Excel, allez dans **Outils > Macros** (ou Alt+F8)
2. Sélectionnez **GenererPlanningAutomatique**
3. Cliquez sur **Exécuter**

#### SOLUTION 2 : Créer un bouton (RECOMMANDÉ pour utilisation régulière)
1. Allez dans l'onglet **Développeur** (si invisible : Fichier > Options > Personnaliser le ruban > Cocher "Développeur")
2. Cliquez sur **Insérer** > **Bouton (Contrôle de formulaire)**
3. Dessinez le bouton dans la feuille Planning (en haut à droite)
4. Dans la fenêtre qui s'ouvre, sélectionnez : **Module_Planning.GenererPlanningAutomatique**
5. Cliquez **OK**
6. Clic droit sur le bouton > **Modifier le texte** > Écrivez "Générer Planning"
7. Sauvegardez (Cmd+S)

#### SOLUTION 3 : Raccourci clavier (PLUS RAPIDE)
1. Outils > Macros
2. Sélectionnez **GenererPlanningAutomatique**
3. Cliquez sur **Options**
4. Assignez un raccourci (par exemple : Ctrl+Shift+G)
5. OK

### Important
Après avoir appliqué les corrections VBA, vous DEVEZ utiliser une de ces méthodes pour lancer la génération du planning. Le système ne se lance pas automatiquement.

