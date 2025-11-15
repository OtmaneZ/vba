# Corrections - Visibilité Spécialisations

## Problème identifié
La feuille "Spécialisations" était masquée (`veryHidden`) et n'apparaissait pas pour l'admin.

## Cause
Dans la fonction `SeDeconnecter()`, **toutes** les feuilles sauf "Accueil" étaient masquées, y compris "Spécialisations".

## Solutions appliquées

### 1. Module_Authentification.bas - Fonction SeDeconnecter()
**Modification ligne ~548-554 :**
```vba
' AVANT : Masquait TOUTES les feuilles sauf Accueil
For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> "Accueil" Then
        ws.Visible = xlSheetVeryHidden
    End If
Next ws

' APRÈS : Exclut aussi Spécialisations du masquage
For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> "Accueil" And ws.Name <> "Spécialisations" Then
        ws.Visible = xlSheetVeryHidden
    End If
Next ws
```

### 2. Module_Authentification.bas - Fonction MasquerFeuillesOriginalesPourGuide()
**Ajout après la ligne de FEUILLE_CONFIG :**
```vba
' Masquer Spécialisations pour les guides (visible uniquement pour admin)
Set ws = Nothing
Set ws = ThisWorkbook.Sheets("Spécialisations")
If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
```

### 3. Fichier Excel PLANNING.xlsm
**Modification avec openpyxl :**
- Feuille "Spécialisations" : `veryHidden` → `visible`

## Comportement final

### Pour l'ADMIN :
1. Connexion → "Spécialisations" reste **VISIBLE** ✅
2. Déconnexion → "Spécialisations" reste **VISIBLE** ✅
3. Prochaine connexion → "Spécialisations" toujours **VISIBLE** ✅

### Pour les GUIDES :
1. Connexion → "Spécialisations" est **MASQUÉE** automatiquement ✅
2. Pas d'accès à cette feuille (sécurité) ✅

## Actions requises
1. ✅ Fichier Excel modifié (Spécialisations visible)
2. ⚠️ **Importer le Module_Authentification.bas corrigé dans Excel**
   - Ouvrir PLANNING.xlsm
   - Alt+F11 (VBA Editor)
   - Supprimer l'ancien Module_Authentification
   - File → Import → Module_Authentification.bas
   - Sauvegarder (Cmd+S)

## Test recommandé
1. Se connecter en tant qu'ADMIN → Vérifier que "Spécialisations" est visible
2. Se déconnecter → Vérifier que "Spécialisations" reste visible
3. Se connecter en tant que GUIDE → Vérifier que "Spécialisations" est masquée

---
**Date :** 14 novembre 2025
**Fichiers modifiés :**
- vba-modules/Module_Authentification.bas
- PLANNING.xlsm (feuille Spécialisations)
