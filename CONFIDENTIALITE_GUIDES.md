# 🔐 CONFIDENTIALITÉ ET FILTRAGE PAR GUIDE

## 🎯 Objectif
Chaque guide ne voit **QUE ses propres données** - pas celles des collègues.

---

## 📋 Ce qu'un guide voit actuellement VS ce qu'il devrait voir

| Feuille | État ACTUEL | État IDÉAL |
|---------|-------------|------------|
| **Guides** | ❌ Voit tous les guides | ⚠️ Voit juste les noms (pas emails/tél privés) |
| **Disponibilités** | ❌ Voit toutes les dispo | ✅ Voit UNIQUEMENT ses dispos |
| **Visites** | ❌ Voit toutes les visites | ✅ Voit UNIQUEMENT ses visites assignées |
| **Planning** | ✅ Filtré (fonction existe) | ✅ Voit UNIQUEMENT son planning |
| **Calculs_Paie** | ✅ Masqué | ✅ Masqué |
| **Contrats** | ✅ Masqué | ✅ Masqué |
| **Configuration** | ✅ Masqué | ✅ Masqué |

---

## ✅ Code déjà en place

### 1. Filtrage du Planning ✓
```vb
' Dans Module_Authentification.bas ligne 108
Sub AfficherPlanningGuide(nomGuide As String)
    ' Crée une feuille "Mon_Planning" avec UNIQUEMENT les visites du guide
    ' Filtre automatique : InStr(nomGuide) dans la colonne Guide_Attribue
```

**✅ DÉJÀ FONCTIONNEL** - Chaque guide voit son planning perso

---

## 🛠️ Améliorations à ajouter

### Option 1 : Filtrage automatique des feuilles (SIMPLE)
Quand un guide se connecte, on applique des filtres Excel automatiques.

**Avantages :**
- ✅ Simple à implémenter
- ✅ Le guide peut enlever le filtre (mais ne devrait pas)
- ✅ Toutes les données restent dans la même feuille

**Inconvénients :**
- ⚠️ Pas 100% sécurisé (filtre enlevable)
- ⚠️ Le guide "technique" peut voir les autres lignes

### Option 2 : Feuilles temporaires filtrées (RECOMMANDÉ)
Créer des feuilles temporaires comme "Mes_Visites", "Mes_Disponibilites".

**Avantages :**
- ✅ 100% sécurisé - impossible de voir les autres données
- ✅ Déjà utilisé pour "Mon_Planning" (fonction existe)
- ✅ Données originales protégées

**Inconvénients :**
- ⚠️ Nécessite synchronisation lors de la modification

### Option 3 : Protection par mot de passe des feuilles (SÉCURITÉ MAX)
Protéger les feuilles et déverrouiller temporairement selon l'utilisateur.

**Avantages :**
- ✅ Sécurité maximale
- ✅ Impossible de modifier les données des autres

**Inconvénients :**
- ⚠️ Plus complexe
- ⚠️ Nécessite gestion des mots de passe par feuille

---

## 💡 Solution recommandée

### Créer 3 nouvelles fonctions (comme AfficherPlanningGuide)

```vb
' 1. Afficher uniquement SES visites
Sub AfficherMesVisites(nomGuide As String)
    ' Créer feuille "Mes_Visites"
    ' Copier uniquement les lignes où Guide_Attribue = nomGuide
End Sub

' 2. Afficher uniquement SES disponibilités
Sub AfficherMesDisponibilites(nomGuide As String)
    ' Créer feuille "Mes_Disponibilites"
    ' Copier uniquement les lignes où ID_Guide = nomGuide
End Sub

' 3. Masquer les infos sensibles des autres guides
Sub AfficherListeGuidesLimitee()
    ' Créer feuille "Annuaire"
    ' Afficher uniquement : Prenom, Nom (pas email, pas téléphone)
End Sub
```

### Appeler ces fonctions à la connexion

```vb
' Modifier SeConnecter() ligne ~85
If estGuide Then
    niveauAcces = "GUIDE"
    utilisateurConnecte = nomGuide

    ' Créer les vues filtrées
    Call AfficherMesVisites(nomGuide)
    Call AfficherMesDisponibilites(nomGuide)
    Call AfficherPlanningGuide(nomGuide)  ' Déjà fait !
    Call AfficherListeGuidesLimitee()

    ' Masquer les feuilles originales
    ThisWorkbook.Sheets(FEUILLE_VISITES).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_GUIDES).Visible = xlSheetVeryHidden

    ' Afficher uniquement ses feuilles perso
    ThisWorkbook.Sheets("Mes_Visites").Activate
End If
```

---

## 🔒 Résultat final

### Guide connecté (ex: "Marie")
**Onglets visibles :**
- ✅ **Accueil** : Page d'accueil
- ✅ **Mes_Visites** : Uniquement ses visites
- ✅ **Mes_Disponibilites** : Uniquement ses dispos
- ✅ **Mon_Planning** : Son planning personnel
- ✅ **Annuaire** : Noms des collègues (sans infos privées)

**Onglets masqués :**
- ❌ Visites (données complètes)
- ❌ Disponibilites (données complètes)
- ❌ Guides (infos complètes)
- ❌ Planning (données complètes)
- ❌ Calculs_Paie
- ❌ Contrats
- ❌ Configuration

### Admin connecté
**Onglets visibles :**
- ✅ **TOUS** les onglets (accès complet)

---

## 📊 Exemple concret

### Données actuelles (feuille Visites)
| Date | Heure | Type | Guide_Attribue |
|------|-------|------|----------------|
| 10/11/2025 | 10h00 | Classique | Marie |
| 10/11/2025 | 14h00 | Premium | Jean |
| 11/11/2025 | 09h00 | Classique | Marie |
| 11/11/2025 | 15h00 | VIP | Sophie |

### Ce que voit Marie (feuille Mes_Visites)
| Date | Heure | Type | Guide_Attribue |
|------|-------|------|----------------|
| 10/11/2025 | 10h00 | Classique | Marie |
| 11/11/2025 | 09h00 | Classique | Marie |

**Marie ne voit PAS les visites de Jean et Sophie !** ✅

---

## ⚡ Tu veux que je code ces fonctions ?

Je peux ajouter :
1. ✅ `AfficherMesVisites()`
2. ✅ `AfficherMesDisponibilites()`
3. ✅ `AfficherListeGuidesLimitee()`
4. ✅ Modifier `SeConnecter()` pour appeler ces fonctions
5. ✅ Masquer les feuilles originales pour les guides

**Dis-moi si tu veux que je l'implémente !** 🚀
