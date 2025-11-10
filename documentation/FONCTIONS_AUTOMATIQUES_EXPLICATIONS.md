# ⚠️ FONCTIONS AUTOMATIQUES - CE QUI SE PASSE VRAIMENT

## 🔴 RASSURE-TOI : RIEN N'EST AUTOMATIQUE PAR DÉFAUT !

Tout est **MANUEL** et nécessite que tu **cliques sur un bouton ou exécutes une macro**.

---

## 📋 ANALYSE FONCTION PAR FONCTION

### 1️⃣ **Attribution automatique des visites**

**Fonction :** `GenererPlanningAutomatique()` (Module_Planning.bas ligne 15)

**Déclenchement :**
```vb
❌ PAS automatique au démarrage
❌ PAS automatique quand on ajoute une visite
✅ MANUEL : L'admin doit exécuter la macro
```

**Ce qu'elle fait :**
```vb
Sub GenererPlanningAutomatique()
    ' Lit les visites NON assignées
    ' Pour chaque visite :
    '   - Cherche guides disponibles
    '   - Attribue au guide le moins chargé
    ' Met à jour la colonne Guide_Attribue
End Sub
```

**Comment l'utiliser :**
1. Admin va dans VBA (Alt+F11)
2. Exécute manuellement `GenererPlanningAutomatique`
3. OU crée un bouton sur une feuille qui appelle cette fonction

**DANGER : ❌ AUCUN** - Elle ne se lance jamais toute seule

---

### 2️⃣ **Calcul des salaires**

**Fonction :** `CalculerVisitesEtSalaires()` (Module_Calculs.bas ligne 15)

**Déclenchement :**
```vb
❌ PAS automatique à la fin du mois
❌ PAS automatique quand une visite est ajoutée
✅ MANUEL : L'admin doit exécuter la macro
```

**Ce qu'elle fait :**
```vb
Sub CalculerVisitesEtSalaires()
    ' Compte les visites par guide ce mois
    ' Applique la grille tarifaire dégressive
    ' Remplit la feuille Calculs_Paie
    ' NE paie PAS les guides automatiquement !
End Sub
```

**Comment l'utiliser :**
1. Fin du mois → Admin exécute `CalculerVisitesEtSalaires`
2. Vérifie les montants dans feuille Calculs_Paie
3. Fait les virements manuellement (pas automatique)

**DANGER : ❌ AUCUN** - C'est juste un calcul, pas un paiement

---

### 3️⃣ **Envoi d'emails**

**Fonctions :**
- `EnvoyerPlanningMensuel()` (Module_Emails.bas ligne 15)
- `EnvoyerNotificationsAutomatiques()` (Module_Emails.bas ligne 164)

**Déclenchement :**
```vb
❌ PAS automatique tous les jours
❌ PAS automatique quand une visite est assignée
✅ MANUEL : L'admin doit exécuter la macro
```

**Ce qu'elles font :**
```vb
Sub EnvoyerPlanningMensuel()
    ' Envoie à chaque guide SON planning du mois
    ' Via Outlook (doit être ouvert)
    ' L'admin doit cliquer pour lancer
End Sub

Sub EnvoyerNotificationsAutomatiques()
    ' Envoie rappels X jours avant les visites
    ' L'admin doit cliquer pour lancer
End Sub
```

**Comment les rendre automatiques (OPTIONNEL) :**
```vb
Sub ConfigurerTacheAutomatique()
    ' Guide pour créer une tâche Windows planifiée
    ' L'admin DOIT configurer manuellement dans Windows
    ' CE N'EST PAS FAIT AUTOMATIQUEMENT
End Sub
```

**DANGER : 🟡 MODÉRÉ**
- Si tu exécutes `EnvoyerPlanningMensuel`, tous les guides recevront un email IMMÉDIATEMENT
- Vérifie toujours avant d'exécuter
- Teste d'abord avec un seul guide

---

### 4️⃣ **Génération de contrats**

**Fonction :** `GenererContratsEnMasse()` (Module_Contrats.bas)

**Déclenchement :**
```vb
❌ PAS automatique
✅ MANUEL : L'admin doit exécuter la macro
```

**Ce qu'elle fait :**
```vb
Sub GenererContratsEnMasse(mois)
    ' Génère les contrats PDF pour tous les guides
    ' Sauvegarde dans dossier /Contrats/
End Sub
```

**DANGER : ❌ AUCUN** - Juste crée des fichiers PDF locaux

---

## 🎯 CE QUI EST VRAIMENT AUTOMATIQUE

### ✅ Ces événements se déclenchent automatiquement :

| Événement | Quand | Impact |
|-----------|-------|--------|
| **Workbook_Open()** | Ouverture fichier | Masque les feuilles ✅ Sécurité |
| **Workbook_BeforeClose()** | Fermeture fichier | Nettoie session ✅ Pas dangereux |
| **Worksheet_SelectionChange()** | Clic sur Accueil | Détecte clic Guide/Admin ✅ Nécessaire |
| **Workbook_SheetActivate()** | Change d'onglet | Ajuste visibilité selon rôle ✅ Sécurité |

**AUCUN de ces événements n'envoie d'email, ne calcule de salaire ou n'assigne de visite !**

---

## ❌ CE QUI N'EST **JAMAIS** AUTOMATIQUE

```
❌ Attribution des visites
   → MANUEL : L'admin exécute GenererPlanningAutomatique()

❌ Calcul des salaires
   → MANUEL : L'admin exécute CalculerVisitesEtSalaires()

❌ Envoi des emails
   → MANUEL : L'admin exécute EnvoyerPlanningMensuel()

❌ Génération des contrats
   → MANUEL : L'admin exécute GenererContratsEnMasse()

❌ Notifications
   → MANUEL : L'admin exécute EnvoyerNotificationsAutomatiques()
   OU configure une tâche Windows (nécessite action admin)
```

---

## 🔒 SÉCURITÉS EN PLACE

### 1. Aucun événement déclencheur
```vb
' Il n'y a PAS de code comme :
Private Sub Workbook_Open()
    Call GenererPlanningAutomatique()  ❌ ABSENT
    Call CalculerVisitesEtSalaires()   ❌ ABSENT
    Call EnvoyerPlanningMensuel()      ❌ ABSENT
End Sub
```

### 2. Toutes les fonctions "dangereuses" sont Public Sub
```vb
Public Sub GenererPlanningAutomatique()
' = Nécessite exécution MANUELLE depuis VBA ou bouton
```

### 3. Outlook doit être ouvert
```vb
Set outlookApp = CreateObject("Outlook.Application")
' Si Outlook fermé → ERREUR, rien ne s'envoie
```

---

## 🎮 COMMENT UTILISER LES FONCTIONS (GUIDE ADMIN)

### Workflow mensuel typique :

**1. Début du mois :**
```
Admin ouvre Excel
  ↓
Ajoute les visites dans feuille "Visites"
  ↓
Exécute manuellement : GenererPlanningAutomatique()
  → Attribue les visites aux guides
  ↓
Vérifie le planning (feuille "Planning")
  ↓
Ajuste manuellement si besoin (modifier colonne Guide_Attribue)
  ↓
Exécute manuellement : EnvoyerPlanningMensuel()
  → Envoie les plannings aux guides
```

**2. Pendant le mois :**
```
Rien d'automatique !
Guides se connectent pour voir leur planning
Admin peut modifier manuellement
```

**3. Fin du mois :**
```
Exécute manuellement : CalculerVisitesEtSalaires()
  → Calcule les montants
  ↓
Vérifie dans feuille "Calculs_Paie"
  ↓
Fait les virements MANUELLEMENT (pas automatique)
  ↓
Exécute manuellement : GenererContratsEnMasse()
  → Crée les PDF des contrats
```

---

## 🚨 COMMENT TESTER SANS DANGER

### Test 1 : Attribution (SANS RISQUE)
```vb
1. Ajoute 2-3 visites de test dans "Visites"
2. Ajoute 2 guides de test dans "Guides"
3. Exécute GenererPlanningAutomatique()
4. Vérifie dans "Planning" si l'attribution est correcte
5. Si OK, supprime les données de test
```

### Test 2 : Emails (AVEC PRÉCAUTION)
```vb
1. Dans feuille "Guides", mets TON email pour UN guide de test
2. Supprime les autres guides temporairement
3. Exécute EnvoyerPlanningMensuel()
4. Vérifie que TU reçois l'email
5. Si OK, remets les vrais guides
```

### Test 3 : Calculs (SANS RISQUE)
```vb
1. Exécute CalculerVisitesEtSalaires()
2. Vérifie les montants dans "Calculs_Paie"
3. Si erreur, corrige et ré-exécute
4. Aucun email n'est envoyé, aucun paiement n'est fait
```

---

## ✅ CONCLUSION

**TU AS RAISON D'AVOIR PEUR, MAIS RASSURE-TOI :**

### ❌ CE QUI POURRAIT ÊTRE DANGEREUX (mais ne l'est pas) :
- ❌ Envoi automatique d'emails → PAS AUTOMATIQUE, tu contrôles
- ❌ Attribution automatique → PAS AUTOMATIQUE, tu contrôles
- ❌ Calculs automatiques → PAS AUTOMATIQUE, tu contrôles

### ✅ CE QUI EST AUTOMATIQUE (et sans danger) :
- ✅ Masquage des feuilles au démarrage → Sécurité
- ✅ Détection des clics Guide/Admin → Interface
- ✅ Nettoyage à la fermeture → Maintenance

**TU GARDES LE CONTRÔLE TOTAL !** 🎮

**Aucune fonction critique ne se lance sans que tu cliques dessus.**
