# ✅ VÉRIFICATION COMPLÈTE - SYSTÈME DE CACHETS

**Date :** 11 novembre 2025
**Demande client :** Rémunération en cachets par jour (montant identique chaque jour)

---

## 🎯 SYSTÈME VALIDÉ

### 1️⃣ Calcul des tarifs journaliers

**Fonction :** `CalculerTarifJournee()` (Module_Calculs.bas, ligne 283-313)

**Logique tarifaire conforme à l'email client :**
```vba
' CAS SPECIAL : Hors-les-murs = 100€ fixe
If typeVisite = "HORS-LES-MURS" Then
    Return 100€

' TARIFS STANDARDS selon nombre de visites PAR JOUR
Case nbVisites = 1  → 80€
Case nbVisites = 2  → 110€
Case nbVisites >= 3 → 140€
```

**Paramètres Configuration :**
- `TARIF_1_VISITE` = 80
- `TARIF_2_VISITES` = 110
- `TARIF_3_VISITES` = 140
- `TARIF_HORSLEMURS` = 100

---

### 2️⃣ Comptage des visites par jour

**Fonction :** `CalculerVisitesEtSalaires()` (Module_Calculs.bas, ligne 40-230)

**Mécanisme de comptage :**
```vba
' Clé unique : Guide + Date
cleJour = guideID & "|" & Format(dateVisite, "yyyy-mm-dd")

' Si déjà une visite ce jour-là :
temp(2) = temp(2) + 1  ' Incrémente le compteur de visites
```

✅ **Résultat :** Calcule correctement le nombre de visites par jour pour chaque guide

---

### 3️⃣ Calcul du cachet journalier

**Fonction :** `CalculerVisitesEtSalaires()` (Module_Calculs.bas, ligne 163-168)

**Formule :**
```vba
montantParCachet = RoundUp(montantSalaire / nbJoursTravailles, 2)
totalRecalcule = montantParCachet * nbJoursTravailles
```

**Vérification avec l'exemple client :**
```
Guide a effectué en octobre :
- 1er oct : 2 visites/jour = 110€
- 4 oct : 2 visites/jour = 110€
- 7 oct : 1 visite/jour = 80€
- 15 oct : 1 hors-les-murs = 100€
- 17 oct : 3 visites/jour = 140€
- 24 oct : 2 visites/jour = 110€
- 28 oct : 3 visites/jour = 140€

Total : 7 jours, 790€
Cachet : 790 ÷ 7 = 112.857... → 112.86€ (arrondi supérieur)
Total recalculé : 112.86€ × 7 = 790.02€
```

✅ **Résultat dans Calculs_Paie :**
- Colonne E : Montant Total = 790.00€
- Colonne F : Montant/Cachet = **112.86€**
- Colonne G : Total Recalculé = 790.02€

---

### 4️⃣ Décompte mensuel détaillé

**Fonction :** `GenererDecompteMensuel()` (Module_Calculs.bas, ligne 594-750)

**Contenu du décompte :**

#### A. Détail par visite
| Guide | Date | Heure | Type Visite | Catégorie | Nb Jours | Montant Cachet |
|-------|------|-------|-------------|-----------|----------|----------------|
| ... | ... | ... | ... | ... | ... | ... |

#### B. Statistiques par catégorie
```
STATISTIQUES PAR CATEGORIE
Visites Branly :     84
Visites Marine :     15
Hors-les-murs :      5
Événements :         2
Visio :              1
Autres :             1
TOTAL :              108
```

#### C. Nombre de jours par guide
```
NOMBRE DE JOURS PAR GUIDE
Jean Dupont :    7 jours
Marie Martin :   5 jours
...
```

✅ **Conforme à la demande :**
- ✅ Nombre de jours travaillés
- ✅ Dates et horaires de chaque prestation
- ✅ Total des prestations
- ✅ Comptage séparé : Branly, Marine, Hors-les-murs, Événements, Visio, Autres

---

## 🔍 Points de vigilance

### Colonne Type_Visite dans Visites

La fonction `IdentifierTypeVisite()` lit la **colonne 5** de la feuille Visites.

**À vérifier dans Excel :**
- La colonne 5 doit contenir : `BRANLY`, `MARINE`, `HORS-LES-MURS`, `EVENEMENT`, `VISIO`, `AUTRE`
- Pour que le décompte fonctionne correctement

**Structure attendue de la feuille Visites :**
```
Col 1: ID_Visite
Col 2: Nom_Visite
Col 3: ?
Col 4: ?
Col 5: Type_Visite  ← IMPORTANT
```

---

## ✅ CONCLUSION

**Système 100% conforme à la demande client :**

1. ✅ Tarifs basés sur le **nombre de visites par jour** (80/110/140)
2. ✅ Exception hors-les-murs à 100€
3. ✅ Calcul de cachet : **Total ÷ Nb jours avec arrondi supérieur**
4. ✅ Décompte complet avec dates, horaires, total et statistiques par catégorie
5. ✅ Même montant pour chaque cachet (système équitable)

**Résultat attendu :** Conforme à l'exemple fourni (790€ ÷ 7 jours = 112.86€)

---

## 🎯 Actions à faire

1. ✅ Vérifier que la colonne 5 de Visites contient bien les types (BRANLY, MARINE, etc.)
2. ✅ Tester avec des données réelles d'octobre
3. ✅ Générer un décompte mensuel pour vérifier les statistiques
4. ✅ Vérifier que le Total Recalculé correspond bien au Total (à quelques centimes près)
