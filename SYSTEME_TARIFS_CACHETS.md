# 📊 SYSTÈME TARIFAIRE & CACHETS - EXPLICATION COMPLÈTE

## 🎯 **Les 2 systèmes expliqués**

### 1️⃣ **TARIFS** (Facturation au musée)
Ce sont les **prix facturés au musée/client** pour chaque type de visite.

| Type de visite | Tarif | Durée standard |
|----------------|-------|----------------|
| Branly | 50€ | 2h |
| Marine | 50€ | 1.5h |
| Hors-les-murs | 55€ | 2h |
| Événements | 60€ | - |
| Visio | 45€ | 1h |
| Autres | 50€ | - |

**Exemple :**
- Guide fait 2 visites Branly dans la journée → 2 × 50€ = **100€**
- Guide fait 1 visite Marine + 1 visio → 50€ + 45€ = **95€**

---

### 2️⃣ **CACHETS** (Rémunération du guide)
Ce sont les **montants payés au guide** en fin de mois.

#### Calcul :
```
TOTAL MENSUEL ÷ NOMBRE DE JOURS TRAVAILLÉS = CACHET JOURNALIER
```

#### Exemple concret :
Un guide travaille 7 jours dans le mois :
- Jour 1 : 2 visites Branly = 100€
- Jour 2 : 1 visite Marine = 50€
- Jour 3 : 3 visites Branly = 150€
- Jour 4 : 1 visio + 1 hors-les-murs = 45€ + 55€ = 100€
- Jour 5 : 2 visites Marine = 100€
- Jour 6 : 1 événement = 60€
- Jour 7 : 2 visites autres = 100€

**TOTAL MENSUEL** = 100 + 50 + 150 + 100 + 100 + 60 + 100 = **660€**

**CACHET JOURNALIER** = 660€ ÷ 7 jours = **94,29€** (arrondi supérieur)

**CONTRAT FIN DE MOIS** :
- Nombre de cachets : 7
- Montant par cachet : 94,29€
- Total à payer : 7 × 94,29€ = **660,03€**

---

## 🔧 **Corrections appliquées aujourd'hui**

### ✅ **Configuration**
- Tarifs corrigés selon formulaire client
- Valeurs converties en nombres (pas texte)

### ✅ **Module_Calculs.bas**
1. **Fonction `CalculerTarifJournee` (ligne 283)** :
   - ❌ AVANT : Utilisait TARIF_1_VISITE, TARIF_BRANLY_2H (n'existent pas)
   - ✅ APRÈS : Utilise TARIF_BRANLY, TARIF_MARINE, etc. (Configuration)
   - ✅ Calcul : Tarif × Nombre de visites
   - ✅ Ajustement proportionnel si durée différente

2. **Fonction `IdentifierTypeVisite` (ligne 250)** :
   - ❌ AVANT : Cherchait dans le NOM de la visite
   - ✅ APRÈS : Lit la colonne **Type_Visite** (colonne 5)

3. **Système de cachets (ligne 159-181)** :
   - ✅ Calcule total mensuel
   - ✅ Divise par nb jours travaillés
   - ✅ Arrondi supérieur (RoundUp)
   - ✅ Enregistre dans colonnes F et G de Calculs_Paie

---

## 📋 **Structure des feuilles**

### **Visites**
| Col | Nom | Description |
|-----|-----|-------------|
| A | ID_Visite | Identifiant unique |
| B | Date | Date de la visite |
| C | Heure_Debut | Heure de début |
| D | Heure_Fin | Heure de fin |
| E | **Type_Visite** | BRANLY, MARINE, etc. |
| F | Musee | Nom du musée |
| G | Langue | Langue de la visite |
| H | Nb_Personnes | Nombre de visiteurs |
| I | Tarif | Tarif calculé |
| J | Guide_Attribue | Guide assigné |
| K | Statut | Statut de la visite |

### **Calculs_Paie**
| Col | Nom | Description |
|-----|-----|-------------|
| A | Prenom | Prénom du guide |
| B | Nom | Nom du guide |
| C | Nb_Visites | Nombre total de visites |
| D | Nb_Heures | Nombre de jours travaillés |
| E | Total_Brut | Total mensuel brut |
| F | **Montant_Par_Cachet** | Cachet journalier |
| G | **Nb_Cachets** | = Nb jours |
| H | **Total_Recalcule** | Cachet × Nb jours |
| I | Mois | Mois concerné |

---

## 🎬 **Comment utiliser le système**

### **En fin de mois** :
1. **Admin** lance `CalculerVisitesEtSalaires()`
2. Le système parcourt le planning
3. Pour chaque guide :
   - Compte les jours travaillés
   - Somme les montants journaliers
   - Calcule le cachet : Total ÷ Nb jours
4. Remplit la feuille **Calculs_Paie**

### **Génération contrats** :
1. Début de mois : Contrat avec **tarif minimum**
2. Fin de mois : Contrat avec **cachets calculés**

---

## ⚠️ **Points d'attention**

1. **Type_Visite doit être correct** dans la feuille Visites
   - Utiliser exactement : BRANLY, MARINE, HORS-LES-MURS, EVENEMENT, VISIO, AUTRE
   - Sinon → tarif "AUTRE" appliqué (50€)

2. **Durées** :
   - Si Heure_Fin fournie → calcul proportionnel
   - Sinon → durée standard utilisée

3. **Configuration** :
   - Ne PAS modifier les noms de paramètres
   - Garder les valeurs en NOMBRES (pas texte)

---

## 📞 **Support**

Si problème de calcul, vérifier :
1. ✅ Type_Visite correct dans Visites
2. ✅ Guide_Attribue rempli dans Planning
3. ✅ Dates valides
4. ✅ Configuration avec nombres (pas texte)

---

*Document créé le 11/11/2025*
*Système de gestion planning guides musées*
