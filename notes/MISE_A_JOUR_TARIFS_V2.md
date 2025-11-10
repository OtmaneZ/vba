# 📊 MISE À JOUR - TARIFS ET NOUVEAUX TYPES VISITES

**Date** : 10 novembre 2025 - 12h15
**Fichier reçu** : FORMULAIRE_CLIENT_PRO V2.xlsx

---

## ✅ NOUVELLES DONNÉES REÇUES

### **1. Tarifs guides enfin fournis !**

**Barème STANDARD** :
- 1 visite/jour : **80€**
- 2 visites/jour : **110€**
- 3 visites/jour : **140€**

**Barème ÉVÉNEMENT BRANLY** :
- 2h : **120€**
- 3h : **150€**
- 4h : **180€**

**Barème HORS-LES-MURS BRANLY** :
- 1 visite/jour : **100€**
- 2 visites/jour (même endroit) : **130€**
- 3 visites/jour (même endroit) : **160€**

**⚠️ Note** : "montant différent selon l'événement ou hors-les-murs au cas par cas"

### **2. Nouveaux types de visites ajoutés**

**Total : 79 types de visites** (vs 20 initialement)

**Ajouts principaux** :
- **Visio contées** (3 types)
- **Hors les murs** (3 types)
- **Temps d'échange** (30 min)
- **Événements avec durées variables** :
  - Dimanche en famille (2h, 3h, 4h)
  - Tous au Musée (1h, 2h, 3h, 4h)
  - Nuit des Musées (1h, 2h, 3h, 4h)
  - Un Autre Noël (1h, 2h, 3h, 4h)
  - Jardin des contes (1h, 2h, 3h, 4h)
  - Week-end de l'ethnologie (1h, 2h, 3h, 4h)
  - Événement Branly (1h, 2h, 3h, 4h)
- **Visites MARINE** :
  - BULLE (45 min)
  - ZOO (1h)
  - A L'ABORDAGE (1h)
  - JOYEUX MERCREDI (1h, 2h, 3h, 4h)
  - JOURNEES DU PATRIMOINE (1h, 2h, 3h, 4h)
  - NUIT DE LA LECTURE (1h, 2h, 3h, 4h)
  - EVENEMENT MARINE (1h, 2h, 3h, 4h)
- **5 slots "AUTRE"** pour flexibilité future

---

## ⚠️ CE QUI MANQUE ENCORE

### **Tarifs horaires individuels par guide**

**Colonne "Tarif horaire (€)" dans feuille Guides : TOUJOURS VIDE**

**Question pour l'appel 14h :**
> "Tous les guides sont-ils payés selon le même barème (80€/110€/140€) ou certains ont des tarifs spécifiques ?"

**Options possibles :**
1. **Tous au même tarif** → On utilise le barème standard pour tous
2. **Tarifs différenciés** → Elle doit remplir la colonne "Tarif horaire" pour chaque guide

---

## 🔧 IMPLICATIONS TECHNIQUES

### **1. Module_Calculs.bas à adapter**

**Actuellement** : Système dégressif basique
```vb
If nbVisites <= 5 Then
    montant = nbVisites * tauxBase
ElseIf nbVisites <= 10 Then
    montant = (5 * tauxBase) + ((nbVisites - 5) * tauxBase * 0.9)
Else
    montant = (5 * tauxBase) + (5 * tauxBase * 0.9) + ((nbVisites - 10) * tauxBase * 0.8)
End If
```

**Nouveau système requis** :
```vb
Function CalculerSalaireGuide(nomGuide As String, mois As String) As Double
    Dim nbVisitesParJour As Collection
    ' Compter combien de visites par jour
    ' Appliquer le bon barème :
    ' - Si 1 visite ce jour → +80€
    ' - Si 2 visites ce jour → +110€
    ' - Si 3+ visites ce jour → +140€
    ' - Si événement BRANLY → tarif selon durée
    ' - Si hors-les-murs BRANLY → tarif selon nb visites
End Function
```

**Complexité** :
- ⚠️ Le calcul n'est plus basé sur le **mois** mais sur le **jour**
- ⚠️ Tarifs différents selon **type de visite** (standard, BRANLY événement, hors-les-murs)
- ⚠️ "Au cas par cas pour certains événements" → gestion manuelle nécessaire

### **2. Feuille "Types Visites" à catégoriser**

**79 types de visites** nécessitent une colonne "Barème" pour automatiser :
- Standard (80/110/140)
- Événement BRANLY (120/150/180)
- Hors-les-murs BRANLY (100/130/160)
- Cas par cas (à gérer manuellement)

### **3. Colonne "Catégorie" à remplir**

**79 visites** à catégoriser en :
- Individuel
- Groupe
- Événement
- Hors-les-murs
- Marine

**Suggestion pour l'appel** :
> "J'ai 79 types de visites. Pour automatiser le code couleur, avez-vous un pattern pour les catégoriser rapidement ?"

**Patterns détectés** :
- Visites contées standards → **Individuel**
- Visites thématiques (Asie, Afrique, etc.) → **Individuel** ou **Groupe**
- Dimanche en famille, Tous au Musée, etc. → **Événement**
- "Hors les murs" dans le nom → **Hors-les-murs**
- BULLE, ZOO, A L'ABORDAGE, MARINE en majuscules → **Marine**

---

## 🎯 PLAN D'ACTION AVANT APPEL 14H

### **1. Créer barème de tarification (5 min)**
Ajouter colonne "Barème" dans feuille Types Visites :
- Standard / Événement BRANLY / Hors-les-murs BRANLY / Cas par cas

### **2. Pré-catégoriser les visites (10 min)**
Remplir automatiquement la colonne "Catégorie" selon patterns détectés

### **3. Adapter Module_Calculs.bas (15 min)**
Coder le nouveau système de calcul par jour avec les 3 barèmes

### **4. Préparer questions pour l'appel**
- Validation du barème (tous les guides au même tarif ?)
- Clarification "cas par cas"
- Validation catégorisation automatique

---

## 📋 QUESTIONS PRIORITAIRES APPEL 14H

### **1. Tarifs guides (CRITIQUE)**
- "Tous les guides sont payés selon ce barème ou il y a des différences ?"
- "Les 'cas par cas' pour événements : comment je sais lesquels ?"
- "Si un guide fait 4 visites dans une journée, c'est toujours 140€ ou ça augmente ?"

### **2. Catégorisation visites**
- "J'ai détecté 79 types. Puis-je les catégoriser automatiquement selon ces règles... ?"
- (Montrer les patterns détectés)

### **3. Spécialisations mises à jour**
- "Vous avez ajouté des visites MARINE (JOYEUX MERCREDI, JOURNEES PATRIMOINE, etc.)"
- "Marianne et Solène font-elles tous les événements MARINE ou seulement BULLE/ZOO/ABORDAGE ?"

---

## ✅ CE QUI EST PRÊT

- ✅ Structure Excel avec feuille Spécialisations
- ✅ Colonne Catégorie dans Visites
- ✅ Modules VBA pour spécialisations et code couleur
- ✅ Système de génération planning avec vérifications

## ⚠️ CE QUI NÉCESSITE ADAPTATION

- ⚠️ Module_Calculs.bas (nouveau barème jour par jour)
- ⚠️ Feuille Types Visites (ajouter colonne Barème)
- ⚠️ Catégorisation des 79 visites
- ⚠️ Mise à jour Spécialisations (nouveaux types MARINE)

---

## ⏱️ TEMPS RESTANT AVANT 14H

**Il est 12h15, appel à 14h = 1h45 disponible**

**Priorisation** :
1. **Créer colonne Barème + pré-remplissage** (10 min) ✅ À FAIRE
2. **Pré-catégoriser les visites** (10 min) ✅ À FAIRE
3. **Adapter Module_Calculs.bas** (30 min) ⚠️ Peut attendre validation
4. **Mise à jour Spécialisations** (15 min) ⚠️ Peut attendre validation
5. **Préparer démo** (20 min) ✅ À FAIRE

**Décision** : Faire 1, 2, 5 maintenant. 3 et 4 après validation pendant l'appel.
