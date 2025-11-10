# QUESTIONS PRIORITAIRES APPEL 14H - MARIE-LAURE SAINT-BONNET

## 🎯 URGENCE HAUTE

### 1. Tarifs guides individuels
**Question :** Tous les guides sont payés selon le même barème ou certains ont des tarifs spécifiques ?
- Si barème unique → OK, on utilise Standard/Événement BRANLY/Hors-les-murs BRANLY
- Si tarifs individuels → À renseigner dans colonne "Tarif horaire" du fichier V2

**Contexte :** Fichier V2 contient 3 barèmes :
- Standard : 80€ (1 visite/jour), 110€ (2 visites/jour), 140€ (3 visites/jour)
- Événement BRANLY : 120€ (2h), 150€ (3h), 180€ (4h)
- Hors-les-murs BRANLY : 100€ (1 visite/jour), 130€ (2 visites/jour), 160€ (3 visites/jour)

### 2. Application des barèmes BRANLY
**Question :** Quelles visites utilisent "Événement BRANLY" et "Hors-les-murs BRANLY" ?
- Actuellement : 0 visites détectées pour ces barèmes
- "Événement BRANLY" → durée en heures (2h/3h/4h) ?
- "Hors-les-murs BRANLY" → hors les murs au musée du Quai Branly uniquement ?

**Action si validation :** Identifier et marquer ces visites dans colonne Barème

### 3. Spécialisations MARINE (Marianne & Solène)
**Question :** Marianne et Solène font-elles TOUS les événements MARINE ou seulement certains ?

**19 événements MARINE détectés :**
- LA BULLE
- LE ZOO
- L'ABORDAGE
- JOYEUX MERCREDI LES PETITS MOUSSES !
- JOURNÉES DU PATRIMOINE
- NUIT DE LA LECTURE
- AUTRE (x5 slots)
- + visites standards mentionnant "Marine"

**Action si confirmation :** Ajouter dans feuille Spécialisations :
```
Marianne | Marine | UNIQUEMENT | [liste des événements] | Autorisée uniquement sur événements Marine
Solène | Marine | UNIQUEMENT | [liste des événements] | Autorisée uniquement sur événements Marine
```

## 📌 URGENCE MOYENNE

### 4. Validation catégories automatiques
**Action :** Montrer la feuille Visites avec code couleur et valider :
- 42 Individuelles (bleu clair) → OK ?
- 15 Événements (rose/orange) → OK ?
- 3 Hors-les-murs (rouge/orange) → OK ?
- 19 Marine (bleu foncé GRAS) → OK ?
- 1 Groupe (bleu très clair) → OK ?

### 5. Les 9 "AUTRE" - cas par cas
**Question :** Ces 9 slots "AUTRE" sont pour des événements ponctuels ?
- Barème actuel : "Cas par cas"
- Besoin d'un système de saisie manuelle du tarif pour ces visites ?

## 💡 INFORMATION (PAS DE QUESTION)

### 6. Outlook pour email automatique
**Rappel :** Le système d'envoi automatique d'emails nécessite Outlook installé
- OVH Mail peut être configuré dans Outlook (même boîte, 2 interfaces)
- Alternative : Export CSV des emails à envoyer manuellement

**Pas urgent**, on peut livrer sans cette fonctionnalité et l'ajouter plus tard

---

## 📋 CHECKLIST POST-APPEL

Après l'appel, actions à faire selon réponses :

- [ ] Mettre à jour colonne Barème selon visites BRANLY identifiées
- [ ] Ajouter spécialisations MARINE (Marianne/Solène) dans feuille Spécialisations
- [ ] Corriger catégorisations si erreurs détectées
- [ ] **Réécrire Module_Calculs.bas** pour nouveau système tarif journalier
- [ ] Tester génération planning complet avec 79 types
- [ ] Documenter utilisation colonne "Cas par cas"

---

## ⏰ TIMING

- **13h00** → Dernière relecture questions
- **13h30** → Ouvrir PLANNING_MUSEE_FINAL.xlsm pour démo visuelle
- **14h00** → APPEL (durée estimée 30-45 min)
- **14h45** → Implémenter modifications selon réponses
- **16h30** → Tests finaux
- **17h00** → Livraison

---

**Objectif appel :** Valider l'approche technique pour éviter refactoring majeur après implémentation Module_Calculs.bas
