# 🎯 CHEAT SHEET DÉMO CLIENT - Système de Gestion Planning Musée

**Date :** 11 novembre 2025
**Client :** Le Bal de Saint-Bonnet
**Système :** Gestion automatisée planning guides + rémunération en cachets

---

## 📋 PLAN DE DÉMO (20 minutes max)

### 🔹 PARTIE 1 : Connexion & Sécurité (3 min)

**CE QUE TU MONTRES :**
1. Ouvrir le fichier → Page Accueil apparaît automatiquement
2. Connexion ADMIN : `admin` / mot de passe admin
3. Montrer qu'on voit TOUTES les feuilles (Planning, Guides, Visites, etc.)
4. Se déconnecter
5. Connexion GUIDE : `ID_guide` / mot de passe guide
6. Montrer qu'on voit SEULEMENT "Mon Planning" (sécurité)

**CE QUI SE PASSE AUTOMATIQUEMENT :**
- ✅ Toutes les feuilles masquées sauf Accueil au démarrage
- ✅ Droits d'accès selon le rôle (admin voit tout, guide voit que son planning)
- ✅ Impossibilité de modifier les feuilles cachées

**MODULE :** `Module_Authentification.bas`

---

### 🔹 PARTIE 2 : Attribution Automatique (5 min)

**CE QUE TU MONTRES :**

**Avant :**
```
Feuille Planning :
| Date     | Heure | Musée  | Type Visite | Guide Attribué |
|----------|-------|--------|-------------|----------------|
| 15/11/25 | 10h   | Branly | 2 visites   | NON ATTRIBUE   |
```

**Action :** Bouton "Attribuer automatiquement"

**Après :**
```
| Date     | Heure | Musée  | Type Visite | Guide Attribué |
|----------|-------|--------|-------------|----------------|
| 15/11/25 | 10h   | Branly | 2 visites   | GUIDE_001      | ✅
```

**CE QUI SE PASSE AUTOMATIQUEMENT :**
- ✅ Cherche les guides disponibles cette date
- ✅ Vérifie les spécialisations (Branly, Marine, etc.)
- ✅ Répartit équitablement entre guides
- ✅ Colore en vert si attribué, rouge si problème
- ✅ Met à jour "Mon Planning" du guide concerné

**MODULE :** `Module_Planning.bas` (fonction `AttribuerGuidesAutomatiquement`)

---

### 🔹 PARTIE 3 : Calcul Cachets Automatique (5 min)

**CE QUE TU MONTRES :**

**Feuille Calculs_Paie (avant calcul) :**
```
| Guide    | Nb Visites | Nb Jours | Total | Cachet | Total Recalculé |
|----------|------------|----------|-------|--------|-----------------|
| (vide)   |            |          |       |        |                 |
```

**Action :** Bouton "Calculer salaires" → Choisir le mois (ex: 10/2025)

**Feuille Calculs_Paie (après calcul) :**
```
| Guide      | Nb Visites | Nb Jours | Total   | Cachet  | Total Recalculé |
|------------|------------|----------|---------|---------|-----------------|
| Jean Dupont| 14         | 7        | 790.00€ | 112.86€ | 790.02€         |
```

**EXPLICATION CLIENT (IMPORTANT) :**

> "Le système calcule automatiquement selon votre mail :
> - 1 visite/jour = 80€
> - 2 visites/jour = 110€
> - 3+ visites/jour = 140€
> - Hors-les-murs = 100€
>
> **Exemple concret :**
> Jean a fait 7 jours en octobre :
> - 1er oct : 2 visites → 110€
> - 4 oct : 2 visites → 110€
> - 7 oct : 1 visite → 80€
> - 15 oct : hors-les-murs → 100€
> - 17 oct : 3 visites → 140€
> - 24 oct : 2 visites → 110€
> - 28 oct : 3 visites → 140€
>
> **Total : 790€ ÷ 7 jours = 112.86€ par cachet** (arrondi supérieur)"

**CE QUI SE PASSE AUTOMATIQUEMENT :**
- ✅ Compte automatiquement le nombre de visites PAR JOUR pour chaque guide
- ✅ Applique le bon tarif selon le nombre de visites
- ✅ Calcule le cachet : Total ÷ Nb jours (arrondi supérieur)
- ✅ Vérifie que Total Recalculé = Cachet × Nb jours

**MODULE :** `Module_Calculs.bas` (fonction `CalculerVisitesEtSalaires`)

---

### 🔹 PARTIE 4 : Décompte Mensuel Détaillé (3 min)

**Action :** Bouton "Générer décompte mensuel" → Choisir le mois

**CE QUE ÇA GÉNÈRE (nouveau fichier Excel) :**

```
=== DÉCOMPTE DÉTAILLÉ - OCTOBRE 2025 ===

| Guide       | Date       | Heure | Type Visite | Catégorie  | Nb Jours | Cachet  |
|-------------|------------|-------|-------------|------------|----------|---------|
| Jean Dupont | 01/10/2025 | 10h   | Branly      | Branly     | 7        | 112.86€ |
| Jean Dupont | 04/10/2025 | 14h   | Marine      | Marine     | 7        | 112.86€ |
| ...         | ...        | ...   | ...         | ...        | ...      | ...     |

STATISTIQUES PAR CATÉGORIE :
Visites Branly :     84
Visites Marine :     15
Hors-les-murs :      5
Événements :         2
Visio :              1
Autres :             1
TOTAL :              108

NOMBRE DE JOURS PAR GUIDE :
Jean Dupont :    7 jours
Marie Martin :   5 jours
```

**CE QUI SE PASSE AUTOMATIQUEMENT :**
- ✅ Liste toutes les visites du mois avec dates et horaires
- ✅ Compte automatiquement par catégorie (Branly, Marine, etc.)
- ✅ Calcule le nombre de jours travaillés par guide
- ✅ Affiche le montant du cachet pour chaque ligne

**MODULE :** `Module_Calculs.bas` (fonction `GenererDecompteMensuel`)

---

### 🔹 PARTIE 5 : Génération Contrats (2 min)

**Action :** Bouton "Générer contrats"

**CE QUE ÇA GÉNÈRE :**
- ✅ Contrat de début de mois (pré-rempli avec planning prévisionnel)
- ✅ Contrat de fin de mois (avec dates réelles + cachets calculés)
- ✅ Génération automatique en format Word/PDF

**CE QUI SE PASSE AUTOMATIQUEMENT :**
- ✅ Remplit automatiquement : nom, prénom, adresse, dates
- ✅ Insère le nombre de cachets et le montant
- ✅ Prêt à signer

**MODULE :** `Module_Contrats.bas`

---

### 🔹 PARTIE 6 : Export DPAE (2 min)

**Action :** Bouton "Exporter DPAE"

**CE QUE ÇA GÉNÈRE :**
- ✅ Fichier Excel avec toutes les infos pour la DPAE
- ✅ Format prêt à copier-coller dans le site gouvernemental
- ✅ Une ligne par guide avec dates de contrat

**CE QUI SE PASSE AUTOMATIQUEMENT :**
- ✅ Récupère automatiquement : nom, prénom, NIR, adresse
- ✅ Calcule dates début/fin de contrat
- ✅ Format conforme DPAE

**MODULE :** `Module_DPAE.bas`

---

## 🤖 AUTOMATISATIONS INVISIBLES

**Ces choses se font TOUTES SEULES (pas besoin de cliquer) :**

### 1. Notifications automatiques (TOUS LES JOURS à 8h-18h)
- ✅ Envoie email aux guides 7 jours avant leur mission
- ✅ Rappel 1 jour avant
- ✅ Pas de doublon (n'envoie qu'une fois)

**MODULE :** `Module_Emails.bas` + `ThisWorkbook.cls` ligne 104-108

---

### 2. Planning mensuel (1er du mois à 9h)
- ✅ Envoie automatiquement le planning du mois à tous les guides
- ✅ Format récapitulatif clair

**MODULE :** `Module_Emails.bas` + `ThisWorkbook.cls` ligne 98-102

---

### 3. Calcul salaires (dernier jour du mois à 17h)
- ✅ Propose automatiquement de calculer les salaires
- ✅ Demande confirmation avant calcul
- ✅ Propose ensuite de générer les contrats

**MODULE :** `ThisWorkbook.cls` ligne 110-127

---

## 🎨 Système de Couleurs (Automatique)

**Dans la feuille Planning :**
- 🟢 **VERT** = Visite attribuée et confirmée
- 🟡 **JAUNE** = Guide disponible mais pas encore attribué
- 🔴 **ROUGE** = Aucun guide disponible cette date
- ⚪ **BLANC** = "NON ATTRIBUE"

**MODULE :** `Module_CodeCouleur.bas`

---

## 📊 RÉCAPITULATIF : Que fait le système ?

| Fonctionnalité | Manuelle | Automatique |
|----------------|----------|-------------|
| Connexion sécurisée | ✅ | - |
| Attribution guides | ✅ (clic bouton) | ✅ (cherche disponibilités) |
| Calcul cachets | ✅ (clic bouton) | ✅ (formule 80/110/140) |
| Décompte mensuel | ✅ (clic bouton) | ✅ (stats par catégorie) |
| Génération contrats | ✅ (clic bouton) | ✅ (remplissage auto) |
| Export DPAE | ✅ (clic bouton) | ✅ (format conforme) |
| Notifications guides | - | ✅ (tous les jours) |
| Planning mensuel | - | ✅ (1er du mois) |
| Proposition calcul salaire | - | ✅ (fin de mois) |
| Mise à jour "Mon Planning" | - | ✅ (instantané) |
| Couleurs visuelles | - | ✅ (automatique) |

---

## 🚨 SCÉNARIOS DE DÉMO (choisis 1 ou 2)

### Scénario 1 : Ajouter une nouvelle mission
```
1. Ouvre feuille Visites
2. Ajoute ligne : "Visite Halloween" | Branly | 31/10/25 | 14h
3. Ouvre feuille Planning
4. Ajoute ligne : 31/10/25 | 14h | Branly | ID visite
5. Clic "Attribuer automatiquement"
6. → Guide attribué automatiquement !
7. Se connecter en tant que ce guide
8. → "Mon Planning" mis à jour automatiquement !
```

### Scénario 2 : Calculer salaire d'un guide
```
1. Remplis quelques lignes dans Planning (mois passé)
2. Attribue des guides
3. Clic "Calculer salaires"
4. Entre le mois (ex: 10/2025)
5. → Feuille Calculs_Paie se remplit
6. → Cachets calculés automatiquement
7. Clic "Générer décompte"
8. → Fichier Excel détaillé généré
```

### Scénario 3 : Ajouter un nouveau guide
```
1. Ouvre feuille Guides
2. Ajoute ligne : GUIDE_005 | Nouveau Nom | email | tel | NIR | adresse
3. Feuille Disponibilites → Ajoute ses dispos
4. Feuille Specialisations → Coche ses musées
5. → Système le prend en compte automatiquement !
6. → Il peut se connecter avec son ID
```

---

## 💡 RÉPONSES AUX QUESTIONS CLIENTS

### "Comment ça marche si j'ajoute une visite ?"
> "Vous ajoutez la ligne dans Planning, vous cliquez sur 'Attribuer automatiquement', et le système cherche le guide disponible avec la bonne spécialisation. C'est instantané."

### "Et si un guide est malade ?"
> "Vous changez le nom du guide dans Planning, le système met à jour automatiquement son planning et celui du nouveau guide."

### "Comment je vérifie les calculs ?"
> "Le décompte détaillé liste TOUTES les visites avec dates, horaires, et le comptage par catégorie. Vous pouvez vérifier manuellement."

### "C'est sûr niveau sécurité ?"
> "Oui, les guides ne voient QUE leur planning. Impossible d'accéder aux autres données. Seul l'admin voit tout."

### "Je peux changer les tarifs ?"
> "Oui, feuille Configuration. Vous changez TARIF_1_VISITE, TARIF_2_VISITES, etc. Le système recalcule automatiquement."

### "Et si j'ai un problème ?"
> "Support 7 jours inclus. Je corrige/améliore si besoin. Après validation, je vous forme et je documente tout."

---

## ✅ CHECKLIST AVANT DÉMO

**À FAIRE MAINTENANT (5 min) :**
- [ ] Ouvrir le fichier et vérifier que ça s'ouvre bien
- [ ] Tester connexion admin/guide
- [ ] Vérifier qu'il y a des données de test dans Planning
- [ ] Préparer un exemple de calcul (790€ / 7 jours)
- [ ] Fermer tous les autres fichiers Excel
- [ ] Désactiver notifications macOS (pour pas être dérangé)
- [ ] Avoir un verre d'eau à côté 😊

**PENDANT LA DÉMO :**
- [ ] Partager écran avec Excel en plein écran
- [ ] Parler lentement et expliquer CHAQUE clic
- [ ] Laisser des silences pour qu'ils posent des questions
- [ ] Noter leurs remarques dans un Notepad à côté
- [ ] NE PAS dire "je sais pas" → dire "je note et je vérifie"

**APRÈS LA DÉMO :**
- [ ] Récapituler ce qui a été validé
- [ ] Proposer support 7 jours
- [ ] Envoyer email de confirmation
- [ ] Respirer ! 🎉

---

## 🎯 PHRASE CLÉ DE CONCLUSION

> "Le système est opérationnel et testé. Je vous propose de le tester avec vos vraies données cette semaine, et je reste disponible 7 jours pour tout ajustement. C'est la procédure standard pour garantir que ça correspond exactement à votre usage quotidien."

---

**BON COURAGE ! TU VAS ASSURER ! 💪🚀**
