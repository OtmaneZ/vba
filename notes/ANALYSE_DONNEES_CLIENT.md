# 📊 ANALYSE DES DONNÉES CLIENT - Le Bal de Saint-Bonnet

**Date** : 10 novembre 2025
**Fichier analysé** : FORMULAIRE_CLIENT_PRO.xlsx

---

## ✅ DONNÉES COMPLÈTES ET EXPLOITABLES

### 🏛️ **Infos Musée**
- **Environnement technique** :
  - Excel : Microsoft 365 ✅
  - Outlook : ❌ NON (utilise OVH mail) → **CRITIQUE : Pas d'automatisation email via VBA**
  - Utilisateurs : 1 personne (+ délégation occasionnelle)

### 👥 **Liste des Guides** (15 guides)
✅ **Complet** : Noms, prénoms, emails fournis
⚠️ **Manquant** : Téléphones et tarifs horaires vides

**Liste complète** :
1. Evelyne MOSER - peyronelle@yahoo.fr
2. Maiko VUILLOD - maikovuillod@gmail.com
3. Peggy GENESTIE - peggy.genestie@gmail.com
4. Mathieu BOULET - bouletmathieu603@gmail.com
5. Pablo CONTESTABILE - pablocontestabile16@gmail.com
6. Gabriela ARANGUIZ - Gabriela.aranguiz.munoz@gmail.com
7. Fanny MILCENT - fanny.comedy@gmail.com
8. Rosaline DESLAURIERS - rosaline.deslauriers@gmail.com
9. Sandrine COLOMBET - sandrinecolombet@free.fr
10. Ange GRAH - bekangety@gmail.com
11. Hanako DANJO - hanako.danjo@gmail.com
12. Silvia MASSEGUR - smassegur@gmail.com
13. Solene ARBEL - solene.arbel@hotmail.fr
14. Shady NAFAR - shadynafar@hotmail.fr
15. Marie-Laure SAINT-BONNET - mlsb@club-internet.fr

### 🎫 **Types de Visites** (environ 20 types)
✅ **Complet** avec code couleur complexe :
- 🔵 Bleu : Individuels
- 🔵 Bleu clair : Groupes (avec établissement + niveau scolaire)
- 🌸 Rose : Événements (Dimanche en famille, Nuit des Musées, etc.)
- 🔴 Rouge : Hors-les-murs (hôpital, prison, médiathèques, etc.)
- 🔵 Bleu foncé gras MAJUSCULES : Visites MARINE + événements spéciaux

**Exemples** :
- Ma Petite Visite Contée Maman Serpent (45 min)
- Ma Petite Visite Contée Petit Ours (45 min)
- Couleurs, Autour du monde, Asie, Afrique, Amérique, Océanie (1h)
- Devins et Sorciers, Orient, Femmes, Carnavals, Mythes de création (1h)

### 📅 **Organisation**
✅ **Complet** :
- Ouvert 7j/7 (Lundi à Dimanche)
- Horaires :
  - Matin : 8h30 - 13h00
  - Après-midi : 13h00 - 18h00
  - Soir : 18h00 - 00h00 (optionnel)
- Fermeture : 11 novembre 2025 → 1er mai 2025 (⚠️ dates incohérentes, à clarifier)

### ⚙️ **Règles et Contraintes**
✅ **Détaillé** :
- Min : 4 visites/guide/mois (avec exceptions)
- Max : 25 visites/guide/mois (avec exceptions)
- Max : 4 visites/guide/jour
- Rotation obligatoire : Oui
- Priorité seniors : Non
- Jours fixes : Non

⚠️ **CONTRAINTES COMPLEXES** (à gérer manuellement) :
- **Peggy** : Ne fait PAS Maman Serpent, Petit Ours, BULLE
- **Hanako** : Fait UNIQUEMENT les visites 3.5 ans + Couleurs + Autour du Monde
- **Silvia** : Fait UNIQUEMENT les visites 3.5 ans + Couleurs + Autour du Monde + Orient
- **Marianne** : Fait SEULEMENT BULLE, ZOO, A L'ABORDAGE + événements MARINE
- **Solène** : Fait SEULEMENT BULLE, ZOO, A L'ABORDAGE + événements MARINE + (à définir)
- **Shady** : À préciser

### 📊 **Volume d'Activité**
✅ **Données claires** :
- Période calme : 30 visites/mois
- Période normale : 100 visites/mois
- Période chargée : 150 visites/mois
- Mois chargés : MAI, JUIN, NOVEMBRE, DÉCEMBRE

⚠️ **Spécificités** :
- Délai réservation standard : 2 semaines
- Délai MARINE : 2 jours avant
- Délai BRANLY : 1 semaine avant
- **Planning fait 15-18 jours avant le mois** (ex : planning décembre fait le 13 novembre)
- Nouvelles réservations possibles J-2 (MARINE) et J-7 (BRANLY)

### 📧 **Communication**
✅ **Besoins définis** :
- Planning mensuel : **EN MILIEU DE MOIS (autour du 13-15)** → Oui
- Rappel J-7 : Oui
- Rappel J-1 : Oui
- Récapitulatif mensuel salaires : Oui
- Objet email : "Planning de vos visites"
- Signature : "L'équipe du musée"

---

## ⚠️ PROBLÈMES CRITIQUES IDENTIFIÉS

### 🚨 **1. PAS D'OUTLOOK INSTALLÉ**
**Impact** : Impossible d'envoyer des emails automatiques via VBA Outlook

**Solutions possibles** :
1. ❌ Installer Outlook (coût + complexité config OVH)
2. ✅ **API email externe** (SendGrid, Mailgun, Brevo) → Python requis
3. ✅ **Génération de brouillons dans OVH Mail** (copier-coller manuel)
4. ✅ **Export CSV des emails à envoyer** → client gère manuellement

**Recommandation** : Solution 3 ou 4 (simplicité + pas de coût)

### ⚠️ **2. TARIFS HORAIRES MANQUANTS**
**Impact** : Impossible de calculer automatiquement les salaires

**Action requise** : Demander les tarifs horaires par guide lors de l'appel 14h

### ⚠️ **3. RÈGLES DE SPÉCIALISATION COMPLEXES**
**Impact** : Impossible à automatiser complètement (6 guides avec contraintes spécifiques)

**Solution** : 
- Système semi-automatique : génération planning avec **alertes visuelles**
- Validation manuelle obligatoire par la cliente

### ⚠️ **4. DATES DE FERMETURE INCOHÉRENTES**
11 novembre 2025 → 1er mai 2025 (impossible, va dans le passé)

**Action requise** : Clarifier lors de l'appel 14h

---

## 📋 ÉTAT D'AVANCEMENT DU PROJET

### ✅ **CE QUI EST FAIT**
1. ✅ Structure Excel complète (8 feuilles)
2. ✅ Système d'authentification Guide/Admin
3. ✅ Collecte disponibilités guides
4. ✅ Génération planning automatique (avec contraintes basiques)
5. ✅ Calcul salaires avec taux dégressif
6. ✅ Génération contrats automatique
7. ✅ Interface Accueil avec navigation

### ⚠️ **CE QUI MANQUE / À ADAPTER**

#### **1. Système de notification emails** 
**Statut** : ⚠️ À REVOIR (pas d'Outlook)

**Options** :
- A. Générer un fichier CSV avec liste emails à envoyer
- B. Créer des modèles d'emails pré-remplis à copier-coller
- C. Intégration API externe (hors scope VBA pur)

**Décision** : À discuter avec cliente (14h)

#### **2. Contraintes de spécialisation guides**
**Statut** : ⚠️ SEMI-AUTOMATISABLE

**LE PROBLÈME** :
6 guides ont des restrictions spécifiques sur les types de visites qu'ils peuvent faire :

| Guide | Peut faire | Ne peut PAS faire |
|-------|-----------|-------------------|
| **Peggy** | Tous les autres | ❌ Maman Serpent, Petit Ours, BULLE |
| **Hanako** | UNIQUEMENT 3.5 ans + Couleurs + Autour du Monde | ❌ Toutes les autres visites |
| **Silvia** | UNIQUEMENT 3.5 ans + Couleurs + Autour du Monde + Orient | ❌ Toutes les autres visites |
| **Marianne** | UNIQUEMENT BULLE, ZOO, A L'ABORDAGE + événements MARINE | ❌ Toutes les autres visites |
| **Solène** | UNIQUEMENT BULLE, ZOO, A L'ABORDAGE + événements MARINE + (à définir) | ❌ Toutes les autres visites |
| **Shady** | À préciser avec cliente | À préciser |

**IMPACT SUR L'AUTOMATISATION** :
- Si on génère automatiquement le planning, risque d'attribuer une visite "Afrique" à Hanako → **IMPOSSIBLE**
- Si on attribue "Maman Serpent" à Peggy → **IMPOSSIBLE**
- Système doit vérifier la compatibilité guide ↔ visite AVANT attribution

**SOLUTIONS POSSIBLES** :

**Option A : Filtre automatique (recommandé)** ✅
1. Créer feuille "Spécialisations" :
   ```
   Guide          | Visite autorisée
   ---------------|-----------------
   Peggy          | Asie
   Peggy          | Afrique
   Peggy          | Amérique
   ...
   Hanako         | Maman Serpent
   Hanako         | Petit Ours
   Hanako         | Couleurs
   ...
   ```
2. Lors de la génération planning :
   - VBA vérifie si le guide peut faire cette visite
   - Si NON → passe au guide suivant
   - Si OUI → attribution possible

**Option B : Alertes visuelles post-génération** ⚠️
1. Planning généré automatiquement SANS filtre
2. Macro de vérification qui colore en ROUGE les attributions impossibles
3. Responsable corrige manuellement

**Option C : Validation manuelle complète** ❌
1. Système propose des guides disponibles
2. Responsable choisit manuellement pour chaque visite
3. Perd l'intérêt de l'automatisation

**RECOMMANDATION** : **Option A** avec vérification automatique
- Effort dev : +2h (création feuille + VBA de filtrage)
- Fiabilité : ✅ Aucune erreur possible
- Maintenance : ✅ Facile d'ajouter/modifier contraintes

**Ce qu'on peut faire** :
- ✅ Créer une feuille "Spécialisations" pour mapper guides ↔ types de visites
- ✅ Lors de la génération planning, **filtrer automatiquement** les guides non compatibles
- ✅ Afficher **alertes visuelles** si aucun guide disponible pour une visite

**Ce qu'on ne peut PAS faire** :
- ❌ Deviner les spécialisations non renseignées (Shady, précisions pour Solène)
- ❌ Gérer des règles changeantes sans mise à jour de la feuille Spécialisations

#### **3. Tarifs horaires guides**
**Statut** : ⚠️ DONNÉES MANQUANTES

**Action** : Demander les tarifs lors de l'appel 14h

#### **4. Code couleur planning complexe**
**Statut** : ⚠️ FAISABLE EN VBA (mais complexe)

**LE PROBLÈME** :
La cliente utilise un système de couleurs très spécifique dans son planning actuel :

| Couleur | Type de visite | Formatage spécial |
|---------|---------------|-------------------|
| 🔵 **Bleu** | Visites individuels | Standard |
| 🔵 **Bleu clair** | Visites groupes | + Colonnes "Établissement" + "Niveau scolaire" |
| 🌸 **Rose** | Événements (Dimanche en famille, Un Autre Noël, Nuit des Musées, etc.) | Standard |
| 🔴 **Rouge** | Hors-les-murs (hôpital, prison, médiathèque, centre culturel, etc.) | Standard |
| 🔵 **Bleu foncé** | Visites MARINE + événements spéciaux | **GRAS + MAJUSCULES** |

**IMPACT SUR L'AUTOMATISATION** :
- VBA doit **identifier automatiquement** le type de visite pour appliquer la bonne couleur
- Problème : Comment savoir qu'une visite est "hors-les-murs" vs "individuel" ?
- Solution : Besoin d'une colonne "Catégorie" dans la feuille Visites

**SOLUTIONS POSSIBLES** :

**Option A : Colonne catégorie dans feuille Visites (recommandé)** ✅
1. Ajouter colonne "Catégorie" dans feuille Visites :
   ```
   Date | Type visite | Guide | Catégorie
   -----|-------------|-------|------------
   15/12 | Asie | Maiko | Individuel
   16/12 | Asie | Pablo | Groupe
   17/12 | BULLE MARINE | Marianne | Marine
   18/12 | Prison Fleury | Ange | Hors-les-murs
   ```

2. VBA applique automatiquement la couleur selon catégorie :
   ```vb
   Select Case categorie
       Case "Individuel"
           cell.Interior.Color = RGB(0, 112, 192) ' Bleu
       Case "Groupe"
           cell.Interior.Color = RGB(155, 194, 230) ' Bleu clair
       Case "Événement"
           cell.Interior.Color = RGB(255, 192, 203) ' Rose
       Case "Hors-les-murs"
           cell.Interior.Color = RGB(255, 0, 0) ' Rouge
       Case "Marine"
           cell.Interior.Color = RGB(0, 32, 96) ' Bleu foncé
           cell.Font.Bold = True
           cell.Value = UCase(cell.Value) ' MAJUSCULES
   End Select
   ```

**Option B : Détection automatique par mots-clés** ⚠️
- VBA détecte "BULLE", "MARINE", "ZOO" → catégorie Marine
- VBA détecte "prison", "hôpital", "médiathèque" → catégorie Hors-les-murs
- **Risque** : Faux positifs, maintenance complexe

**Option C : Pas de code couleur automatique** ❌
- Responsable colore manuellement après génération planning
- **Perd l'intérêt de l'automatisation**

**RECOMMANDATION** : **Option A** avec colonne catégorie
- Effort dev : +2h (ajout colonne + VBA formatage)
- Fiabilité : ✅ 100% fiable
- Maintenance : ✅ Facile à gérer
- Clarté : ✅ Catégorie visible dans les données

**Ce qu'on doit coder** :
- ✅ Colonne "Catégorie" dans feuille Visites (si pas déjà présente)
- ✅ Menu déroulant : Individuel / Groupe / Événement / Hors-les-murs / Marine
- ✅ VBA qui applique automatiquement :
  - 🔵 Bleu : Individuels
  - 🔵 Bleu clair : Groupes
  - 🌸 Rose : Événements
  - 🔴 Rouge : Hors-les-murs
  - 🔵 Bleu foncé GRAS MAJUSCULES : MARINE

**Effort** : 2-3h de dev

**QUESTION À POSER LORS DE L'APPEL** :
> "Votre système de couleurs est très précis. Pour l'automatiser, j'ai besoin de savoir comment vous identifiez qu'une visite est 'hors-les-murs' ou 'marine'. Vous le notez quelque part ou c'est juste visuel ?"

#### **5. Dates de fermeture**
**Statut** : ⚠️ DONNÉES INCOHÉRENTES

**Action** : Clarifier lors de l'appel 14h

---

## ❓ QUESTIONS À POSER LORS DE L'APPEL (14H)

### **1. Notifications emails (CRITIQUE)**
- "Vous n'avez pas Outlook installé. Comment souhaitez-vous gérer l'envoi des plannings et rappels ?"
  - Option A : Je génère un fichier avec tous les emails pré-rédigés, vous les copiez-collez dans OVH Mail
  - Option B : Je crée un bouton qui ouvre votre client email avec le message pré-rempli
  - Option C : Vous acceptez d'installer Outlook (payant, config OVH)

### **2. Tarifs horaires guides**
- "Quel est le tarif horaire de chaque guide ? Sont-ils tous au même tarif ?"
- "Y a-t-il des majorations (week-end, soir, événements spéciaux) ?"

### **3. Dates de fermeture**
- "J'ai noté une fermeture du 11 novembre 2025 au 1er mai 2025, c'est bien ça ?"
- (Probablement erreur de saisie, à corriger)

### **4. Spécialisations guides**
- "Les contraintes de spécialisation (Peggy, Hanako, Silvia, Marianne, Solène, Shady) sont-elles complètes ?"
- "Faut-il ajouter d'autres guides avec des contraintes ?"
- "Acceptez-vous un système semi-automatique avec validation manuelle ?"

### **5. Planning MARINE et BRANLY**
- "Les réservations MARINE (J-2) et BRANLY (J-7) doivent-elles être traitées différemment dans le planning ?"
- "Faut-il un code couleur spécifique pour ces visites ?"

### **6. Code couleur**
- "Vous avez un système de couleurs complexe. Voulez-vous que je l'implémente exactement ou on simplifie ?"
- (Bleu, bleu clair, rose, rouge, bleu foncé gras majuscules)

### **7. Validation du projet existant**
- "J'ai créé un système avec 8 feuilles Excel + authentification + planning automatique + calculs salaires. Voulez-vous que je vous montre rapidement pour valider l'approche ?"

### **8. Délai et budget**
- "Vu la complexité (15 guides, contraintes spécifiques, pas d'Outlook), je vous confirme la livraison pour [DATE] avec [BUDGET]. Ça vous convient ?"

---

## 🎯 PLAN D'ACTION APRÈS L'APPEL

### **Scénario A : Elle accepte les limitations (recommandé)**
1. ✅ Implémenter code couleur planning (2h)
2. ✅ Créer feuille Spécialisations guides (1h)
3. ✅ Intégrer tarifs horaires (30 min)
4. ✅ Système d'export emails CSV (1h30)
5. ✅ Documentation utilisateur (1h)
6. ✅ Tests et livraison (1h)

**Total** : ~7h de travail restant

### **Scénario B : Elle veut Outlook absolument**
1. ⚠️ L'aider à installer Outlook (1-2h support)
2. ⚠️ Configurer SMTP OVH (complexe, risque d'échec)
3. ✅ Intégrer code email VBA Outlook (déjà fait dans le projet actuel)
4. ✅ Reste des tâches (5h)

**Total** : ~10-12h (avec risque technique)

---

## 💰 ÉVALUATION BUDGET

**Travail déjà effectué** : ~40h (structure complète, VBA, authentification, planning, calculs)

**Travail restant** :
- Scénario A (limitations acceptées) : 7h
- Scénario B (Outlook requis) : 10-12h

**Budget recommandé** : 
- Si forfait déjà négocié : tenir le budget
- Si à renégocier : +500-800€ pour complexités supplémentaires

---

## ✅ CONCLUSION : ON A CE QU'IL FAUT ?

### **OUI, les données sont complètes pour :**
✅ Créer la structure Excel
✅ Gérer 15 guides avec leurs emails
✅ Définir ~20 types de visites
✅ Paramétrer horaires et jours d'ouverture
✅ Implémenter règles de base (min/max visites)
✅ Calculer volumes d'activité

### **NON, il manque pour finaliser :**
❌ Tarifs horaires guides (calcul salaires)
⚠️ Solution technique emails (pas d'Outlook)
⚠️ Clarification dates fermeture
⚠️ Validation approche semi-automatique pour spécialisations

### **CE QUI RESTE À FAIRE :**
1. **Appel 14h** : Clarifier les 8 questions ci-dessus
2. **Développement** : 7-12h selon scénario choisi
3. **Tests** : 1-2h avec données réelles
4. **Formation** : 1h avec la cliente
5. **Support post-livraison** : 2-3h (corrections/ajustements)

---

## 🎯 RECOMMANDATION FINALE

**Le projet est RÉALISABLE et on a 90% des données.**

**Approche recommandée pour l'appel 14h** :
1. ✅ Montrer ce qui est déjà fait (impressionner)
2. ⚠️ Expliquer les limitations techniques (Outlook)
3. ✅ Proposer des solutions pragmatiques (export CSV emails)
4. ✅ Valider les contraintes de spécialisation (semi-auto)
5. ✅ Récupérer les tarifs horaires
6. ✅ Confirmer délai et budget final

**Prévision** : Si elle accepte les adaptations, **livraison possible sous 3-5 jours ouvrés**.
