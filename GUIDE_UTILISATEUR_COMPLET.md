# 📘 Guide Utilisateur - Système de Planning Musée

## 🎯 Bienvenue !

Ce guide vous explique comment utiliser votre nouveau système de gestion de planning pour les guides du musée.

**Tout est automatisé pour vous faciliter la vie :**
- ✅ Les guides déclarent leurs disponibilités
- ✅ Vous créez les visites et attribuez les guides
- ✅ Les emails sont envoyés automatiquement
- ✅ Les calculs de paie se font tout seuls
- ✅ Les contrats sont générés en un clic

---

## 📋 Table des matières

1. [Premier démarrage](#premier-démarrage)
2. [Connexion au système](#connexion-au-système)
3. [Gestion des guides](#gestion-des-guides)
4. [Système de disponibilités](#système-de-disponibilités)
5. [Créer et attribuer des visites](#créer-et-attribuer-des-visites)
6. [Calculs de paie et défraiements](#calculs-de-paie-et-défraiements)
7. [Générer les contrats](#générer-les-contrats)
8. [Configuration email](#configuration-email)
9. [Questions fréquentes](#questions-fréquentes)
10. [Aide et support](#aide-et-support)

---

## 🚀 Premier démarrage

### Étape 1 : Ouvrir le fichier

1. Double-cliquez sur **`PLANNING_MUSEE_FINAL_PROPRE.xlsm`**
2. Excel s'ouvre avec un bandeau jaune en haut :
   ```
   ⚠️ AVERTISSEMENT DE SÉCURITÉ
   Les macros ont été désactivées.
   [Activer le contenu]
   ```
3. **Cliquez sur "Activer le contenu"** → C'est indispensable !

### Étape 2 : Première configuration

1. Allez sur l'onglet **"Accueil"**
2. Vous voyez deux boutons :
   - 🔑 **[Admin]** → Pour vous
   - 👤 **[Guide]** → Pour les guides

3. Allez sur l'onglet **"Configuration"**
4. Modifiez les informations :
   ```
   Ligne 2, Colonne A : Email_Expediteur
   Ligne 2, Colonne B : votre-email@outlook.fr

   Ligne 3, Colonne A : Nom_Association
   Ligne 3, Colonne B : Votre nom d'association
   ```

### Étape 3 : Ajouter vos guides

1. Allez sur l'onglet **"Guides"**
2. Remplissez pour chaque guide :
   ```
   Colonne A : Prénom (ex: Marie)
   Colonne B : Nom (ex: Dupont)
   Colonne C : Email (ex: marie.dupont@email.com)
   Colonne D : Téléphone (ex: 06 12 34 56 78)
   Colonne E : Spécialisations (ex: Art moderne, Histoire)
   Colonne F : Mot de passe (ex: marie123)
   ```

**Vous êtes prêt !** 🎉

---

## 🔐 Connexion au système

### Pour vous (Administrateur)

1. Cliquez sur l'onglet **"Accueil"**
2. Cliquez sur le bouton **[Admin]**
3. Entrez le mot de passe : `admin123`
4. **Tous les onglets s'affichent** :
   - Guides
   - Disponibilités
   - Visites
   - Planning
   - Calculs_Paie
   - Contrats
   - Configuration

### Pour les guides

1. Cliquez sur l'onglet **"Accueil"**
2. Cliquez sur le bouton **[Guide]**
3. Liste des guides s'affiche
4. Guide sélectionne son nom
5. Guide entre son mot de passe
6. **Guide voit uniquement** :
   - Mes_Disponibilites (pour déclarer ses dispos)
   - Mon_Planning (lecture seule, voit ses visites)

**Important :** Chaque guide ne voit QUE ses propres données. Les autres guides restent invisibles (confidentialité).

---

## 👥 Gestion des guides

### Ajouter un nouveau guide

1. Connectez-vous en admin
2. Allez sur **"Guides"**
3. Ajoutez une nouvelle ligne :
   ```
   A: Prénom
   B: Nom
   C: Email
   D: Téléphone
   E: Spécialisations (séparées par virgule)
   F: Mot de passe (simple, le guide pourra le changer)
   ```

### Modifier un guide

- Modifiez directement dans la feuille "Guides"
- Changements effectifs immédiatement

### Supprimer un guide

**⚠️ ATTENTION : Ne supprimez JAMAIS une ligne !**

Pour désactiver un guide :
1. Ajoutez une colonne "Actif"
2. Mettez "NON" pour les guides inactifs
3. Ou laissez la ligne vide

**Pourquoi ?** Les historiques de planning et de paie sont liés aux numéros de ligne.

---

## 📅 Système de disponibilités

### Comment ça marche ?

1. **Guides déclarent leurs disponibilités** à l'avance
2. **Système détecte automatiquement** qui est dispo pour chaque visite
3. **Vous voyez directement** la liste des guides disponibles
4. **Vous attribuez** en connaissance de cause

### Guide déclare ses disponibilités

**Le guide fait ça lui-même :**

1. Se connecte avec son login
2. Va sur **"Mes_Disponibilites"**
3. Remplit ligne par ligne :
   ```
   Colonne A : Son numéro de guide (1, 2, 3...)
   Colonne B : Date (format : 15/11/2025)
   Colonne C : Disponible (valeurs : OUI ou NON)
   Colonne D : Commentaire optionnel (ex: "Préfère matin")
   ```

**Exemple :**
```
1 | 15/11/2025 | OUI | Disponible toute la journée
1 | 16/11/2025 | NON | Rendez-vous médical
1 | 17/11/2025 | OUI | Matin uniquement
1 | 20/11/2025 | OUI |
```

**Points importants :**
- ✅ Guide peut remplir **quand il veut** (pas besoin d'attendre une date précise)
- ✅ Guide peut ajouter **autant de lignes** qu'il veut
- ✅ Guide peut **modifier** ses dispos n'importe quand
- ✅ **Une seule ligne par date** pour éviter les confusions

### Vous consultez les disponibilités

En tant qu'admin, vous voyez **TOUTES** les disponibilités :

1. Allez sur **"Disponibilites"** (onglet visible pour vous uniquement)
2. Vous voyez toutes les déclarations de tous les guides
3. Colonne **"Guides_Disponibles"** dans Planning se remplit automatiquement

**Le système fait le travail pour vous !** 🎯

---

## 🎫 Créer et attribuer des visites

### Étape 1 : Créer une visite

1. Connectez-vous en admin
2. Allez sur **"Visites"**
3. Ajoutez une nouvelle ligne :
   ```
   Colonne A : ID_Visite (numéro unique, ex: 1, 2, 3...)
   Colonne B : Date (format: 15/11/2025)
   Colonne C : Heure_Debut (format: 10:00)
   Colonne D : Heure_Fin (format: 12:00)
   Colonne E : Musée (ex: Louvre)
   Colonne F : Type_Visite (ex: Visite guidée, Atelier enfants)
   Colonne G : Nombre_Visiteurs (ex: 25)
   ```

### Étape 2 : Attribuer un guide

1. Allez sur **"Planning"**
2. La visite apparaît automatiquement
3. **Regardez la colonne "Guides_Disponibles"** → Le système affiche qui a dit OUI pour cette date !
   ```
   Exemple : "Marie Dupont, Jean Martin, Sophie Dubois"
   ```

4. **Choisissez un guide** dans la colonne "Guide_Attribue"
   - Liste déroulante avec tous vos guides
   - **Si vous choisissez quelqu'un qui n'est PAS dispo** → ⚠️ Message d'alerte :
     ```
     ⚠️ ATTENTION !
     Ce guide a déclaré ne PAS être disponible pour cette date.
     Voulez-vous quand même l'attribuer ?
     ```

5. **Confirmez l'attribution**

### Étape 3 : Email automatique envoyé

**Dès que vous attribuez**, le guide reçoit un email :

```
De: votre-email@outlook.fr
À: marie.dupont@email.com
Objet: Nouvelle visite attribuée - 15/11/2025

Bonjour Marie,

Une visite vous a été attribuée :

📅 Date : 15 novembre 2025
🕐 Heure : 10h00 - 12h00
🏛️ Musée : Louvre
📝 Type : Visite guidée
👥 Nombre : 25 personnes

Cette visite vous est attribuée automatiquement.
Pour toute modification, contactez l'administrateur.

Cordialement,
Musée des Guides

---
⚠️ NE PAS REPONDRE À CET EMAIL
Pour toute question, contactez-nous au : [votre téléphone]
```

### Étape 4 : Guide consulte son planning

1. Guide se connecte
2. Va sur **"Mon_Planning"**
3. **Voit ses visites en lecture seule** :
   ```
   Date       | Heure      | Musée  | Type           | Statut
   15/11/2025 | 10:00-12:00| Louvre | Visite guidée  | Confirmée
   20/11/2025 | 14:00-16:00| Orsay  | Atelier enfants| Confirmée
   ```

**Le guide ne peut PAS refuser** depuis l'interface. Tout changement passe par vous.

---

## 💰 Calculs de paie et défraiements

### Comment ça marche ?

À la fin du mois, le système calcule automatiquement :
- ✅ Nombre de visites par guide
- ✅ Nombre d'heures travaillées
- ✅ Montant total (système de cachets)
- ✅ Défraiements (remboursements de frais)
- ✅ Total final avec frais

### Lancer le calcul de paie

1. Connectez-vous en admin
2. Allez sur **"Calculs_Paie"**
3. Cliquez sur le bouton **"Calculer Paie du Mois"**
4. Choisissez le mois (ex: Novembre 2025)
5. **Le système remplit automatiquement** :
   ```
   Colonne A : Guide (nom complet)
   Colonne B : Nombre_Visites
   Colonne C : Total_Heures
   Colonne D : Montant_Salaire
   ... (autres colonnes de détails)
   Colonne I : Total_Brut (salaire avant défraiements)
   Colonne N : Défraiements (à remplir manuellement)
   Colonne O : Total_Avec_Frais (calculé automatiquement)
   ```

### Ajouter les défraiements

**Les défraiements = Remboursements de frais (transport, repas, etc.)**

1. Dans la colonne **N "Défraiements"**, entrez le montant pour chaque guide
   ```
   Exemple :
   Marie Dupont : 45.50 €
   Jean Martin : 0.00 €
   Sophie Dubois : 32.00 €
   ```

2. La colonne **O "Total_Avec_Frais"** se met à jour automatiquement
   ```
   Formule : Total_Brut + Défraiements = Total_Avec_Frais
   ```

### Vérifier les calculs

- **Colonne J** : Montant_Par_Cachet (montant unitaire par prestation)
- **Colonne K** : Nb_Cachets (nombre de prestations)
- **Colonne L** : Total_Recalcule (vérification : Nb_Cachets × Montant_Par_Cachet)
- **Colonne M** : Mois (mois concerné)

**Tout est vérifié et recalculé automatiquement !**

---

## 📄 Générer les contrats

### Contrat provisoire (début de mois)

1. Allez sur **"Contrats"**
2. Cliquez sur **"Générer Contrat Provisoire"**
3. Sélectionnez le guide
4. Choisissez la période
5. **Le contrat se génère automatiquement** avec :
   - Informations du guide
   - Liste des visites prévues
   - Montant estimé (sans défraiements)

### Contrat final (fin de mois)

1. **Calculez d'abord la paie** (étape précédente)
2. Cliquez sur **"Générer Contrat Final"**
3. Sélectionnez le guide
4. **Le contrat final inclut** :
   - Nombre réel de visites
   - Montant par cachet
   - **Sous-total cachets**
   - **Défraiements** (si montant > 0)
   - **MONTANT TOTAL** (avec frais inclus)
   - Détail du calcul

**Exemple d'affichage dans le contrat :**
```
Montant par cachet : 150,00 €
Sous-total cachets : 750,00 €
Défraiements : 45,50 €
MONTANT TOTAL DU : 795,50 €

Calcul : 5 cachets × 150,00 € = 750,00 € + Défraiements 45,50 € = 795,50 €
```

### Exporter le contrat

1. Contrat généré dans un nouvel onglet
2. **Clic droit** sur l'onglet → **"Déplacer ou copier..."**
3. Cochez **"Créer une copie"** → Nouveau classeur
4. **Fichier → Enregistrer sous** → Format PDF
5. Envoyez au guide par email

---

## 📧 Configuration email

### Prérequis

**Vous avez besoin de :**
- ✅ Microsoft Outlook installé sur votre Mac
- ✅ Un compte email configuré dans Outlook
- ✅ Outlook doit être **votre client mail par défaut**

### Créer une adresse email pour le planning (RECOMMANDÉ)

**Pourquoi ?** Séparer les emails du planning de votre boîte personnelle.

1. **Créez un compte Outlook.com** (gratuit) :
   - Allez sur https://outlook.com
   - Cliquez sur "Créer un compte"
   - Choisissez : `planning-musee@outlook.fr` (ou autre nom disponible)
   - Mot de passe : un mot de passe fort

2. **Ajoutez le compte dans Outlook** (app Mac) :
   - Ouvrez Outlook
   - Menu **Outils** → **Comptes**
   - Cliquez sur **"+"** (Ajouter un compte)
   - Entrez : `planning-musee@outlook.fr`
   - Entrez le mot de passe
   - Cliquez **"Ajouter un compte"**
   - Attendez la synchronisation

3. **Configurez Excel** :
   - Dans Excel, onglet **"Configuration"**
   - Ligne 2, Colonne B : `planning-musee@outlook.fr`

### Configuration de la réponse automatique

**Important** : Les guides vont recevoir des emails de cette adresse, mais vous ne voulez pas lire leurs réponses.

1. Dans **Outlook**, menu **Outils** → **Absence du bureau** (ou **Réponses automatiques**)
2. Activez les réponses automatiques
3. Message :
   ```
   Bonjour,

   Cette boîte mail est utilisée uniquement pour l'envoi de notifications
   automatiques du planning.

   Vos réponses ne seront pas consultées.

   Pour toute question, contactez directement :
   [Votre nom]
   Email : [votre-email-principal@...]
   Téléphone : [votre numéro]

   Cordialement,
   [Nom de votre association]
   ```

### Premier envoi : Autorisation de sécurité

**Au premier email envoyé depuis Excel**, Outlook affichera :

```
⚠️ Microsoft Excel souhaite accéder à Outlook
pour envoyer des emails.

[Refuser] [Autoriser]
```

**Cliquez sur "Autoriser"** → Cochez "Toujours autoriser"

**C'est tout !** Les emails partiront automatiquement ensuite.

---

## ❓ Questions fréquentes

### 1. "Le bouton Admin/Guide ne fait rien"

**Cause** : Les macros ne sont pas activées

**Solution** :
1. Fermez Excel
2. Rouvrez le fichier
3. Cliquez sur **"Activer le contenu"** dans le bandeau jaune

### 2. "Les emails ne partent pas"

**Causes possibles** :

**a) Outlook n'est pas installé**
- Vérifiez que Outlook est bien installé
- Ouvrez Outlook manuellement, vérifiez qu'il fonctionne

**b) Compte email non configuré**
- Ouvrez Outlook → Menu **Outils** → **Comptes**
- Vérifiez que votre compte email est présent et synchronisé

**c) Excel n'a pas l'autorisation**
- Au prochain envoi, cliquez sur **"Autoriser"**
- Cochez "Toujours autoriser"

**d) Adresse email incorrecte dans Configuration**
- Vérifiez l'onglet **"Configuration"**, ligne 2, colonne B
- Doit correspondre EXACTEMENT au compte configuré dans Outlook

### 3. "Un guide ne voit pas ses visites"

**Causes** :

**a) Guide ne s'est pas connecté**
- Guide doit cliquer sur **[Guide]** et se connecter
- Vérifiez login/mot de passe

**b) Visite attribuée à un autre guide**
- Vérifiez l'onglet **"Planning"**, colonne "Guide_Attribue"
- Le nom doit correspondre EXACTEMENT (majuscules/minuscules)

### 4. "La colonne Guides_Disponibles est vide"

**Cause** : Aucun guide n'a déclaré être disponible pour cette date

**Solution** :
1. Rappelez aux guides de remplir leurs disponibilités
2. Vérifiez l'onglet **"Disponibilites"** (en admin)
3. Si vide, les guides doivent ajouter des lignes

### 5. "Les calculs de paie sont incorrects"

**Vérifications** :

1. **Dates correctes** dans "Visites" ?
2. **Guide bien attribué** dans "Planning" ?
3. **Tarif par cachet** correct dans Configuration ?
4. **Recalculer** : Cliquez sur "Calculer Paie du Mois" à nouveau

### 6. "Guide dit qu'il s'est trompé dans ses dispos"

**Workflow** :

1. Guide vous contacte (téléphone, email, WhatsApp)
2. **Vous décidez** :
   - Option A : **Maintenir l'attribution** (guide doit honorer)
   - Option B : **Réattribuer** à un autre guide

3. **Pour réattribuer** :
   - Onglet **"Planning"**
   - Trouvez la ligne de la visite
   - Utilisez la fonction **"Refuser et Réattribuer"** (bouton ou macro)
   - Système cherche automatiquement un autre guide dispo
   - Email envoyé au nouveau guide

### 7. "Les contrats ne s'affichent pas correctement"

**Cause** : Colonnes masquées ou feuille mal formatée

**Solution** :
1. Supprimez l'onglet du contrat généré
2. Relancez la génération
3. Si problème persiste : vérifiez que **"Calculs_Paie"** est bien remplie

### 8. "Comment changer le mot de passe admin ?"

1. Ouvrez Excel
2. Appuyez sur **Alt + F11** (ou **Option + F11** sur Mac) → Éditeur VBA
3. Double-cliquez sur **"Module_Config"** (à gauche)
4. Cherchez la ligne :
   ```vb
   Public Const MOT_DE_PASSE_ADMIN As String = "admin123"
   ```
5. Changez `"admin123"` par votre nouveau mot de passe
6. **Fichier → Enregistrer** (Cmd+S)
7. Fermez l'éditeur VBA

### 9. "Un guide a oublié son mot de passe"

1. Allez sur l'onglet **"Guides"**
2. Trouvez la ligne du guide
3. Colonne F : changez le mot de passe
4. Communiquez-lui le nouveau mot de passe

### 10. "Puis-je utiliser un autre logiciel qu'Outlook ?"

**Actuellement : NON**

Le système est conçu pour Outlook uniquement. Pour utiliser Gmail, Thunderbird, etc., il faudrait modifier le code VBA (service payant).

**Alternative** : Désactiver les emails automatiques et envoyer manuellement.

---

## 🔧 Personnalisation

### Changer les tarifs

1. Onglet **"Configuration"**
2. Ajoutez des lignes pour vos paramètres :
   ```
   Tarif_Base_Horaire | 25.00
   Tarif_Par_Cachet | 150.00
   Majoration_Weekend | 1.25
   ```

3. Les calculs utilisent ces valeurs automatiquement

### Ajouter des types de visites

1. Onglet **"Visites"**
2. Colonne F "Type_Visite" : liste déroulante
3. Pour modifier la liste :
   - Sélectionnez une cellule de la colonne
   - Menu **Données** → **Validation des données**
   - Modifiez la liste : `Visite guidée, Atelier enfants, Conférence, ...`

### Ajouter des musées

Même principe que les types de visites :
- Colonne E "Musée"
- Validation des données
- Liste personnalisée

---

## 🆘 Aide et support

### Pendant les 7 premiers jours

**Support inclus pendant 7 jours après la livraison.**

**Contact :**
- Email : [Email du développeur]
- Délai de réponse : 24-48h
- Support disponible pour :
  - ✅ Bugs ou erreurs
  - ✅ Questions sur le fonctionnement
  - ✅ Petites modifications
  - ✅ Configuration email

### Après les 7 jours

**Vous êtes autonome !**

Ce guide contient tout ce dont vous avez besoin pour :
- Utiliser le système au quotidien
- Résoudre les problèmes courants
- Former vos guides

**En cas de problème majeur** :
- Contactez le développeur pour un support payant

---

## 📊 Récapitulatif du workflow complet

### 📅 Début de mois

1. **Guides remplissent leurs disponibilités** (quand ils veulent)
   - Se connectent
   - Onglet "Mes_Disponibilites"
   - Ajoutent des lignes : Date | OUI/NON

2. **Vous créez les visites**
   - Onglet "Visites"
   - Ajoutez : Date, Heure, Musée, Type

3. **Vous attribuez les guides**
   - Onglet "Planning"
   - Regardez "Guides_Disponibles" (rempli automatiquement)
   - Choisissez un guide
   - Email envoyé automatiquement

4. **Guides consultent leur planning**
   - Se connectent
   - Onglet "Mon_Planning"
   - Voient leurs visites en lecture seule

### 📆 Pendant le mois

5. **Ajustements si nécessaire**
   - Guide vous contacte pour changement
   - Vous réattribuez si besoin
   - Nouvel email envoyé automatiquement

### 💰 Fin de mois

6. **Calcul de la paie**
   - Onglet "Calculs_Paie"
   - Bouton "Calculer Paie du Mois"
   - Tout se remplit automatiquement

7. **Ajout des défraiements**
   - Colonne N : entrez les montants manuellement
   - Colonne O : total se calcule automatiquement

8. **Génération des contrats**
   - Onglet "Contrats"
   - "Générer Contrat Final"
   - Contrats incluent cachets + défraiements
   - Exportez en PDF
   - Envoyez aux guides

9. **Archivage**
   - Copiez la feuille "Calculs_Paie" → nouveau fichier
   - Nommez : `Paie_Novembre_2025.xlsx`
   - Conservez pour vos archives

---

## ✅ Checklist pour bien démarrer

### Configuration initiale (à faire une seule fois)

- [ ] Fichier Excel ouvert avec "Activer le contenu"
- [ ] Onglet "Configuration" rempli (email, nom association)
- [ ] Compte email créé (ex: planning-musee@outlook.fr)
- [ ] Compte ajouté dans Outlook
- [ ] Outlook configuré en réponse automatique
- [ ] Premier test d'email réussi (autorisation accordée)
- [ ] Onglet "Guides" rempli (tous vos guides)
- [ ] Mot de passe admin changé (optionnel)

### Utilisation mensuelle

- [ ] Demander aux guides de remplir leurs dispos
- [ ] Créer les visites du mois
- [ ] Attribuer les guides (vérifier colonne "Guides_Disponibles")
- [ ] Vérifier que les emails sont partis
- [ ] En fin de mois : calculer la paie
- [ ] Ajouter les défraiements
- [ ] Générer les contrats finaux
- [ ] Archiver les données du mois

---

## 🎉 Félicitations !

Vous êtes maintenant prêt à utiliser votre système de planning !

**Rappelez-vous :**
- ✅ Le système fait 90% du travail pour vous
- ✅ Les guides sont autonomes pour leurs dispos
- ✅ Les emails et calculs sont automatiques
- ✅ Tout est sécurisé et confidentiel

**Support disponible pendant 7 jours pour toute question.**

**Bonne gestion de vos plannings !** 🚀

---

*Guide utilisateur - Version finale - Novembre 2025*
*Système de gestion de planning pour guides de musée*
