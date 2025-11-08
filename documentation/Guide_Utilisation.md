# 📗 Guide d'Utilisation - Système de Gestion des Guides

## 📋 Table des matières

1. [Vue d'ensemble](#vue-densemble)
2. [Gestion des guides](#gestion-des-guides)
3. [Saisie des disponibilités](#saisie-des-disponibilités)
4. [Gestion des visites](#gestion-des-visites)
5. [Génération du planning](#génération-du-planning)
6. [Envoi des emails](#envoi-des-emails)
7. [Calculs de paie](#calculs-de-paie)
8. [Génération des contrats](#génération-des-contrats)
9. [Workflows complets](#workflows-complets)

---

## 🎯 Vue d'ensemble

Le système permet de :
- ✅ Gérer les disponibilités des guides de manière **confidentielle**
- ✅ Attribuer automatiquement les guides aux visites
- ✅ Envoyer les plannings par email
- ✅ Notifier automatiquement les guides (J-7 et J-1)
- ✅ Calculer les salaires automatiquement
- ✅ Générer les contrats pré-remplis

---

## 👥 Gestion des guides

### Ajouter un nouveau guide

1. Aller dans la feuille **"Guides"**
2. Ajouter une nouvelle ligne avec les informations :

| Colonne | Description | Exemple |
|---------|-------------|---------|
| ID_Guide | Identifiant unique | G001, G002, etc. |
| Nom | Nom de famille | DUPONT |
| Prénom | Prénom | Marie |
| Email | Email professionnel | marie.dupont@email.fr |
| Téléphone | Numéro de téléphone | 0601020304 |

⚠️ **Important :** L'ID_Guide doit être unique et ne jamais changer.

### Modifier un guide

1. Localiser la ligne du guide dans la feuille "Guides"
2. Modifier les informations nécessaires
3. **NE PAS modifier l'ID_Guide** (sinon perte de l'historique)

### Désactiver un guide

Ne pas supprimer la ligne, mais :
1. Ajouter "(INACTIF)" après le nom
2. Ou déplacer dans une section "Anciens guides"

---

## 📅 Saisie des disponibilités

### Pour un guide (saisie individuelle)

**Méthode 1 : Macro assistée (recommandée)**

1. Onglet **Développeur** → **Macros**
2. Sélectionner : **`SaisirDisponibilites`**
3. Cliquer sur **Exécuter**
4. Suivre les instructions :
   - Entrer votre ID Guide (exemple : G001)
   - Sélectionner la période (date début et fin)
   - Pour chaque jour, indiquer si vous êtes disponible (Oui/Non)
   - Ajouter un commentaire si nécessaire

✅ **Avantage :** Interface guidée, pas besoin de toucher aux feuilles

**Méthode 2 : Saisie manuelle**

1. Aller dans la feuille **"Disponibilités"**
2. Ajouter vos disponibilités :

| ID_Guide | Date | Disponible | Commentaire |
|----------|------|------------|-------------|
| G001 | 15/12/2025 | OUI | |
| G001 | 16/12/2025 | NON | Congés |
| G001 | 17/12/2025 | OUI | |

⚠️ **Attention :** Cette méthode expose les disponibilités des autres guides.

### Exporter ses propres disponibilités

1. Macro : **`ExporterMesDisponibilites`**
2. Entrer votre ID Guide
3. Choisir l'emplacement de sauvegarde
4. Un fichier Excel avec uniquement vos disponibilités est créé

### Import en masse (administrateur)

Si les guides remplissent un fichier externe :

1. Préparer un fichier Excel avec les colonnes : ID_Guide, Date, Disponible, Commentaire
2. Macro : **`ImporterDisponibilitesMasse`**
3. Sélectionner le fichier
4. Les données sont importées automatiquement

---

## 🏛️ Gestion des visites

### Ajouter une visite

1. Aller dans la feuille **"Visites"**
2. Ajouter une nouvelle ligne :

| Colonne | Description | Format | Exemple |
|---------|-------------|--------|---------|
| ID_Visite | Identifiant unique | V001, V002... | V042 |
| Date | Date de la visite | jj/mm/aaaa | 20/12/2025 |
| Heure_Debut | Heure de début | hh:mm | 10:00 |
| Heure_Fin | Heure de fin | hh:mm | 12:00 |
| Musée | Nom du musée | Texte | Musée du Louvre |
| Type_Visite | Type | Texte | Visite guidée |
| Nombre_Visiteurs | Nombre | Nombre | 25 |

⚠️ **Important :** Les horaires permettent de calculer la durée et donc le salaire.

### Modifier une visite

1. Localiser la visite dans la feuille "Visites"
2. Modifier les informations
3. **Régénérer le planning** pour prendre en compte les changements

### Supprimer une visite

1. Supprimer la ligne dans "Visites"
2. Supprimer la ligne correspondante dans "Planning" (si elle existe)

---

## 🗓️ Génération du planning

### Génération automatique

**C'est la fonctionnalité principale !**

1. Macro : **`GenererPlanningAutomatique`**
2. Le système :
   - Parcourt toutes les visites
   - Cherche les guides disponibles pour chaque date
   - Attribue automatiquement un guide libre
   - Identifie les visites sans guide disponible (en rouge)

✅ **Résultat :** La feuille "Planning" est remplie automatiquement

**Codes couleur :**
- 🟢 Vert : Visite attribuée avec succès
- 🔴 Rouge : Aucun guide disponible

### Voir les guides disponibles pour une date

1. Macro : **`AfficherGuidesDisponiblesPourVisite`**
2. Entrer la date souhaitée
3. Une liste des guides disponibles s'affiche

### Modifier une attribution manuellement

Si l'attribution automatique ne convient pas :

1. Macro : **`ModifierAttribution`**
2. Entrer l'ID de la visite
3. Voir les informations actuelles
4. Entrer le nouvel ID du guide
5. Confirmation

**Ou directement dans la feuille "Planning" :**
- Modifier la colonne "ID_Guide"
- Le nom se met à jour automatiquement (si formule présente)

### Exporter le planning

1. Macro : **`ExporterPlanning`**
2. Choisir l'emplacement
3. Un fichier Excel séparé est créé

---

## 📧 Envoi des emails

### Envoyer le planning mensuel à tous les guides

**Fréquence recommandée :** Une fois par mois (début du mois)

1. Macro : **`EnvoyerPlanningMensuel`**
2. Entrer le mois concerné (format MM/AAAA, exemple : 12/2025)
3. Le système :
   - Groupe les visites par guide
   - Envoie un email personnalisé à chaque guide
   - Affiche un résumé

**Contenu de l'email :**
- Planning personnel du guide
- Dates et horaires de chaque visite
- Nombre total de visites
- Mention des rappels automatiques

### Envoyer les notifications automatiques (J-7 et J-1)

**Configuration recommandée :** Tâche planifiée quotidienne

**Méthode manuelle :**
1. Macro : **`EnvoyerNotificationsAutomatiques`**
2. Le système :
   - Vérifie toutes les visites
   - Envoie un email aux guides concernés si :
     - La visite est dans 7 jours (première notification)
     - La visite est demain (rappel)

**Méthode automatique (Windows) :**
1. Planificateur de tâches Windows
2. Créer une tâche quotidienne (par exemple à 9h00)
3. Action : Ouvrir le fichier Excel et exécuter la macro
4. Voir le guide d'installation pour les détails

### Tester l'envoi d'emails

1. Macro : **`TestEnvoiEmail`**
2. Entrer votre email de test
3. Vérifier la réception
4. Si ça ne fonctionne pas, vérifier la configuration Outlook

---

## 💰 Calculs de paie

### Calculer les salaires pour une période

1. Macro : **`CalculerVisitesEtSalaires`**
2. Choisir :
   - Filtrer par mois (MM/AAAA) pour une période spécifique
   - Laisser vide pour calculer sur toute la période

✅ **Résultat :** La feuille "Calculs_Paie" est remplie avec :
- ID et nom du guide
- Nombre de visites effectuées
- Montant total du salaire

**Calcul du salaire :**
```
Salaire = Nombre d'heures × Tarif horaire
```

Le tarif horaire est défini dans la feuille "Configuration".

### Générer une fiche de paie individuelle

Pour un guide spécifique :

1. Macro : **`GenererFichePaieGuide`**
2. Entrer l'ID du guide
3. Entrer le mois (MM/AAAA)
4. Un fichier Excel détaillé est créé avec :
   - Informations du guide
   - Liste de toutes les visites du mois
   - Calcul détaillé des heures et du montant
   - Totaux

### Exporter un récapitulatif général

Pour l'ensemble des guides :

1. Macro : **`ExporterRecapitulatifPaie`**
2. Choisir la période si demandé
3. Un fichier Excel avec le récapitulatif complet est généré

---

## 📄 Génération des contrats

### Générer un contrat pour un guide

1. Macro : **`GenererContratGuide`**
2. Entrer l'ID du guide
3. Entrer le mois du contrat (MM/AAAA)
4. Le système :
   - Collecte toutes les visites du guide pour ce mois
   - Calcule les heures et le montant total
   - Génère un document de contrat pré-rempli

✅ **Contenu du contrat :**
- Informations de l'association
- Informations du guide (nom, email, téléphone)
- Période du contrat
- Liste complète des dates de visite
- Horaires détaillés
- Calcul de la rémunération
- Clauses contractuelles
- Zones de signature

⚠️ **À faire manuellement :**
- Vérifier et adapter les clauses juridiques
- Compléter l'adresse de l'association
- Faire signer les deux parties

### Générer tous les contrats d'un mois

1. Macro : **`GenererContratsEnMasse`**
2. Entrer le mois (MM/AAAA)
3. Sélectionner un dossier de destination
4. Tous les contrats sont générés automatiquement

### Voir l'historique des contrats

1. Macro : **`AfficherContratsGeneres`**
2. Une liste de tous les contrats générés s'affiche

Ou consulter directement la feuille **"Contrats"**.

---

## 🔄 Workflows complets

### Workflow 1 : Nouveau mois

**Au début de chaque mois :**

1. ✅ **Ajouter les visites du mois** dans la feuille "Visites"
2. ✅ **Collecter les disponibilités** :
   - Envoyer un email aux guides
   - Chaque guide utilise `SaisirDisponibilites`
3. ✅ **Générer le planning** : `GenererPlanningAutomatique`
4. ✅ **Vérifier et ajuster** :
   - Identifier les visites non attribuées (en rouge)
   - Utiliser `ModifierAttribution` si nécessaire
   - Contacter des guides supplémentaires si besoin
5. ✅ **Envoyer les plannings** : `EnvoyerPlanningMensuel`
6. ✅ **Configurer les notifications** (si pas déjà fait)

### Workflow 2 : Fin de mois (paie)

**À la fin du mois :**

1. ✅ **Calculer les salaires** : `CalculerVisitesEtSalaires` (pour le mois écoulé)
2. ✅ **Vérifier les calculs** dans la feuille "Calculs_Paie"
3. ✅ **Générer les contrats** : `GenererContratsEnMasse` (pour le mois)
4. ✅ **Générer les fiches de paie** : `GenererFichePaieGuide` (pour chaque guide)
5. ✅ **Envoyer les documents** aux guides
6. ✅ **Archiver** les fichiers générés

### Workflow 3 : Gestion quotidienne

**Chaque jour (automatisé ou manuel) :**

1. ✅ **Notifications automatiques** : `EnvoyerNotificationsAutomatiques`
   - Les guides sont notifiés 7 jours et 1 jour avant leurs visites

**Si besoin :**
- Vérifier les disponibilités
- Ajuster le planning
- Ajouter des visites de dernière minute

### Workflow 4 : Ajout d'un nouveau guide

1. ✅ Ajouter le guide dans la feuille **"Guides"** (avec ID unique)
2. ✅ Le guide saisit ses disponibilités : `SaisirDisponibilites`
3. ✅ Régénérer le planning si nécessaire : `GenererPlanningAutomatique`
4. ✅ Le guide est maintenant inclus dans les attributions

---

## 🔐 Confidentialité et sécurité

### Bonnes pratiques

1. **Protéger les feuilles sensibles** :
   - Disponibilités
   - Calculs_Paie
   - Configuration

2. **Limiter les accès** :
   - Les guides ne doivent accéder qu'à la macro `SaisirDisponibilites`
   - Créer un fichier séparé pour eux si nécessaire

3. **Sauvegarder régulièrement** :
   - Copie de sécurité hebdomadaire
   - Avant chaque opération importante

4. **Emails confidentiels** :
   - Ne jamais mettre tous les emails en destinataire
   - Le système envoie automatiquement des emails individuels

---

## 📊 Astuces et conseils

### Pour gagner du temps

- ✅ Créer des boutons sur une feuille "Accueil" pour les macros fréquentes
- ✅ Utiliser des raccourcis clavier pour les macros (via Options)
- ✅ Automatiser les notifications avec le Planificateur Windows
- ✅ Créer des vues filtrées dans les feuilles (filtres Excel)

### Pour éviter les erreurs

- ✅ Toujours vérifier le planning avant envoi
- ✅ Faire un test d'envoi email avant la première utilisation
- ✅ Vérifier les horaires des visites (impacts salaires)
- ✅ Ne jamais modifier les ID (guides, visites)

### Pour personnaliser

- ✅ Adapter les textes des emails dans le code VBA
- ✅ Modifier le modèle de contrat selon vos besoins légaux
- ✅ Ajouter des colonnes personnalisées dans les feuilles
- ✅ Changer les couleurs dans le module Configuration

---

## 🆘 FAQ - Questions fréquentes

### Q : Un guide ne reçoit pas les emails, pourquoi ?

**R :** Vérifier :
- Son adresse email dans la feuille "Guides" (pas de faute)
- Qu'il a bien des visites assignées dans le planning
- Qu'Outlook est bien configuré
- Les spams/courrier indésirable

### Q : Comment annuler une visite ?

**R :**
1. Supprimer la ligne dans "Visites"
2. Supprimer la ligne dans "Planning"
3. Prévenir le guide concerné (email manuel ou nouveau planning)

### Q : Le planning automatique ne trouve pas de guide, mais j'en vois de disponibles

**R :** Vérifier :
- Les dates correspondent exactement
- Le format des dates est correct (jj/mm/aaaa)
- La colonne "Disponible" contient bien "OUI" (en majuscules)
- Le guide n'a pas déjà une autre visite ce jour-là

### Q : Comment modifier le tarif horaire en cours de mois ?

**R :**
1. Modifier dans la feuille "Configuration"
2. Recalculer les salaires : `CalculerVisitesEtSalaires`
3. ⚠️ Attention : cela affecte tous les calculs rétroactifs

### Q : Puis-je utiliser le système sans Outlook ?

**R :** Oui, mais il faut modifier le code VBA pour utiliser :
- Une autre application email
- Un service SMTP (Gmail, etc.)
- Ou désactiver l'envoi automatique

---

## 📞 Support et maintenance

### Sauvegarde

**Sauvegarde automatique :**
- Excel sauvegarde automatiquement les versions récentes
- Fichier → Informations → Gérer le classeur → Récupérer des classeurs non enregistrés

**Sauvegarde manuelle :**
- Copier le fichier `.xlsm` régulièrement
- Renommer avec la date : `Planning_Guides_2025_12_01.xlsm`

### Mise à jour

Si de nouveaux modules sont fournis :
1. Sauvegarder le fichier actuel
2. Ouvrir l'éditeur VBA
3. Supprimer l'ancien module
4. Importer le nouveau module
5. Tester

### Journal des modifications

Tenir un journal dans une feuille "Historique" :
- Date
- Action effectuée
- Par qui
- Remarques

---

## ✨ Félicitations !

Vous maîtrisez maintenant toutes les fonctionnalités du système de gestion des guides.

**Pour aller plus loin :**
- Personnaliser les emails
- Ajouter des statistiques
- Créer des rapports visuels (graphiques)
- Automatiser davantage avec le Planificateur de tâches

**Bon courage dans la gestion de vos visites guidées ! 🏛️**
