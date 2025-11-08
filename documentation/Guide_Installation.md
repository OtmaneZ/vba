# 📘 Guide d'Installation - Système de Gestion des Guides

## 🎯 Prérequis

- **Microsoft Excel** (version 2016 ou supérieure recommandée)
- **Microsoft Outlook** (pour l'envoi automatique d'emails)
- **Windows** ou **macOS** (avec Excel installé)
- **Macros activées** dans Excel

---

## 📥 Installation Étape par Étape

### ÉTAPE 1 : Créer le fichier Excel

1. Ouvrir **Microsoft Excel**
2. Créer un nouveau classeur vierge
3. Enregistrer sous le nom : **`Planning_Guides_Musee.xlsm`**
   - ⚠️ Important : Le format doit être `.xlsm` (Excel avec macros)
   - File → Save As → Format : "Excel Macro-Enabled Workbook (.xlsm)"

---

### ÉTAPE 2 : Activer l'onglet Développeur

Si l'onglet "Développeur" n'est pas visible :

**Sur Windows :**
1. Fichier → Options
2. Personnaliser le ruban
3. Cocher "Développeur" dans la liste de droite
4. OK

**Sur Mac :**
1. Excel → Préférences
2. Ruban et barre d'outils
3. Cocher "Développeur"
4. Enregistrer

---

### ÉTAPE 3 : Ouvrir l'éditeur VBA

1. Cliquer sur l'onglet **Développeur**
2. Cliquer sur **Visual Basic** (ou appuyer sur `Alt + F11` sur Windows, `Option + F11` sur Mac)

Vous voyez maintenant l'**éditeur VBA** (Visual Basic Editor)

---

### ÉTAPE 4 : Importer les modules VBA

Pour chaque fichier `.bas` du dossier `vba-modules/`, suivre ces étapes :

1. Dans l'éditeur VBA, clic droit sur **"VBAProject (Planning_Guides_Musee.xlsm)"**
2. Sélectionner **Insertion → Module**
3. Un nouveau module vide s'affiche (Module1, Module2, etc.)
4. **Double-cliquer** sur le module pour l'ouvrir
5. **Ouvrir le fichier `.bas`** correspondant dans un éditeur de texte :
   - `Module_Config.bas`
   - `Module_Disponibilites.bas`
   - `Module_Planning.bas`
   - `Module_Emails.bas`
   - `Module_Calculs.bas`
   - `Module_Contrats.bas`
6. **Copier tout le contenu** du fichier
7. **Coller** dans la fenêtre du module VBA
8. Répéter pour chaque fichier `.bas`

**Résultat :** Vous devez avoir 6 modules au total dans votre projet VBA.

---

### ÉTAPE 5 : Initialiser les feuilles Excel

1. Revenir dans Excel (fermer l'éditeur VBA ou `Alt + F11`)
2. Onglet **Développeur** → Cliquer sur **Macros**
3. Sélectionner la macro : **`InitialiserApplication`**
4. Cliquer sur **Exécuter**

✅ **Résultat :** 7 feuilles sont automatiquement créées :
- Guides
- Disponibilités
- Visites
- Planning
- Calculs_Paie
- Contrats
- Configuration

---

### ÉTAPE 6 : Configurer les paramètres

1. Aller dans la feuille **"Configuration"**
2. Modifier les valeurs selon vos besoins :

| Paramètre | Valeur | Description |
|-----------|--------|-------------|
| Email_Expediteur | `votre@email.fr` | Email de l'association |
| Nom_Association | `Nom de votre asso` | Nom complet |
| Tarif_Heure | `50` | Tarif horaire en € |
| Notification_J7 | `OUI` | Activer notification J-7 |
| Notification_J1 | `OUI` | Activer notification J-1 |

---

### ÉTAPE 7 : Configuration Outlook (pour l'envoi d'emails)

#### Option A : Outlook installé localement (recommandé)

1. Outlook doit être installé et configuré avec votre compte
2. Tester l'envoi avec la macro : **`TestEnvoiEmail`**
   - Développeur → Macros → `TestEnvoiEmail` → Exécuter
   - Entrer votre email de test
   - Vérifier que l'email s'affiche
   - Envoyer

#### Option B : Problèmes avec Outlook

Si Outlook n'est pas disponible :
- Modifier le code pour utiliser une autre méthode (Gmail API, SMTP)
- Contacter le support technique

---

### ÉTAPE 8 : Remplir les données de base

#### 1. Feuille "Guides"

Ajouter vos guides (à partir de la ligne 2) :

| ID_Guide | Nom | Prénom | Email | Téléphone |
|----------|-----|--------|-------|-----------|
| G001 | DUPONT | Marie | marie.dupont@email.fr | 0601020304 |
| G002 | MARTIN | Pierre | pierre.martin@email.fr | 0605060708 |

#### 2. Feuille "Visites"

Ajouter les visites prévues :

| ID_Visite | Date | Heure_Debut | Heure_Fin | Musée | Type_Visite | Nombre_Visiteurs |
|-----------|------|-------------|-----------|-------|-------------|------------------|
| V001 | 15/12/2025 | 10:00 | 12:00 | Louvre | Guidée | 20 |
| V002 | 16/12/2025 | 14:00 | 16:00 | Orsay | Atelier | 15 |

---

### ÉTAPE 9 : Créer des boutons (optionnel mais recommandé)

Pour faciliter l'utilisation, créer des boutons sur une feuille "Accueil" :

1. Créer une nouvelle feuille appelée **"Accueil"**
2. Onglet **Développeur** → **Insérer** → **Bouton (Contrôle de formulaire)**
3. Dessiner le bouton sur la feuille
4. Dans la boîte de dialogue, sélectionner la macro à associer
5. Nommer le bouton (exemple : "Saisir mes disponibilités")

**Boutons recommandés :**
- 📝 Saisir mes disponibilités → `SaisirDisponibilites`
- 📅 Générer le planning → `GenererPlanningAutomatique`
- 📧 Envoyer les plannings → `EnvoyerPlanningMensuel`
- 🔔 Envoyer notifications → `EnvoyerNotificationsAutomatiques`
- 💰 Calculer les salaires → `CalculerVisitesEtSalaires`
- 📄 Générer un contrat → `GenererContratGuide`

---

### ÉTAPE 10 : Sécurité et confidentialité

#### Protéger les feuilles sensibles

1. Clic droit sur l'onglet de la feuille **"Disponibilités"**
2. **Protéger la feuille...**
3. Cocher : "Sélectionner les cellules verrouillées" et "Sélectionner les cellules déverrouillées"
4. Définir un mot de passe
5. OK

Répéter pour les feuilles : Configuration, Calculs_Paie

---

### ÉTAPE 11 : Sauvegarder et tester

1. **Enregistrer le fichier** (`Ctrl + S`)
2. **Fermer Excel**
3. **Rouvrir le fichier**
4. Si demandé, **Activer les macros**

---

## ✅ Vérification de l'installation

Cocher chaque élément :

- [ ] Fichier `.xlsm` créé et enregistré
- [ ] 6 modules VBA importés
- [ ] 7 feuilles Excel créées automatiquement
- [ ] Configuration remplie (email, tarifs)
- [ ] Test d'envoi email réussi
- [ ] Données de test ajoutées (guides et visites)
- [ ] Boutons créés (optionnel)
- [ ] Feuilles protégées

---

## 🆘 Résolution des problèmes courants

### Problème : "Les macros sont désactivées"

**Solution :**
1. Fichier → Options → Centre de gestion de la confidentialité
2. Paramètres du Centre de gestion
3. Paramètres des macros
4. Sélectionner : "Activer toutes les macros"
5. OK et redémarrer Excel

### Problème : "Outlook n'est pas disponible"

**Solution :**
1. Vérifier qu'Outlook est installé
2. Ouvrir Outlook et configurer un compte
3. Réessayer le test d'envoi

### Problème : "Erreur lors de l'initialisation"

**Solution :**
1. Vérifier que tous les modules sont bien importés
2. Vérifier qu'il n'y a pas de fautes de frappe dans le code
3. Exécuter la macro `InitialiserApplication` à nouveau

### Problème : "Les disponibilités ne s'enregistrent pas"

**Solution :**
1. Vérifier que la feuille "Disponibilités" existe
2. Vérifier que les ID des guides existent dans la feuille "Guides"
3. Désactiver temporairement la protection de la feuille

---

## 📞 Support

Pour toute question ou problème :
- Consulter le Guide d'Utilisation
- Vérifier les commentaires dans le code VBA
- Contacter l'administrateur système

---

## 🔄 Mises à jour

**Version actuelle :** 1.0
**Date :** Novembre 2025

Pour mettre à jour le système :
1. Sauvegarder le fichier actuel
2. Copier les nouvelles versions des modules
3. Remplacer dans l'éditeur VBA

---

**✨ Installation terminée ! Vous êtes prêt à utiliser le système. ✨**

Passez maintenant au **Guide d'Utilisation** pour apprendre à utiliser chaque fonctionnalité.
