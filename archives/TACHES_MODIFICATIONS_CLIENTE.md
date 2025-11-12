# 📋 TÂCHES - MODIFICATIONS DEMANDÉES PAR LA CLIENTE

**Date:** 12 novembre 2025
**Projet:** Planning Guides Musée
**Estimation totale:** ~4h30
**Fichier cliente reçu:** ✅ Screenshot du planning actuel

---

## 📊 STRUCTURE DE SON PLANNING ACTUEL (reçu)

| Colonne | Nom | Exemple de données |
|---------|-----|-------------------|
| A | DATE | "samedi 1 novembre 2025" |
| B | HEURE DEBUT | "10:00" |
| C | HEURE FIN | "10:45" |
| D | NOMBRE DE PARTICIPANTS | "18" |
| E | TYPE DE PRESTATIONS | "VISITE CONTEE BRANLY", "HORS LES MURS", "VISIO", "EVENEMENT BRANLY" |
| F | NOM DE LA STRUCTURE | "Maison des seniors Bondy", "CY PARIS UNIVERSITE" |
| G | NIVEAU | "CP", "G-VC \"Afrique\"" |
| H | THEME | "femmes", "Primaire/CE1" |
| I | COMMENTAIRES | "01 43 11 80.. Email dans dossier" |

**Types de prestations identifiés:**
- VISITE CONTEE BRANLY → Tarif Branly
- VISITE CONTEE MARINE → Tarif Marine
- HORS LES MURS → Tarif Hors-les-murs
- VISIO → Tarif Visio
- EVENEMENT BRANLY → Tarif Événement

**Codes couleurs:** 🔴 Rouge (Hors-les-murs + Événement) | 🟣 Violet (Visio) | Noir (Visites standards)

---

## 🔴 **BLOQUANTS CRITIQUES** (nécessitent modifications)

### ✅ **TÂCHE 1: Adapter structure Visites pour correspondre à son planning** ⏱️ **2h**

**Problème:**
L'onglet Visites actuel ne correspond PAS DU TOUT à sa structure. Impossible de copier-coller son planning.

**SA structure vs NOTRE structure:**
- Elle: DATE (texte) | Nous: Date (format Excel)
- Elle: HEURE DEBUT + HEURE FIN (2 colonnes) | Nous: Heure (1 colonne) + Durée
- Elle: TYPE DE PRESTATIONS (crucial pour tarif) | Nous: Type_Visite (différent)
- Elle: NOM DE LA STRUCTURE (école/client) | Nous: Musée (nom du musée)
- Elle: NIVEAU + THEME + COMMENTAIRES | Nous: Rien

**Actions à faire:**

1. **Réorganiser complètement l'onglet Visites dans PLANNING.xlsm**

   **Nouvelles colonnes (ordre proche du sien):**
   - Colonne A: `ID_Visite` (auto-généré) - GARDER
   - Colonne B: `Date` (format date Excel) - GARDER mais adapter format affichage
   - Colonne C: `Heure_Debut` (HH:MM) - EXISTE DÉJÀ colonne K, déplacer en C
   - Colonne D: `Heure_Fin` (HH:MM) - EXISTE DÉJÀ colonne L, déplacer en D
   - Colonne E: `Nombre_Participants` (nombre) - EXISTE colonne N, renommer et déplacer
   - Colonne F: `Type_Prestation` (liste: VISITE CONTEE BRANLY / VISITE CONTEE MARINE / HORS LES MURS / VISIO / EVENEMENT BRANLY)
   - Colonne G: `Nom_Structure` (texte libre) - ex: "École Massenet", "Maison seniors Bondy"
   - Colonne H: `Niveau` (texte) - ex: "CP", "CE1", "G-VC Afrique"
   - Colonne I: `Theme` (texte) - ex: "femmes", "Orient", "Primaire/CE1"
   - Colonne J: `Commentaires` (texte long) - RENOMMER colonne Q
   - Colonne K: `Statut` (liste: Confirmée/En attente/Annulée) - GARDER colonne H
   - Colonne L: `Guide_Attribue` (texte) - GARDER colonne P
   - Colonne M: `Tarif` (calculé auto) - GARDER colonne O

   **Colonnes obsolètes à supprimer:**
   - ❌ Musée (colonne D) - Remplacé par Nom_Structure
   - ❌ Type_Visite (colonne E) - Remplacé par Type_Prestation
   - ❌ Durée_Heures (colonne F) - Calculé auto depuis Heure_Fin - Heure_Debut
   - ❌ Heure (colonne C) - Dédoublonné avec Heure_Debut2. **Adapter les macros VBA**
   - `Module_Planning.bas` : Mettre à jour les références de colonnes
   - `Module_Calculs.bas` : Lire la colonne Type_Prestation au lieu de deviner
   - `Module_Emails.bas` : Inclure les nouvelles colonnes dans les emails

3. **Tester**
   - Vérifier que GenererPlanningAutomatique fonctionne
   - Vérifier que les calculs de paie lisent bien Type_Prestation

---

### ✅ **TÂCHE 2: Corriger système de détection type visite** ⏱️ **30min**

**Problème:**
Actuellement, le système devine le type de visite (Standard/Branly/Hors-les-murs) depuis le nom de la visite. Pas fiable et source d'erreurs.

**Question cliente:**
*"comment le système reconnaitra que tel guide a fait un hors les murs ou événement ou simple visite ? et donc mettra automatiquement le bon montant de salaire ?"*

**Actions à faire:**

1. **Modifier `Module_Calculs.bas` fonction `IdentifierTypeVisite`**
   ```vba
   ' AVANT: Cherche dans le nom de la visite
   ' APRÈS: Lit directement la colonne S (Type_Prestation)
   ```

2. **Modifier `Module_Calculs.bas` fonction `CalculerTarifJournee`**
   - Si Type_Prestation = "STANDARD" → utiliser TARIF_1_VISITE, TARIF_2_VISITES, TARIF_3_VISITES
   - Si Type_Prestation = "HORSLEMURS" → utiliser TARIF_HORSLEMURS_1, TARIF_HORSLEMURS_2, TARIF_HORSLEMURS_3
   - Si Type_Prestation = "EVENEMENT" → utiliser TARIF_BRANLY_2H, TARIF_BRANLY_3H, TARIF_BRANLY_4H selon durée
   - Si Type_Prestation = "VISIO" → utiliser TARIF_VISIO

3. **Documenter dans le guide**
   - Expliquer comment remplir la colonne Type_Prestation
   - Expliquer l'impact sur les calculs de paie

---

### ✅ **TÂCHE 3: Configuration spécialisations guides** ⏱️ **30min**

**Problème:**
*"Il manque aussi la configuration de l'outil car tous les guides ne font pas toutes les visites."*

L'onglet Spécialisations existe mais n'est pas rempli et pas exploité par les macros.

**Actions à faire:**

1. **Remplir l'onglet Spécialisations**
   - Demander à la cliente la liste des spécialisations par guide
   - OU lui expliquer comment le remplir elle-même

2. **Modifier `Module_Planning.bas` fonction `GenererPlanningAutomatique`**
   - Ajouter un filtre sur les spécialisations
   - Ne proposer que les guides ayant la bonne spécialisation pour chaque visite

3. **Ajouter une colonne Specialisation_Requise dans Visites**
   - Pour indiquer quelle spécialisation est nécessaire
   - Faire le matching automatique

4. **Alternative simple:**
   - Si trop complexe, juste documenter comment assigner manuellement les guides selon leurs spécialisations

---

## 🟡 **AMÉLIORATIONS UX** (moyens)

### ✅ **TÂCHE 4: Améliorer interface disponibilités guide** ⏱️ **15min**

**Problèmes:**
- *"A quoi correspond le numéro de guide dans colonne A Mes dispos du compte guide ?"*
- *"comment le guide met-il des précisions comme libre jusqu'à 16h par exemple"*
- *"s'il n'a aucune dispo, comment le signifier"*

**Actions à faire:**

1. **Nettoyer onglet Mes_Disponibilites**
   - Supprimer colonne A (Guide = numéro, source de confusion)
   - Garder uniquement : Date | Disponible | Commentaire | Prénom | Nom

2. **Améliorer colonne Disponible**
   - Option 1: Liste déroulante (OUI / NON / PARTIEL)
   - Option 2: Garder OUI/NON et utiliser Commentaire pour précisions

3. **Colonne Commentaire**
   - Renommer en "Précisions" pour que ce soit plus clair
   - Exemples : "Libre jusqu'à 16h", "Seulement matin", "Pas disponible"

4. **Documenter**
   - Ajouter instructions claires dans le guide utilisateur
   - Créer une section FAQ sur les disponibilités

---

### ✅ **TÂCHE 5: Supprimer encart gênant sur onglet Visites** ⏱️ **5min**

**Problème:**
*"Colonne B un encart blanc apparaît en haut à droite indiquant : 'attribution automatique activée Ajoutez une visite (ID en colonne A) le guide sera assigné. cet encart me gêne car il cache les 3 premières lignes de la colonne B"*

**Actions à faire:**

1. **Ouvrir PLANNING.xlsm**
2. **Aller onglet Visites**
3. **Chercher et supprimer:**
   - Commentaire Excel (clic droit → Supprimer le commentaire)
   - OU Validation de données avec message
   - OU Note/Post-it

4. **Vérifier** que rien n'est caché dans les 3 premières lignes

---

## 🟢 **FACILES / À DOCUMENTER** (pas de dev)

### ✅ **TÂCHE 6: Documentation tarifs modifiables** ⏱️ **5min**

**Question:**
*"est-ce que je peux modifier titres de la colonne A à partir de la ligne 12 des tarifs ? et aussi colonne C ?"*

**Réponse à lui donner:**

Oui, vous pouvez modifier :
- **Colonne A (Paramètre)** : Vous pouvez renommer (ex: changer "TARIF_1_VISITE" en "TARIF_VISITE_SIMPLE")
- **Colonne B (Valeur)** : Les montants en euros
- **Colonne C (Description)** : Les descriptions pour votre compréhension

⚠️ **Attention:** Si vous renommez colonne A, il faut aussi modifier les références dans le code VBA (Module_Calculs.bas).

**Recommandation:** Modifier seulement colonne B (valeurs) et C (descriptions). Ne pas toucher colonne A sauf si nécessaire.

---

### ✅ **TÂCHE 7: Documentation durées 45min** ⏱️ **2min**

**Question:**
*"Les visites qui durent 45 minutes sont payées le même tarif que celles d'1H (donc on peut toutes mettre 1 dans la colonne f ?)"*

**Réponse à lui donner:**

Oui, exactement. Si les visites de 45 minutes sont payées comme celles d'1h :
- Mettez simplement `1` dans la colonne F (Durée_Heures)
- Le système calculera le salaire en fonction du nombre de visites par jour, pas de la durée exacte

Si vous voulez quand même distinguer, vous pouvez mettre `0.75` (45 min = 0,75h) mais ça n'affectera pas les calculs de paie actuellement.

---

### ✅ **TÂCHE 8: Clarifier onglet Disponibilités admin** ⏱️ **5min**

**Question:**
*"dans ADMIN pourquoi y a -t-il l'onglet DISPONIBILITES avec la mention à remplir ?"*

**Réponse à lui donner:**

Il y a deux onglets pour les disponibilités :

1. **"Disponibilités"** (vue Admin)
   - C'est la base de données CENTRALISÉE de toutes les disponibilités
   - Visible par l'admin
   - Rempli automatiquement quand les guides saisissent leurs dispos

2. **"Mes_Disponibilites"** (vue Guide)
   - C'est la vue PERSONNELLE de chaque guide
   - Chaque guide voit uniquement SES propres disponibilités
   - C'est ici qu'ils saisissent

**Vous n'avez PAS à remplir "Disponibilités" manuellement.** Les guides le font via leur interface, et ça se remplit automatiquement dans cet onglet.

---

## 📊 **SYNTHÈSE ESTIMATIONS**

| Catégorie | Tâches | Temps total |
|-----------|--------|-------------|
| 🔴 Bloquants critiques | 3 tâches | 2h30 |
| 🟡 Améliorations UX | 2 tâches | 20min |
| 🟢 Documentation | 3 tâches | 12min |
| ⚙️ Tests finaux | - | 30min |
| **TOTAL** | **8 tâches** | **~3h15** |

---

## 📝 **ORDRE D'EXÉCUTION RECOMMANDÉ**

1. ✅ **Tâche 5** (5min) - Supprimer encart gênant → Quick win
2. ✅ **Tâche 1** (1h30) - Ajouter colonnes Visites → Bloquant principal
3. ✅ **Tâche 2** (30min) - Corriger détection type visite → Lié à Tâche 1
4. ✅ **Tâche 4** (15min) - Améliorer interface disponibilités → UX important
5. ✅ **Tâche 3** (30min) - Config spécialisations → Peut être reporté
6. ✅ **Tâches 6-8** (12min) - Documentation → Réponses par email
7. ⚙️ **Tests complets** (30min) - Vérifier que tout fonctionne

---

## 🎯 **DÉCISION À PRENDRE**

### **Scénario A : Facturation supplémentaire**
- Temps de dev : ~3h15
- Tarif proposé : 100-150€
- Justification : "Adaptations structure hors scope initial"

### **Scénario B : Geste commercial**
- Faire les modifs gratuitement
- Clôturer définitivement le projet après
- Conditions : plus aucune demande après ça

### **Scénario C : Compromis**
- Faire Tâches 1, 2, 4, 5 gratuitement (les critiques, 2h15)
- Tâche 3 (spécialisations) en option payante
- Tâches 6-8 (documentation) → réponses par email

---

## 📧 **TEMPLATES MESSAGES**

### **Si facturation (Scénario A)**

*"Bonjour,*

*J'ai analysé vos retours. Certains points sont des questions de documentation que je peux clarifier immédiatement.*

*D'autres nécessitent des adaptations de la structure (ajout de colonnes pour nom groupe, niveau, thème, etc.) car votre planning utilise un format différent du système livré.*

*Voici ce que je propose :*

**Modifications structure + adaptations macros (3h de dev) : 100€**
- Ajout colonnes Nom_Groupe, Niveau, Theme, Type_Prestation
- Adaptation calculs automatiques selon type de visite
- Interface disponibilités améliorée

**OU juste import de vos données dans la structure actuelle : Gratuit**
- Je prends votre fichier Excel
- J'importe vos visites dans le système
- Sans modifier la structure

*Que préférez-vous ?*

*Cordialement*"

---

### **Si geste commercial (Scénario B)**

*"Bonjour,*

*J'ai bien compris vos besoins. Je vais adapter le système pour qu'il colle exactement à votre structure de planning.*

*Je fais les modifications nécessaires (ajout colonnes, adaptation macros) et je vous renvoie le fichier complet d'ici 48h.*

*Après cette livraison, le projet sera définitivement clôturé.*

*Cordialement*"

---

### **Si compromis (Scénario C)**

*"Bonjour,*

*Je vais faire les adaptations critiques pour que vous puissiez importer votre planning (ajout colonnes, correction calculs).*

*Pour la configuration avancée des spécialisations par guide, cela nécessite un paramétrage spécifique qui peut être fait ultérieurement si besoin (prestation supplémentaire).*

*Je vous renvoie le fichier adapté sous 48h.*

*Cordialement*"

---

## ✅ **CHECKLIST AVANT LIVRAISON FINALE**

- [ ] Colonnes ajoutées dans Visites (Nom_Groupe, Niveau, Theme, Type_Prestation)
- [ ] Module_Calculs.bas adapté pour lire Type_Prestation
- [ ] Module_Planning.bas mis à jour (références colonnes)
- [ ] Encart gênant supprimé
- [ ] Interface Mes_Disponibilites nettoyée
- [ ] Tests complets effectués :
  - [ ] Import de données test
  - [ ] Génération planning automatique
  - [ ] Calculs de paie corrects
  - [ ] Emails fonctionnels
- [ ] Documentation mise à jour
- [ ] Réponses aux questions envoyées
- [ ] Fichier PLANNING.xlsm final envoyé
- [ ] Projet clôturé sur Malt

---

**Fin du document**
