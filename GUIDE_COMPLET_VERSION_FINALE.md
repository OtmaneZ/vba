# 🎉 SYSTÈME PLANNING - VERSION COMPLÈTE

## ✅ CORRECTIONS APPLIQUÉES

### 1️⃣ **Planning Automatique** (Problèmes résolus)
- ✅ **Colonne HEURE** : Affiche maintenant `10:30` au lieu de `0.4375`
- ✅ **Colonne DATE** : Format uniforme `16/11/2025`
- ✅ **Colonne GUIDES_DISPONIBLES** : Liste complète des guides disponibles
- ✅ **Feuille SPÉCIALISATIONS** : Ne disparaît plus
- ✅ **Colonnes réorganisées** : Type_Visite, Nb_Participants, Durée corrects

---

## 🆕 NOUVELLES FONCTIONNALITÉS AJOUTÉES

### Interface Admin - 6 Boutons

**LIGNE 1 - Gestion Planning :**
1. **[X] Déconnexion Admin** - Se déconnecter
2. **[!] Refuser et Réattribuer** - Changer guide assigné
3. **[+] Générer Planning** - Créer planning automatique

**LIGNE 2 - Paie & Export :**
4. **[$] Calculer Paie Mois** - Calcule salaires de tous les guides
5. **[=] Fiche Paie Guide** - Génère fiche individuelle
6. **[PDF] Export Planning** - Archive planning en PDF

---

## 📋 MODULES VBA À IMPORTER

### 🔴 OBLIGATOIRES (3 modules)
```
1. Module_Config.bas
2. Module_Calculs.bas
3. Module_Authentification.bas (MODIFIÉ - 6 boutons)
```

### ✅ DÉJÀ IMPORTÉS (2 modules)
```
4. Module_Planning_CORRECTED.bas (renommer en Module_Planning)
5. Module_Specialisations_CORRECTED.bas (renommer en Module_Specialisations)
```

---

## 📝 PROCÉDURE D'IMPORT

### Étape 1 : Ouvrir l'éditeur VBA
- Mac : `Option + F11`
- Windows : `Alt + F11`

### Étape 2 : Réimporter Module_Authentification (MODIFIÉ)

**⚠️ IMPORTANT : Ce module a été mis à jour avec les nouveaux boutons**

1. Supprimer l'ancien `Module_Authentification` :
   - Clic droit sur le module → **Supprimer**

2. Importer le nouveau :
   - Clic droit sur `VBAProject (PLANNING.xlsm)`
   - **Fichier** → **Importer un fichier...**
   - Sélectionner `vba-modules/Module_Authentification.bas`

### Étape 3 : Importer les nouveaux modules

**Si pas déjà importés, importer :**

3. `Module_Config.bas` (Constantes globales)
4. `Module_Calculs.bas` (Calculs paie - 883 lignes)

### Étape 4 : Sauvegarder
- `Ctrl+S` (Windows) ou `Cmd+S` (Mac)
- Fermer l'éditeur VBA

---

## 🚀 UTILISATION

### Workflow Complet

#### 1️⃣ **PRÉPARER** (Saisie données)
- Aller dans `Disponibilites` : Saisir dispos guides
- Aller dans `Visites` : Importer visites depuis emails
- Aller dans `Specialisations` : Vérifier qui fait quoi

#### 2️⃣ **GÉNÉRER PLANNING**
- Se connecter en tant qu'Admin
- Cliquer **[+] Générer Planning**
- Vérifier résultats dans feuille `Planning`

#### 3️⃣ **CALCULER PAIE** 💰
- Cliquer **[$] Calculer Paie Mois**
- Entrer mois (ex: `12/2025` pour décembre)
- Résultats dans feuille `Calculs_Paie`

#### 4️⃣ **FICHE PAIE INDIVIDUELLE**
- Cliquer **[=] Fiche Paie Guide**
- Entrer ID guide (ex: `HANAKO DANJO`)
- Entrer mois (ex: `12/2025`)
- Fiche générée dans nouveau fichier Excel

#### 5️⃣ **EXPORTER PDF**
- Cliquer **[PDF] Export Planning**
- Fichier PDF créé sur le Bureau

---

## 💰 CALCUL DE PAIE

### Système automatique par JOURNÉE

Le système calcule automatiquement selon le nombre de visites **le même jour** :

#### Visites Standards (45min)
- 1 visite/jour = 80 €
- 2 visites/jour = 110 €
- 3+ visites/jour = 140 €

#### Événements Branly (selon durée)
- 2 heures = 120 €
- 3 heures = 150 €
- 4 heures = 180 €

#### Hors-les-murs (déplacements)
- 1 visite/jour = 100 €
- 2 visites/jour = 130 €
- 3+ visites/jour = 160 €

**Le calcul est 100% automatique** basé sur les données du Planning.

---

## 📊 FEUILLES UTILISÉES

| Feuille | Utilité | Accès |
|---------|---------|-------|
| **Disponibilites** | Saisir dispos guides | Tout le monde |
| **Visites** | Importer visites | Admin |
| **Planning** | Planning généré | Admin + Guides |
| **Calculs_Paie** | Salaires calculés | Admin uniquement |
| **Specialisations** | Qui fait quoi | Admin |
| **Guides** | Liste guides | Admin |

---

## ⚠️ NOTES IMPORTANTES

### Calcul Paie
- Le calcul se fait sur les visites **confirmées et effectuées**
- Basé sur la colonne `Guide_Attribué` du Planning
- Groupement automatique par journée

### Export PDF
- Fichier créé sur le **Bureau**
- Nom : `Planning_Export_[Date].pdf`

### Fiche Paie Guide
- Crée un **nouveau fichier Excel**
- Contient : détail journées, total visites, montant total
- Nom : `FichePaie_[Guide]_[Mois].xlsx`

---

## 🆘 AIDE

### Les boutons ne s'affichent pas
→ Vérifier que vous êtes connecté en tant qu'**Admin**

### Erreur "Feuille non trouvée"
→ Vérifier que `Module_Config.bas` est bien importé

### Calcul paie incorrect
→ Vérifier les dates dans colonne Date du Planning
→ Vérifier que Guide_Attribué est rempli

### Export PDF ne fonctionne pas
→ Vérifier les droits d'écriture sur le Bureau

---

## 📞 RÉSUMÉ

**Vous avez maintenant :**
- ✅ Génération planning automatique
- ✅ Calcul salaires automatique
- ✅ Génération fiches de paie
- ✅ Export PDF
- ✅ Toutes les heures/dates correctes
- ✅ Guides disponibles affichés

**Prêt pour les plannings de décembre ! 🎄**

