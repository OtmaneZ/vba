# 📥 GUIDE : Comment importer vos visites depuis votre planning Excel existant

**Date :** 11 novembre 2025
**Problème :** Vous avez déjà un fichier Excel avec vos visites planifiées et vous ne voulez pas les ressaisir à la main.

---

## 🎯 **3 SOLUTIONS POSSIBLES**

### ✅ **SOLUTION 1 : Copier-Coller Direct (LA PLUS SIMPLE)**

#### Étapes :

1. **Ouvrez votre fichier Excel existant** avec vos visites planifiées
2. **Sélectionnez les colonnes** correspondant à :
   - Date de la visite
   - Heure de début
   - Musée
   - Type de visite
   - Durée
   - Nombre de visiteurs

3. **Copiez** (Ctrl+C ou Cmd+C)

4. **Ouvrez PLANNING.xlsm**
5. **Allez sur l'onglet "Visites"**
6. **Cliquez sur la cellule B2** (première cellonne après ID_Visite, ligne 2)
7. **Collez** (Ctrl+V ou Cmd+V)

#### 📋 Correspondance des colonnes :

| Votre fichier | → | PLANNING.xlsm | Colonne |
|---------------|---|---------------|---------|
| Date visite | → | Date | B |
| Heure début | → | Heure | C |
| Nom musée | → | Musée | D |
| Type/Titre visite | → | Type_Visite | E |
| Durée (en heures) | → | Durée_Heures | F |
| Nb visiteurs | → | Nombre_Visiteurs | G |

#### ⚠️ Important :
- **Colonne A (ID_Visite)** : Ne touchez pas, elle contient déjà V001, V002, etc.
- **Colonne H (Statut)** : Laissez vide, sera rempli automatiquement
- Les autres colonnes se rempliront automatiquement

---

### ✅ **SOLUTION 2 : Script Python d'Import Automatique (RECOMMANDÉ)**

Je vous ai préparé un script qui fait tout automatiquement !

#### Étapes :

1. **Préparez votre fichier Excel** avec vos visites :
   - Nommez-le : `MES_VISITES_A_IMPORTER.xlsx`
   - Placez-le dans le même dossier que `PLANNING.xlsm`

2. **Structure minimale requise dans votre fichier** :
   ```
   Colonne A : Date (format JJ/MM/AAAA ou AAAA-MM-JJ)
   Colonne B : Heure (format HH:MM)
   Colonne C : Musée
   Colonne D : Type de visite
   Colonne E : Durée (en heures, ex: 1, 2, 3)
   Colonne F : Nombre de visiteurs
   ```

3. **Lancez le script** (je vous l'envoie séparément)

4. **Le script va :**
   - ✅ Lire votre fichier
   - ✅ Vérifier les données
   - ✅ Les copier dans PLANNING.xlsm onglet Visites
   - ✅ Générer les ID automatiquement (V001, V002, etc.)
   - ✅ Mettre le statut "Planifié"
   - ✅ Créer une sauvegarde avant modification

---

### ✅ **SOLUTION 3 : Macro VBA d'Import (Dans PLANNING.xlsm)**

Si vous préférez tout faire dans Excel sans Python :

#### J'ai ajouté une macro dans votre fichier :

1. **Ouvrez PLANNING.xlsm**
2. **Appuyez sur Alt+F11** (ou Option+F11 sur Mac) pour ouvrir VBA
3. **Exécutez la macro : `ImporterVisitesDepuisFichier`**

#### La macro vous demandera :
- Le chemin de votre fichier Excel
- L'onglet où sont vos visites
- La ligne où commencent les données

#### Elle importera automatiquement toutes vos visites !

---

## 🔧 **AIDE : Quelle colonne correspond à quoi ?**

### Dans PLANNING.xlsm, l'onglet Visites contient :

| Col | Nom | Description | Obligatoire |
|-----|-----|-------------|-------------|
| A | ID_Visite | Identifiant unique (V001, V002...) | ✅ Auto |
| B | Date | Date de la visite (JJ/MM/AAAA) | ✅ OUI |
| C | Heure | Heure de début (HH:MM) | ✅ OUI |
| D | Musée | Nom du musée | ✅ OUI |
| E | Type_Visite | Type/Titre de la visite | ✅ OUI |
| F | Durée_Heures | Durée en heures (1, 2, 3, 4...) | ✅ OUI |
| G | Nombre_Visiteurs | Nombre de personnes | ⚠️ Recommandé |
| H | Statut | À planifier / Planifié / Confirmé | 🤖 Auto |
| I+ | Autres | Remplis automatiquement | 🤖 Auto |

---

## 📧 **Formats acceptés**

### Dates :
- ✅ `15/12/2025`
- ✅ `2025-12-15`
- ✅ `15-12-2025`

### Heures :
- ✅ `14:00`
- ✅ `14h00`
- ✅ `14:30`

### Durées :
- ✅ `1` (1 heure)
- ✅ `2` (2 heures)
- ✅ `1.5` (1h30)
- ✅ `45 minutes` (converti en 0.75)

---

## 🆘 **En cas de problème**

### Problème : "Les dates ne s'affichent pas correctement"
**Solution :** Sélectionnez la colonne B → Clic droit → Format de cellule → Date → Choisir format JJ/MM/AAAA

### Problème : "J'ai des colonnes en plus/en moins"
**Solution :** Pas grave ! Copiez seulement les colonnes qui correspondent. Le reste se remplira automatiquement.

### Problème : "J'ai 200 visites à importer"
**Solution :** Utilisez la **Solution 2** (script Python) ou la **Solution 3** (macro VBA), c'est fait pour ça !

---

## ✅ **Après l'import, que se passe-t-il ?**

1. ✅ Vos visites sont dans l'onglet "Visites"
2. ✅ Vous pouvez lancer la macro **"GenererPlanningAutomatique"**
3. ✅ Le système va croiser avec les disponibilités des guides
4. ✅ Les guides seront attribués automatiquement
5. ✅ Les emails partiront automatiquement

---

## 💡 **Astuce Pro**

**Gardez votre fichier Excel original** comme référence, et utilisez PLANNING.xlsm uniquement pour :
- L'attribution des guides
- Les calculs de paie
- La génération de contrats
- L'envoi d'emails

Vous pouvez importer de nouvelles visites chaque mois avec la même méthode !

---

## 📞 **Besoin d'aide ?**

Si vous avez des difficultés, envoyez-moi :
1. Une capture d'écran de votre fichier Excel (les en-têtes)
2. Le nombre de visites à importer
3. La solution que vous préférez (1, 2 ou 3)

Je vous guiderai pas à pas ! 🎯
