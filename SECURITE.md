# 🔒 SÉCURITÉ ET CONTRÔLE D'ACCÈS

## 📊 État par défaut (avant connexion)

### ✅ Feuilles VISIBLES (tout le monde)
- **Accueil** : Page de connexion
- **Guides** : Liste des guides (lecture seule pour guides)
- **Disponibilites** : Saisie des disponibilités
- **Visites** : Visites programmées
- **Planning** : Planning général

### ❌ Feuilles MASQUÉES (xlSheetVeryHidden)
- **Calculs_Paie** : Calculs salaires (sensible !)
- **Contrats** : Génération contrats (sensible !)
- **Configuration** : Paramètres système (sensible !)

> **xlSheetVeryHidden** = Invisible même via clic droit "Afficher"

---

## 👤 Connexion GUIDE

### Ce qu'un guide peut voir :
- ✅ **Accueil** : Page de bienvenue
- ✅ **Guides** : Liste des collègues (lecture seule)
- ✅ **Disponibilites** : Saisir ses disponibilités
- ✅ **Visites** : Voir ses visites assignées uniquement
- ✅ **Planning** : Son planning personnel filtré

### Ce qu'un guide NE PEUT PAS voir :
- ❌ **Calculs_Paie** : Reste masqué
- ❌ **Contrats** : Reste masqué
- ❌ **Configuration** : Reste masqué

### Code activé :
```vb
niveauAcces = "GUIDE"
' Les feuilles sensibles restent masquées
' Filtrage automatique : ne voir que SES propres données
```

---

## 👑 Connexion ADMIN

### Ce qu'un admin peut voir :
- ✅ **TOUTES les feuilles** démasquées automatiquement
- ✅ **Calculs_Paie** : Visible + éditable
- ✅ **Contrats** : Visible + éditable
- ✅ **Configuration** : Visible + éditable
- ✅ Accès complet à toutes les données

### Code activé :
```vb
niveauAcces = "ADMIN"
Call AfficherToutesFeuillesAdmin()
' Toutes les feuilles deviennent xlSheetVisible
```

---

## 🛡️ Protection supplémentaire (optionnel)

### Protection VBA (empêcher modification du code)
1. Dans VBA : **Outils** → **Propriétés de VBAProject**
2. Onglet **Protection**
3. Cocher "Verrouiller le projet pour l'affichage"
4. Définir un mot de passe

### Protection des feuilles (empêcher modification cellules)
Ajouter dans `Module_Config.bas` après création des feuilles :
```vb
' Protéger les feuilles sensibles
ThisWorkbook.Sheets(FEUILLE_CALCULS).Protect Password:="admin123", UserInterfaceOnly:=True
ThisWorkbook.Sheets(FEUILLE_CONFIG).Protect Password:="admin123", UserInterfaceOnly:=True
```

### Masquer l'onglet VBA (utilisateurs normaux)
Dans le Registre Windows (avancé) ou désactiver l'accès à VBA dans Excel.

---

## ⚙️ Configuration actuelle

### Mot de passe admin
- **Fichier** : `Module_Authentification.bas`
- **Variable** : `mdpAdmin = "admin123"`
- **À changer** : Modifier ligne ~15

### Feuilles masquées par défaut
- **Fichier** : `Module_Config.bas`
- **Fonction** : `MasquerFeuillesSensibles()`
- Appelée automatiquement par `InitialiserApplication()`

---

## 🚨 Avertissements

### ⚠️ Limites de sécurité VBA
- Un utilisateur avec accès VBA peut **toujours** voir le code
- La protection VBA peut être contournée avec des outils
- Les mots de passe Excel ne sont **pas cryptés** de manière forte

### 🔐 Pour une vraie sécurité
Si données très sensibles :
1. **Backend séparé** (base de données SQL avec authentification)
2. **Application Web** avec serveur sécurisé
3. **Fichiers séparés** : un par guide (sans accès aux autres)

---

## ✅ Résumé

| Utilisateur | Feuilles visibles | Feuilles masquées | Droits |
|------------|------------------|-------------------|--------|
| **Visiteur** (non connecté) | Accueil | Toutes les autres | Aucun |
| **Guide** (connecté) | Accueil + 4 feuilles métier | 3 feuilles admin | Lecture/Écriture filtré |
| **Admin** (connecté) | **Toutes** (7 feuilles) | Aucune | Lecture/Écriture complet |

**Sécurité actuelle** : 🟡 Moyenne (suffisant pour usage interne)
**Sécurité recommandée** : 🔐 Ajouter protection VBA + mots de passe forts
