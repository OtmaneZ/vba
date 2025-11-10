# 📋 DONNÉES À COMPLÉTER - PLANNING MUSÉE

**Fichier nettoyé créé :** `PLANNING_MUSEE_FINAL_PROPRE.xlsm`

---

## ✅ CE QUI EST DÉJÀ FAIT (95%)

### 1. **Guides** (15 guides importés)
- ✅ Noms et prénoms
- ✅ Adresses emails
- 🔴 **MANQUE :** Tarifs horaires + Mots de passe

### 2. **Types de Visites** (80 visites configurées)
- ✅ Noms des visites
- ✅ Musées
- ✅ Durées
- ✅ Catégories (Groupe/Individuel/Événement/Hors-les-murs)
- ✅ Barèmes tarifaires (Standard/BRANLY Event/BRANLY Hors-les-murs)
- ✅ Codes couleurs automatiques
- 🔴 **MANQUE :** Dates et heures programmées

### 3. **Système VBA** (100% fonctionnel)
- ✅ Authentification guides + admin
- ✅ Plannings personnalisés
- ✅ Confirmation/Refus de visites
- ✅ Réattribution automatique
- ✅ Codes couleurs par spécialisation
- ✅ Export PDF
- ✅ Calculs de paie (en attente validation tarifs)

---

## 🔴 À COMPLÉTER OBLIGATOIREMENT

### **ONGLET "GUIDES"** (15 lignes à compléter)

| Colonne | Nom | Exemple | Obligatoire |
|---------|-----|---------|-------------|
| E | **Tarif_Horaire** | `30` ou `35` | ✅ OUI |
| F | **Mot_De_Passe** | `guide123` | ✅ OUI |

**Instructions :**
1. Ouvrir l'onglet "Guides"
2. Pour chaque guide (lignes 2 à 16) :
   - Colonne E : Saisir le tarif horaire (en euros)
   - Colonne F : Choisir un mot de passe (le guide l'utilisera pour se connecter)

**⚠️ Important :** Sans ces données, les guides ne pourront pas se connecter !

---

### **ONGLET "VISITES"** (80 lignes à compléter)

| Colonne | Nom | Exemple | Obligatoire |
|---------|-----|---------|-------------|
| B | **Date** | `15/12/2025` | ✅ OUI (si visite programmée) |
| C | **Heure** | `14:00` | ✅ OUI (si visite programmée) |
| G | **Nombre_Visiteurs** | `20` | ⚠️ Recommandé |

**Instructions :**
1. Ouvrir l'onglet "Visites"
2. Pour chaque visite programmée :
   - Colonne B : Date au format JJ/MM/AAAA
   - Colonne C : Heure au format HH:MM
   - Colonne G : Nombre de visiteurs attendus

**💡 Astuce :** Vous pouvez laisser certaines visites sans date (= types de visites disponibles mais pas encore programmés)

---

### **ONGLET "DISPONIBILITÉS"** (à créer entièrement)

**Structure :**

| Colonne A | Colonne B | Colonne C | Colonne D |
|-----------|-----------|-----------|-----------|
| Guide | Date | Disponible | Commentaire |
| Sophie Durand | 15/12/2025 | OUI | |
| Marc Martin | 15/12/2025 | NON | Congé |

**Instructions :**
1. Chaque guide doit renseigner ses disponibilités
2. Format date : JJ/MM/AAAA
3. Disponible : OUI ou NON
4. Commentaire optionnel (ex: "Congé", "Occupé")

**⚠️ Important :** Sans disponibilités, l'attribution automatique ne fonctionnera pas !

---

### **ONGLET "CONFIGURATION"** (3 paramètres à modifier)

| Paramètre | Valeur actuelle (TEST) | À remplacer par |
|-----------|------------------------|-----------------|
| **Email_Expediteur** | admin@musee.fr | Votre email professionnel |
| **Nom_Association** | Musée des Guides | Nom réel de votre association |
| **MotDePasseAdmin** | admin123 | Mot de passe sécurisé pour l'admin |

**Instructions :**
1. Ouvrir l'onglet "Configuration"
2. Colonne B : Remplacer les valeurs test par vos vraies données
3. **⚠️ Notez bien le mot de passe admin !**

---

## ⚠️ DONNÉES DÉJÀ SUPPRIMÉES (étaient fausses)

- ❌ 180 disponibilités fictives (novembre 2025)
- ❌ Tarifs horaires test (28-35€)
- ❌ Mots de passe test (`guide123`)
- ❌ Email test (`admin@musee.fr`)

---

## 📊 RÉCAPITULATIF

| Élément | Statut | À faire |
|---------|--------|---------|
| Noms guides | ✅ OK | - |
| Emails guides | ✅ OK | - |
| Tarifs guides | 🔴 VIDE | **15 à remplir** |
| Mots de passe guides | 🔴 VIDE | **15 à remplir** |
| Types visites | ✅ OK | - |
| Dates visites | 🔴 VIDE | **À remplir selon besoin** |
| Disponibilités | 🔴 VIDE | **À créer** |
| Configuration | 🟡 TEST | **3 à modifier** |
| Code VBA | ✅ OK | - |

---

## 🎯 ORDRE DE PRIORITÉ

### **AVANT LA DÉMO CLIENT :**
1. ✅ Vérifier les 15 guides (noms/emails corrects)
2. ✅ Vérifier les 80 types de visites
3. 🔴 Remplir 3 paramètres Configuration
4. 🔴 Remplir tarifs des guides (colonne E)

### **POUR UTILISATION RÉELLE :**
5. 🔴 Créer mots de passe guides (colonne F)
6. 🔴 Saisir disponibilités de tous les guides
7. 🔴 Programmer les visites (dates/heures)
8. ⚠️ Valider les 3 barèmes tarifaires avec le client

---

## 📞 QUESTIONS POUR LE CLIENT

### **Tarification (URGENT)**
1. Quelles visites utilisent le barème "Standard" ? (actuellement : 77 visites)
2. Quelles visites utilisent "BRANLY Event" ? (actuellement : 3 visites)
3. Quelles visites utilisent "BRANLY Hors-les-murs" ? (actuellement : 0 visite)

**💡 Aide :** Voir onglet "Visites" colonne I et J pour les catégories actuelles

---

## 🚀 FICHIERS DISPONIBLES

1. **PLANNING_MUSEE_FINAL_PROPRE.xlsm** ← Fichier nettoyé prêt à compléter
2. **PLANNING_MUSEE_FINAL_COMPLET.xlsm** ← Ancien fichier avec données test (sauvegarde)
3. **PLANNING_MUSEE_DEMO_V2.xlsx** ← Fichier Excel sans macros (référence données)

---

**Date du nettoyage :** 10 novembre 2025  
**Temps estimé pour compléter :** 2-3 heures  
**Prêt pour démo :** ✅ OUI (avec données test visibles)  
**Prêt pour production :** 🔴 NON (compléter d'abord les données ci-dessus)
