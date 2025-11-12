# 📊 ANALYSE D'IMPACT - RESTRUCTURATION VISITES

**Date:** 12 novembre 2025
**Branche:** update-client
**Durée estimée:** 6h

---

## 🎯 OBJECTIF

Adapter PLANNING.xlsm pour qu'il corresponde exactement au planning actuel de la cliente.

---

## 📋 PARTIE 1 : CE QU'ON A ACTUELLEMENT

### Onglet Visites - Structure actuelle (17 colonnes)

| Col | Nom | Usage actuel | Garder? |
|-----|-----|-------------|---------|
| A | ID_Visite | V0001, V0002... | ✅ OUI |
| B | Date | Format date Excel | ✅ OUI (adapter affichage) |
| C | Heure | UNE seule heure | ❌ NON (doublon avec K) |
| D | Musée | "Musée du Quai Branly" | ❌ NON (remplacer par Structure) |
| E | Type_Visite | "Visite guidée", "Atelier" | ❌ NON (remplacer par Type_Prestation) |
| F | Durée_Heures | 1, 1.5, 2 | ⚠️  CALCULER (Heure_Fin - Heure_Debut) |
| G | Nombre_Visiteurs | 15, 20, 30 | ⚠️  RENOMMER (Nb_Participants) |
| H | Statut | "Confirmée", "Annulée" | ✅ OUI (déplacer) |
| K | Heure_Debut | "10:00", "14:30" | ✅ OUI (déplacer en C) |
| L | Heure_Fin | "11:30", "16:00" | ✅ OUI (déplacer en D) |
| M | Langue | "Français", "Anglais" | ⚠️  Optionnel (déplacer fin) |
| N | Nb_Personnes | DOUBLON de G | ❌ SUPPRIMER |
| O | Tarif | Calculé auto | ✅ OUI (déplacer) |
| P | Guide_Attribue | "Marie Dupont" | ✅ OUI (déplacer) |
| Q | Notes | Commentaires | ✅ OUI (renommer Commentaires) |

### Onglet Specialisations - Structure actuelle

| Col | Nom | État |
|-----|-----|------|
| A | ID_Specialisation | Vide |
| B | Guide | Vide |
| C | Specialisation | Vide |

**⚠️ PROBLÈME:** Structure inadaptée, il faut refaire.

---

## 📋 PARTIE 2 : CE QUE VEUT LA CLIENTE

### Son planning actuel (9 colonnes essentielles)

| Col | Nom | Exemples | Critique? |
|-----|-----|----------|-----------|
| A | DATE | "samedi 1 novembre 2025" | ✅ |
| B | HEURE DEBUT | "10:00", "14:30" | ✅ |
| C | HEURE FIN | "10:45", "18:30" | ✅ |
| D | NOMBRE DE PARTICIPANTS | "18", "5" | ✅ |
| E | **TYPE DE PRESTATIONS** | "VISITE CONTEE BRANLY", "HORS LES MURS", "VISIO", "EVENEMENT BRANLY" | 🔴 CRUCIAL |
| F | NOM DE LA STRUCTURE | "Maison des seniors Bondy", "École Massenet" | 🔴 CRUCIAL |
| G | NIVEAU | "CP", "CE1", "G-VC Afrique" | ✅ |
| H | THEME | "femmes", "Orient", "Primaire/CE1" | ✅ |
| I | COMMENTAIRES | "01 43 11 80.. Email", "contact : Karine" | ✅ |

### Types de prestations (détermine TARIF automatique)

1. **VISITE CONTEE BRANLY** → Tarif Branly (TARIF_BRANLY_2H, _3H, _4H)
2. **VISITE CONTEE MARINE** → Tarif Marine (TARIF_MARINE)
3. **HORS LES MURS** → Tarif Hors-les-murs (TARIF_HORSLEMURS_1, _2, _3)
4. **VISIO** → Tarif Visio (TARIF_VISIO)
5. **EVENEMENT BRANLY** → Tarif Événement (TARIF_BRANLY selon durée)

---

## 🔄 PARTIE 3 : NOUVELLE STRUCTURE (à implémenter)

### Onglet Visites - Structure finale (14 colonnes)

| Col | Nom | Type | Source | Notes |
|-----|-----|------|--------|-------|
| A | ID_Visite | Auto | Existant | V0001, V0002... |
| B | Date | Date | Existant | Format date Excel |
| C | Heure_Debut | Heure | Déplacer col K | HH:MM |
| D | Heure_Fin | Heure | Déplacer col L | HH:MM |
| E | Nb_Participants | Nombre | Renommer col G | Nombre de personnes |
| F | **Type_Prestation** | Liste | 🆕 NOUVEAU | 5 types (voir ci-dessus) |
| G | Nom_Structure | Texte | 🆕 NOUVEAU | Client/École/Institution |
| H | Niveau | Texte | 🆕 NOUVEAU | CP, CE1, etc. |
| I | Theme | Texte | 🆕 NOUVEAU | femmes, Orient, etc. |
| J | Commentaires | Texte | Renommer col Q | Notes diverses |
| K | Statut | Liste | Déplacer col H | Confirmée/Annulée |
| L | Guide_Attribue | Texte | Déplacer col P | Nom du guide |
| M | Tarif | Calculé | Déplacer col O | Auto calculé |
| N | Duree_Heures | Calculé | Formule | =D-C converti en heures |

**Validation colonne F (Type_Prestation):**
```
Liste déroulante: VISITE CONTEE BRANLY, VISITE CONTEE MARINE, HORS LES MURS, VISIO, EVENEMENT BRANLY
```

---

## ⚙️ PARTIE 4 : MODULES VBA IMPACTÉS

### 🔴 CRITIQUE - À MODIFIER OBLIGATOIREMENT

#### 1. **Module_Calculs.bas** (866 lignes) - IMPACT MAJEUR ⏱️ 1h

**Fonction `IdentifierTypeVisite` (ligne ~450)**
```vba
' AVANT
Function IdentifierTypeVisite(nomVisite As String) As String
    If InStr(LCase(nomVisite), "branly") > 0 Then
        IdentifierTypeVisite = "BRANLY"
    ...
End Function

' APRÈS
Function IdentifierTypeVisite(typePrestation As String) As String
    ' Lit directement colonne F (Type_Prestation)
    Select Case UCase(typePrestation)
        Case "VISITE CONTEE BRANLY"
            IdentifierTypeVisite = "BRANLY"
        Case "VISITE CONTEE MARINE"
            IdentifierTypeVisite = "MARINE"
        Case "HORS LES MURS"
            IdentifierTypeVisite = "HORSLEMURS"
        Case "VISIO"
            IdentifierTypeVisite = "VISIO"
        Case "EVENEMENT BRANLY"
            IdentifierTypeVisite = "EVENEMENT"
        Case Else
            IdentifierTypeVisite = "STANDARD"
    End Select
End Function
```

**Fonction `CalculerVisitesEtSalaires` (ligne ~50-200)**

Références colonnes à adapter:
```vba
' AVANT → APRÈS
wsPlanning.Cells(i, 2) ' Date → Cells(i, 2) ✅ OK (pas de changement)
wsPlanning.Cells(i, 4) ' Musée → Cells(i, 7) ' Nom_Structure
wsPlanning.Cells(i, 5) ' Type_Visite → Cells(i, 6) ' Type_Prestation
wsPlanning.Cells(i, 7) ' Nombre_Visiteurs → Cells(i, 5) ' Nb_Participants
wsPlanning.Cells(i, 16) ' Guide → Cells(i, 12) ' Guide_Attribue
```

**Ligne par ligne à modifier:**
- Ligne 63: `guideID = Trim(wsPlanning.Cells(i, 5).Value)` → Vérifier quelle colonne
- Ligne 68: `dateVisite = CDate(wsPlanning.Cells(i, 2).Value)` → ✅ OK
- Lignes 200-250: Boucle lecture visites → Adapter toutes les références colonnes

#### 2. **Module_Planning.bas** (403 lignes) - IMPACT MAJEUR ⏱️ 45min

**Fonction `GenererPlanningAutomatique` (ligne ~50)**

Références colonnes Planning:
```vba
' Ligne où il lit les infos visites
' AVANT → APRÈS
.Cells(row, 2) ' Date → .Cells(row, 2) ✅ OK
.Cells(row, 3) ' Heure → .Cells(row, 3) ✅ Heure_Debut
.Cells(row, 4) ' Musée → .Cells(row, 7) ' Nom_Structure
.Cells(row, 5) ' Type_Visite → .Cells(row, 6) ' Type_Prestation
```

**Fonction `AssignerGuideAutomatiquement` (ligne ~180)**

Adapter lecture colonnes + ajouter filtre spécialisations:
```vba
' NOUVEAU CODE À AJOUTER
Dim typePrestation As String
typePrestation = wsVisites.Cells(row, 6).Value ' Colonne F

' Vérifier si guide autorisé pour ce type
If EstGuideAutorise(nomGuide, typePrestation) Then
    ' Assigner
End If
```

**🆕 NOUVELLE FONCTION À CRÉER: `EstGuideAutorise`**
```vba
Function EstGuideAutorise(nomGuide As String, typePrestation As String) As Boolean
    ' Vérifie dans onglet Specialisations
    ' Si onglet vide → tout le monde autorisé
    ' Sinon → filtrer selon table
End Function
```

#### 3. **Module_Emails.bas** (562 lignes) - IMPACT MOYEN ⏱️ 30min

**Fonctions envoi emails notifications**

Templates emails à enrichir:
```vba
' AJOUTER dans corps email:
body = body & "Client: " & ws.Cells(row, 7).Value & vbCrLf ' Nom_Structure
body = body & "Niveau: " & ws.Cells(row, 8).Value & vbCrLf ' Niveau
body = body & "Thème: " & ws.Cells(row, 9).Value & vbCrLf ' Theme
body = body & "Type: " & ws.Cells(row, 6).Value & vbCrLf ' Type_Prestation
```

### 🟡 MOYEN - À VÉRIFIER

#### 4. **Module_Contrats.bas** (665 lignes) ⏱️ 20min

Génération contrats PDF/Word → Adapter références colonnes

#### 5. **Module_DPAE.bas** (217 lignes) ⏱️ 15min

Export DPAE → Adapter références colonnes

#### 6. **Feuille_Visites.cls** (60 lignes) ⏱️ 10min

Événements feuille Visites (changements cellules, etc.)

### 🟢 FAIBLE - PAS D'IMPACT

- Module_Authentification.bas ✅ (login/logout guides)
- Module_CodeCouleur.bas ✅ (coloration cellules)
- Module_Config.bas ✅ (lecture Configuration)
- Module_Disponibilites.bas ✅ (gestion dispos)
- Module_Accueil.bas ✅ (page accueil)
- Feuille_Accueil.cls ✅
- Feuille_Mon_Planning.cls ✅
- ThisWorkbook.cls ✅

---

## 📝 PARTIE 5 : PLAN D'EXÉCUTION (ordre logique)

### Phase 1: Restructuration Excel (1h30)

**Étape 1.1: Backup (5min)**
```python
import shutil
shutil.copy('PLANNING.xlsm', 'PLANNING_BACKUP_AVANT_RESTRUCTURATION.xlsm')
```

**Étape 1.2: Script restructuration (30min)**

Script `phase1_restructurer_visites.py`:
```python
1. Ouvrir PLANNING.xlsm
2. Lire onglet Visites actuel
3. Créer nouvelles colonnes F, G, H, I (Type_Prestation, Nom_Structure, Niveau, Theme)
4. Déplacer colonnes existantes:
   - K (Heure_Debut) → C
   - L (Heure_Fin) → D
   - G (Nombre_Visiteurs) → E (renommer Nb_Participants)
   - H (Statut) → K
   - P (Guide_Attribue) → L
   - O (Tarif) → M
   - Q (Notes) → J (renommer Commentaires)
5. Supprimer colonnes obsolètes: C (Heure), D (Musée), E (Type_Visite), F (Durée_Heures), N (Nb_Personnes)
6. Ajouter validation liste colonne F (Type_Prestation)
7. Ajouter formule colonne N (Durée_Heures)
8. Formater colonnes (dates, heures, nombres)
9. Sauvegarder
```

**Étape 1.3: Tests manuels (15min)**
- Ouvrir PLANNING.xlsm
- Vérifier structure visuelle
- Tester liste déroulante Type_Prestation
- Vérifier formule Durée_Heures

**Étape 1.4: Initialiser Specialisations (30min)**

Script `phase1_initialiser_specialisations.py`:
```python
1. Ouvrir onglet Specialisations
2. Effacer structure actuelle
3. Créer nouvelle structure:
   - A: ID_Specialisation
   - B: Prenom_Guide
   - C: Nom_Guide
   - D: Type_Prestation (liste déroulante)
   - E: Autorise (OUI/NON)
4. Lire onglet Guides
5. Pour chaque guide, créer 5 lignes (1 par type de prestation)
6. Par défaut: Autorise = OUI (tous les guides font tout)
7. Sauvegarder
```

**Étape 1.5: Commit Git (10min)**
```bash
git add -A
git commit -m "Phase 1: Restructuration onglet Visites + Spécialisations initialisées"
git push
```

### Phase 2: Adaptation VBA (2h)

**Ordre d'adaptation (du plus critique au moins):**

1. **Module_Calculs.bas** (1h)
   - IdentifierTypeVisite
   - CalculerVisitesEtSalaires
   - Toutes références colonnes

2. **Module_Planning.bas** (30min)
   - GenererPlanningAutomatique
   - AssignerGuideAutomatiquement
   - Fonction EstGuideAutorise (nouvelle)

3. **Module_Emails.bas** (20min)
   - Templates emails

4. **Modules secondaires** (10min)
   - Module_Contrats.bas
   - Module_DPAE.bas
   - Feuille_Visites.cls

**Commit après chaque module modifié**

### Phase 3: Script import (45min)

`phase3_importer_planning_cliente.py`

### Phase 4: Corrections mineures (15min)

- Supprimer encart gênant
- Nettoyer Mes_Disponibilites

### Phase 5: Tests complets (30min)

### Phase 6: Documentation + Livraison (30min)

---

## ✅ RÉPONSE À TA QUESTION

### "Faudra-t-il changer les .bas et .cls ?"

**OUI, absolument ! Voici lesquels:**

#### 🔴 OBLIGATOIRE (3 modules critiques)
1. **Module_Calculs.bas** - Change 100% (fonction + références colonnes)
2. **Module_Planning.bas** - Change 80% (références colonnes + nouvelle fonction)
3. **Module_Emails.bas** - Change 30% (enrichir templates)

#### 🟡 RECOMMANDÉ (3 modules)
4. **Module_Contrats.bas** - Adapter références colonnes
5. **Module_DPAE.bas** - Adapter références colonnes
6. **Feuille_Visites.cls** - Vérifier événements

#### 🟢 PAS DE CHANGEMENT (10 modules)
- Module_Authentification.bas ✅
- Module_CodeCouleur.bas ✅
- Module_Config.bas ✅
- Module_Disponibilites.bas ✅
- Module_Accueil.bas ✅
- Module_Specialisations.bas ✅ (nouveau mais pas à modifier)
- Feuille_Accueil.cls ✅
- Feuille_Mon_Planning.cls ✅
- ThisWorkbook.cls ✅
- Module_Emails_SMTP.bas ✅

---

## 🎯 CONCLUSION

**Restructurer Visites = OUI, mais ça implique:**
- ✅ Modifier structure Excel (colonnes, validations)
- ✅ Adapter 3 modules VBA critiques
- ✅ Vérifier 3 modules VBA secondaires
- ✅ Initialiser onglet Specialisations

**C'est pour ça qu'on a estimé 1h30 pour Phase 1 (Excel) + 2h pour Phase 2 (VBA) = 3h30 total.**

**On commence quand tu veux ! 💪**
