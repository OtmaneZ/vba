# 🎯 PLAN D'ACTION - ADAPTATION AU PLANNING RÉEL DE LA CLIENTE

**Date:** 12 novembre 2025
**Fichier reçu:** ✅ Screenshot planning actuel
**Estimation totale:** ~4h30

---

## 📊 CE QU'ELLE A (Structure reçue)

| Col | Nom | Exemples |
|-----|-----|----------|
| A | DATE | "samedi 1 novembre 2025", "lundi 3 novembre 2025" |
| B | HEURE DEBUT | "10:00", "14:30" |
| C | HEURE FIN | "10:45", "18:30" |
| D | NOMBRE DE PARTICIPANTS | "18", "5" |
| E | **TYPE DE PRESTATIONS** | "VISITE CONTEE BRANLY", "HORS LES MURS", "VISIO", "EVENEMENT BRANLY" |
| F | NOM DE LA STRUCTURE | "Maison des seniors Bondy", "CY PARIS UNIVERSITE", "FOLIES" |
| G | NIVEAU | "CP", "CE1", "G-VC \"Afrique\"", "Visio Contée \"Contes des Amériques\"" |
| H | THEME | "femmes", "Primaire/CE1", "Orient" |
| I | COMMENTAIRES | "01 43 11 80.. Email dans dossier", "Responsable Local" |

**🎨 Codes couleurs:**
- 🔴 ROUGE: HORS LES MURS + EVENEMENT BRANLY
- 🟣 VIOLET: VISIO
- ⚪ NOIR/Normal: VISITE CONTEE BRANLY / MARINE

**💰 Types de prestations = Tarifs:**
1. VISITE CONTEE BRANLY → Tarif Branly (50€ + progressifs)
2. VISITE CONTEE MARINE → Tarif Marine (50€ + progressifs)
3. HORS LES MURS → Tarif Hors-les-murs (55€ + progressifs)
4. VISIO → Tarif Visio
5. EVENEMENT BRANLY → Tarif Événement Branly (selon durée)

---

## 📋 CE QU'ON A (PLANNING.xlsm actuel)

**Onglet Visites - Colonnes actuelles:**
- A: ID_Visite
- B: Date
- C: Heure (1 seule, pas début/fin)
- D: Musée (nom musée, pas client)
- E: Type_Visite (Visite guidée/Atelier - pas pareil que Type_Prestation)
- F: Durée_Heures (nombre, pas heures fin)
- G: Nombre_Visiteurs
- H: Statut
- K: Heure_Debut
- L: Heure_Fin
- M: Langue
- N: Nb_Personnes
- O: Tarif
- P: Guide_Attribue
- Q: Notes

**❌ PROBLÈMES:**
1. Colonne "Musée" ≠ "Nom_Structure" (client/école)
2. Colonne "Type_Visite" ≠ "Type_Prestation" (crucial pour tarif!)
3. Pas de colonnes: Niveau, Theme
4. Structure désorganisée (doublons: Heure/Heure_Debut, Nombre_Visiteurs/Nb_Personnes)

---

## 🔥 PLAN D'ACTION (dans l'ordre)

### **PHASE 1: RESTRUCTURATION ONGLET VISITES** ⏱️ **1h30**

#### Étape 1.1: Créer nouvelle structure propre

**Ordre des colonnes (proche de son planning):**

| Col | Nom | Type | Description |
|-----|-----|------|-------------|
| A | ID_Visite | Auto | V0001, V0002... |
| B | Date | Date | Format date Excel |
| C | Heure_Debut | Heure | Format HH:MM |
| D | Heure_Fin | Heure | Format HH:MM |
| E | Nb_Participants | Nombre | Nombre de personnes |
| F | **Type_Prestation** | Liste | **VISITE CONTEE BRANLY** / VISITE CONTEE MARINE / HORS LES MURS / VISIO / EVENEMENT BRANLY |
| G | Nom_Structure | Texte | Client/École/Institution |
| H | Niveau | Texte | CP, CE1, CE2, G-VC "Afrique", etc. |
| I | Theme | Texte | femmes, Orient, Primaire/CE1, etc. |
| J | Commentaires | Texte | Notes diverses |
| K | Statut | Liste | Confirmée / En attente / Annulée |
| L | Guide_Attribue | Texte | Nom du guide |
| M | Tarif | Nombre | Calculé auto |
| N | Durée_Heures | Calculé | (Heure_Fin - Heure_Debut) |

**Validation données colonne F (Type_Prestation):**
- Liste déroulante avec les 5 types exacts
- Valeur par défaut: "VISITE CONTEE BRANLY"
- Obligatoire (pas de cellule vide)

#### Étape 1.2: Modifier PLANNING.xlsm avec openpyxl

Python script `restructurer_visites.py`:
```python
# 1. Backup PLANNING.xlsm
# 2. Lire onglet Visites actuel
# 3. Réorganiser colonnes selon nouvelle structure
# 4. Ajouter validations (listes déroulantes)
# 5. Formater colonnes (dates, heures, nombres)
# 6. Sauvegarder
```

**Actions:**
- Renommer en-têtes colonnes
- Déplacer données existantes si possible
- Supprimer colonnes obsolètes
- Ajouter nouvelles colonnes vides (Nom_Structure, Niveau, Theme)
- Créer liste déroulante Type_Prestation

---

### **PHASE 2: ADAPTER TOUTES LES MACROS VBA** ⏱️ **2h**

#### Étape 2.1: Module_Calculs.bas (CRITIQUE)

**Fonction à réécrire: `IdentifierTypeVisite`**

AVANT (ligne ~450):
```vba
Function IdentifierTypeVisite(nomVisite As String) As String
    ' Devine depuis le nom de la visite
    If InStr(LCase(nomVisite), "branly") > 0 Then
        IdentifierTypeVisite = "BRANLY"
    ElseIf InStr(LCase(nomVisite), "marine") > 0 Then
        IdentifierTypeVisite = "MARINE"
    ...
End Function
```

APRÈS:
```vba
Function IdentifierTypeVisite(typePrestation As String) As String
    ' Lit directement depuis colonne F (Type_Prestation)
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

**Fonction à adapter: `CalculerVisitesEtSalaires`**
- Ligne ~200: Lire colonne F au lieu de E
- Ligne ~250: Appeler IdentifierTypeVisite avec Type_Prestation
- Ligne ~300: Adapter références colonnes (C→C, D→G, etc.)

**Toutes les références colonnes à mettre à jour:**
```vba
' AVANT → APRÈS
ws.Cells(i, 4) ' Musée → ws.Cells(i, 7) ' Nom_Structure
ws.Cells(i, 5) ' Type_Visite → ws.Cells(i, 6) ' Type_Prestation
ws.Cells(i, 7) ' Nombre_Visiteurs → ws.Cells(i, 5) ' Nb_Participants
ws.Cells(i, 17) ' Notes → ws.Cells(i, 10) ' Commentaires
```

#### Étape 2.2: Module_Planning.bas

**Fonction `GenererPlanningAutomatique` (ligne ~50)**
- Adapter toutes les références colonnes
- Lire Type_Prestation pour filtrer guides spécialisés

**Fonction `AssignerGuideAutomatiquement` (ligne ~180)**
- Mettre à jour lecture colonnes

#### Étape 2.3: Module_Emails.bas

**Templates emails (ligne ~100+)**
```vba
' Ajouter dans le corps email:
body = body & "Client: " & ws.Cells(row, 7).Value & vbCrLf ' Nom_Structure
body = body & "Niveau: " & ws.Cells(row, 8).Value & vbCrLf ' Niveau
body = body & "Thème: " & ws.Cells(row, 9).Value & vbCrLf ' Theme
body = body & "Type: " & ws.Cells(row, 6).Value & vbCrLf ' Type_Prestation
```

#### Étape 2.4: Module_Contrats.bas, Module_DPAE.bas

- Mettre à jour références colonnes dans génération contrats
- Adapter exports DPAE

#### Étape 2.5: TESTS UNITAIRES

Créer `test_nouveau_systeme.bas`:
```vba
Sub TestCalculsTarifs()
    ' Tester chaque type de prestation
    Debug.Print IdentifierTypeVisite("VISITE CONTEE BRANLY") ' → BRANLY
    Debug.Print IdentifierTypeVisite("HORS LES MURS") ' → HORSLEMURS
    Debug.Print IdentifierTypeVisite("VISIO") ' → VISIO
End Sub
```

---

### **PHASE 3: SCRIPT IMPORT PYTHON** ⏱️ **45min**

#### Étape 3.1: Créer `importer_planning_cliente.py`

```python
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import re

# Lire son fichier Excel
wb_source = load_workbook('PLANNING_CLIENTE.xlsx')
ws_source = wb_source.active

# Ouvrir PLANNING.xlsm
wb_dest = load_workbook('PLANNING.xlsm', keep_vba=True)
ws_dest = wb_dest['Visites']

next_id = 1
next_row = 2  # Ligne 1 = en-têtes

for row in range(2, ws_source.max_row + 1):
    # Lire ses données
    date_str = ws_source.cell(row, 1).value  # "samedi 1 novembre 2025"
    heure_debut = ws_source.cell(row, 2).value  # "10:00"
    heure_fin = ws_source.cell(row, 3).value  # "10:45"
    nb_participants = ws_source.cell(row, 4).value  # "18"
    type_presta = ws_source.cell(row, 5).value  # "VISITE CONTEE BRANLY"
    nom_structure = ws_source.cell(row, 6).value  # "École Massenet"
    niveau = ws_source.cell(row, 7).value  # "CP"
    theme = ws_source.cell(row, 8).value  # "femmes"
    commentaires = ws_source.cell(row, 9).value  # "..."

    # Parser date française → date Excel
    date_obj = parser_date_francaise(date_str)

    # Écrire dans PLANNING.xlsm
    ws_dest.cell(next_row, 1).value = f"V{next_id:04d}"  # ID_Visite
    ws_dest.cell(next_row, 2).value = date_obj  # Date
    ws_dest.cell(next_row, 3).value = heure_debut  # Heure_Debut
    ws_dest.cell(next_row, 4).value = heure_fin  # Heure_Fin
    ws_dest.cell(next_row, 5).value = int(nb_participants or 0)  # Nb_Participants
    ws_dest.cell(next_row, 6).value = type_presta  # Type_Prestation
    ws_dest.cell(next_row, 7).value = nom_structure  # Nom_Structure
    ws_dest.cell(next_row, 8).value = niveau  # Niveau
    ws_dest.cell(next_row, 9).value = theme  # Theme
    ws_dest.cell(next_row, 10).value = commentaires  # Commentaires
    ws_dest.cell(next_row, 11).value = "Confirmée"  # Statut

    # Calculer durée
    if heure_debut and heure_fin:
        duree = calculer_duree(heure_debut, heure_fin)
        ws_dest.cell(next_row, 14).value = duree  # Durée_Heures

    next_id += 1
    next_row += 1

# Sauvegarder
wb_dest.save('PLANNING.xlsm')
print(f"✅ {next_id-1} visites importées !")
```

#### Étape 3.2: Fonctions helper

```python
def parser_date_francaise(date_str):
    """Convertir 'samedi 1 novembre 2025' → date Excel"""
    mois_fr = {
        'janvier': 1, 'février': 2, 'mars': 3, 'avril': 4,
        'mai': 5, 'juin': 6, 'juillet': 7, 'août': 8,
        'septembre': 9, 'octobre': 10, 'novembre': 11, 'décembre': 12
    }

    # Regex: "samedi 1 novembre 2025"
    match = re.search(r'(\d+)\s+(\w+)\s+(\d{4})', date_str)
    if match:
        jour = int(match.group(1))
        mois_nom = match.group(2).lower()
        annee = int(match.group(3))
        mois = mois_fr.get(mois_nom, 1)
        return datetime(annee, mois, jour)
    return None

def calculer_duree(heure_debut, heure_fin):
    """Calculer durée en heures: '10:00' → '11:30' = 1.5"""
    # Parser heures
    h1, m1 = map(int, heure_debut.split(':'))
    h2, m2 = map(int, heure_fin.split(':'))

    minutes_total = (h2 * 60 + m2) - (h1 * 60 + m1)
    return round(minutes_total / 60, 2)
```

---

### **PHASE 4: CORRECTIONS MINEURES** ⏱️ **15min**

#### Tâche 4.1: Supprimer encart gênant (5min)

*"Colonne B un encart blanc apparaît en haut à droite"*

- Ouvrir PLANNING.xlsm
- Onglet Visites → chercher commentaire/note cellule B1-B3
- Supprimer

#### Tâche 4.2: Améliorer Mes_Disponibilites (10min)

*"A quoi correspond le numéro de guide dans colonne A ?"*

- Supprimer colonne A (ID_Guide numérique)
- Garder: Date | Disponible | Précisions | Prénom | Nom

---

### **PHASE 5: SYSTÈME SPÉCIALISATIONS GUIDES** ⏱️ **1h**

#### Problème identifié

*"Il manque aussi la configuration de l'outil car tous les guides ne font pas toutes les visites."*

Tous les guides n'ont pas les mêmes compétences:
- Guide A: VISITE CONTEE BRANLY + MARINE
- Guide B: HORS LES MURS uniquement
- Guide C: Toutes les visites
- etc.

Le système doit filtrer automatiquement les guides disponibles selon le type de visite.

#### Étape 5.1: Définir structure onglet Spécialisations (15min)

**Structure actuelle** (existe mais vide):
L'onglet `Specialisations` existe déjà mais n'est pas exploité.

**Nouvelle structure à implémenter:**

| Col | Nom | Type | Description |
|-----|-----|------|-------------|
| A | ID_Specialisation | Auto | S0001, S0002... |
| B | Prenom_Guide | Texte | Prénom du guide |
| C | Nom_Guide | Texte | Nom du guide |
| D | Type_Prestation | Liste | VISITE CONTEE BRANLY / MARINE / HORS LES MURS / VISIO / EVENEMENT BRANLY |
| E | Autorise | Oui/Non | OUI = le guide peut faire ce type |

**Validation colonne D:**
- Liste déroulante avec les 5 types de prestations
- Même liste que colonne F de Visites

**Exemple de données:**

| ID | Prénom | Nom | Type_Prestation | Autorisé |
|----|--------|-----|-----------------|----------|
| S0001 | Marie | Dupont | VISITE CONTEE BRANLY | OUI |
| S0002 | Marie | Dupont | VISITE CONTEE MARINE | OUI |
| S0003 | Marie | Dupont | HORS LES MURS | NON |
| S0004 | Pierre | Martin | VISITE CONTEE BRANLY | OUI |
| S0005 | Pierre | Martin | HORS LES MURS | OUI |
| S0006 | Pierre | Martin | VISIO | OUI |

**Script Python `initialiser_specialisations.py`:**
```python
import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

wb = load_workbook('PLANNING.xlsm', keep_vba=True)
ws_spec = wb['Specialisations']
ws_guides = wb['Guides']

# En-têtes
headers = ['ID_Specialisation', 'Prenom_Guide', 'Nom_Guide', 'Type_Prestation', 'Autorise']
for col, header in enumerate(headers, 1):
    ws_spec.cell(1, col).value = header

# Liste déroulante Type_Prestation (colonne D)
types_presta = '"VISITE CONTEE BRANLY,VISITE CONTEE MARINE,HORS LES MURS,VISIO,EVENEMENT BRANLY"'
dv_type = DataValidation(type="list", formula1=types_presta)
ws_spec.add_data_validation(dv_type)
dv_type.add(f'D2:D1000')

# Liste déroulante Autorisé (colonne E)
dv_autorise = DataValidation(type="list", formula1='"OUI,NON"')
ws_spec.add_data_validation(dv_autorise)
dv_autorise.add(f'E2:E1000')

# Pré-remplir pour tous les guides (tous autorisés par défaut)
types_prestations = [
    "VISITE CONTEE BRANLY",
    "VISITE CONTEE MARINE",
    "HORS LES MURS",
    "VISIO",
    "EVENEMENT BRANLY"
]

next_row = 2
spec_id = 1

for row in range(2, ws_guides.max_row + 1):
    prenom = ws_guides.cell(row, 2).value  # Colonne B
    nom = ws_guides.cell(row, 3).value     # Colonne C

    if not prenom or not nom:
        continue

    # Créer 5 lignes par guide (1 par type de prestation)
    for type_presta in types_prestations:
        ws_spec.cell(next_row, 1).value = f"S{spec_id:04d}"
        ws_spec.cell(next_row, 2).value = prenom
        ws_spec.cell(next_row, 3).value = nom
        ws_spec.cell(next_row, 4).value = type_presta
        ws_spec.cell(next_row, 5).value = "OUI"  # Par défaut tous autorisés

        spec_id += 1
        next_row += 1

wb.save('PLANNING.xlsm')
print(f"✅ Spécialisations initialisées: {spec_id-1} lignes créées")
```

#### Étape 5.2: Adapter Module_Planning.bas (30min)

**Fonction `GenererPlanningAutomatique` - Ajouter filtre spécialisations**

Modifier ligne ~180 (boucle des guides disponibles):

AVANT:
```vba
' Parcourir tous les guides disponibles
For Each guideDispo In guidesDisponibles
    ' Assigner le guide
    ...
Next guideDispo
```

APRÈS:
```vba
' Parcourir tous les guides disponibles
For Each guideDispo In guidesDisponibles
    ' NOUVEAU: Vérifier si guide autorisé pour ce type de prestation
    Dim typePrestation As String
    typePrestation = wsVisites.Cells(i, 6).Value ' Colonne F: Type_Prestation

    If EstGuideAutorise(guideDispo, typePrestation) Then
        ' Assigner le guide
        ...
    End If
Next guideDispo
```

**Nouvelle fonction à ajouter: `EstGuideAutorise`**

```vba
Function EstGuideAutorise(nomGuide As String, typePrestation As String) As Boolean
    ' Vérifie si un guide est autorisé pour un type de prestation
    ' Recherche dans l'onglet Specialisations

    Dim wsSpec As Worksheet
    Dim derniereLigne As Long
    Dim i As Long
    Dim prenom As String, nom As String
    Dim prenomGuide As String, nomGuide As String

    Set wsSpec = ThisWorkbook.Worksheets("Specialisations")
    derniereLigne = wsSpec.Cells(wsSpec.Rows.Count, 1).End(xlUp).Row

    ' Parser le nom complet "Prénom Nom"
    If InStr(nomGuide, " ") > 0 Then
        prenomGuide = Split(nomGuide, " ")(0)
        nomGuide = Split(nomGuide, " ")(1)
    Else
        prenomGuide = nomGuide
        nomGuide = ""
    End If

    ' Parcourir les spécialisations
    For i = 2 To derniereLigne
        prenom = wsSpec.Cells(i, 2).Value      ' Colonne B
        nom = wsSpec.Cells(i, 3).Value         ' Colonne C
        Dim typeSpec As String
        typeSpec = wsSpec.Cells(i, 4).Value    ' Colonne D
        Dim autorise As String
        autorise = wsSpec.Cells(i, 5).Value    ' Colonne E

        ' Vérifier correspondance
        If UCase(prenom) = UCase(prenomGuide) And _
           UCase(nom) = UCase(nomGuide) And _
           UCase(typeSpec) = UCase(typePrestation) And _
           UCase(autorise) = "OUI" Then
            EstGuideAutorise = True
            Exit Function
        End If
    Next i

    ' Si pas trouvé ou pas autorisé
    EstGuideAutorise = False
End Function
```

**Alternative simple si onglet Specialisations vide:**
```vba
Function EstGuideAutorise(nomGuide As String, typePrestation As String) As Boolean
    ' Si onglet Specialisations vide, autoriser tout le monde (comportement par défaut)
    Dim wsSpec As Worksheet
    Set wsSpec = ThisWorkbook.Worksheets("Specialisations")

    If wsSpec.Cells(2, 1).Value = "" Then
        ' Onglet vide = pas de restrictions
        EstGuideAutorise = True
    Else
        ' Appliquer les restrictions
        ' [Code ci-dessus]
    End If
End Function
```

#### Étape 5.3: Tests spécialisations (15min)

1. **Test 1: Tous guides autorisés**
   - Laisser onglet Specialisations vide
   - Générer planning → Tous les guides doivent être proposés

2. **Test 2: Restrictions actives**
   - Remplir spécialisations (Guide A = seulement Branly)
   - Ajouter visite Branly → Guide A proposé ✅
   - Ajouter visite Hors-les-murs → Guide A PAS proposé ✅

3. **Test 3: Multiple guides**
   - Guide A = Branly
   - Guide B = Branly + Marine
   - Visite Branly → A et B proposés ✅
   - Visite Marine → Seulement B proposé ✅

---

### **PHASE 6: TESTS FINAUX & LIVRAISON** ⏱️ **30min**

#### Tests complets

1. ✅ Import de quelques lignes de son planning
2. ✅ Calculs de paie avec nouveaux types
3. ✅ Génération planning automatique
4. ✅ Envoi email test
5. ✅ Export contrat DPAE

#### Documentation réponses

Créer `REPONSES_QUESTIONS.md`:

**Q: Peut modifier colonnes A et C des tarifs ?**
R: Oui colonne B (valeurs) et C (descriptions). Pas toucher A sauf si tu modifies aussi le code VBA.

**Q: 45min = 1h pour salaire ?**
R: Oui, le système calcule selon nombre de visites/jour, pas durée exacte.

**Q: Comment guide met précisions "libre jusqu'à 16h" ?**
R: Colonne "Précisions" dans Mes_Disponibilites.

**Q: Guide doit remettre nom/prénom ?**
R: Non, rempli automatiquement selon connexion.

**Q: Aucune dispo comment signifier ?**
R: Ne rien saisir = pas dispo.

**Q: Pourquoi onglet Disponibilites admin ?**
R: Vue centralisée. Pas à remplir, se remplit auto.

---

## 📧 EMAIL À LUI ENVOYER MAINTENANT

```
Objet: Adaptation du système - Planning reçu

Bonjour Marie-Laure,

Merci pour votre planning exemple, c'est exactement ce qu'il me fallait !

J'ai analysé votre structure et je vais adapter le système pour qu'il corresponde parfaitement :

✅ Ajout des colonnes manquantes (Nom structure, Niveau, Thème)
✅ Colonne Type de prestation (Branly/Marine/Hors-les-murs/Visio/Événement)
✅ Adaptation des calculs automatiques de salaire selon type
✅ Import automatique de vos données

**Livraison prévue: Demain matin (13 novembre)**

Je vous tiens informée ce soir de l'avancement.

Pour vos autres questions (tarifs modifiables, disponibilités, etc.), je prépare un document récapitulatif avec toutes les réponses.

Cordialement,
Otmane
```

---

## 📊 RÉCAPITULATIF TIMING

| Phase | Tâches | Temps |
|-------|--------|-------|
| Phase 1 | Restructurer Visites | 1h30 |
| Phase 2 | Adapter macros VBA | 2h00 |
| Phase 3 | Script import Python | 0h45 |
| Phase 4 | Corrections mineures | 0h15 |
| Phase 5 | Spécialisations guides | 1h00 |
| Phase 6 | Tests finaux & doc | 0h30 |
| **TOTAL** | **6 phases** | **6h00** |

**Répartition réaliste:**
- Ce soir (12 nov 18h-22h): Phases 1-2 = 3h30
- Demain matin (13 nov 9h-12h): Phases 3-6 = 2h30
- **Livraison: 13 novembre midi**

---

## ✅ CHECKLIST FINALE

### Développement
- [ ] PLANNING.xlsm restructuré (onglet Visites)
- [ ] Module_Calculs.bas adapté (IdentifierTypeVisite + références colonnes)
- [ ] Module_Planning.bas adapté (références colonnes + spécialisations)
- [ ] Module_Emails.bas adapté (templates)
- [ ] Module_Contrats.bas + DPAE adapté
- [ ] Onglet Specialisations initialisé
- [ ] Fonction EstGuideAutorise créée
- [ ] Script import créé (importer_planning_cliente.py)
- [ ] Script spécialisations créé (initialiser_specialisations.py)

### Tests
- [ ] Quelques données importées pour test
- [ ] Calculs paie vérifiés (5 types de prestations)
- [ ] Spécialisations guides testées
- [ ] Génération planning automatique OK
- [ ] Emails fonctionnels
- [ ] Encart gênant supprimé
- [ ] Mes_Disponibilites nettoyé

### Documentation & Livraison
- [ ] Document REPONSES_QUESTIONS.md créé
- [ ] Tests complets OK
- [ ] Backup de l'ancien PLANNING.xlsm
- [ ] Fichier final envoyé
- [ ] Email de livraison envoyé
- [ ] Projet clôturé sur Malt

---

## 📧 EMAIL DE LIVRAISON

```
Objet: ✅ Système adapté - Prêt pour import planning

Bonjour Marie-Laure,

Le système est maintenant parfaitement adapté à votre structure de planning !

**✅ Ce qui a été fait:**

1. **Structure Visites adaptée** - Colonnes identiques à votre planning:
   - Date, Heure début, Heure fin, Participants
   - Type de prestation (Branly/Marine/Hors-les-murs/Visio/Événement)
   - Nom structure, Niveau, Thème, Commentaires

2. **Import automatique** - Script Python qui importe votre planning Excel en 1 clic

3. **Calculs automatiques** - Le système reconnaît maintenant automatiquement:
   - Visite Contée Branly → Tarif Branly
   - Hors les murs → Tarif Hors-les-murs
   - Visio → Tarif Visio
   - Événement → Tarif Événement

4. **Spécialisations guides** - Onglet configuré pour définir qui fait quoi
   (actuellement tous les guides font tout - vous pouvez restreindre si besoin)

5. **Corrections** - Encart gênant supprimé, interface disponibilités simplifiée

**📎 Fichiers joints:**
- PLANNING.xlsm (version adaptée)
- REPONSES_QUESTIONS.pdf (réponses à toutes vos questions)
- importer_planning.py (script d'import - je peux l'utiliser pour vous)

**🚀 Prochaine étape:**
Envoyez-moi votre fichier Excel de planning complet, je l'importe et vous renvoie le PLANNING.xlsm rempli avec toutes vos données.

Cordialement,
Otmane
```

---

**FIN DU PLAN**
