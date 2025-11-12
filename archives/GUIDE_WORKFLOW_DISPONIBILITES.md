# 📋 GUIDE COMPLET : Workflow du Système de Planning

## 🎯 Vue d'ensemble

Le système permet de **gérer automatiquement les disponibilités des guides** et d'attribuer les visites en fonction de ces disponibilités.

---

## 📊 Architecture des données

### **Feuille "Disponibilites" (cachée - vue admin uniquement)**
```
Colonne A: ID_Guide (numéro du guide, ex: 1, 2, 3...)
Colonne B: Date (format: jj/mm/aaaa)
Colonne C: Disponible (valeurs: "OUI", "NON", "DISPONIBLE")
Colonne D: Commentaire (optionnel, ex: "Préfère matin")
```

**Exemple de données :**
```
1 | 15/11/2025 | OUI  | Matin préféré
1 | 16/11/2025 | NON  | RDV médical
2 | 15/11/2025 | OUI  | Toute la journée
3 | 15/11/2025 | OUI  |
1 | 17/11/2025 | OUI  |
2 | 17/11/2025 | NON  | Congé
3 | 17/11/2025 | OUI  |
4 | 15/11/2025 | NON  | Déjà pris
```

### **Feuille "Planning" (vue admin)**
```
Colonne A: ID_Visite
Colonne B: Date
Colonne C: Heure
Colonne D: Type_Visite
Colonne E: Guide_Attribue
Colonne F: Guides_Disponibles ← REMPLIE AUTOMATIQUEMENT
Colonne G: Statut_Confirmation
Colonne H: Historique
```

---

## 🔄 Workflow étape par étape

### **ÉTAPE 1 : Les guides déclarent leurs disponibilités**

#### **Méthode 1 : Saisie manuelle (RECOMMANDÉE)**
1. Guide se connecte avec son login/mot de passe
2. Va sur l'onglet **"Mes_Disponibilites"**
3. Voit ses propres dispos uniquement (confidentialité)
4. Ajoute des lignes manuellement :
   ```
   Ligne 2: 1 | 15/11/2025 | OUI | Je préfère le matin
   Ligne 3: 1 | 20/11/2025 | OUI | Disponible toute la journée
   Ligne 4: 1 | 25/11/2025 | NON | Rendez-vous médical
   ```

#### **Méthode 2 : Via macro (LONGUE - pas recommandée)**
- Macro `SaisirDisponibilites` disponible mais fastidieuse
- Affiche un popup pour CHAQUE jour individuellement
- Exemple : pour 30 jours = 30 popups à valider

**💡 Meilleure pratique :**
- Guides remplissent leurs dispos **au fur et à mesure** qu'ils connaissent leurs contraintes
- Peuvent ajouter autant de lignes qu'ils veulent
- Peuvent modifier/compléter n'importe quand

---

### **ÉTAPE 2 : Admin crée une visite**

1. **Admin se connecte** avec le mot de passe admin
2. Va sur l'onglet **"Planning"**
3. **Crée une nouvelle visite** :
   - Date : 15/11/2025
   - Heure : 10h00
   - Musée : Louvre
   - Type : Visite guidée
   - Nombre de personnes : 25

4. **Le système remplit automatiquement "Guides_Disponibles"**

---

### **ÉTAPE 3 : Système détecte automatiquement qui est disponible**

**Code exécuté automatiquement : `ObtenirGuidesDisponiblesPourDate()`**

```vb
' Pseudo-code du processus
Pour chaque guide dans la base :
    estDisponible = FAUX

    Pour chaque ligne de Disponibilites :
        Si ligne.ID_Guide = guide_actuel ET
           ligne.Date = date_visite ET
           ligne.Disponible = "OUI" ALORS
            estDisponible = VRAI
            Sortir de la boucle
        Fin Si
    Fin Pour

    Si estDisponible = VRAI ALORS
        Ajouter guide à la liste "Guides_Disponibles"
    Fin Si
Fin Pour
```

**Exemple concret pour visite du 15/11/2025 :**

Données Disponibilités :
```
Ligne 2: 1 | 15/11/2025 | OUI  ← Guide 1 DISPO ✅
Ligne 3: 1 | 16/11/2025 | NON  ← Date différente, ignoré
Ligne 4: 2 | 15/11/2025 | OUI  ← Guide 2 DISPO ✅
Ligne 5: 3 | 15/11/2025 | OUI  ← Guide 3 DISPO ✅
Ligne 6: 4 | 15/11/2025 | NON  ← Guide 4 PAS DISPO ❌
```

**Résultat dans Planning, colonne F :**
```
"Marie Dupont, Jean Martin, Sophie Dubois"
```

---

### **ÉTAPE 4 : Admin attribue la visite**

1. **Admin voit la colonne "Guides_Disponibles"** remplie automatiquement
2. **Admin choisit un guide** dans la liste déroulante (colonne E)
3. **Système vérifie :**
   - ✅ Si guide est dans la liste des dispos → Attribution OK
   - ⚠️ Si guide n'est PAS dispo → Message d'alerte :
     ```
     ⚠️ ATTENTION !
     Ce guide a déclaré ne PAS être disponible pour cette date.
     Voulez-vous quand même l'attribuer ?
     [Oui] [Non]
     ```

4. **Admin confirme l'attribution**

---

### **ÉTAPE 5 : Notification automatique du guide**

**Code exécuté : `EnvoyerNotificationReattribution()` ou `EnvoyerPlanningMensuel()`**

1. **Email automatique envoyé** via Outlook :
   ```
   De: planning@musee.fr
   À: guide@email.com
   Sujet: Nouvelle visite attribuée - 15/11/2025

   Bonjour Marie,

   Une visite vous a été attribuée :

   📅 Date : 15 novembre 2025
   🕐 Heure : 10h00
   🏛️ Musée : Louvre
   📝 Type : Visite guidée
   👥 Nombre : 25 personnes

   Cette visite vous est attribuée automatiquement.
   Pour toute modification, contactez l'administrateur.

   Cordialement,
   Musée des Guides

   ---
   ⚠️ NE PAS REPONDRE À CET EMAIL
   Cette boîte mail n'est pas consultée.
   ```

2. **Guide reçoit l'email**

---

### **ÉTAPE 6 : Guide consulte son planning**

1. **Guide se connecte** à l'Excel
2. Va sur **"Mon_Planning"**
3. **Voit ses visites en LECTURE SEULE** :
   ```
   Date       | Heure | Musée  | Type          | Statut
   15/11/2025 | 10h00 | Louvre | Visite guidée | Confirmée
   20/11/2025 | 14h00 | Orsay  | Visite privée | Confirmée
   ```

4. **Aucun bouton pour refuser** → Planning non modifiable
5. **Message affiché :**
   ```
   📋 Votre planning

   Pour toute modification, contactez l'administrateur.
   Vous ne pouvez pas refuser une visite attribuée.
   ```

---

### **ÉTAPE 7 : Si erreur → Guide contacte admin**

**Scénario :** Guide reçoit une visite mais s'est trompé dans ses dispos

1. **Guide contacte admin** (téléphone, email, WhatsApp)
   ```
   "Bonjour, j'ai reçu la visite du 15/11 mais je ne peux finalement pas,
   j'ai un rendez-vous médical imprévu."
   ```

2. **Admin décide** :
   - Option A : **Réattribuer la visite** → Fonction `RefuserEtReattribuerVisite()`
   - Option B : **Forcer le guide à honorer** → Aucune action

3. **Si réattribution :**
   ```vb
   ' Admin utilise la fonction dans Planning
   RefuserEtReattribuerVisite(ligneVisite, raisonRefus)
   ```
   - Système cherche **automatiquement un autre guide dispo**
   - Envoie email au nouveau guide
   - Met à jour le planning
   - Archive l'historique

---

## 🔍 Détails techniques : Fonction `ObtenirGuidesDisponiblesPourDate()`

### **Code complet :**

```vb
Function ObtenirGuidesDisponiblesPourDate(dateVisite As Date, heureVisite As String, guideExclu As String) As String
    Dim wsDisponibilites As Worksheet
    Dim wsGuides As Worksheet
    Dim lastRowDispo As Long
    Dim lastRowGuides As Long
    Dim i As Long
    Dim j As Long
    Dim listeGuides As String
    Dim nomGuide As String
    Dim estDisponible As Boolean

    ' Récupérer les feuilles
    Set wsDisponibilites = ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES)
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)

    listeGuides = ""
    lastRowGuides = wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
    lastRowDispo = wsDisponibilites.Cells(wsDisponibilites.Rows.Count, 1).End(xlUp).Row

    ' BOUCLE PRINCIPALE : Pour chaque guide
    For i = 2 To lastRowGuides
        nomGuide = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value
        nomGuide = Trim(nomGuide)

        ' Exclure le guide qui a refusé (si réattribution)
        If UCase(nomGuide) <> UCase(guideExclu) And nomGuide <> "" Then
            estDisponible = False

            ' BOUCLE SECONDAIRE : Scanner TOUTES les lignes de Disponibilites
            For j = 2 To lastRowDispo
                Dim guideDispoNom As String
                guideDispoNom = wsDisponibilites.Cells(j, 1).Value

                ' Vérifier si c'est le bon guide
                If InStr(1, UCase(guideDispoNom), UCase(nomGuide), vbTextCompare) > 0 Then
                    Dim dateDispo As Date
                    dateDispo = CDate(wsDisponibilites.Cells(j, 2).Value)

                    ' Vérifier si c'est la bonne date
                    If dateDispo = dateVisite Then
                        ' Vérifier si disponible (colonne C)
                        If UCase(wsDisponibilites.Cells(j, 3).Value) = "OUI" Or _
                           UCase(wsDisponibilites.Cells(j, 3).Value) = "DISPONIBLE" Then
                            estDisponible = True
                            Exit For  ' Trouvé ! Pas besoin de chercher plus
                        End If
                    End If
                End If
            Next j

            ' Ajouter à la liste si disponible
            If estDisponible Then
                If listeGuides = "" Then
                    listeGuides = nomGuide
                Else
                    listeGuides = listeGuides & ", " & nomGuide
                End If
            End If
        End If
    Next i

    ' Retourner la liste complète
    ObtenirGuidesDisponiblesPourDate = listeGuides
End Function
```

### **Exemple d'exécution :**

**Données :**
- Date visite : 15/11/2025
- Guides dans base : Marie Dupont, Jean Martin, Sophie Dubois, Pierre Leroy

**Feuille Disponibilites :**
```
ID | Date       | Disponible | Commentaire
1  | 15/11/2025 | OUI        | Matin préféré
2  | 15/11/2025 | OUI        | Toute la journée
3  | 15/11/2025 | OUI        |
4  | 15/11/2025 | NON        | Congé
1  | 16/11/2025 | NON        | RDV médical
```

**Traitement :**
```
i=2 (Marie Dupont) :
  j=2: ID=1, Date=15/11, Dispo=OUI → estDisponible = TRUE ✅
  → Ajouter "Marie Dupont" à listeGuides

i=3 (Jean Martin) :
  j=3: ID=2, Date=15/11, Dispo=OUI → estDisponible = TRUE ✅
  → Ajouter "Jean Martin" à listeGuides

i=4 (Sophie Dubois) :
  j=4: ID=3, Date=15/11, Dispo=OUI → estDisponible = TRUE ✅
  → Ajouter "Sophie Dubois" à listeGuides

i=5 (Pierre Leroy) :
  j=5: ID=4, Date=15/11, Dispo=NON → estDisponible = FALSE ❌
  → Ne pas ajouter

Résultat final : "Marie Dupont, Jean Martin, Sophie Dubois"
```

---

## ✅ Avantages du système

### **Pour les guides :**
- ✅ Déclarent leurs dispos quand ils veulent
- ✅ Peuvent ajouter/modifier n'importe quand
- ✅ Pas de popup fastidieux
- ✅ Reçoivent email automatique
- ✅ Voient planning lecture seule
- ✅ Pas de pression pour "accepter/refuser" dans l'interface

### **Pour l'admin :**
- ✅ Voit automatiquement qui est dispo
- ✅ Alerte si attribution à quelqu'un non dispo
- ✅ Peut forcer l'attribution si nécessaire
- ✅ Peut réattribuer facilement en cas d'erreur
- ✅ Historique complet des changements

### **Pour le système :**
- ✅ Flexible : fonctionne avec 1 ligne ou 1000 lignes de dispo
- ✅ Temps réel : dès qu'un guide ajoute une dispo, admin la voit
- ✅ Automatique : colonne "Guides_Disponibles" se remplit seule
- ✅ Fiable : scanne TOUTES les lignes pour être sûr

---

## 🚨 Cas particuliers

### **Cas 1 : Guide n'a pas encore déclaré ses dispos**
- Colonne "Guides_Disponibles" ne le liste pas
- Admin peut quand même l'attribuer (pas de blocage)
- Pas de message d'alerte

### **Cas 2 : Guide a dit NON pour cette date**
- N'apparaît pas dans "Guides_Disponibles"
- Si admin l'attribue quand même → Message d'alerte
- Admin peut forcer

### **Cas 3 : Guide a plusieurs lignes pour la même date**
```
1 | 15/11/2025 | OUI | Matin
1 | 15/11/2025 | NON | Après-midi
```
- Système s'arrête à la **première correspondance**
- Dans cet exemple : verra "OUI" et dira que le guide est dispo
- **Recommandation** : une seule ligne par guide par jour

### **Cas 4 : Réattribution après refus**
```vb
RefuserEtReattribuerVisite(ligneVisite, "Guide malade")
```
- Système appelle `ObtenirGuidesDisponiblesPourDate()` avec `guideExclu`
- Exclut automatiquement le guide qui a refusé
- Cherche dans les autres guides dispos
- Si trouvé → Attribution automatique + email
- Si personne de dispo → Message à l'admin

---

## 📝 Checklist de livraison

### **Fichiers modifiés :**
- ✅ `PLANNING_MUSEE_FINAL_PROPRE.xlsm` (Excel principal)
- ✅ `Module_Authentification.bas` (suppression bouton refuser)
- ✅ `Module_Calculs.bas` (ajout colonnes défraiements)
- ✅ `Module_Config.bas` (structure mise à jour)
- ✅ `Module_Contrats.bas` (défraiements dans contrats)
- ✅ `Feuille_Mon_Planning.cls` (lecture seule)

### **Nouvelles colonnes Calculs_Paie :**
- ✅ Colonne N : Défraiements (€0 par défaut, saisie manuelle)
- ✅ Colonne O : Total_Avec_Frais (formule : =I+N)

### **Configuration requise :**
- ✅ Microsoft Excel pour Mac (ou Windows)
- ✅ Microsoft Outlook installé et configuré
- ✅ Compte email Outlook.com/Outlook.fr (recommandé pour simplicité)
- ✅ Macros activées ("Activer le contenu" à l'ouverture)

### **Premier démarrage :**
1. Ouvrir `PLANNING_MUSEE_FINAL_PROPRE.xlsm`
2. Cliquer "Activer le contenu" (macros)
3. Aller sur onglet "Accueil"
4. Connecter Outlook si demandé
5. Modifier Configuration (B2 = email expéditeur)
6. Créer les comptes guides dans "Guides"
7. Demander aux guides de remplir leurs dispos
8. Commencer à créer des visites !

---

## 🎓 Résumé pour la cliente

**Ce que fait le système :**
1. ✅ Guides remplissent leurs dispos dans leur onglet personnel
2. ✅ Quand vous créez une visite, le système montre automatiquement qui est dispo
3. ✅ Vous attribuez la visite à un guide (avec alerte si pas dispo)
4. ✅ Guide reçoit un email automatique
5. ✅ Guide voit son planning mais ne peut pas refuser
6. ✅ Si erreur, guide vous contacte et vous réattribuez

**Ce qui est automatique :**
- 📧 Envoi des emails
- 📋 Calcul de qui est disponible
- ⚠️ Alertes si attribution incorrecte
- 💰 Calculs de paie (cachets + défraiements)
- 📄 Génération des contrats

**Ce qui est manuel :**
- Guides remplissent leurs dispos (quand ils veulent)
- Vous créez les visites
- Vous choisissez le guide (parmi ceux suggérés)
- Vous saisissez les défraiements en fin de mois

**Support pendant 7 jours disponible pour tout ajustement !**

---

*Document généré le 11 novembre 2025*
*Version du système : FINAL avec défraiements*
