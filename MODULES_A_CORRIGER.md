# MODULES VBA À RÉIMPORTER DANS EXCEL

## ✅ Modules corrigés dans vba-modules/

### 1. **Module_Planning.bas** (CRITIQUE)
**Problèmes corrigés :**
- `ObtenirGuidesDisponibles()` : Lecture incorrecte des colonnes Disponibilites
  - ❌ AVANT : Date en col 2, Guide en col 1
  - ✅ APRÈS : Date en col 1, Dispo en col 2, Prenom en col 4, Nom en col 5
- Formatage des heures : `Format(time, "hh:mm")` au lieu de concaténation
- Colonnes Visites corrigées :
  - Col 4 = Durée (pas col 14)
  - Col 5 = Type_Visite (pas col 6)
  - Col 6 = Musée (pas col 7)
  - Col 8 = Niveau, Col 9 = Thème

**Impact :** Résout heures affichées en nombres + colonne Guides_Disponibles vide

### 2. **Module_Specialisations.bas** (CRITIQUE)
**Problèmes corrigés :**
- `GuideAutoriseVisite()` : Lecture structure Spécialisations
  - ❌ AVANT : Col1=Guide, Col2=Visite, Col3=Notes (ancien format)
  - ✅ APRÈS : Col1=Prenom, Col2=Nom, Col3=Type_Visite, Col4=Autorise (OUI/NON)
- Logique simplifiée : OUI/NON au lieu de texte complexe

**Impact :** Résout attribution selon spécialisations

---

## 📝 PROCÉDURE DE RÉIMPORT

### Option 1 : Manuelle (RECOMMANDÉE - 2 min)
```
1. Ouvre PLANNING.xlsm dans Excel
2. Alt+F11 (ou Cmd+F11 sur Mac) pour ouvrir VBA
3. Double-clique sur 'Module_Planning' dans la liste
4. Sélectionne TOUT le code (Cmd+A)
5. Ouvre vba-modules/Module_Planning.bas dans VSCode
6. Copie TOUT le contenu
7. Colle dans Excel VBA (remplace tout)
8. Sauvegarde (Cmd+S)
9. RÉPÈTE pour Module_Specialisations
10. Ferme VBA et Excel
11. Rouvre PLANNING.xlsm
```

### Option 2 : Via script Python
```bash
# Nécessite installation xlwings (marche sur Mac)
pip install xlwings
python3 reimporter_vba_complet.py
```

---

## 🧪 TESTS APRÈS RÉIMPORT

1. **Ouvre PLANNING.xlsm**
2. **Connecte en ADMIN**
3. **Va dans VBA** (Alt+F11) et exécute `GenererPlanningAutomatique`
4. **Vérifie :**
   - ✅ Colonne HEURE affiche "10:30", "13:00" (PAS 0.4375)
   - ✅ Colonne GUIDES_DISPONIBLES se remplit
   - ✅ Guides attribués selon spécialisations
   - ✅ Feuille Spécialisations visible

---

## 📊 RÉSUMÉ

| Module | Statut | Impact |
|--------|--------|--------|
| Module_Planning.bas | ⚠️ À réimporter | CRITIQUE |
| Module_Specialisations.bas | ⚠️ À réimporter | CRITIQUE |
| Module_Authentification.bas | ✅ OK | Déjà corrigé |
| Module_Emails.bas | ✅ OK | Déjà corrigé |
| Module_Config.bas | ✅ OK | Déjà corrigé |

