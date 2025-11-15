#!/usr/bin/env python3
"""
Ajouter un bouton 'Générer Planning' dans la feuille Planning
Nécessite que les macros soient activées dans Excel
"""
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import shutil
from pathlib import Path

fichier = Path("PLANNING.xlsm")

print("="*80)
print("AJOUT BOUTON 'GÉNÉRER PLANNING'")
print("="*80)

# Backup
backup = Path("PLANNING_backup_avant_bouton.xlsm")
shutil.copy2(fichier, backup)
print(f"\n✅ Backup: {backup.name}")

print("\n⚠️  Les boutons VBA ne peuvent pas être créés via Python")
print("   Il faut les créer manuellement dans Excel")

print("\n📝 PROCÉDURE MANUELLE:")
print("="*80)
print("1. Ouvre PLANNING.xlsm dans Excel")
print("2. Va dans la feuille 'Planning'")
print("3. Dans le ruban, clique sur 'Développeur'")
print("   (Si invisible: Fichier > Options > Ruban > Cocher 'Développeur')")
print("4. Clique sur 'Insérer' > 'Bouton (Contrôle de formulaire)'")
print("5. Dessine le bouton en haut à droite (à côté de 'Déconnexion' si visible)")
print("6. Dans la fenêtre qui s'ouvre, sélectionne la macro:")
print("   'Module_Planning.GenererPlanningAutomatique'")
print("7. Clique OK")
print("8. Clique droit sur le bouton > 'Modifier le texte'")
print("9. Écris: 'Générer Planning'")
print("10. Sauvegarde (Cmd+S)")
print("="*80)

print("\n💡 ALTERNATIVE PLUS SIMPLE:")
print("   Dans VBA (Alt+F11), tu peux aussi:")
print("   - Créer un UserForm avec un bouton")
print("   - Ou lancer la macro depuis le menu Outils > Macros")
print("   - Ou assigner un raccourci clavier (Outils > Macros > Options)")

print("\n🔧 RACCOURCI CLAVIER RECOMMANDÉ:")
print("   1. Alt+F11 pour ouvrir VBA")
print("   2. Outils > Macros")
print("   3. Sélectionne 'GenererPlanningAutomatique'")
print("   4. Clique 'Options'")
print("   5. Assigne: Ctrl+Shift+G (ou autre)")
print("   6. OK")

print("\n✅ Terminé")
print("="*80)
