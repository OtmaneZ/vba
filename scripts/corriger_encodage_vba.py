"""
Corriger tous les problèmes d'encodage dans les fichiers VBA
"""

import os
import glob

print("=" * 80)
print("CORRECTION ENCODAGE FICHIERS VBA")
print("=" * 80)

# Corrections à appliquer
corrections = {
    "acces": "accès",
    "Acces": "Accès",
    "activee": "activée",
    "cree": "créé",
    "creee": "créée",
    "Creer": "Créer",
    "executee": "exécutée",
    "Connecte": "Connecté",
    "succes": "succès",
}

# Parcourir tous les fichiers VBA
vba_files = glob.glob("vba-modules/*.bas") + glob.glob("vba-modules/*.cls")

nb_total_corrections = 0

for filepath in vba_files:
    filename = os.path.basename(filepath)
    
    try:
        # Lire le fichier
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        nb_corrections = 0
        
        # Appliquer les corrections
        for old, new in corrections.items():
            if old in content:
                count = content.count(old)
                content = content.replace(old, new)
                nb_corrections += count
        
        # Sauvegarder si modifié
        if content != original_content:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"✅ {filename:<35} : {nb_corrections} corrections")
            nb_total_corrections += nb_corrections
        else:
            print(f"   {filename:<35} : OK")
    
    except Exception as e:
        print(f"❌ {filename:<35} : ERREUR - {e}")

print("\n" + "=" * 80)
print(f"✅ TERMINÉ : {nb_total_corrections} corrections appliquées")
print("=" * 80)
