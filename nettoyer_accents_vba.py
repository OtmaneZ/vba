#!/usr/bin/env python3
"""
Nettoyer tous les accents et caractères spéciaux des fichiers VBA
"""
import os
from pathlib import Path
import shutil

# Mapping des caractères accentués
REPLACEMENTS = {
    'à': 'a', 'â': 'a', 'ä': 'a',
    'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
    'î': 'i', 'ï': 'i',
    'ô': 'o', 'ö': 'o',
    'ù': 'u', 'û': 'u', 'ü': 'u',
    'ç': 'c',
    'À': 'A', 'Â': 'A', 'Ä': 'A',
    'É': 'E', 'È': 'E', 'Ê': 'E', 'Ë': 'E',
    'Î': 'I', 'Ï': 'I',
    'Ô': 'O', 'Ö': 'O',
    'Ù': 'U', 'Û': 'U', 'Ü': 'U',
    'Ç': 'C',
    # Caractères problématiques d'encodage
    '√©': 'e',
    '√®': 'i',
    '√´': 'o',
    '√¢': 'a',
    '√™': 'u',
    '√ß': 'c',
    '√‰': 'E',
    # Guillemets
    '"': '"', '"': '"', ''': "'", ''': "'",
    '«': '"', '»': '"',
    '…': '...',
}

def nettoyer_texte(texte):
    """Remplace tous les accents et caractères spéciaux"""
    for old, new in REPLACEMENTS.items():
        texte = texte.replace(old, new)
    return texte

def nettoyer_fichier(filepath):
    """Nettoie un fichier VBA"""
    print(f"\n📄 {filepath.name}")
    
    # Lire avec différents encodages
    content = None
    for encoding in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
        try:
            with open(filepath, 'r', encoding=encoding) as f:
                content = f.read()
            print(f"   ✅ Lu avec {encoding}")
            break
        except:
            continue
    
    if content is None:
        print(f"   ❌ Impossible de lire le fichier")
        return False
    
    # Compter les caractères problématiques
    nb_accents = sum(content.count(char) for char in REPLACEMENTS.keys())
    
    if nb_accents == 0:
        print(f"   ✓ Aucun accent à nettoyer")
        return True
    
    print(f"   🔧 {nb_accents} caractère(s) à nettoyer")
    
    # Nettoyer
    content_clean = nettoyer_texte(content)
    
    # Backup
    backup = filepath.parent / f"{filepath.stem}_backup{filepath.suffix}"
    shutil.copy2(filepath, backup)
    
    # Sauvegarder en UTF-8 propre
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content_clean)
    
    print(f"   ✅ Nettoyé et sauvegardé (backup: {backup.name})")
    return True

def main():
    print("="*80)
    print("NETTOYAGE ACCENTS ET ENCODAGE - FICHIERS VBA")
    print("="*80)
    
    vba_dir = Path("vba-modules")
    
    if not vba_dir.exists():
        print(f"\n❌ Dossier {vba_dir} introuvable")
        return
    
    # Trouver tous les fichiers .bas et .cls
    fichiers = list(vba_dir.glob("*.bas")) + list(vba_dir.glob("*.cls"))
    
    print(f"\n📦 {len(fichiers)} fichier(s) trouvé(s)")
    
    nettoyes = 0
    for fichier in sorted(fichiers):
        if nettoyer_fichier(fichier):
            nettoyes += 1
    
    print("\n" + "="*80)
    print(f"✅ NETTOYAGE TERMINÉ : {nettoyes}/{len(fichiers)} fichiers")
    print("="*80)
    
    print("\n📝 PROCHAINES ÉTAPES:")
    print("   1. Les fichiers dans vba-modules/ sont maintenant propres")
    print("   2. Réimporte-les dans PLANNING.xlsm (copier-coller dans VBA)")
    print("   3. Ou utilise un script pour automatiser l'import")

if __name__ == "__main__":
    main()
