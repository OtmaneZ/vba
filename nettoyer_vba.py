#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script pour nettoyer tous les caractères spéciaux dans les modules VBA
Remplace les accents et emojis par des caractères ASCII simples
"""

import os
import re

# Dossier contenant les modules VBA
VBA_DIR = "/Users/otmaneboulahia/Documents/Excel-Auto/vba-modules"

# Dictionnaire de remplacement
REPLACEMENTS = {
    # Accents français
    'é': 'e',
    'è': 'e',
    'ê': 'e',
    'ë': 'e',
    'à': 'a',
    'â': 'a',
    'ä': 'a',
    'î': 'i',
    'ï': 'i',
    'ô': 'o',
    'ö': 'o',
    'ù': 'u',
    'û': 'u',
    'ü': 'u',
    'ç': 'c',
    'É': 'E',
    'È': 'E',
    'Ê': 'E',
    'À': 'A',
    'Ô': 'O',
    'Ù': 'U',
    'Ç': 'C',

    # Emojis et symboles Unicode
    '✅': '[OK]',
    '❌': '[X]',
    '🔐': '>>>',
    '⚠️': '[!]',
    'ℹ️': '[i]',
    '🚪': '[>]',
    '✓': '[OK]',
    '✗': '[X]',
    '•': '-',
    '→': '->',
    '←': '<-',
    '…': '...',
    '"': '"',
    '"': '"',
    ''': "'",
    ''': "'",
    '–': '-',
    '—': '-',

    # Caractères corrompus spécifiques
    '√©': 'e',
    '√®': 'e',
    '√†': 'a',
    '√¥': 'e',
    '√´': 'e',
    '‚Ä¢': '-',
}

def clean_vba_file(filepath):
    """Nettoie un fichier VBA de ses caractères spéciaux"""
    print(f"Traitement: {os.path.basename(filepath)}")

    try:
        # Lire le fichier
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        original_content = content
        changes_count = 0

        # Appliquer les remplacements
        for old_char, new_char in REPLACEMENTS.items():
            if old_char in content:
                count = content.count(old_char)
                content = content.replace(old_char, new_char)
                changes_count += count
                if count > 0:
                    print(f"  - '{old_char}' -> '{new_char}' ({count} fois)")

        # Sauvegarder si des changements ont été faits
        if content != original_content:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"  ✓ {changes_count} remplacements effectués\n")
            return changes_count
        else:
            print(f"  ✓ Aucun changement nécessaire\n")
            return 0

    except Exception as e:
        print(f"  ✗ ERREUR: {e}\n")
        return 0

def main():
    """Parcourt tous les fichiers .bas et .cls et les nettoie"""
    print("=" * 60)
    print("NETTOYAGE DES MODULES VBA")
    print("=" * 60)
    print()

    total_changes = 0
    files_processed = 0

    # Parcourir tous les fichiers VBA
    for filename in os.listdir(VBA_DIR):
        if filename.endswith(('.bas', '.cls')):
            filepath = os.path.join(VBA_DIR, filename)
            changes = clean_vba_file(filepath)
            total_changes += changes
            files_processed += 1

    print("=" * 60)
    print(f"TERMINÉ !")
    print(f"Fichiers traités: {files_processed}")
    print(f"Total remplacements: {total_changes}")
    print("=" * 60)

if __name__ == "__main__":
    main()
