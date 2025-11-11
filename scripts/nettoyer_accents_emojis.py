#!/usr/bin/env python3
"""
Script pour nettoyer TOUS les accents et emojis des fichiers VBA
Pour éviter les problèmes d'encodage
"""

import os
import re

def nettoyer_texte(texte):
    """Enlève tous les accents et emojis"""

    # Mapping complet des caractères accentués
    replacements = {
        # Minuscules
        'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'å': 'a',
        'è': 'e', 'é': 'e', 'ê': 'e', 'ë': 'e',
        'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i',
        'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
        'ù': 'u', 'ú': 'u', 'û': 'u', 'ü': 'u',
        'ý': 'y', 'ÿ': 'y',
        'ñ': 'n', 'ç': 'c',

        # Majuscules
        'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A', 'Ä': 'A', 'Å': 'A',
        'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E',
        'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I',
        'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O', 'Ö': 'O',
        'Ù': 'U', 'Ú': 'U', 'Û': 'U', 'Ü': 'U',
        'Ý': 'Y', 'Ÿ': 'Y',
        'Ñ': 'N', 'Ç': 'C',

        # Caractères spéciaux français
        'œ': 'oe', 'Œ': 'OE',
        'æ': 'ae', 'Æ': 'AE',

        # Guillemets
        '«': '"', '»': '"',
        ''': "'", ''': "'",
        '"': '"', '"': '"',

        # Tirets
        '–': '-', '—': '-',

        # Autres
        '…': '...',
        '€': 'EUR',
        '°': ' degres',
    }

    # Remplacer les caractères
    for old, new in replacements.items():
        texte = texte.replace(old, new)

    # Supprimer les emojis et autres caractères Unicode > 127
    texte = re.sub(r'[^\x00-\x7F]+', ' ', texte)

    return texte


def nettoyer_fichiers_vba():
    """Nettoie tous les fichiers VBA"""

    vba_dir = 'vba-modules'

    if not os.path.exists(vba_dir):
        print(f"❌ Dossier {vba_dir} introuvable")
        return

    fichiers_modifies = []
    total_changements = 0

    print("🔧 NETTOYAGE ACCENTS & EMOJIS")
    print("=" * 50)

    for fichier in sorted(os.listdir(vba_dir)):
        if fichier.endswith(('.bas', '.cls')):
            chemin = os.path.join(vba_dir, fichier)

            try:
                with open(chemin, 'r', encoding='utf-8') as f:
                    contenu_original = f.read()

                contenu_nettoye = nettoyer_texte(contenu_original)

                if contenu_original != contenu_nettoye:
                    with open(chemin, 'w', encoding='utf-8') as f:
                        f.write(contenu_nettoye)

                    nb_diff = sum(1 for a, b in zip(contenu_original, contenu_nettoye) if a != b)
                    fichiers_modifies.append(fichier)
                    total_changements += nb_diff
                    print(f"✅ {fichier:<35} ({nb_diff} changements)")
                else:
                    print(f"⚪ {fichier:<35} (deja propre)")

            except Exception as e:
                print(f"❌ {fichier}: {e}")

    print("=" * 50)
    print(f"\n📊 RÉSUMÉ:")
    print(f"   Fichiers modifiés: {len(fichiers_modifies)}")
    print(f"   Total changements: {total_changements}")
    print(f"   Fichiers traités: {len([f for f in os.listdir(vba_dir) if f.endswith(('.bas', '.cls'))])}")

    if fichiers_modifies:
        print(f"\n✅ Nettoyage terminé avec succès!")
        print(f"\n📝 Fichiers modifiés:")
        for f in fichiers_modifies:
            print(f"   - {f}")
    else:
        print(f"\n✅ Aucun accent ou emoji trouvé - code déjà propre!")


if __name__ == "__main__":
    nettoyer_fichiers_vba()
