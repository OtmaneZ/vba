#!/usr/bin/env python3
"""
Extraire et vérifier le VRAI code VBA du fichier PLANNING.xlsm
"""

import sys
from pathlib import Path
import zipfile
import io

fichier = Path(__file__).parent / "PLANNING.xlsm"

try:
    # Ouvrir le XLSM comme un ZIP
    with zipfile.ZipFile(fichier, 'r') as z:
        # Chercher les fichiers VBA
        vba_files = [f for f in z.namelist() if 'vbaProject' in f.lower() or '.bin' in f]

        print("Fichiers VBA trouvés:")
        for f in vba_files:
            print(f"  - {f}")

        # Extraire vbaProject.bin pour analyse
        if 'xl/vbaProject.bin' in z.namelist():
            print("\nExtraction de vbaProject.bin...")
            vba_data = z.read('xl/vbaProject.bin')

            # Sauvegarder pour analyse
            output = Path(__file__).parent / "vbaProject_extracted.bin"
            with open(output, 'wb') as f:
                f.write(vba_data)
            print(f"Sauvegardé dans: {output}")

            # Chercher les strings "Spécialisations" dans le binaire
            vba_str = vba_data.decode('latin-1', errors='ignore')

            if 'FEUILLE_SPECIALISATIONS' in vba_str:
                print("\n✅ FEUILLE_SPECIALISATIONS trouvée dans le VBA")
            else:
                print("\n❌ FEUILLE_SPECIALISATIONS NOT FOUND dans le VBA!")

            if 'AfficherToutesFeuillesAdmin' in vba_str:
                print("✅ AfficherToutesFeuillesAdmin trouvée")

                # Chercher le code autour
                idx = vba_str.find('AfficherToutesFeuillesAdmin')
                if idx > 0:
                    snippet = vba_str[idx:idx+2000]
                    if 'FEUILLE_SPECIALISATIONS' in snippet:
                        print("   ✅ Utilise FEUILLE_SPECIALISATIONS")
                    elif 'Spécialisations' in snippet:
                        print("   ⚠️  Utilise le nom en dur 'Spécialisations'")
                    else:
                        print("   ❌ NE MENTIONNE PAS Spécialisations du tout!")
            else:
                print("❌ AfficherToutesFeuillesAdmin NOT FOUND!")

except Exception as e:
    print(f"ERREUR: {e}")
    import traceback
    traceback.print_exc()
