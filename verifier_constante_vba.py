#!/usr/bin/env python3
"""
Vérifier si FEUILLE_SPECIALISATIONS existe vraiment dans le VBA du fichier Excel
"""

import sys
from pathlib import Path
import zipfile

fichier = Path(__file__).parent / "PLANNING.xlsm"

try:
    with zipfile.ZipFile(fichier, 'r') as z:
        vba_data = z.read('xl/vbaProject.bin')
        vba_str = vba_data.decode('latin-1', errors='ignore')

        print("🔍 Recherche dans le VBA du fichier Excel...")
        print("="*80)

        # Chercher la définition de la constante
        if 'FEUILLE_SPECIALISATIONS' in vba_str:
            print("✅ FEUILLE_SPECIALISATIONS trouvée")

            # Chercher la définition exacte
            idx = vba_str.find('FEUILLE_SPECIALISATIONS')
            if idx > 0:
                # Extraire le contexte
                start = max(0, idx - 200)
                end = min(len(vba_str), idx + 300)
                context = vba_str[start:end]

                # Chercher si c'est une définition de constante
                if 'Const' in context and 'String' in context:
                    print("   ✅ Définition de constante trouvée")

                    # Extraire la valeur
                    if '"Spécialisations"' in context or '"Specialisations"' in context:
                        print('   ✅ Valeur = "Spécialisations"')
                    else:
                        print("   ❌ Valeur incorrecte ou manquante")
                else:
                    print("   ❌ Pas de définition Const trouvée!")
                    print(f"\nContexte:\n{context}\n")
        else:
            print("❌ FEUILLE_SPECIALISATIONS NOT FOUND dans le fichier Excel!")
            print("\n🔴 PROBLÈME: La constante n'a PAS été ajoutée dans Module_Config du fichier Excel")
            print("\n📝 ACTION REQUISE:")
            print("1. Ouvrez PLANNING.xlsm")
            print("2. Alt+F11 pour VBA")
            print("3. Ouvrez Module_Config")
            print("4. Après la ligne:")
            print('   Public Const FEUILLE_CONFIG As String = "Configuration"')
            print("5. Ajoutez:")
            print('   Public Const FEUILLE_SPECIALISATIONS As String = "Spécialisations"')
            print("6. Ctrl+S et fermez")

        print("\n" + "="*80)
        print("\n🔍 Vérification de l'utilisation dans AfficherToutesFeuillesAdmin...")

        idx = vba_str.find('AfficherToutesFeuillesAdmin')
        if idx > 0:
            # Extraire la fonction complète
            snippet = vba_str[idx:idx+3000]

            # Compter les utilisations
            uses_const = snippet.count('FEUILLE_SPECIALISATIONS')
            uses_string = snippet.count('"Spécialisations"')

            print(f"   Utilise FEUILLE_SPECIALISATIONS: {uses_const} fois")
            print(f"   Utilise \"Spécialisations\" en dur: {uses_string} fois")

            if uses_const > 0:
                print("   ✅ La fonction utilise la constante")
            elif uses_string > 0:
                print("   ❌ La fonction utilise ENCORE le nom en dur!")
            else:
                print("   ❌ La fonction ne mentionne PAS du tout Spécialisations!")

        print("="*80)

except Exception as e:
    print(f"ERREUR: {e}")
    import traceback
    traceback.print_exc()
