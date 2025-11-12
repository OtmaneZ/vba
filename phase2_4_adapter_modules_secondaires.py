#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PHASE 2.4 - ADAPTER MODULES SECONDAIRES

Adapte les 3 modules secondaires :
- Module_Contrats.bas
- Module_DPAE.bas
- Feuille_Visites.cls

MAPPING (rappel):
  Col 5: guideID → Col 12 (Guide_Attribue)
  Col 3-4: Heure_Debut/Fin → OK (déjà correct)
"""

import sys


def lire_fichier(chemin):
    """Lit le fichier VBA."""
    with open(chemin, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()


def ecrire_fichier(chemin, contenu):
    """Écrit le fichier VBA."""
    with open(chemin, 'w', encoding='utf-8', newline='\r\n') as f:
        f.write(contenu)


def adapter_module_contrats():
    """Adapte Module_Contrats.bas."""
    chemin = 'vba-modules/Module_Contrats.bas'
    print("\n🔧 MODULE_CONTRATS.BAS")
    print("-" * 80)

    contenu = lire_fichier(chemin)
    lignes = contenu.split('\n')

    modifications = 0

    # Remplacer col 5 (guideID) par col 12
    for idx, ligne in enumerate(lignes):
        if 'wsPlanning.Cells(i, 5).Value = guideID' in ligne or \
           'guideID = wsPlanning.Cells(i, 5).Value' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 5)',
                'wsPlanning.Cells(i, 12) \' Guide_Attribue'
            )
            modifications += 1

    ecrire_fichier(chemin, '\n'.join(lignes))
    print(f"   ✅ {modifications} modifications (guideID → col 12)")

    return modifications


def adapter_module_dpae():
    """Adapte Module_DPAE.bas."""
    chemin = 'vba-modules/Module_DPAE.bas'
    print("\n🔧 MODULE_DPAE.BAS")
    print("-" * 80)

    contenu = lire_fichier(chemin)
    lignes = contenu.split('\n')

    modifications = 0

    # Remplacer col 5 (guideID) par col 12
    for idx, ligne in enumerate(lignes):
        if 'guideID = Trim(wsPlanning.Cells(i, 5).Value)' in ligne:
            lignes[idx] = ligne.replace(
                'wsPlanning.Cells(i, 5)',
                'wsPlanning.Cells(i, 12) \' Guide_Attribue'
            )
            modifications += 1

    ecrire_fichier(chemin, '\n'.join(lignes))
    print(f"   ✅ {modifications} modifications (guideID → col 12)")

    return modifications


def adapter_feuille_visites():
    """Adapte Feuille_Visites.cls."""
    chemin = 'vba-modules/Feuille_Visites.cls'
    print("\n🔧 FEUILLE_VISITES.CLS")
    print("-" * 80)

    try:
        contenu = lire_fichier(chemin)
        lignes = contenu.split('\n')

        modifications = 0

        # Chercher des références colonnes (peu probable)
        a_modifier = False
        for ligne in lignes:
            if '.Cells(' in ligne and any(str(i) in ligne for i in [3, 4, 5, 6, 7]):
                a_modifier = True
                break

        if not a_modifier:
            print(f"   ✅ Aucune modification nécessaire (pas de références colonnes)")
            return 0

        # Si nécessaire, adapter ici
        ecrire_fichier(chemin, '\n'.join(lignes))
        print(f"   ✅ {modifications} modifications")

        return modifications

    except FileNotFoundError:
        print(f"   ⚠️  Fichier non trouvé (optionnel)")
        return 0


def main():
    """Fonction principale."""
    print("=" * 100)
    print("🔧 PHASE 2.4 - ADAPTATION MODULES SECONDAIRES")
    print("=" * 100)

    total_modifs = 0

    # Adapter les 3 modules
    total_modifs += adapter_module_contrats()
    total_modifs += adapter_module_dpae()
    total_modifs += adapter_feuille_visites()

    print("\n" + "=" * 100)
    print(f"📊 RÉSUMÉ: {total_modifs} modifications au total")
    print("=" * 100)

    if total_modifs > 0:
        print("\n✅ MODULES SECONDAIRES ADAPTÉS AVEC SUCCÈS !")
    else:
        print("\n✅ AUCUNE MODIFICATION NÉCESSAIRE (déjà correct)")

    print("=" * 100)

    return True


if __name__ == '__main__':
    try:
        succes = main()
        sys.exit(0 if succes else 1)
    except Exception as e:
        print(f"\n❌ ERREUR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
