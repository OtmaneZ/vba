#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrait le VBA de PLANNING.xlsm pour voir ce qu'il contient vraiment
"""

import zipfile
import os
import tempfile
import shutil

FICHIER = "PLANNING.xlsm"

def extraire_vba():
    """Extrait le contenu VBA du fichier Excel"""

    print("=" * 80)
    print("🔍 EXTRACTION VBA DE PLANNING.xlsm")
    print("=" * 80)
    print()

    # Créer un dossier temporaire
    temp_dir = tempfile.mkdtemp()

    try:
        # Décompresser le fichier Excel (c'est un ZIP)
        with zipfile.ZipFile(FICHIER, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        vba_path = os.path.join(temp_dir, 'xl', 'vbaProject.bin')

        if os.path.exists(vba_path):
            vba_size = os.path.getsize(vba_path)
            print(f"✅ VBA trouvé : {vba_size:,} octets")

            # Vérifier si le VBA contient les modifications Phase 2
            with open(vba_path, 'rb') as f:
                vba_content = f.read()

            # Rechercher des marqueurs de Phase 2
            marqueurs = {
                "Module_Specialisations": b"Module_Specialisations",
                "AttribuerGuideParSpecialisation": b"AttribuerGuideParSpecialisation",
                "Type_Prestation": b"Type_Prestation",
                "Nom_Structure": b"Nom_Structure"
            }

            print("\n📋 Vérification des modifications Phase 2 :")
            for nom, marqueur in marqueurs.items():
                if marqueur in vba_content:
                    print(f"  ✅ {nom} trouvé")
                else:
                    print(f"  ❌ {nom} NON trouvé")

            # Extraire vbaProject.bin pour inspection manuelle
            output_path = "vbaProject_extracted.bin"
            shutil.copy(vba_path, output_path)
            print(f"\n💾 VBA extrait vers : {output_path}")
            print(f"   (Pour inspection avec un outil comme oletools)")

        else:
            print("❌ Pas de VBA trouvé dans le fichier !")

        # Lister les fichiers XML
        xl_path = os.path.join(temp_dir, 'xl')
        if os.path.exists(xl_path):
            print(f"\n📂 Contenu du dossier xl/ :")
            for item in os.listdir(xl_path):
                item_path = os.path.join(xl_path, item)
                if os.path.isfile(item_path):
                    size = os.path.getsize(item_path)
                    print(f"  - {item} ({size:,} octets)")

    finally:
        # Nettoyer
        shutil.rmtree(temp_dir)

    print("\n" + "=" * 80)

if __name__ == "__main__":
    extraire_vba()
