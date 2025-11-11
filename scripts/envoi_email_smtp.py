#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script d'envoi d'emails via SMTP (OVH, Gmail, etc.)
Alternative à Outlook pour l'envoi automatique d'emails
"""

import smtplib
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def envoyer_email_smtp(expediteur, mot_de_passe, destinataire, sujet, corps,
                       serveur_smtp="ssl0.ovh.net", port_smtp=587):
    """
    Envoie un email via SMTP

    Args:
        expediteur: Adresse email expéditeur
        mot_de_passe: Mot de passe email
        destinataire: Adresse email destinataire
        sujet: Sujet de l'email
        corps: Contenu de l'email (texte brut)
        serveur_smtp: Serveur SMTP (défaut: OVH)
        port_smtp: Port SMTP (défaut: 587 pour TLS)

    Returns:
        True si envoi réussi, False sinon
    """
    try:
        # Créer le message
        message = MIMEMultipart()
        message['From'] = expediteur
        message['To'] = destinataire
        message['Subject'] = sujet

        # Ajouter le corps du message
        message.attach(MIMEText(corps, 'plain', 'utf-8'))

        # Connexion au serveur SMTP
        print(f"📡 Connexion au serveur SMTP {serveur_smtp}:{port_smtp}...")
        server = smtplib.SMTP(serveur_smtp, port_smtp)
        server.starttls()  # Activer TLS

        # Authentification
        print(f"🔐 Authentification avec {expediteur}...")
        server.login(expediteur, mot_de_passe)

        # Envoi de l'email
        print(f"📧 Envoi de l'email à {destinataire}...")
        server.send_message(message)

        # Fermeture de la connexion
        server.quit()

        print("✅ Email envoyé avec succès !")
        return True

    except smtplib.SMTPAuthenticationError:
        print("❌ ERREUR: Authentification échouée. Vérifiez l'email et le mot de passe.")
        return False
    except smtplib.SMTPException as e:
        print(f"❌ ERREUR SMTP: {e}")
        return False
    except Exception as e:
        print(f"❌ ERREUR: {e}")
        return False


if __name__ == "__main__":
    """
    Utilisation depuis VBA:

    Dim cmd As String
    cmd = "python3 /chemin/vers/envoi_email_smtp.py " & _
          "expediteur@ovh.com " & _
          "motdepasse " & _
          "destinataire@email.com " & _
          "'Sujet du message' " & _
          "'Corps du message'"
    Shell cmd, vbHide
    """

    # Test rapide
    if len(sys.argv) == 1:
        print("=" * 60)
        print("📧 CONFIGURATION SMTP POUR OVH")
        print("=" * 60)
        print("\n📋 Paramètres OVH:")
        print("   Serveur SMTP: ssl0.ovh.net")
        print("   Port: 587 (TLS) ou 465 (SSL)")
        print("   Authentification: Oui")
        print("   Email: votre-email@domaine.com")
        print("   Mot de passe: Mot de passe email OVH")
        print("\n💡 Test:")
        print("   python3 envoi_email_smtp.py email@ovh.com motdepasse dest@email.com 'Test' 'Bonjour'")
        sys.exit(0)

    # Envoi depuis ligne de commande
    if len(sys.argv) < 6:
        print("❌ Usage: python3 envoi_email_smtp.py EXPEDITEUR MOTDEPASSE DESTINATAIRE SUJET CORPS")
        sys.exit(1)

    expediteur = sys.argv[1]
    mot_de_passe = sys.argv[2]
    destinataire = sys.argv[3]
    sujet = sys.argv[4]
    corps = sys.argv[5]

    # Serveur SMTP optionnel (défaut: OVH)
    serveur = sys.argv[6] if len(sys.argv) > 6 else "ssl0.ovh.net"
    port = int(sys.argv[7]) if len(sys.argv) > 7 else 587

    # Envoi
    succes = envoyer_email_smtp(expediteur, mot_de_passe, destinataire, sujet, corps, serveur, port)
    sys.exit(0 if succes else 1)
