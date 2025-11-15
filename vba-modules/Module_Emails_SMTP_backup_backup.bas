Attribute VB_Name = "Module_Emails_SMTP"
' ============================================
' MODULE: ENVOI EMAILS VIA SMTP (SANS OUTLOOK)
' ============================================
' Alternative a Outlook pour l'envoi d'emails
' Utilise Python + SMTP (compatible OVH, Gmail, etc.)
'
' CONFIGURATION REQUISE:
' 1. Python 3 installe sur le Mac
' 2. Script envoi_email_smtp.py dans le dossier scripts/
' 3. Parametres SMTP dans Configuration (nouvelles lignes a ajouter):
'    - SMTP_Serveur (ex: ssl0.ovh.net pour OVH)
'    - SMTP_Port (ex: 587 pour TLS)
'    - SMTP_Mot_De_Passe (mot de passe email)

Option Explicit

' ============================================
' FONCTION: Envoyer email via SMTP Python
' ============================================
Public Function EnvoyerEmailSMTP(emailDest As String, sujet As String, corps As String) As Boolean
    On Error GoTo GestionErreur

    ' Recuperer les parametres SMTP depuis Configuration
    Dim emailExp As String
    Dim motDePasse As String
    Dim serveurSMTP As String
    Dim portSMTP As String

    emailExp = ObtenirConfig("Email_Expediteur")
    motDePasse = ObtenirConfig("SMTP_Mot_De_Passe")
    serveurSMTP = ObtenirConfig("SMTP_Serveur")
    portSMTP = ObtenirConfig("SMTP_Port")

    ' Valeurs par defaut si non configurees
    If serveurSMTP = "" Then serveurSMTP = "ssl0.ovh.net"
    If portSMTP = "" Then portSMTP = "587"

    ' Verifications
    If emailExp = "" Or motDePasse = "" Then
        MsgBox "Configuration SMTP incomplete !" & vbCrLf & vbCrLf & _
               "Ajoutez dans Configuration:" & vbCrLf & _
               "- SMTP_Mot_De_Passe" & vbCrLf & _
               "- SMTP_Serveur (optionnel)" & vbCrLf & _
               "- SMTP_Port (optionnel)", _
               vbExclamation, "Configuration manquante"
        EnvoyerEmailSMTP = False
        Exit Function
    End If

    ' Chemin du script Python
    Dim cheminScript As String
    cheminScript = ThisWorkbook.Path & "/scripts/envoi_email_smtp.py"

    ' Verifier que le script existe
    If Dir(cheminScript) = "" Then
        MsgBox "Script Python introuvable !" & vbCrLf & vbCrLf & _
               "Fichier attendu: " & cheminScript, _
               vbCritical, "Erreur"
        EnvoyerEmailSMTP = False
        Exit Function
    End If

    ' Nettoyer les apostrophes et guillemets dans le contenu
    sujet = Replace(sujet, "'", "\'")
    sujet = Replace(sujet, """", "\""")
    corps = Replace(corps, "'", "\'")
    corps = Replace(corps, """", "\""")

    ' Construire la commande Shell
    Dim cmd As String
    cmd = "python3 """ & cheminScript & """ " & _
          emailExp & " " & _
          motDePasse & " " & _
          emailDest & " " & _
          "'" & sujet & "' " & _
          "'" & corps & "' " & _
          serveurSMTP & " " & _
          portSMTP

    ' Executer la commande
    Debug.Print "Envoi email SMTP a: " & emailDest
    Dim resultat As Long
    resultat = Shell(cmd, vbHide)

    ' Attendre un peu pour laisser Python envoyer
    Application.Wait (Now + TimeValue("0:00:02"))

    EnvoyerEmailSMTP = True
    Exit Function

GestionErreur:
    MsgBox "Erreur lors de l'envoi de l'email via SMTP:" & vbCrLf & _
           Err.Description, vbCritical, "Erreur SMTP"
    EnvoyerEmailSMTP = False
End Function

' ============================================
' FONCTION: Envoyer planning guide (version SMTP)
' ============================================
Public Sub EnvoyerPlanningSMTP(guideID As Long, emailDest As String, nomMois As String)
    On Error GoTo GestionErreur

    Dim wsPlanning As Worksheet
    Dim wsGuides As Worksheet
    Dim corpsEmail As String
    Dim nomGuide As String
    Dim i As Long

    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)

    ' Recuperer le nom du guide
    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            nomGuide = wsGuides.Cells(i, 2).Value & " " & wsGuides.Cells(i, 3).Value
            Exit For
        End If
    Next i

    ' Construire le corps de l'email
    corpsEmail = "Bonjour " & nomGuide & "," & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Voici votre planning pour le mois de " & nomMois & "." & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Connectez-vous a l'application Excel pour consulter vos visites." & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Pour toute modification, contactez l'administrateur." & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Cordialement," & vbCrLf
    corpsEmail = corpsEmail & ObtenirConfig("Nom_Association") & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "---" & vbCrLf
    corpsEmail = corpsEmail & "Cet email a ete genere automatiquement. Ne pas repondre."

    ' Envoyer via SMTP
    If EnvoyerEmailSMTP(emailDest, "Planning du mois de " & nomMois, corpsEmail) Then
        Debug.Print "✓ Planning envoye a " & emailDest
    Else
        Debug.Print "✗ Echec envoi planning a " & emailDest
    End If

    Exit Sub

GestionErreur:
    MsgBox "Erreur lors de l'envoi du planning:" & vbCrLf & Err.Description, _
           vbExclamation, "Erreur"
End Sub

' ============================================
' FONCTION: Tester la configuration SMTP
' ============================================
Public Sub TesterConfigurationSMTP()
    Dim emailTest As String
    emailTest = InputBox("Entrez votre email pour recevoir un email de test:", "Test SMTP")

    If emailTest = "" Then Exit Sub

    Dim sujet As String
    Dim corps As String

    sujet = "Test configuration SMTP - Planning Musee"
    corps = "Ceci est un email de test." & vbCrLf & vbCrLf & _
            "Si vous recevez cet email, la configuration SMTP fonctionne correctement !" & vbCrLf & vbCrLf & _
            "Configuration utilisee:" & vbCrLf & _
            "- Serveur: " & ObtenirConfig("SMTP_Serveur") & vbCrLf & _
            "- Port: " & ObtenirConfig("SMTP_Port") & vbCrLf & _
            "- Expediteur: " & ObtenirConfig("Email_Expediteur")

    If EnvoyerEmailSMTP(emailTest, sujet, corps) Then
        MsgBox "Email de test envoye !" & vbCrLf & vbCrLf & _
               "Verifiez votre boite de reception (et spams).", _
               vbInformation, "Test SMTP"
    Else
        MsgBox "L'envoi du test a echoue." & vbCrLf & vbCrLf & _
               "Verifiez la configuration SMTP.", _
               vbExclamation, "Test SMTP"
    End If
End Sub
