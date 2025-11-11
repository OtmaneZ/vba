Attribute VB_Name = "Module_Emails"
'===============================================================================
' MODULE: Gestion des Emails
' DESCRIPTION: Envoi automatique des plannings et notifications
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

'===============================================================================
' FONCTION: EnvoyerPlanningMensuel
' DESCRIPTION: Envoie le planning du mois a chaque guide par email
'===============================================================================
Public Sub EnvoyerPlanningMensuel()
    Dim wsPlanning As Worksheet
    Dim wsGuides As Worksheet
    Dim dictGuides As Object ' Dictionary
    Dim i As Long
    Dim guideID As String
    Dim guideMail As String
    Dim guideNom As String
    Dim mois As Integer
    Dim annee As Integer
    Dim compteur As Integer

    On Error GoTo Erreur

    ' Demander le mois concerne
    Dim moisStr As String
    moisStr = InputBox("Entrez le mois (format MM/AAAA):", "Planning mensuel", Format(Date, "mm/yyyy"))
    If moisStr = "" Then Exit Sub

    On Error Resume Next
    mois = CInt(Left(moisStr, 2))
    annee = CInt(Right(moisStr, 4))
    On Error GoTo Erreur

    If mois < 1 Or mois > 12 Then
        MsgBox "Mois invalide.", vbExclamation
        Exit Sub
    End If

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set dictGuides = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Regrouper les visites par guide
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = wsPlanning.Cells(i, 5).Value

        ' Ignorer si non attribue
        If guideID <> "NON ATTRIBUE" And guideID <> "" Then
            On Error Resume Next
            Dim dateVisite As Date
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 Then
                If Month(dateVisite) = mois And Year(dateVisite) = annee Then
                    ' Ajouter au dictionnaire
                    If Not dictGuides.exists(guideID) Then
                        dictGuides.Add guideID, New Collection
                    End If

                    ' Ajouter les infos de la visite
                    Dim infoVisite As String
                    infoVisite = Format(dateVisite, "dd/mm/yyyy") & " | " & _
                                wsPlanning.Cells(i, 3).Value & " | " & _
                                wsPlanning.Cells(i, 4).Value

                    dictGuides(guideID).Add infoVisite
                End If
            End If
            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    ' Envoyer un email a chaque guide
    compteur = 0
    Dim key As Variant
    For Each key In dictGuides.Keys
        guideID = CStr(key)
        guideNom = ObtenirNomGuideEmail(guideID)
        guideMail = ObtenirEmailGuide(guideID)

        If guideMail <> "" Then
            Call EnvoyerEmailPlanning(guideMail, guideNom, dictGuides(guideID), mois, annee)
            compteur = compteur + 1
        End If
    Next key

    Application.ScreenUpdating = True

    MsgBox "Plannings envoyes a " & compteur & " guide(s) !", vbInformation
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'envoi : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: EnvoyerEmailPlanning
' DESCRIPTION: Envoie un email avec le planning personnalise
'===============================================================================
Private Sub EnvoyerEmailPlanning(emailDest As String, nomGuide As String, visites As Collection, mois As Integer, annee As Integer)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim corpsEmail As String
    Dim visite As Variant
    Dim nomMois As String

    On Error GoTo Erreur

    ' Creer l'objet Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) ' 0 = olMailItem

    ' Nom du mois en francais
    nomMois = Format(DateSerial(annee, mois, 1), "mmmm yyyy")

    ' Construire le corps de l'email
    corpsEmail = "Bonjour " & nomGuide & "," & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Voici votre planning pour le mois de " & nomMois & " :" & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "" & vbCrLf

    For Each visite In visites
        corpsEmail = corpsEmail & visite & vbCrLf
    Next visite

    corpsEmail = corpsEmail & "" & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Nombre total de visites : " & visites.Count & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Vous recevrez des rappels automatiques 7 jours et 1 jour avant chaque visite." & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Cordialement," & vbCrLf
    corpsEmail = corpsEmail & "L'equipe de gestion" & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "" & vbCrLf
    corpsEmail = corpsEmail & "Cet email a ete genere automatiquement. Ne pas repondre."

    ' Configurer l'email
    With OutlookMail
        .To = emailDest
        .Subject = "Planning du mois de " & nomMois
        .Body = corpsEmail
        .Send ' Utiliser .Display pour voir avant envoi (mode test)
    End With

    ' Nettoyer
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

    Exit Sub

Erreur:
    MsgBox "Erreur envoi email a " & emailDest & " : " & Err.Description, vbExclamation
End Sub

'===============================================================================
' FONCTION: EnvoyerNotificationsAutomatiques
' DESCRIPTION: Envoie les notifications J-7 et J-1 pour toutes les visites
'===============================================================================
Public Sub EnvoyerNotificationsAutomatiques()
    Dim wsPlanning As Worksheet
    Dim i As Long
    Dim dateVisite As Date
    Dim dateAujourdhui As Date
    Dim joursDifference As Long
    Dim guideID As String
    Dim guideMail As String
    Dim guideNom As String
    Dim infoVisite As String
    Dim compteurJ7 As Integer
    Dim compteurJ1 As Integer

    On Error GoTo Erreur

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    dateAujourdhui = Date
    compteurJ7 = 0
    compteurJ1 = 0

    Application.ScreenUpdating = False

    ' Parcourir toutes les visites planifiees
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = wsPlanning.Cells(i, 5).Value

        ' Ignorer si non attribue
        If guideID <> "NON ATTRIBUE" And guideID <> "" Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 Then
                joursDifference = DateDiff("d", dateAujourdhui, dateVisite)

                ' Notification J-7
                If joursDifference = DELAI_NOTIFICATION_1 Then
                    guideNom = wsPlanning.Cells(i, 6).Value
                    guideMail = ObtenirEmailGuide(guideID)

                    infoVisite = "Date : " & Format(dateVisite, "dd/mm/yyyy") & vbCrLf & _
                                "Heure : " & wsPlanning.Cells(i, 3).Value & vbCrLf & _
                                "Lieu : " & wsPlanning.Cells(i, 4).Value

                    If guideMail <> "" Then
                        Call EnvoyerNotificationVisite(guideMail, guideNom, infoVisite, "J-7")
                        compteurJ7 = compteurJ7 + 1
                    End If
                End If

                ' Notification J-1
                If joursDifference = DELAI_NOTIFICATION_2 Then
                    guideNom = wsPlanning.Cells(i, 6).Value
                    guideMail = ObtenirEmailGuide(guideID)

                    infoVisite = "Date : " & Format(dateVisite, "dd/mm/yyyy") & vbCrLf & _
                                "Heure : " & wsPlanning.Cells(i, 3).Value & vbCrLf & _
                                "Lieu : " & wsPlanning.Cells(i, 4).Value

                    If guideMail <> "" Then
                        Call EnvoyerNotificationVisite(guideMail, guideNom, infoVisite, "J-1")
                        compteurJ1 = compteurJ1 + 1
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Notifications envoyees :" & vbCrLf & _
           "- J-7 : " & compteurJ7 & " notification(s)" & vbCrLf & _
           "- J-1 : " & compteurJ1 & " notification(s)", vbInformation

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'envoi des notifications : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: EnvoyerNotificationVisite
' DESCRIPTION: Envoie une notification pour une visite
'===============================================================================
Private Sub EnvoyerNotificationVisite(emailDest As String, nomGuide As String, infoVisite As String, typeNotif As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim corpsEmail As String
    Dim importance As String

    On Error GoTo Erreur

    ' Creer l'objet Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    ' Message selon le type
    If typeNotif = "J-7" Then
        importance = "dans 7 jours"
    Else
        importance = "DEMAIN"
    End If

    ' Construire le corps de l'email
    corpsEmail = "Bonjour " & nomGuide & "," & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "[!] RAPPEL : Vous avez une visite " & importance & " !" & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "Details de la visite :" & vbCrLf
    corpsEmail = corpsEmail & "" & vbCrLf
    corpsEmail = corpsEmail & infoVisite & vbCrLf
    corpsEmail = corpsEmail & "" & vbCrLf & vbCrLf

    If typeNotif = "J-1" Then
        corpsEmail = corpsEmail & "N'oubliez pas de preparer votre visite !" & vbCrLf & vbCrLf
    End If

    corpsEmail = corpsEmail & "Cordialement," & vbCrLf
    corpsEmail = corpsEmail & "L'equipe de gestion" & vbCrLf & vbCrLf
    corpsEmail = corpsEmail & "" & vbCrLf
    corpsEmail = corpsEmail & "Cet email a ete genere automatiquement."

    ' Configurer l'email
    With OutlookMail
        .To = emailDest
        .Subject = "Rappel Visite " & typeNotif
        .Body = corpsEmail

        ' Importance haute pour J-1
        If typeNotif = "J-1" Then
            .Importance = 2 ' olImportanceHigh
        End If

        .Send ' Utiliser .Display pour voir avant envoi (mode test)
    End With

    ' Nettoyer
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

    Exit Sub

Erreur:
    ' Ne pas bloquer le processus en cas d'erreur sur un email
    Debug.Print "Erreur envoi notification a " & emailDest & " : " & Err.Description
End Sub

'===============================================================================
' FONCTION: ObtenirEmailGuide
' DESCRIPTION: Retourne l'email d'un guide
'===============================================================================
Private Function ObtenirEmailGuide(guideID As String) As String
    Dim wsGuides As Worksheet
    Dim i As Long

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    ObtenirEmailGuide = ""

    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            ObtenirEmailGuide = wsGuides.Cells(i, 4).Value ' Colonne Email
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' FONCTION: ObtenirNomGuideEmail
' DESCRIPTION: Retourne le nom complet d'un guide
'===============================================================================
Private Function ObtenirNomGuideEmail(guideID As String) As String
    Dim wsGuides As Worksheet
    Dim i As Long

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    ObtenirNomGuideEmail = ""

    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            ObtenirNomGuideEmail = wsGuides.Cells(i, 2).Value & " " & wsGuides.Cells(i, 3).Value
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' FONCTION: TestEnvoiEmail
' DESCRIPTION: Fonction de test pour verifier la configuration Outlook
'===============================================================================
Public Sub TestEnvoiEmail()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim emailTest As String

    On Error GoTo Erreur

    emailTest = InputBox("Entrez votre email pour le test:", "Test envoi email")
    If emailTest = "" Then Exit Sub

    ' Creer l'objet Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail
        .To = emailTest
        .Subject = "Test - Systeme de gestion planning guides"
        .Body = "Ceci est un email de test." & vbCrLf & vbCrLf & _
                "Si vous recevez cet email, la configuration est correcte !" & vbCrLf & vbCrLf & _
                "Date/Heure : " & Now
        .Display ' Afficher au lieu d'envoyer directement
    End With

    MsgBox "Email de test prepare ! Verifiez et envoyez.", vbInformation

    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Exit Sub

Erreur:
    MsgBox "Erreur lors du test : " & Err.Description & vbCrLf & vbCrLf & _
           "Verifiez qu'Outlook est installe et configure.", vbCritical
End Sub

'===============================================================================
' FONCTION: ConfigurerTacheAutomatique
' DESCRIPTION: Cree une tache planifiee pour les notifications (necessite config Windows)
'===============================================================================
Public Sub ConfigurerTacheAutomatique()
    MsgBox "Pour automatiser l'envoi des notifications quotidiennes :" & vbCrLf & vbCrLf & _
           "1. Ouvrez le Planificateur de taches Windows" & vbCrLf & _
           "2. Creez une nouvelle tache de base" & vbCrLf & _
           "3. Programme : Excel.exe" & vbCrLf & _
           "4. Arguments : /x ""[Chemin du fichier]"" /e" & vbCrLf & _
           "5. Macro a executer : Module_Emails.EnvoyerNotificationsAutomatiques" & vbCrLf & vbCrLf & _
           "Ou consultez le guide d'installation pour plus de details.", _
           vbInformation, "Configuration tache automatique"
End Sub

'===============================================================================
' FONCTION: EnvoyerContratParEmail
' DESCRIPTION: Envoie un contrat genere par email au guide
'===============================================================================
Public Sub EnvoyerContratParEmail()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim guideID As String, guideNom As String, guideMail As String
    Dim wsGuides As Worksheet
    Dim fichierContrat As String
    Dim moisFiltre As String
    Dim i As Long

    On Error GoTo Erreur

    ' Demander le guide
    guideID = InputBox("Entrez l'ID du guide:", "Envoi contrat par email")
    If guideID = "" Then Exit Sub

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)

    ' Recuperer infos guide
    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            guideNom = wsGuides.Cells(i, 2).Value & " " & wsGuides.Cells(i, 3).Value
            guideMail = wsGuides.Cells(i, 4).Value
            Exit For
        End If
    Next i

    If guideNom = "" Then
        MsgBox "Guide non trouve.", vbExclamation
        Exit Sub
    End If

    If guideMail = "" Then
        MsgBox "Email du guide non renseigne.", vbExclamation
        Exit Sub
    End If

    ' Demander le mois
    moisFiltre = InputBox("Mois du contrat (MM/AAAA):", "Periode", Format(Date, "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    ' Demander le type de contrat
    Dim typeContrat As String
    typeContrat = InputBox("Type de contrat (PROVISOIRE ou FINAL):", "Type", "PROVISOIRE")
    If typeContrat = "" Then Exit Sub

    typeContrat = UCase(Trim(typeContrat))

    ' Demander le fichier contrat
    fichierContrat = Application.GetOpenFilename("Fichiers Excel (*.xlsx; *.xls), *.xlsx; *.xls", , "Selectionner le contrat a envoyer")
    If fichierContrat = "False" Or fichierContrat = "" Then Exit Sub

    ' Creer l'email
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail
        .To = guideMail

        If typeContrat = "PROVISOIRE" Then
            .Subject = "Contrat provisoire - " & moisFiltre & " - " & guideNom
            .Body = "Bonjour " & guideNom & "," & vbCrLf & vbCrLf & _
                    "Veuillez trouver ci-joint votre contrat provisoire pour le mois de " & moisFiltre & "." & vbCrLf & vbCrLf & _
                    "Ce contrat contient le pre-planning avec le tarif minimum." & vbCrLf & _
                    "Un contrat final avec les dates et montants exacts vous sera envoye en fin de mois." & vbCrLf & vbCrLf & _
                    "Cordialement," & vbCrLf & _
                    "L'Association"
        Else
            .Subject = "Contrat final - " & moisFiltre & " - " & guideNom
            .Body = "Bonjour " & guideNom & "," & vbCrLf & vbCrLf & _
                    "Veuillez trouver ci-joint votre contrat final pour le mois de " & moisFiltre & "." & vbCrLf & vbCrLf & _
                    "Ce contrat contient les dates reelles et le montant exact de votre remuneration." & vbCrLf & _
                    "Merci de le signer et de nous le retourner." & vbCrLf & vbCrLf & _
                    "Cordialement," & vbCrLf & _
                    "L'Association"
        End If

        .Attachments.Add fichierContrat
        .Display ' Afficher pour verification avant envoi
    End With

    MsgBox "Email prepare avec succes !" & vbCrLf & vbCrLf & _
           "Destinataire : " & guideMail & vbCrLf & _
           "Contrat : " & fichierContrat & vbCrLf & vbCrLf & _
           "Verifiez le contenu et cliquez sur Envoyer.", _
           vbInformation

    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Exit Sub

Erreur:
    MsgBox "Erreur lors de l'envoi : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: EnvoyerNotificationReattribution
' DESCRIPTION: Notifie un guide qu'une visite lui a ete reattribuee
'===============================================================================
Public Sub EnvoyerNotificationReattribution(nouveauGuide As String, dateVisite As Date, heureVisite As String, typeVisite As String, ancienGuide As String)
    On Error GoTo Erreur

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim wsGuides As Worksheet
    Dim guideMail As String
    Dim lastRow As Long
    Dim i As Long

    ' Recuperer l'email du nouveau guide
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    lastRow = wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        Dim nomComplet As String
        nomComplet = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value
        If InStr(1, UCase(nomComplet), UCase(nouveauGuide), vbTextCompare) > 0 Then
            guideMail = wsGuides.Cells(i, 3).Value ' Email en colonne C
            Exit For
        End If
    Next i

    If guideMail = "" Then
        ' Pas d'email trouve - notification silencieuse
        Exit Sub
    End If

    ' Creer l'email via Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail
        .To = guideMail
        .Subject = "[NOUVELLE VISITE] Reattribution - " & Format(dateVisite, "dd/mm/yyyy")
        .Body = "Bonjour " & nouveauGuide & "," & vbCrLf & vbCrLf & _
                "Une visite vous a ete reattribuee suite au refus de " & ancienGuide & "." & vbCrLf & vbCrLf & _
                "DETAILS DE LA VISITE :" & vbCrLf & _
                "- Date : " & Format(dateVisite, "dddd dd mmmm yyyy") & vbCrLf & _
                "- Heure : " & heureVisite & vbCrLf & _
                "- Type : " & typeVisite & vbCrLf & vbCrLf & _
                "Merci de confirmer votre disponibilite des que possible en vous connectant au systeme de planning." & vbCrLf & vbCrLf & _
                "Cordialement," & vbCrLf & _
                "L'equipe de gestion"

        ' Envoyer automatiquement (notification rapide)
        .Send
    End With

    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Exit Sub

Erreur:
    ' Erreur silencieuse - ne pas bloquer la reattribution
    Debug.Print "Erreur email reattribution : " & Err.Description
End Sub
