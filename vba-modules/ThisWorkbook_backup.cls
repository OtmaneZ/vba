Option Explicit

' Variables pour eviter envois multiples (avec stockage date)
Private planningEnvoyeCeMois As Boolean
Private notificationsEnvoyeesAujourdhui As Boolean
Private salaireCalculeCeMois As Boolean
Private derniereVerificationDate As Date

Private Sub Workbook_Open()
    On Error Resume Next

    ' Masquer TOUTES les feuilles sauf Accueil par defaut
    Call MasquerToutesFeuillesParDefaut

    ' Activer la page d'accueil UNIQUEMENT si aucun utilisateur connecte
    If utilisateurConnecte = "" Then
        ThisWorkbook.Sheets("Accueil").Activate
    End If

    ' AUTOMATISATION : Verifier si actions a faire
    Call VerifierActionsAutomatiques

    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If niveauAcces <> "" Then
        utilisateurConnecte = ""
        niveauAcces = ""
        emailUtilisateur = ""
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Optimisation: Controle d'acces gere au login dans Module_Authentification
    ' Plus besoin de verifier a chaque changement d'onglet (performance)
    ' Les feuilles sont deja masquees/affichees selon le role au moment de la connexion
End Sub

' Masquer toutes les feuilles sauf Accueil au demarrage
Private Sub MasquerToutesFeuillesParDefaut()
    On Error Resume Next

    ' Si un utilisateur est deja connecte, ne rien faire
    If utilisateurConnecte <> "" Then
        Exit Sub
    End If

    Dim ws As Worksheet

    ' Boucle sur TOUTES les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        ' Masquer TOUTES les feuilles sauf Accueil
        If ws.Name <> "Accueil" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' S'assurer que Accueil est visible (securise)
    On Error Resume Next
    ThisWorkbook.Sheets("Accueil").Visible = xlSheetVisible
    If Err.Number <> 0 Then
        ' Feuille Accueil introuvable - creer une nouvelle
        MsgBox "ATTENTION : Feuille Accueil introuvable !" & vbCrLf & _
               "Veuillez executer InitialiserPageAccueil() pour la recreer.", _
               vbExclamation, "Feuille manquante"
    End If
    On Error GoTo 0
End Sub

' ============================================
' AUTOMATISATION : Verifier actions a faire
' ============================================
Private Sub VerifierActionsAutomatiques()
    On Error Resume Next

    Dim jourActuel As Integer
    Dim moisActuel As Integer
    Dim dernierJourDuMois As Date
    Dim dateAujourdhui As Date

    jourActuel = Day(Date)
    moisActuel = Month(Date)
    dateAujourdhui = Date
    dernierJourDuMois = DateSerial(Year(Date), Month(Date) + 1, 0)

    ' Reinitialiser flags si changement de date
    If derniereVerificationDate <> dateAujourdhui Then
        derniereVerificationDate = dateAujourdhui

        ' Reinitialiser le flag planning si nouveau mois
        If jourActuel <> 1 Then
            planningEnvoyeCeMois = False
        End If

        ' Reinitialiser notifications chaque jour
        notificationsEnvoyeesAujourdhui = False

        ' Reinitialiser salaire si nouveau mois
        If jourActuel <> Day(dernierJourDuMois) Then
            salaireCalculeCeMois = False
        End If
    End If

    ' 1. ENVOI PLANNING MENSUEL (1er du mois a 9h) - AUTOMATIQUE
    If jourActuel = 1 And Hour(Now) >= 9 And Not planningEnvoyeCeMois Then
        ' Envoi automatique du planning mensuel
        Call EnvoyerPlanningMensuel
        planningEnvoyeCeMois = True
    End If

    ' 2. NOTIFICATIONS QUOTIDIENNES (8h-18h) - AUTOMATIQUE
    If Hour(Now) >= 8 And Hour(Now) < 18 And Not notificationsEnvoyeesAujourdhui Then
        ' Envoi automatique sans confirmation
        Call EnvoyerNotificationsAutomatiques
        notificationsEnvoyeesAujourdhui = True
    End If

    ' 3. CALCUL SALAIRES (dernier jour du mois a 17h)
    If Date = dernierJourDuMois And Hour(Now) >= 17 And Not salaireCalculeCeMois Then
        If MsgBox("C'est le dernier jour du mois !" & vbCrLf & vbCrLf & _
                  "Voulez-vous calculer les salaires automatiquement ?", _
                  vbQuestion + vbYesNo, "Calcul salaires") = vbYes Then
            Call CalculerVisitesEtSalaires
            salaireCalculeCeMois = True

            ' Proposer generation contrats
            If MsgBox("Salaires calcules !" & vbCrLf & vbCrLf & _
                      "Generer les contrats maintenant ?", _
                      vbQuestion + vbYesNo, "Generation contrats") = vbYes Then
                Call GenererContratsEnMasse
            End If
        Else
            ' Si l'utilisateur refuse, marquer comme fait pour eviter re-demande
            salaireCalculeCeMois = True
        End If
    End If

    On Error GoTo 0
End Sub
