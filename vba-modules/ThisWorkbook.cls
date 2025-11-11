Option Explicit

' Variable pour eviter envois multiples
Private planningEnvoyeCeMois As Boolean
Private notificationsEnvoyeesAujourdhui As Boolean

Private Sub Workbook_Open()
    On Error Resume Next

    ' Masquer TOUTES les feuilles sauf Accueil par defaut
    Call MasquerToutesFeuillesParDefaut

    ' Activer la page d'accueil
    ThisWorkbook.Sheets("Accueil").Activate

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
    On Error Resume Next
    If niveauAcces = "GUIDE" Then
        ThisWorkbook.Sheets("Calculs_Paie").Visible = xlSheetVeryHidden
        ThisWorkbook.Sheets("Configuration").Visible = xlSheetVeryHidden
    ElseIf niveauAcces = "ADMIN" Then
        ThisWorkbook.Sheets("Calculs_Paie").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Configuration").Visible = xlSheetVisible
    End If
    On Error GoTo 0
End Sub

' Masquer toutes les feuilles sauf Accueil au demarrage
Private Sub MasquerToutesFeuillesParDefaut()
    On Error Resume Next

    Dim ws As Worksheet

    ' Boucle sur TOUTES les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        ' Masquer toutes sauf Accueil
        If ws.Name <> "Accueil" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' S'assurer que Accueil est visible
    ThisWorkbook.Sheets("Accueil").Visible = xlSheetVisible

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

    jourActuel = Day(Date)
    moisActuel = Month(Date)
    dernierJourDuMois = DateSerial(Year(Date), Month(Date) + 1, 0)

    ' 1. ENVOI PLANNING MENSUEL (1er du mois a 9h) - AUTOMATIQUE
    If jourActuel = 1 And Hour(Now) >= 9 And Not planningEnvoyeCeMois Then
        ' Envoi automatique du planning mensuel
        Call EnvoyerPlanningMensuel
        planningEnvoyeCeMois = True
    End If

    ' Reinitialiser le flag si changement de mois
    If jourActuel <> 1 Then
        planningEnvoyeCeMois = False
    End If

    ' 2. NOTIFICATIONS QUOTIDIENNES (8h-18h) - AUTOMATIQUE
    If Hour(Now) >= 8 And Hour(Now) < 18 And Not notificationsEnvoyeesAujourdhui Then
        ' Envoi automatique sans confirmation
        Call EnvoyerNotificationsAutomatiques
        notificationsEnvoyeesAujourdhui = True
    End If

    ' Reinitialiser notifications chaque jour
    If Hour(Now) < 8 Then
        notificationsEnvoyeesAujourdhui = False
    End If

    ' 3. CALCUL SALAIRES (dernier jour du mois a 17h)
    If Date = dernierJourDuMois And Hour(Now) >= 17 Then
        If MsgBox("C'est le dernier jour du mois !" & vbCrLf & vbCrLf & _
                  "Voulez-vous calculer les salaires automatiquement ?", _
                  vbQuestion + vbYesNo, "Calcul salaires") = vbYes Then
            Call CalculerVisitesEtSalaires

            ' Proposer generation contrats
            If MsgBox("Salaires calcules !" & vbCrLf & vbCrLf & _
                      "Generer les contrats maintenant ?", _
                      vbQuestion + vbYesNo, "Generation contrats") = vbYes Then
                Call GenererContratsEnMasse
            End If
        End If
    End If

    On Error GoTo 0
End Sub
