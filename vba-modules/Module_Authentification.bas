Attribute VB_Name = "Module_Authentification"
' ============================================
' MODULE D'AUTHENTIFICATION
' Gestion des acces par mot de passe et consultation planning
' ============================================

Option Explicit

' Variables globales de session
Public utilisateurConnecte As String
Public niveauAcces As String ' "ADMIN" ou "GUIDE"
Public emailUtilisateur As String

' ============================================
' Fonction de connexion principale
' ============================================
Sub SeConnecter()
    Dim mdp As String
    Dim nomGuide As String
    Dim wsGuides As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim mdpAdmin As String

    ' Verifier que les feuilles existent
    On Error Resume Next
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    On Error GoTo 0

    If wsGuides Is Nothing Then
        MsgBox "Erreur : La feuille Guides n'existe pas." & vbCrLf & _
               "Veuillez initialiser le systeme d'abord.", vbCritical
        Exit Sub
    End If

    ' Recuperer le mot de passe admin depuis Configuration
    mdpAdmin = ObtenirConfig("MotDePasseAdmin", "admin123")

    ' Boite de dialogue pour le nom
    nomGuide = InputBox("Entrez votre nom de famille :" & vbCrLf & vbCrLf & _
                        "Pour l'administrateur, tapez : ADMIN", _
                        ">>> Connexion au systeme")

    If nomGuide = "" Then Exit Sub
    nomGuide = Trim(nomGuide)

    ' Boite de dialogue pour le mot de passe
    mdp = InputBox("Entrez votre mot de passe :", ">>> Authentification")
    If mdp = "" Then Exit Sub

    ' Verifier si c'est l'admin
    If UCase(nomGuide) = "ADMIN" Then
        If mdp = mdpAdmin Then
            utilisateurConnecte = "ADMIN"
            niveauAcces = "ADMIN"
            emailUtilisateur = ObtenirConfig("EmailAdmin", "")

            ' Afficher toutes les feuilles pour l'admin
            Call AfficherToutesFeuillesAdmin

            MsgBox "[OK] Bienvenue Administrateur !" & vbCrLf & vbCrLf & _
                   "Acces complet au systeme." & vbCrLf & _
                   "Vous pouvez gerer tous les plannings.", _
                   vbInformation, "Connexion reussie"

            AfficherInterfaceAdmin
            Exit Sub
        Else
            MsgBox "[ERREUR] Mot de passe administrateur incorrect.", vbCritical, "Erreur d'authentification"
            Exit Sub
        End If
    End If

    ' Verifier dans la liste des guides
    lastRow = wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        ' Comparer avec le nom (colonne B)
        If UCase(wsGuides.Cells(i, 2).Value) = UCase(nomGuide) Then
            ' Verifier le mot de passe (colonne E)
            If wsGuides.Cells(i, 5).Value = mdp Then
                utilisateurConnecte = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value ' Prenom + Nom
                niveauAcces = "GUIDE"
                emailUtilisateur = wsGuides.Cells(i, 3).Value ' Email

                MsgBox "[OK] Bienvenue " & utilisateurConnecte & " !" & vbCrLf & vbCrLf & _
                       "Acces a votre planning personnel.", _
                       vbInformation, "Connexion reussie"

                ' Afficher les vues filtrees du guide
                Call AfficherMesVisites(utilisateurConnecte)
                Call AfficherMesDisponibilites(utilisateurConnecte)
                Call AfficherPlanningGuide(utilisateurConnecte)
                Call AfficherListeGuidesLimitee

                ' Masquer les feuilles originales (securite)
                Call MasquerFeuillesOriginalesPourGuide

                Exit Sub
            Else
                MsgBox "[ERREUR] Mot de passe incorrect pour " & nomGuide & ".", _
                       vbCritical, "Erreur d'authentification"
                Exit Sub
            End If
        End If
    Next i

    MsgBox "[ERREUR] Utilisateur non trouve : " & nomGuide & vbCrLf & vbCrLf & _
           "Verifiez l'orthographe de votre nom.", _
           vbCritical, "Erreur"
End Sub

' ============================================
' Afficher le planning personnel d'un guide
' ============================================
Sub AfficherPlanningGuide(nomGuide As String)
    Dim wsPlanning As Worksheet
    Dim wsPlanningGuide As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ligneDestination As Long
    Dim dateVisite As Date
    Dim aujourdhui As Date

    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)
    aujourdhui = Date

    ' Supprimer l'ancienne feuille temporaire si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Mon_Planning").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Creer la feuille temporaire du guide
    Set wsPlanningGuide = ThisWorkbook.Sheets.Add
    wsPlanningGuide.Name = "Mon_Planning"

    ' Copier l'en-tete du planning
    wsPlanning.Rows(1).Copy wsPlanningGuide.Rows(1)

    ' Ajouter une colonne "Statut" et "Action"
    wsPlanningGuide.Cells(1, 7).Value = "Statut"
    wsPlanningGuide.Cells(1, 8).Value = "Action"

    ' Copier uniquement les lignes du guide (visites futures)
    lastRow = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
    ligneDestination = 2

    For i = 2 To lastRow
        ' Verifier si le guide est attribue a cette visite
        If InStr(1, UCase(wsPlanning.Cells(i, 5).Value), UCase(nomGuide), vbTextCompare) > 0 Then
            ' Verifier si la visite est dans le futur
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 1).Value)
            On Error GoTo 0

            If dateVisite >= aujourdhui Then
                ' Copier la ligne
                wsPlanning.Range(wsPlanning.Cells(i, 1), wsPlanning.Cells(i, 6)).Copy _
                    wsPlanningGuide.Range(wsPlanningGuide.Cells(ligneDestination, 1), wsPlanningGuide.Cells(ligneDestination, 6))

                ' Recuperer le statut de confirmation (colonne G du planning principal)
                Dim statutConfirmation As String
                statutConfirmation = wsPlanning.Cells(i, 7).Value

                If statutConfirmation = "" Then statutConfirmation = "En attente"

                wsPlanningGuide.Cells(ligneDestination, 7).Value = statutConfirmation

                ' Ajouter un bouton de confirmation selon le statut
                If statutConfirmation = "Confirme" Then
                    wsPlanningGuide.Cells(ligneDestination, 8).Value = "[OK] Confirme"
                    wsPlanningGuide.Cells(ligneDestination, 8).Interior.Color = RGB(198, 239, 206)
                ElseIf statutConfirmation = "Refuse" Then
                    wsPlanningGuide.Cells(ligneDestination, 8).Value = "[X] Refuse"
                    wsPlanningGuide.Cells(ligneDestination, 8).Interior.Color = RGB(255, 199, 206)
                Else
                    wsPlanningGuide.Cells(ligneDestination, 8).Value = "[!] A confirmer"
                    wsPlanningGuide.Cells(ligneDestination, 8).Interior.Color = RGB(255, 235, 156)
                End If

                ligneDestination = ligneDestination + 1
            End If
        End If
    Next i

    ' Mise en forme
    With wsPlanningGuide
        .Columns("A:A").ColumnWidth = 12  ' ID
        .Columns("B:B").ColumnWidth = 15  ' Date
        .Columns("C:C").ColumnWidth = 20  ' Heure
        .Columns("D:D").ColumnWidth = 25  ' Musee
        .Columns("E:E").ColumnWidth = 30  ' Type
        .Columns("F:F").ColumnWidth = 12  ' Duree
        .Columns("G:G").ColumnWidth = 15  ' Statut
        .Columns("H:H").ColumnWidth = 20  ' Action
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(70, 173, 71)
        .Range("A1:H1").Font.Color = RGB(255, 255, 255)
        .Range("A1:H1").HorizontalAlignment = xlCenter

        ' Bordures
        If ligneDestination > 2 Then
            .Range("A1:H" & ligneDestination - 1).Borders.LineStyle = xlContinuous
            .Range("A1:H" & ligneDestination - 1).Borders.Weight = xlThin
        End If
    End With

    ' Ajouter des boutons d'action
    AjouterBoutonsGuide wsPlanningGuide

    ' Activer la feuille
    wsPlanningGuide.Activate

    If ligneDestination = 2 Then
        MsgBox "[i] Vous n'avez aucune visite programmee a venir.", vbInformation, "Planning vide"
    Else
        MsgBox "[OK] Voici votre planning personnel." & vbCrLf & vbCrLf & _
               "Nombre de visites a venir : " & (ligneDestination - 2) & vbCrLf & vbCrLf & _
               "[!] Confirmez ou refusez chaque visite en cliquant sur les cellules de la colonne 'Action'.", _
               vbInformation, "Mon Planning"
    End If
End Sub

' ============================================
' Ajouter les boutons d'action pour le guide
' ============================================
Sub AjouterBoutonsGuide(ws As Worksheet)
    Dim btnConfirmer As Button
    Dim btnDeconnexion As Button
    Dim btnExporter As Button

    ' Calculer la largeur des colonnes en pixels (approximatif)
    Dim leftPos As Double
    leftPos = ws.Range("I1").Left  ' Position apres la colonne H

    ' Bouton Confirmer toutes les visites
    Set btnConfirmer = ws.Buttons.Add(leftPos, 10, 180, 30)
    With btnConfirmer
        .Caption = "[OK] Confirmer TOUTES mes visites"
        .OnAction = "ConfirmerToutesVisites"
    End With

    ' Bouton Deconnexion
    Set btnDeconnexion = ws.Buttons.Add(leftPos + 190, 10, 120, 30)
    With btnDeconnexion
        .Caption = "[>] Deconnexion"
        .OnAction = "SeDeconnecter"
    End With

    ' Bouton Exporter mon planning
    Set btnExporter = ws.Buttons.Add(leftPos + 320, 10, 140, 30)
    With btnExporter
        .Caption = " Exporter en PDF"
        .OnAction = "ExporterPlanningGuide"
    End With
End Sub

' ============================================
' Confirmer ou refuser une visite (clic sur cellule)
' ============================================
Sub ConfirmerOuRefuserVisite()
    Dim ws As Worksheet
    Dim wsPlanning As Worksheet
    Dim ligneSelectionnee As Long
    Dim dateVisite As String
    Dim heureVisite As String
    Dim typeVisite As String
    Dim reponse As VbMsgBoxResult
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    ' Verifier qu'on est sur la bonne feuille
    If ws.Name <> "Mon_Planning" Then
        MsgBox "Cette action n'est disponible que depuis votre planning personnel.", vbExclamation
        Exit Sub
    End If

    ligneSelectionnee = ActiveCell.Row

    If ligneSelectionnee < 2 Then Exit Sub

    ' Recuperer les infos de la visite
    dateVisite = ws.Cells(ligneSelectionnee, 1).Value
    heureVisite = ws.Cells(ligneSelectionnee, 2).Value
    typeVisite = ws.Cells(ligneSelectionnee, 3).Value

    ' Demander confirmation ou refus
    reponse = MsgBox("Visite du " & dateVisite & " a " & heureVisite & vbCrLf & _
                     "Type : " & typeVisite & vbCrLf & vbCrLf & _
                     "Voulez-vous CONFIRMER cette visite ?" & vbCrLf & _
                     "(Cliquez Non pour REFUSER)", _
                     vbYesNoCancel + vbQuestion, "Confirmation de visite")

    If reponse = vbCancel Then Exit Sub

    ' Mettre a jour dans le planning principal
    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)
    lastRow = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If wsPlanning.Cells(i, 1).Value = dateVisite And _
           wsPlanning.Cells(i, 2).Value = heureVisite And _
           InStr(1, UCase(wsPlanning.Cells(i, 5).Value), UCase(utilisateurConnecte), vbTextCompare) > 0 Then

            If reponse = vbYes Then
                wsPlanning.Cells(i, 7).Value = "Confirme"
                ws.Cells(ligneSelectionnee, 7).Value = "Confirme"
                ws.Cells(ligneSelectionnee, 8).Value = "[OK] Confirme"
                ws.Cells(ligneSelectionnee, 8).Interior.Color = RGB(198, 239, 206)
                MsgBox "[OK] Visite confirmee !" & vbCrLf & _
                       "L'administrateur en sera informe.", vbInformation
            Else
                ' REFUS -> REATTRIBUTION AUTOMATIQUE
                wsPlanning.Cells(i, 7).Value = "Refuse par " & utilisateurConnecte
                ws.Cells(ligneSelectionnee, 7).Value = "Refuse"
                ws.Cells(ligneSelectionnee, 8).Value = "[X] Refuse"
                ws.Cells(ligneSelectionnee, 8).Interior.Color = RGB(255, 199, 206)

                ' Lancer la reattribution automatique
                Dim nouveauGuide As String
                nouveauGuide = ReattribuerVisiteAutomatiquement(i, wsPlanning, utilisateurConnecte)

                If nouveauGuide <> "" Then
                    MsgBox "[X] Visite refusee." & vbCrLf & vbCrLf & _
                           "[OK] Le systeme a automatiquement reattribue cette visite a :" & vbCrLf & _
                           "   " & nouveauGuide & vbCrLf & vbCrLf & _
                           "Un email de notification lui sera envoye.", vbInformation, "Reattribution automatique"
                Else
                    MsgBox "[X] Visite refusee." & vbCrLf & vbCrLf & _
                           "[!] Aucun autre guide n'est disponible pour cette date." & vbCrLf & _
                           "L'administrateur en sera informe.", vbExclamation, "Pas de reattribution possible"
                End If
            End If

            Exit For
        End If
    Next i
End Sub

' ============================================
' Confirmer toutes les visites en attente
' ============================================
Sub ConfirmerToutesVisites()
    Dim ws As Worksheet
    Dim wsPlanning As Worksheet
    Dim lastRowGuide As Long
    Dim lastRowPlanning As Long
    Dim i As Long
    Dim j As Long
    Dim dateVisite As String
    Dim heureVisite As String
    Dim nbConfirmations As Long

    Set ws = ActiveSheet

    If ws.Name <> "Mon_Planning" Then
        MsgBox "Cette action n'est disponible que depuis votre planning personnel.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Voulez-vous confirmer TOUTES vos visites en attente ?", _
              vbYesNo + vbQuestion, "Confirmation globale") <> vbYes Then
        Exit Sub
    End If

    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)
    lastRowGuide = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastRowPlanning = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
    nbConfirmations = 0

    For i = 2 To lastRowGuide
        If ws.Cells(i, 7).Value <> "Confirme" And ws.Cells(i, 7).Value <> "Refuse" Then
            dateVisite = ws.Cells(i, 1).Value
            heureVisite = ws.Cells(i, 2).Value

            ' Trouver la ligne correspondante dans le planning principal
            For j = 2 To lastRowPlanning
                If wsPlanning.Cells(j, 1).Value = dateVisite And _
                   wsPlanning.Cells(j, 2).Value = heureVisite And _
                   InStr(1, UCase(wsPlanning.Cells(j, 5).Value), UCase(utilisateurConnecte), vbTextCompare) > 0 Then

                    wsPlanning.Cells(j, 7).Value = "Confirme"
                    ws.Cells(i, 7).Value = "Confirme"
                    ws.Cells(i, 8).Value = "[OK] Confirme"
                    ws.Cells(i, 8).Interior.Color = RGB(198, 239, 206)
                    nbConfirmations = nbConfirmations + 1
                    Exit For
                End If
            Next j
        End If
    Next i

    MsgBox "[OK] " & nbConfirmations & " visite(s) confirmee(s) !", vbInformation
End Sub

' ============================================
' Exporter le planning du guide en PDF
' ============================================
Sub ExporterPlanningGuide()
    Dim ws As Worksheet
    Dim cheminFichier As String

    Set ws = ActiveSheet

    If ws.Name <> "Mon_Planning" Then
        MsgBox "Cette action n'est disponible que depuis votre planning personnel.", vbExclamation
        Exit Sub
    End If

    cheminFichier = ThisWorkbook.Path & "\Planning_" & Replace(utilisateurConnecte, " ", "_") & "_" & Format(Date, "yyyymmdd") & ".pdf"

    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminFichier, Quality:=xlQualityStandard

    If Err.Number = 0 Then
        MsgBox "[OK] Planning exporte avec succes :" & vbCrLf & vbCrLf & _
               cheminFichier, vbInformation, "Export reussi"
    Else
        MsgBox "[X] Erreur lors de l'export PDF.", vbCritical
    End If
    On Error GoTo 0
End Sub

' ============================================
' Afficher l'interface administrateur
' ============================================
Sub AfficherInterfaceAdmin()
    Dim wsPlanning As Worksheet
    Dim btnDeconnexion As Button

    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)

    ' Supprimer l'ancien bouton de deconnexion s'il existe
    On Error Resume Next
    wsPlanning.Buttons("BtnDeconnexionAdmin").Delete
    On Error GoTo 0

    ' Creer le bouton de deconnexion pour l'admin
    Set btnDeconnexion = wsPlanning.Buttons.Add(10, 10, 150, 30)
    With btnDeconnexion
        .Name = "BtnDeconnexionAdmin"
        .Caption = "[X] Deconnexion Admin"
        .OnAction = "SeDeconnecter"
        .Font.Bold = True
    End With

    wsPlanning.Activate

    MsgBox "Interface administrateur activee." & vbCrLf & vbCrLf & _
           "Vous avez acces a :" & vbCrLf & _
           "- Tous les plannings" & vbCrLf & _
           "- Generation automatique" & vbCrLf & _
           "- Envoi d'emails" & vbCrLf & _
           "- Gestion des guides" & vbCrLf & _
           "- Statistiques et calculs" & vbCrLf & vbCrLf & _
           "Bouton [X] Deconnexion Admin disponible en haut a gauche.", _
           vbInformation, "Acces Admin"
End Sub

' ============================================
' Deconnexion
' ============================================
Sub SeDeconnecter()
    Dim ws As Worksheet

    ' Reinitialiser les variables de session
    utilisateurConnecte = ""
    niveauAcces = ""
    emailUtilisateur = ""

    ' Supprimer les boutons de deconnexion
    On Error Resume Next
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Buttons("BtnDeconnexionAdmin").Delete
    On Error GoTo 0

    ' Supprimer la feuille temporaire si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Mon_Planning").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Masquer toutes les feuilles sauf Accueil
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Accueil" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' Retourner a la feuille d'accueil
    ThisWorkbook.Sheets("Accueil").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Accueil").Activate

    MsgBox "Vous etes deconnecte(e)." & vbCrLf & _
           "A bientot !", vbInformation, "Deconnexion"
End Sub

' ============================================
' Verifier si l'utilisateur est admin
' ============================================
Function EstAdmin() As Boolean
    EstAdmin = (niveauAcces = "ADMIN")
End Function

' ============================================
' Obtenir une valeur de configuration
' ============================================
Function ObtenirConfig(nomParam As String, valeurDefaut As String) As String
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim i As Long

    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(FEUILLE_CONFIG)
    On Error GoTo 0

    If wsConfig Is Nothing Then
        ObtenirConfig = valeurDefaut
        Exit Function
    End If

    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If wsConfig.Cells(i, 1).Value = nomParam Then
            ObtenirConfig = wsConfig.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    ObtenirConfig = valeurDefaut
End Function

' ============================================
' Reattribuer automatiquement une visite refusee
' ============================================
Function ReattribuerVisiteAutomatiquement(ligneVisite As Long, wsPlanning As Worksheet, guideRefus As String) As String
    Dim wsDisponibilites As Worksheet
    Dim wsGuides As Worksheet
    Dim dateVisite As Date
    Dim heureVisite As String
    Dim guidesDisponibles As String
    Dim tabGuides() As String
    Dim i As Integer
    Dim nouveauGuide As String
    Dim lastRowDispo As Long
    Dim j As Long
    Dim guideTrouve As Boolean

    On Error Resume Next
    Set wsDisponibilites = ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES)
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    On Error GoTo 0

    If wsDisponibilites Is Nothing Or wsGuides Is Nothing Then
        ReattribuerVisiteAutomatiquement = ""
        Exit Function
    End If

    ' Recuperer les infos de la visite
    dateVisite = wsPlanning.Cells(ligneVisite, 1).Value
    heureVisite = wsPlanning.Cells(ligneVisite, 2).Value
    guidesDisponibles = wsPlanning.Cells(ligneVisite, 6).Value ' Colonne "Guides_Disponibles"

    ' Si pas de liste de guides disponibles, utiliser la fonction de recherche
    If guidesDisponibles = "" Or IsEmpty(guidesDisponibles) Then
        ' Chercher tous les guides disponibles pour cette date
        guidesDisponibles = ObtenirGuidesDisponiblesPourDate(dateVisite, heureVisite, guideRefus)
    End If

    ' Retirer le guide qui a refuse de la liste
    guidesDisponibles = Replace(guidesDisponibles, guideRefus, "")
    guidesDisponibles = Replace(guidesDisponibles, ",,", ",")
    guidesDisponibles = Trim(guidesDisponibles)
    If Left(guidesDisponibles, 1) = "," Then guidesDisponibles = Mid(guidesDisponibles, 2)
    If Right(guidesDisponibles, 1) = "," Then guidesDisponibles = Left(guidesDisponibles, Len(guidesDisponibles) - 1)

    ' Verifier s'il reste des guides disponibles
    If guidesDisponibles = "" Then
        ReattribuerVisiteAutomatiquement = ""
        Exit Function
    End If

    ' Separer les guides disponibles
    tabGuides = Split(guidesDisponibles, ",")

    ' Trouver le guide avec le moins de visites attribuees ce mois-ci
    nouveauGuide = ""
    Dim nbVisitesMin As Integer
    nbVisitesMin = 999

    For i = LBound(tabGuides) To UBound(tabGuides)
        Dim guideCourant As String
        guideCourant = Trim(tabGuides(i))

        If guideCourant <> "" Then
            Dim nbVisites As Integer
            nbVisites = CompterVisitesGuide(guideCourant, wsPlanning, dateVisite)

            If nbVisites < nbVisitesMin Then
                nbVisitesMin = nbVisites
                nouveauGuide = guideCourant
            End If
        End If
    Next i

    ' Si un nouveau guide est trouve, mettre a jour le planning
    If nouveauGuide <> "" Then
        wsPlanning.Cells(ligneVisite, 5).Value = nouveauGuide ' Colonne "Guide_Attribue"
        wsPlanning.Cells(ligneVisite, 7).Value = "En attente" ' Statut

        ' Marquer l'historique de reattribution
        wsPlanning.Cells(ligneVisite, 8).Value = "Reattribue de " & guideRefus & " a " & nouveauGuide & " le " & Format(Now, "dd/mm/yyyy hh:nn")

        ' TODO: Envoyer un email au nouveau guide (a implementer dans Module_Emails)
        ' Call EnvoyerNotificationReattribution(nouveauGuide, dateVisite, heureVisite, wsPlanning.Cells(ligneVisite, 3).Value)
    End If

    ReattribuerVisiteAutomatiquement = nouveauGuide
End Function

' ============================================
' Compter le nombre de visites d'un guide ce mois
' ============================================
Function CompterVisitesGuide(nomGuide As String, wsPlanning As Worksheet, dateReference As Date) As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim compteur As Integer
    Dim dateVisite As Date
    Dim moisReference As Integer
    Dim anneeReference As Integer

    moisReference = Month(dateReference)
    anneeReference = Year(dateReference)
    compteur = 0

    lastRow = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        On Error Resume Next
        dateVisite = CDate(wsPlanning.Cells(i, 1).Value)
        On Error GoTo 0

        ' Compter uniquement les visites du meme mois et confirmees ou en attente
        If Month(dateVisite) = moisReference And Year(dateVisite) = anneeReference Then
            If InStr(1, UCase(wsPlanning.Cells(i, 5).Value), UCase(nomGuide), vbTextCompare) > 0 Then
                Dim statut As String
                statut = wsPlanning.Cells(i, 7).Value
                If statut = "Confirme" Or statut = "En attente" Then
                    compteur = compteur + 1
                End If
            End If
        End If
    Next i

    CompterVisitesGuide = compteur
End Function

' ============================================
' Obtenir les guides disponibles pour une date/heure
' ============================================
Function ObtenirGuidesDisponiblesPourDate(dateVisite As Date, heureVisite As String, guideExclu As String) As String
    Dim wsDisponibilites As Worksheet
    Dim wsGuides As Worksheet
    Dim lastRowDispo As Long
    Dim lastRowGuides As Long
    Dim i As Long
    Dim j As Long
    Dim listeGuides As String
    Dim nomGuide As String
    Dim estDisponible As Boolean

    On Error Resume Next
    Set wsDisponibilites = ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES)
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    On Error GoTo 0

    If wsDisponibilites Is Nothing Or wsGuides Is Nothing Then
        ObtenirGuidesDisponiblesPourDate = ""
        Exit Function
    End If

    listeGuides = ""
    lastRowGuides = wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
    lastRowDispo = wsDisponibilites.Cells(wsDisponibilites.Rows.Count, 1).End(xlUp).Row

    ' Parcourir tous les guides
    For i = 2 To lastRowGuides
        nomGuide = wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value ' Prenom + Nom
        nomGuide = Trim(nomGuide)

        ' Exclure le guide qui a refuse
        If UCase(nomGuide) <> UCase(guideExclu) And nomGuide <> "" Then
            estDisponible = False

            ' Verifier la disponibilite dans la feuille Disponibilites
            For j = 2 To lastRowDispo
                Dim guideDispoNom As String
                guideDispoNom = wsDisponibilites.Cells(j, 1).Value ' Colonne avec nom du guide

                If InStr(1, UCase(guideDispoNom), UCase(nomGuide), vbTextCompare) > 0 Then
                    Dim dateDispo As Date
                    On Error Resume Next
                    dateDispo = CDate(wsDisponibilites.Cells(j, 2).Value)
                    On Error GoTo 0

                    If dateDispo = dateVisite Then
                        ' Verifier si disponible (colonne 3)
                        If UCase(wsDisponibilites.Cells(j, 3).Value) = "OUI" Or _
                           UCase(wsDisponibilites.Cells(j, 3).Value) = "DISPONIBLE" Then
                            estDisponible = True
                            Exit For
                        End If
                    End If
                End If
            Next j

            ' Ajouter a la liste si disponible
            If estDisponible Then
                If listeGuides = "" Then
                    listeGuides = nomGuide
                Else
                    listeGuides = listeGuides & "," & nomGuide
                End If
            End If
        End If
    Next i

    ObtenirGuidesDisponiblesPourDate = listeGuides
End Function

'===============================================================================
' FONCTION: AfficherToutesFeuillesAdmin
' DESCRIPTION: Affiche toutes les feuilles pour l'administrateur
'===============================================================================
Private Sub AfficherToutesFeuillesAdmin()
    On Error Resume Next

    ' Afficher toutes les feuilles pour l'admin
    ThisWorkbook.Sheets(FEUILLE_GUIDES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_VISITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CALCULS).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CONTRATS).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CONFIG).Visible = xlSheetVisible

    On Error GoTo 0
End Sub

'===============================================================================
' FONCTION: AfficherMesVisites
' DESCRIPTION: Affiche uniquement les visites assignees au guide connecte
'===============================================================================
Sub AfficherMesVisites(nomGuide As String)
    Dim wsVisites As Worksheet
    Dim wsMesVisites As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ligneDestination As Long

    On Error Resume Next
    Set wsVisites = ThisWorkbook.Sheets(FEUILLE_VISITES)
    On Error GoTo 0

    If wsVisites Is Nothing Then Exit Sub

    ' Supprimer l'ancienne feuille si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Mes_Visites").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Creer la nouvelle feuille
    Set wsMesVisites = ThisWorkbook.Sheets.Add
    wsMesVisites.Name = "Mes_Visites"
    wsMesVisites.Tab.Color = RGB(70, 173, 71)

    ' Copier l'en-tete
    wsVisites.Rows(1).Copy wsMesVisites.Rows(1)

    ' Formater l'en-tete
    With wsMesVisites.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(70, 173, 71)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Copier uniquement les visites du guide
    lastRow = wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
    ligneDestination = 2

    For i = 2 To lastRow
        ' Colonne 5 = Guide_Attribue (ou verifier selon votre structure)
        If InStr(1, UCase(wsVisites.Cells(i, 5).Value), UCase(nomGuide), vbTextCompare) > 0 Then
            wsVisites.Rows(i).Copy wsMesVisites.Rows(ligneDestination)
            ligneDestination = ligneDestination + 1
        End If
    Next i

    ' Ajuster les colonnes
    wsMesVisites.Columns.AutoFit

    ' Message si aucune visite
    If ligneDestination = 2 Then
        wsMesVisites.Range("A2").Value = "Aucune visite assignee pour le moment"
        wsMesVisites.Range("A2").Font.Italic = True
        wsMesVisites.Range("A2").Font.Color = RGB(150, 150, 150)
    End If
End Sub

'===============================================================================
' FONCTION: AfficherMesDisponibilites
' DESCRIPTION: Affiche uniquement les disponibilites du guide connecte
'===============================================================================
Sub AfficherMesDisponibilites(nomGuide As String)
    Dim wsDispos As Worksheet
    Dim wsMesDispos As Worksheet
    Dim wsGuides As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ligneDestination As Long
    Dim idGuide As Long

    On Error Resume Next
    Set wsDispos = ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES)
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    On Error GoTo 0

    If wsDispos Is Nothing Or wsGuides Is Nothing Then Exit Sub

    ' Trouver l'ID du guide (numero de ligne dans la feuille Guides)
    lastRow = wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
    idGuide = 0

    For i = 2 To lastRow
        If InStr(1, UCase(wsGuides.Cells(i, 1).Value & " " & wsGuides.Cells(i, 2).Value), UCase(nomGuide), vbTextCompare) > 0 Then
            idGuide = i - 1 ' ID commence a 1 (ligne 2 = ID 1)
            Exit For
        End If
    Next i

    If idGuide = 0 Then Exit Sub

    ' Supprimer l'ancienne feuille si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Mes_Disponibilites").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Creer la nouvelle feuille
    Set wsMesDispos = ThisWorkbook.Sheets.Add
    wsMesDispos.Name = "Mes_Disponibilites"
    wsMesDispos.Tab.Color = RGB(52, 152, 219)

    ' Copier l'en-tete
    wsDispos.Rows(1).Copy wsMesDispos.Rows(1)

    ' Formater l'en-tete
    With wsMesDispos.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Copier uniquement les disponibilites du guide
    lastRow = wsDispos.Cells(wsDispos.Rows.Count, 1).End(xlUp).Row
    ligneDestination = 2

    For i = 2 To lastRow
        ' Colonne 1 = ID_Guide
        If wsDispos.Cells(i, 1).Value = idGuide Then
            wsDispos.Rows(i).Copy wsMesDispos.Rows(ligneDestination)
            ligneDestination = ligneDestination + 1
        End If
    Next i

    ' Ajuster les colonnes
    wsMesDispos.Columns.AutoFit

    ' Message si aucune disponibilite
    If ligneDestination = 2 Then
        wsMesDispos.Range("A2").Value = "Aucune disponibilite enregistree"
        wsMesDispos.Range("A2").Font.Italic = True
        wsMesDispos.Range("A2").Font.Color = RGB(150, 150, 150)
    End If
End Sub

'===============================================================================
' FONCTION: AfficherListeGuidesLimitee
' DESCRIPTION: Affiche uniquement les noms des guides (pas infos privees)
'===============================================================================
Sub AfficherListeGuidesLimitee()
    Dim wsGuides As Worksheet
    Dim wsAnnuaire As Worksheet
    Dim lastRow As Long
    Dim i As Long

    On Error Resume Next
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    On Error GoTo 0

    If wsGuides Is Nothing Then Exit Sub

    ' Supprimer l'ancienne feuille si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Annuaire").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Creer la nouvelle feuille
    Set wsAnnuaire = ThisWorkbook.Sheets.Add
    wsAnnuaire.Name = "Annuaire"
    wsAnnuaire.Tab.Color = RGB(155, 89, 182)

    ' Creer l'en-tete (seulement Prenom et Nom)
    wsAnnuaire.Range("A1").Value = "Prenom"
    wsAnnuaire.Range("B1").Value = "Nom"

    ' Formater l'en-tete
    With wsAnnuaire.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(155, 89, 182)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Copier uniquement les prenoms et noms
    lastRow = wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        wsAnnuaire.Cells(i, 1).Value = wsGuides.Cells(i, 1).Value ' Prenom
        wsAnnuaire.Cells(i, 2).Value = wsGuides.Cells(i, 2).Value ' Nom
    Next i

    ' Ajuster les colonnes
    wsAnnuaire.Columns.AutoFit

    ' Ajouter un message informatif
    wsAnnuaire.Range("A" & lastRow + 2).Value = "[i] Pour contacter un collegue, demandez a l'administrateur"
    wsAnnuaire.Range("A" & lastRow + 2).Font.Italic = True
    wsAnnuaire.Range("A" & lastRow + 2).Font.Color = RGB(150, 150, 150)
End Sub

'===============================================================================
' FONCTION: MasquerFeuillesOriginalesPourGuide
' DESCRIPTION: Masque les feuilles sensibles pour les guides
'===============================================================================
Sub MasquerFeuillesOriginalesPourGuide()
    On Error Resume Next

    ' Masquer les feuilles originales (donnees completes)
    ThisWorkbook.Sheets(FEUILLE_VISITES).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_GUIDES).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Visible = xlSheetVeryHidden

    ' Les feuilles admin restent masquees (deja fait dans Module_Config)
    ThisWorkbook.Sheets(FEUILLE_CALCULS).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_CONTRATS).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_CONFIG).Visible = xlSheetVeryHidden

    ' Activer la feuille Mes_Visites
    ThisWorkbook.Sheets("Mes_Visites").Activate

    On Error GoTo 0
End Sub


