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

                ' Afficher les vues filtrees du guide (optimise - batch)
                Application.ScreenUpdating = False
                Application.EnableEvents = False

                ' Reafficher d'abord toutes les feuilles necessaires avant de creer les vues
                Call ReafficherFeuillesAvantConnexionGuide

                Call AfficherMesVisites(utilisateurConnecte)
                Call AfficherMesDisponibilites(utilisateurConnecte)
                Call AfficherPlanningGuide(utilisateurConnecte)
                ' Annuaire supprime pour raisons de confidentialite
                ' Call AfficherListeGuidesLimitee

                ' Masquer les feuilles originales (securite) et activer Mon_Planning
                Call MasquerFeuillesOriginalesPourGuide

                Application.EnableEvents = True
                Application.ScreenUpdating = True

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

    ' Utiliser la feuille Mon_Planning existante par son CodeName (Feuil3)
    On Error Resume Next
    Set wsPlanningGuide = Feuil3  ' Utiliser le CodeName direct au lieu du Name
    On Error GoTo 0

    ' Si la feuille n'existe pas, afficher une erreur (elle devrait toujours exister dans le fichier)
    If wsPlanningGuide Is Nothing Then
        MsgBox "ERREUR : La feuille Mon_Planning (Feuil3) n'existe pas dans le fichier !" & vbCrLf & _
               "Veuillez contacter l'administrateur.", vbCritical, "Erreur systeme"
        Exit Sub
    Else
        ' IMPORTANT : Rendre la feuille visible AVANT de l'utiliser
        wsPlanningGuide.Visible = xlSheetVisible

        ' S'assurer que le nom d'onglet est correct
        On Error Resume Next
        wsPlanningGuide.Name = "Mon_Planning"
        On Error GoTo 0

        ' Vider le contenu de la feuille existante (sauf les en-tetes)
        Application.EnableEvents = False
        If wsPlanningGuide.UsedRange.Rows.Count > 1 Then
            wsPlanningGuide.Rows("2:" & wsPlanningGuide.UsedRange.Rows.Count).Delete
        End If
        Application.EnableEvents = True
    End If

    ' Creer/Verifier les en-tetes (structure attendue - LECTURE SEULE)
    With wsPlanningGuide
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "Heure"
        .Cells(1, 3).Value = "Musee"
        .Cells(1, 4).Value = "Type_Visite"
        .Cells(1, 5).Value = "Duree"
        .Cells(1, 6).Value = "Langue"
        .Cells(1, 7).Value = "Nb_Personnes"
    End With

    ' Copier uniquement les lignes du guide (visites futures)
    lastRow = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
    ligneDestination = 2

    For i = 2 To lastRow
        ' Trouver la colonne Guide_Attribue (colonne 7 selon structure Excel)
        Dim guideAttribue As String
        guideAttribue = wsPlanning.Cells(i, 7).Value

        ' Verifier si le guide est attribue a cette visite
        If InStr(1, UCase(guideAttribue), UCase(nomGuide), vbTextCompare) > 0 Then
            ' Verifier si la visite est dans le futur
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)  ' COLONNE 2 = Date
            On Error GoTo 0

            If dateVisite >= aujourdhui Then
                ' Copier les donnees dans Mon_Planning (LECTURE SEULE - pas de confirmation requise)
                ' Structure Mon_Planning : Date, Heure, Musee, Type_Visite, Duree, Langue, Nb_Personnes
                ' Structure Planning REELLE : ID_Visite(1), Date(2), Heure(3), Musee(4), Type_Visite(5), Duree(6), Guide_Attribue(7), Niveau(8), Theme(9), Guides_Disponibles(10), Statut_Confirmation(11), Historique(12), Heure_Debut(13), Heure_Fin(14), Langue(15), Nb_Personnes(16)

                wsPlanningGuide.Cells(ligneDestination, 1).Value = wsPlanning.Cells(i, 2).Value ' Date
                wsPlanningGuide.Cells(ligneDestination, 2).Value = wsPlanning.Cells(i, 3).Value ' Heure
                wsPlanningGuide.Cells(ligneDestination, 3).Value = wsPlanning.Cells(i, 4).Value ' Musee
                wsPlanningGuide.Cells(ligneDestination, 4).Value = wsPlanning.Cells(i, 5).Value ' Type_Visite
                wsPlanningGuide.Cells(ligneDestination, 5).Value = wsPlanning.Cells(i, 6).Value ' Duree
                wsPlanningGuide.Cells(ligneDestination, 6).Value = wsPlanning.Cells(i, 15).Value ' Langue (col 13→15)
                wsPlanningGuide.Cells(ligneDestination, 7).Value = wsPlanning.Cells(i, 16).Value ' Nb_Personnes (col 14→16)

                ligneDestination = ligneDestination + 1
            End If
        End If
    Next i

    ' Mise en forme
    With wsPlanningGuide
        .Columns("A:A").ColumnWidth = 12  ' Date
        .Columns("B:B").ColumnWidth = 10  ' Heure
        .Columns("C:C").ColumnWidth = 25  ' Musee
        .Columns("D:D").ColumnWidth = 30  ' Type_Visite
        .Columns("E:E").ColumnWidth = 10  ' Duree
        .Columns("F:F").ColumnWidth = 10  ' Langue
        .Columns("G:G").ColumnWidth = 12  ' Nb_Personnes
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(70, 173, 71)
        .Range("A1:G1").Font.Color = RGB(255, 255, 255)
        .Range("A1:G1").HorizontalAlignment = xlCenter

        ' Bordures
        If ligneDestination > 2 Then
            .Range("A1:G" & ligneDestination - 1).Borders.LineStyle = xlContinuous
            .Range("A1:G" & ligneDestination - 1).Borders.Weight = xlThin
        End If
    End With

    ' Ajouter des boutons d'action
    AjouterBoutonsGuide wsPlanningGuide

    ' S'assurer que la feuille est visible (activation sera faite dans SeConnecter)
    wsPlanningGuide.Visible = xlSheetVisible

    If ligneDestination = 2 Then
        MsgBox "[i] Vous n'avez aucune visite programmee a venir.", vbInformation, "Planning vide"
    Else
        MsgBox "[OK] Voici votre planning personnel." & vbCrLf & vbCrLf & _
               "Nombre de visites a venir : " & (ligneDestination - 2) & vbCrLf & vbCrLf & _
               "[i] Pour toute modification, contactez l'administrateur.", _
               vbInformation, "Mon Planning"
    End If
End Sub

' ============================================
' Ajouter les boutons d'action pour le guide
' ============================================
Sub AjouterBoutonsGuide(ws As Worksheet)
    Dim btnDeconnexion As Button
    Dim btnExporter As Button
    Dim btn As Button

    ' Supprimer tous les anciens boutons de cette feuille
    On Error Resume Next
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0

    ' Calculer la largeur des colonnes en pixels (approximatif)
    Dim leftPos As Double
    leftPos = ws.Range("I1").Left  ' Position apres la colonne H (ne cache plus G et H)

    ' Bouton Deconnexion
    Set btnDeconnexion = ws.Buttons.Add(leftPos, 10, 120, 30)
    With btnDeconnexion
        .Caption = "[>] Deconnexion"
        .OnAction = "SeDeconnecter"
    End With

    ' Bouton Exporter mon planning
    Set btnExporter = ws.Buttons.Add(leftPos + 130, 10, 140, 30)
    With btnExporter
        .Caption = " Exporter en PDF"
        .OnAction = "ExporterPlanningGuide"
    End With
End Sub

' ============================================
' FONCTIONS DE CONFIRMATION SUPPRIMEES
' ============================================
' Les guides n'ont plus besoin de confirmer leurs visites.
' L'attribution par l'admin = engagement automatique.
' Pour toute modification, le guide doit contacter l'admin directement.

'===============================================================================
' FONCTION: RefuserEtReattribuerVisite (ADMIN SEULEMENT)
' DESCRIPTION: Permet a l'admin de refuser une visite et la reattribuer automatiquement
'===============================================================================
Sub RefuserEtReattribuerVisite()
    Dim ws As Worksheet
    Dim ligneSelectionnee As Long
    Dim dateVisite As Date
    Dim heureVisite As String
    Dim guideActuel As String
    Dim typeVisite As String
    Dim reponse As VbMsgBoxResult
    Dim nouveauGuide As String

    ' VERIFICATION : Seulement pour l'admin
    If niveauAcces <> "ADMIN" Then
        MsgBox "Cette fonction est reservee a l'administrateur." & vbCrLf & vbCrLf & _
               "Les guides ne peuvent pas modifier le planning." & vbCrLf & _
               "Contactez l'admin pour toute modification.", vbExclamation, "Acces refuse"
        Exit Sub
    End If

    Set ws = ActiveSheet

    ' Verifier qu'on est sur la feuille Planning
    If ws.Name <> FEUILLE_PLANNING Then
        MsgBox "Cette action doit etre effectuee depuis la feuille Planning.", vbExclamation
        Exit Sub
    End If

    ligneSelectionnee = ActiveCell.Row

    If ligneSelectionnee < 2 Then
        MsgBox "Veuillez selectionner une visite dans le planning.", vbInformation
        Exit Sub
    End If

    ' Recuperer les infos de la visite
    On Error Resume Next
    dateVisite = ws.Cells(ligneSelectionnee, 2).Value
    heureVisite = ws.Cells(ligneSelectionnee, 3).Value
    guideActuel = ws.Cells(ligneSelectionnee, 7).Value
    typeVisite = ws.Cells(ligneSelectionnee, 5).Value
    Dim musee As String
    musee = ws.Cells(ligneSelectionnee, 4).Value
    On Error GoTo 0

    If guideActuel = "" Or guideActuel = "NON ATTRIBUE" Then
        MsgBox "AUCUN GUIDE attribue pour cette visite :" & vbCrLf & vbCrLf & _
               "Date : " & Format(dateVisite, "dd/mm/yyyy") & vbCrLf & _
               "Heure : " & heureVisite & vbCrLf & _
               "Musee : " & musee & vbCrLf & _
               "Type : " & typeVisite & vbCrLf & vbCrLf & _
               "Impossible de refuser une visite sans guide." & vbCrLf & _
               "Veuillez d'abord attribuer un guide.", _
               vbExclamation, "Visite non attribuee"
        Exit Sub
    End If

    ' Confirmer le refus
    reponse = MsgBox("REFUSER cette visite et la REATTRIBUER automatiquement ?" & vbCrLf & vbCrLf & _
                     "Date : " & Format(dateVisite, "dd/mm/yyyy") & vbCrLf & _
                     "Heure : " & heureVisite & vbCrLf & _
                     "Type : " & typeVisite & vbCrLf & _
                     "Guide actuel : " & guideActuel & vbCrLf & vbCrLf & _
                     "Le systeme cherchera un autre guide disponible.", _
                     vbYesNo + vbQuestion, "Refus et Reattribution")

    If reponse <> vbYes Then Exit Sub

    ' Marquer comme refuse
    ws.Cells(ligneSelectionnee, 11).Value = "Refuse par Admin - " & guideActuel  ' Col 9→11 (Statut)
    ws.Cells(ligneSelectionnee, 11).Interior.Color = RGB(255, 199, 206) ' Colonne Statut en rouge

    ' Lancer la reattribution automatique
    nouveauGuide = ReattribuerVisiteAutomatiquement(ligneSelectionnee, ws, guideActuel)

    If nouveauGuide <> "" Then
        MsgBox "[OK] Visite reattribuee avec succes !" & vbCrLf & vbCrLf & _
               "Ancien guide : " & guideActuel & vbCrLf & _
               "Nouveau guide : " & nouveauGuide & vbCrLf & vbCrLf & _
               "Un email de notification sera envoye au nouveau guide.", _
               vbInformation, "Reattribution reussie"
    Else
        MsgBox "[!] ATTENTION : Aucun autre guide disponible trouve !" & vbCrLf & vbCrLf & _
               "La visite reste NON ATTRIBUEE." & vbCrLf & _
               "Vous devrez l'attribuer manuellement.", _
               vbExclamation, "Reattribution impossible"

        ws.Cells(ligneSelectionnee, 7).Value = "NON ATTRIBUE"
        ws.Cells(ligneSelectionnee, 11).Interior.Color = RGB(255, 100, 100)  ' Col 15→11 (Statut)
    End If
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
    Dim btnRefuser As Button

    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)

    ' Supprimer les anciens boutons s'ils existent
    On Error Resume Next
    wsPlanning.Buttons("BtnDeconnexionAdmin").Delete
    wsPlanning.Buttons("BtnRefuserReattribuer").Delete
    On Error GoTo 0

    ' Creer le bouton de deconnexion pour l'admin
    Set btnDeconnexion = wsPlanning.Buttons.Add(10, 800, 150, 30)
    With btnDeconnexion
        .Name = "BtnDeconnexionAdmin"
        .Caption = "[X] Deconnexion Admin"
        .OnAction = "SeDeconnecter"
        .Font.Bold = True
    End With

    ' Creer le bouton Refuser et Reattribuer
    Set btnRefuser = wsPlanning.Buttons.Add(170, 800, 200, 30)
    With btnRefuser
        .Name = "BtnRefuserReattribuer"
        .Caption = "[!] Refuser et Reattribuer"
        .OnAction = "RefuserEtReattribuerVisite"
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
           "[NOUVEAU] Bouton [!] Refuser et Reattribuer :" & vbCrLf & _
           "  Selectionnez une visite dans le Planning," & vbCrLf & _
           "  puis cliquez pour refuser et reattribuer automatiquement." & vbCrLf & vbCrLf & _
           "Bouton [X] Deconnexion Admin disponible en haut a gauche.", _
           vbInformation, "Acces Admin"
End Sub

' ============================================
' Deconnexion
' ============================================
Sub SeDeconnecter()
    Dim ws As Worksheet
    Dim wsPlanningGuide As Worksheet

    ' Reinitialiser les variables de session
    utilisateurConnecte = ""
    niveauAcces = ""
    emailUtilisateur = ""

    ' Supprimer les boutons de deconnexion
    On Error Resume Next
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Buttons("BtnDeconnexionAdmin").Delete
    On Error GoTo 0

    ' Vider la feuille Mon_Planning au lieu de la supprimer (pour conserver le code VBA)
    On Error Resume Next
    Set wsPlanningGuide = Feuil3  ' Utiliser le CodeName direct
    If Not wsPlanningGuide Is Nothing Then
        Application.EnableEvents = False
        wsPlanningGuide.Cells.Clear
        ' Recreer les en-tetes pour la prochaine connexion
        wsPlanningGuide.Cells(1, 1).Value = "Date"
        wsPlanningGuide.Cells(1, 2).Value = "Heure"
        wsPlanningGuide.Cells(1, 3).Value = "Musee"
        wsPlanningGuide.Cells(1, 4).Value = "Type_Visite"
        wsPlanningGuide.Cells(1, 5).Value = "Duree"
        wsPlanningGuide.Cells(1, 6).Value = "Langue"
        wsPlanningGuide.Cells(1, 7).Value = "Nb_Personnes"
        Application.EnableEvents = True

        ' Masquer la feuille apres deconnexion
        wsPlanningGuide.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    ' Masquer toutes les feuilles sauf Accueil - SECURISE
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Accueil" Then
            ' Verifier si la feuille n'est pas deja masquee pour eviter l'erreur 1004
            If ws.Visible <> xlSheetVeryHidden Then
                ws.Visible = xlSheetVeryHidden
            End If
        End If
    Next ws
    On Error GoTo 0

    ' Retourner a la feuille d'accueil
    On Error Resume Next
    ThisWorkbook.Sheets("Accueil").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Accueil").Activate
    On Error GoTo 0

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
    ' Variable guideTrouve supprimee (non utilisee)

    On Error Resume Next
    Set wsDisponibilites = ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES)
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    On Error GoTo 0

    If wsDisponibilites Is Nothing Or wsGuides Is Nothing Then
        ReattribuerVisiteAutomatiquement = ""
        Exit Function
    End If

    ' Recuperer les infos de la visite
    dateVisite = wsPlanning.Cells(ligneVisite, 2).Value
    heureVisite = wsPlanning.Cells(ligneVisite, 3).Value
    guidesDisponibles = wsPlanning.Cells(ligneVisite, 10).Value ' Colonne "Guides_Disponibles" (col 8→10)

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
        wsPlanning.Cells(ligneVisite, 7).Value = nouveauGuide ' Colonne "Guide_Attribue"
        wsPlanning.Cells(ligneVisite, 11).Value = "En attente" ' Statut (col 9→11)

        ' Marquer l'historique de reattribution
        wsPlanning.Cells(ligneVisite, 12).Value = "Reattribue de " & guideRefus & " a " & nouveauGuide & " le " & Format(Now, "dd/mm/yyyy hh:nn")  ' Historique (col 10→12)

        ' Notifier le nouveau guide par email
        EnvoyerNotificationReattribution nouveauGuide, dateVisite, heureVisite, wsPlanning.Cells(ligneVisite, 5).Value, guideRefus
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
        dateVisite = CDate(wsPlanning.Cells(i, 2).Value)
        On Error GoTo 0

        ' Compter uniquement les visites du meme mois et confirmees ou en attente
        If Month(dateVisite) = moisReference And Year(dateVisite) = anneeReference Then
            If InStr(1, UCase(wsPlanning.Cells(i, 7).Value), UCase(nomGuide), vbTextCompare) > 0 Then
                Dim statut As String
                statut = wsPlanning.Cells(i, 11).Value  ' Statut (col 9→11)
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
                guideDispoNom = wsDisponibilites.Cells(j, 1).Value ' Colonne Guide (contient le nom complet)

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

    Dim ws As Worksheet

    ' Afficher toutes les feuilles pour l'admin
    ThisWorkbook.Sheets(FEUILLE_GUIDES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_VISITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CALCULS).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CONTRATS).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_CONFIG).Visible = xlSheetVisible

    ' FORCER l'affichage de Specialisations
    On Error GoTo 0
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Specialisations")
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
    End If

    ' Aussi avec la constante au cas ou
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(FEUILLE_SPECIALISATIONS)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
    End If

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

    ' Utiliser la feuille existante au lieu de supprimer/recreer
    On Error Resume Next
    Set wsMesVisites = ThisWorkbook.Sheets("Mes_Visites")
    On Error GoTo 0

    ' Si la feuille n'existe pas, la creer
    If wsMesVisites Is Nothing Then
        Set wsMesVisites = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMesVisites.Name = "Mes_Visites"
    End If

    ' Rendre visible et vider le contenu
    wsMesVisites.Visible = xlSheetVisible
    wsMesVisites.Cells.Clear
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
        ' Structure Visites : ID_Visite(1), Date(2), Heure(3), Musee(4), Type_Visite(5), Duree_Heures(6), Nombre_Visiteurs(7), Statut(8)
        ' Note : La feuille Visites n'a pas de colonne Guide_Attribue, on ne peut pas filtrer ici
        ' Cette fonction devrait probablement utiliser la feuille Planning au lieu de Visites
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

    ' IMPORTANT : Rendre la feuille visible
    wsMesVisites.Visible = xlSheetVisible
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

    ' Utiliser la feuille existante au lieu de supprimer/recreer
    On Error Resume Next
    Set wsMesDispos = ThisWorkbook.Sheets("Mes_Disponibilites")
    On Error GoTo 0

    ' Si la feuille n'existe pas, la creer
    If wsMesDispos Is Nothing Then
        Set wsMesDispos = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMesDispos.Name = "Mes_Disponibilites"
    End If

    ' Rendre visible et vider le contenu
    wsMesDispos.Visible = xlSheetVisible
    wsMesDispos.Cells.Clear
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

    ' IMPORTANT : Rendre la feuille visible
    wsMesDispos.Visible = xlSheetVisible
End Sub

'===============================================================================
' FONCTION: AfficherListeGuidesLimitee
' DESCRIPTION: SUPPRIMEE - Confidentialite : les guides ne doivent pas voir la liste des autres guides
'===============================================================================
' Sub AfficherListeGuidesLimitee()
'     ' FONCTION DESACTIVEE POUR RAISONS DE CONFIDENTIALITE
'     ' La cliente ne souhaite pas que les guides voient la liste des autres guides
' End Sub

'===============================================================================
' FONCTION: ReafficherFeuillesAvantConnexionGuide
' DESCRIPTION: Reaffiche toutes les feuilles necessaires avant la connexion guide
'===============================================================================
Sub ReafficherFeuillesAvantConnexionGuide()
    On Error Resume Next

    ' Reafficher temporairement les feuilles source pour permettre la lecture des donnees
    ThisWorkbook.Sheets(FEUILLE_VISITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_GUIDES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Visible = xlSheetVisible

    ' Reafficher aussi les feuilles de destination pour guides
    Feuil3.Visible = xlSheetVisible  ' Mon_Planning
    Feuil3.Name = "Mon_Planning"

    ' Si les feuilles Mes_Visites et Mes_Disponibilites existent deja, les rendre visibles
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Mes_Visites" Or ws.Name = "Mes_Disponibilites" Then
            ws.Visible = xlSheetVisible
        End If
    Next ws

    On Error GoTo 0
End Sub

'===============================================================================
' FONCTION: MasquerFeuillesOriginalesPourGuide
' DESCRIPTION: Masque les feuilles sensibles pour les guides
'===============================================================================
Sub MasquerFeuillesOriginalesPourGuide()
    Dim ws As Worksheet

    On Error Resume Next

    ' Masquer la feuille d'accueil (plus necessaire apres connexion)
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("Accueil")
    If Not ws Is Nothing Then ws.Visible = xlSheetVeryHidden

    ' Masquer les feuilles originales (donnees completes) - SECURISE
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(FEUILLE_VISITES)
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES)
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(FEUILLE_PLANNING)
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    ' Les feuilles admin restent masquees (deja fait dans Module_Config)
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(FEUILLE_CALCULS)
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(FEUILLE_CONTRATS)
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(FEUILLE_CONFIG)
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    ' Masquer Specialisations pour les guides (visible uniquement pour admin)
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("Specialisations")
    If Not ws Is Nothing And ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

    ' S'assurer que toutes les feuilles du guide sont visibles
    Dim wsGuide As Worksheet
    For Each wsGuide In ThisWorkbook.Worksheets
        Select Case wsGuide.Name
            Case "Mon_Planning", "Mes_Visites", "Mes_Disponibilites"
                If wsGuide.Visible <> xlSheetVisible Then wsGuide.Visible = xlSheetVisible
        End Select
    Next wsGuide

    ' Supprimer l'Annuaire s'il existe (confidentialite)
    On Error Resume Next
    Dim wsAnnuaireTemp As Worksheet
    Set wsAnnuaireTemp = ThisWorkbook.Sheets("Annuaire")
    If Not wsAnnuaireTemp Is Nothing Then
        Application.DisplayAlerts = False
        wsAnnuaireTemp.Delete
        Application.DisplayAlerts = True
    End If
    Set wsAnnuaireTemp = ThisWorkbook.Sheets("Annuaire_Guides")
    If Not wsAnnuaireTemp Is Nothing Then
        Application.DisplayAlerts = False
        wsAnnuaireTemp.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Rendre Mon_Planning visible (feuille principale pour les guides)
    Set ws = Nothing
    Set ws = Feuil3
    If Not ws Is Nothing Then
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
        ws.Name = "Mon_Planning"
        ws.Activate
    End If

    On Error GoTo 0
End Sub


