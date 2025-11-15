Attribute VB_Name = "Module_Config"
'===============================================================================
' MODULE: Configuration Generale
' DESCRIPTION: Parametres globaux de l'application
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

' ===== CONSTANTES GLOBALES =====

' Noms des feuilles Excel
Public Const FEUILLE_ACCUEIL As String = "Accueil"
Public Const FEUILLE_GUIDES As String = "Guides"
Public Const FEUILLE_DISPONIBILITES As String = "Disponibilites"
Public Const FEUILLE_VISITES As String = "Visites"
Public Const FEUILLE_PLANNING As String = "Planning"
Public Const FEUILLE_CALCULS As String = "Calculs_Paie"
Public Const FEUILLE_CONTRATS As String = "Contrats"
Public Const FEUILLE_CONFIG As String = "Configuration"
Public Const FEUILLE_SPECIALISATIONS As String = "Specialisations"

' Configuration Email
Public Const DELAI_NOTIFICATION_1 As Integer = 7  ' Jours avant (premiere notification)
Public Const DELAI_NOTIFICATION_2 As Integer = 1  ' Jours avant (rappel)

' Configuration Paie (systeme de cachets - calcule automatiquement)

' Couleurs
Public Const COULEUR_DISPONIBLE As Long = 5296274     ' Vert clair
Public Const COULEUR_OCCUPE As Long = 15395562        ' Rouge clair
Public Const COULEUR_ASSIGNE As Long = 16777164       ' Bleu clair

'===============================================================================
' FONCTION: InitialiserApplication
' DESCRIPTION: Initialise la structure du classeur
'===============================================================================
Public Sub InitialiserApplication()
    On Error GoTo Erreur

    Application.ScreenUpdating = False

    ' Verifier/Creer les feuilles necessaires
    Call CreerFeuillesSiNonExistantes

    ' Configurer les plages nommees
    Call ConfigurerPlagesNommees

    ' Masquer les feuilles sensibles par defaut
    Call MasquerFeuillesSensibles

    ' Message de confirmation
    MsgBox "Application initialisee avec succes !" & vbCrLf & _
           "Toutes les feuilles sont pretes.", vbInformation, "Initialisation"

    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'initialisation : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: CreerFeuillesSiNonExistantes
' DESCRIPTION: Cree les feuilles manquantes
'===============================================================================
Private Sub CreerFeuillesSiNonExistantes()
    Dim feuilles() As String
    Dim i As Integer
    Dim ws As Worksheet
    Dim existe As Boolean

    ' Liste des feuilles necessaires
    feuilles = Split(FEUILLE_GUIDES & "," & FEUILLE_DISPONIBILITES & "," & _
                     FEUILLE_VISITES & "," & FEUILLE_PLANNING & "," & _
                     FEUILLE_CALCULS & "," & FEUILLE_CONTRATS & "," & _
                     FEUILLE_CONFIG, ",")

    ' Creer chaque feuille si elle n'existe pas
    For i = LBound(feuilles) To UBound(feuilles)
        existe = False

        ' Verifier si la feuille existe
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = feuilles(i) Then
                existe = True
                Exit For
            End If
        Next ws

        ' Creer si elle n'existe pas
        If Not existe Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = feuilles(i)
            Call InitialiserFeuille(ws)
        End If
    Next i
End Sub

'===============================================================================
' FONCTION: InitialiserFeuille
' DESCRIPTION: Configure les en-tetes d'une feuille selon son type
'===============================================================================
Private Sub InitialiserFeuille(ws As Worksheet)
    Dim rng As Range

    With ws
        Select Case .Name
            Case FEUILLE_GUIDES
                .Range("A1:F1").Value = Array("Prenom", "Nom", "Email", "Telephone", "Specialisations", "Mot_De_Passe")

            Case FEUILLE_DISPONIBILITES
                .Range("A1:D1").Value = Array("ID_Guide", "Date", "Disponible", "Commentaire")

            Case FEUILLE_VISITES
                .Range("A1:G1").Value = Array("ID_Visite", "Date", "Heure_Debut", "Heure_Fin", "Musee", "Type_Visite", "Nombre_Visiteurs")

            Case FEUILLE_PLANNING
                .Range("A1:H1").Value = Array("ID_Visite", "Date", "Heure", "Type_Visite", "Guide_Attribue", "Guides_Disponibles", "Statut_Confirmation", "Historique")

            Case FEUILLE_CALCULS
                ' Structure existante dans Excel (15 colonnes) - ne pas ecraser !
                ' A:Guide, B:Nombre_Visites, C:Total_Heures, D:Montant_Salaire, E:Prenom, F:Nom
                ' G:Nb_Visites, H:Nb_Heures, I:Total_Brut, J:Montant_Par_Cachet, K:Nb_Cachets
                ' L:Total_Recalcule, M:Mois, N:Defraiements, O:Total_Avec_Frais
                ' (En-tetes deja presents - ne rien faire)

            Case FEUILLE_CONTRATS
                .Range("A1:H1").Value = Array("ID_Guide", "Nom", "Mois", "Type_Contrat", "Dates_Visites", "Nb_Cachets", "Montant_Cachet", "Total")

            Case FEUILLE_CONFIG
                .Range("A1").Value = "Parametre"
                .Range("B1").Value = "Valeur"
                .Range("A2:B20").Value = Application.Transpose(Array( _
                    Array("Email_Expediteur", "votre.email@association.fr"), _
                    Array("Nom_Association", "Association des Guides"), _
                    Array("Notification_J7", "OUI"), _
                    Array("Notification_J1", "OUI"), _
                    Array("MotDePasseAdmin", "admin123"), _
                    Array("TARIF_1_VISITE", "80"), _
                    Array("TARIF_2_VISITES", "110"), _
                    Array("TARIF_3_VISITES", "140"), _
                    Array("TARIF_BRANLY_2H", "120"), _
                    Array("TARIF_BRANLY_3H", "150"), _
                    Array("TARIF_BRANLY_4H", "180"), _
                    Array("TARIF_HORSLEMURS_1", "100"), _
                    Array("TARIF_HORSLEMURS_2", "130"), _
                    Array("TARIF_HORSLEMURS_3", "160") _
                ))
        End Select

        ' Formater les en-tetes
        Set rng = .Range(.Cells(1, 1), .Cells(1, .UsedRange.Columns.Count))
        With rng
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

        .Columns.AutoFit
    End With
End Sub

'===============================================================================
' FONCTION: ConfigurerPlagesNommees
' DESCRIPTION: Cree les plages nommees pour faciliter les formules
'===============================================================================
Private Sub ConfigurerPlagesNommees()
    On Error Resume Next

    ' Supprimer les anciennes plages
    ThisWorkbook.Names("Liste_Guides").Delete
    ThisWorkbook.Names("Liste_Visites").Delete

    ' Creer les nouvelles plages (a adapter selon les donnees)
    With ThisWorkbook.Worksheets(FEUILLE_GUIDES)
        If .Range("A2").Value <> "" Then
            ThisWorkbook.Names.Add Name:="Liste_Guides", _
                RefersTo:=.Range("A2:F" & .Cells(.Rows.Count, 1).End(xlUp).Row)
        End If
    End With

    On Error GoTo 0
End Sub

'===============================================================================
' FONCTION: ObtenirConfigEmail
' DESCRIPTION: Retourne l'email de l'expediteur depuis Configuration
'===============================================================================
Public Function ObtenirConfigEmail() As String
    Dim ws As Worksheet
    Dim rng As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CONFIG)
    Set rng = ws.Range("A:A").Find("Email_Expediteur", LookIn:=xlValues, LookAt:=xlWhole)

    If Not rng Is Nothing Then
        ObtenirConfigEmail = ws.Cells(rng.Row, 2).Value
    Else
        ObtenirConfigEmail = ""
    End If
    On Error GoTo 0
End Function

'===============================================================================
' NOTE: ObtenirTarifHeure() supprimee - systeme de cachets remplace tarif horaire
' Utiliser CalculerVisitesEtSalaires() pour calcul automatique des cachets
'===============================================================================

'===============================================================================
' FONCTION: MasquerFeuillesSensibles
' DESCRIPTION: Masque les feuilles sensibles par defaut (securite)
'===============================================================================
Public Sub MasquerFeuillesSensibles()
    On Error Resume Next

    ' Masquer les feuilles sensibles (visibles uniquement pour l'admin)
    ThisWorkbook.Sheets(FEUILLE_CALCULS).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_CONFIG).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets(FEUILLE_CONTRATS).Visible = xlSheetVeryHidden

    ' Feuilles de travail visibles (mais protegees par code)
    ThisWorkbook.Sheets(FEUILLE_GUIDES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_DISPONIBILITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_VISITES).Visible = xlSheetVisible
    ThisWorkbook.Sheets(FEUILLE_PLANNING).Visible = xlSheetVisible

    ' Feuille d'accueil toujours visible
    If Not ThisWorkbook.Sheets(FEUILLE_ACCUEIL) Is Nothing Then
        ThisWorkbook.Sheets(FEUILLE_ACCUEIL).Visible = xlSheetVisible
    End If

    On Error GoTo 0
End Sub
