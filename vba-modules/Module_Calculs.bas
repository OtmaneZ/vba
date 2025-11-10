Attribute VB_Name = "Module_Calculs"
'===============================================================================
' MODULE: Calculs de Paie
' DESCRIPTION: Calcul automatique des visites et des salaires selon grille tarifaire client
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================
' LOGIQUE DE CALCUL:
' - Tarifs par JOURNEE selon nombre de visites effectuees le meme jour
' - 3 types de visites : Standards / Branly / Hors-les-murs
' - Standards (45min): 1 visite=80€, 2 visites=110€, 3 visites=140€
' - Branly (evenements): 2h=120€, 3h=150€, 4h=180€
' - Hors-les-murs (deplacements): 1 visite=100€, 2 visites=130€, 3 visites=160€
'===============================================================================

Option Explicit

'===============================================================================
' FONCTION: CalculerVisitesEtSalaires
' DESCRIPTION: Calcule le nombre de visites et le salaire pour chaque guide
'===============================================================================
Public Sub CalculerVisitesEtSalaires()
    Dim wsPlanning As Worksheet
    Dim wsCalculs As Worksheet
    Dim wsGuides As Worksheet
    Dim wsVisites As Worksheet
    Dim dictGuides As Object ' Dictionary pour chaque guide
    Dim dictJours As Object ' Dictionary pour chaque jour
    Dim i As Long
    Dim guideID As String
    Dim guideNom As String
    Dim nbVisitesTotal As Integer
    Dim montantSalaire As Double
    Dim ligneCalcul As Long
    Dim moisFiltre As String
    Dim dateVisite As Date
    Dim cleJour As String
    Dim typeVisite As String
    Dim dureeHeures As Double

    On Error GoTo Erreur

    ' Demander le mois a calculer (optionnel)
    moisFiltre = InputBox("Filtrer par mois (MM/AAAA) ou laisser vide pour tout:", "Periode de calcul", Format(Date, "mm/yyyy"))

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsCalculs = ThisWorkbook.Worksheets(FEUILLE_CALCULS)
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set dictGuides = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Effacer les anciens calculs (conserver les en-tetes)
    Dim derLigneCalcul As Long
    derLigneCalcul = wsCalculs.Cells(wsCalculs.Rows.Count, 1).End(xlUp).Row
    If derLigneCalcul > 1 Then
        wsCalculs.Range("A2:E" & derLigneCalcul).ClearContents
    End If

    ' Parcourir le planning et grouper par guide + jour
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = Trim(wsPlanning.Cells(i, 5).Value)

        ' Ignorer si non attribue
        If guideID <> "NON ATTRIBUE" And guideID <> "" Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            ' Filtrer par mois si specifie
            Dim inclure As Boolean
            inclure = True

            If moisFiltre <> "" And Err.Number = 0 Then
                Dim moisCible As Integer, anneeCible As Integer
                moisCible = CInt(Left(moisFiltre, 2))
                anneeCible = CInt(Right(moisFiltre, 4))

                If Month(dateVisite) <> moisCible Or Year(dateVisite) <> anneeCible Then
                    inclure = False
                End If
            End If

            If inclure And Err.Number = 0 Then
                ' Cle unique : Guide + Date
                cleJour = guideID & "|" & Format(dateVisite, "yyyy-mm-dd")
                
                ' Obtenir type et duree de la visite
                Dim idVisite As String
                idVisite = wsPlanning.Cells(i, 1).Value
                typeVisite = IdentifierTypeVisite(idVisite)
                dureeHeures = ObtenirDureeVisite(idVisite)

                ' Creer dictionnaire pour ce guide si n'existe pas
                If Not dictGuides.exists(guideID) Then
                    Set dictJours = CreateObject("Scripting.Dictionary")
                    dictGuides.Add guideID, dictJours
                Else
                    Set dictJours = dictGuides(guideID)
                End If

                ' Ajouter ou mettre a jour ce jour
                If Not dictJours.exists(cleJour) Then
                    ' Creer nouvelle journee: [date, type, nb_visites, duree_totale]
                    Dim infoJour As Variant
                    infoJour = Array(dateVisite, typeVisite, 1, dureeHeures)
                    dictJours.Add cleJour, infoJour
                Else
                    ' Incrementer les compteurs de cette journee
                    Dim temp As Variant
                    temp = dictJours(cleJour)
                    temp(2) = temp(2) + 1 ' Nombre de visites
                    temp(3) = temp(3) + dureeHeures ' Duree totale
                    dictJours(cleJour) = temp
                End If
            End If

            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    ' Calculer les salaires pour chaque guide
    ligneCalcul = 2
    Dim keyGuide As Variant
    Dim keyJour As Variant

    For Each keyGuide In dictGuides.Keys
        guideID = CStr(keyGuide)
        guideNom = ObtenirNomCompletGuide(guideID)
        
        Set dictJours = dictGuides(guideID)
        nbVisitesTotal = 0
        montantSalaire = 0

        ' Calculer le salaire pour chaque jour
        For Each keyJour In dictJours.Keys
            Dim infoJour As Variant
            infoJour = dictJours(keyJour)
            
            Dim nbVisitesJour As Integer
            Dim typeVisiteJour As String
            Dim dureeJour As Double
            Dim montantJour As Double
            
            typeVisiteJour = infoJour(1)
            nbVisitesJour = infoJour(2)
            dureeJour = infoJour(3)
            
            ' Calculer le montant pour cette journee
            montantJour = CalculerTarifJournee(typeVisiteJour, nbVisitesJour, dureeJour)
            
            nbVisitesTotal = nbVisitesTotal + nbVisitesJour
            montantSalaire = montantSalaire + montantJour
        Next keyJour

        ' Remplir la ligne dans Calculs_Paie
        wsCalculs.Cells(ligneCalcul, 1).Value = guideID
        wsCalculs.Cells(ligneCalcul, 2).Value = guideNom
        wsCalculs.Cells(ligneCalcul, 3).Value = nbVisitesTotal
        wsCalculs.Cells(ligneCalcul, 4).Value = dictJours.Count ' Nombre de jours travailles
        wsCalculs.Cells(ligneCalcul, 5).Value = montantSalaire
        wsCalculs.Cells(ligneCalcul, 5).NumberFormat = "#,##0.00 €"

        ' Formater
        If nbVisitesTotal > 0 Then
            wsCalculs.Rows(ligneCalcul).Interior.Color = COULEUR_DISPONIBLE
        End If

        ligneCalcul = ligneCalcul + 1
    Next keyGuide

    ' Ajouter une ligne de total
    If ligneCalcul > 2 Then
        wsCalculs.Cells(ligneCalcul, 2).Value = "TOTAL"
        wsCalculs.Cells(ligneCalcul, 2).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 3).Formula = "=SUM(C2:C" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 3).Font.Bold = True
        
        wsCalculs.Cells(ligneCalcul, 4).Formula = "=SUM(D2:D" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 4).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 5).Formula = "=SUM(E2:E" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 5).NumberFormat = "#,##0.00 €"
        wsCalculs.Cells(ligneCalcul, 5).Font.Bold = True

        wsCalculs.Rows(ligneCalcul).Interior.Color = RGB(255, 242, 204)
    End If

    wsCalculs.Columns.AutoFit
    Application.ScreenUpdating = True

    Dim msgPeriode As String
    If moisFiltre <> "" Then
        msgPeriode = " pour " & moisFiltre
    Else
        msgPeriode = " (toutes periodes)"
    End If

    MsgBox "Calculs effectues avec succès" & msgPeriode & " !" & vbCrLf & vbCrLf & _
           "Nombre de guides : " & dictGuides.Count, _
           vbInformation, "Calculs Paie"

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors des calculs : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: IdentifierTypeVisite
' DESCRIPTION: Identifie si c'est une visite Standard, Branly ou Hors-les-murs
'===============================================================================
Private Function IdentifierTypeVisite(idVisite As String) As String
    Dim wsVisites As Worksheet
    Dim i As Long
    Dim nomVisite As String
    
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    IdentifierTypeVisite = "STANDARD" ' Par defaut
    
    ' Chercher la visite
    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
        If wsVisites.Cells(i, 1).Value = idVisite Then
            nomVisite = UCase(Trim(wsVisites.Cells(i, 2).Value))
            
            ' Identifier le type
            If InStr(nomVisite, "BRANLY") > 0 Or _
               InStr(nomVisite, "EVENEMENT BRANLY") > 0 Then
                IdentifierTypeVisite = "BRANLY"
            ElseIf InStr(nomVisite, "HORS-LES-MURS") > 0 Or _
                   InStr(nomVisite, "HORS LES MURS") > 0 Or _
                   InStr(nomVisite, "HORSLEMURS") > 0 Or _
                   InStr(nomVisite, "VISIO") > 0 Then
                IdentifierTypeVisite = "HORSLEMURS"
            End If
            
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' FONCTION: CalculerTarifJournee
' DESCRIPTION: Calcule le tarif pour une journee selon le type et nb de visites
'===============================================================================
Private Function CalculerTarifJournee(typeVisite As String, nbVisites As Integer, dureeHeures As Double) As Double
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Worksheets("Configuration")
    
    ' Valeurs par defaut si parametres non trouves
    CalculerTarifJournee = 0
    
    Select Case UCase(typeVisite)
        Case "STANDARD"
            ' Tarifs standards: 80€/110€/140€
            Select Case nbVisites
                Case 1
                    CalculerTarifJournee = LireParametreConfig("TARIF_1_VISITE", 80)
                Case 2
                    CalculerTarifJournee = LireParametreConfig("TARIF_2_VISITES", 110)
                Case Is >= 3
                    CalculerTarifJournee = LireParametreConfig("TARIF_3_VISITES", 140)
            End Select
            
        Case "BRANLY"
            ' Tarifs Branly selon duree: 2h=120€, 3h=150€, 4h=180€
            If dureeHeures <= 2 Then
                CalculerTarifJournee = LireParametreConfig("TARIF_BRANLY_2H", 120)
            ElseIf dureeHeures <= 3 Then
                CalculerTarifJournee = LireParametreConfig("TARIF_BRANLY_3H", 150)
            Else
                CalculerTarifJournee = LireParametreConfig("TARIF_BRANLY_4H", 180)
            End If
            
        Case "HORSLEMURS"
            ' Tarifs hors-les-murs: 100€/130€/160€
            Select Case nbVisites
                Case 1
                    CalculerTarifJournee = LireParametreConfig("TARIF_HORSLEMURS_1", 100)
                Case 2
                    CalculerTarifJournee = LireParametreConfig("TARIF_HORSLEMURS_2", 130)
                Case Is >= 3
                    CalculerTarifJournee = LireParametreConfig("TARIF_HORSLEMURS_3", 160)
            End Select
    End Select
End Function

'===============================================================================
' FONCTION: LireParametreConfig
' DESCRIPTION: Lit un parametre dans la feuille Configuration
'===============================================================================
Private Function LireParametreConfig(nomParam As String, valeurDefaut As Double) As Double
    Dim wsConfig As Worksheet
    Dim i As Long
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("Configuration")
    
    If wsConfig Is Nothing Then
        LireParametreConfig = valeurDefaut
        Exit Function
    End If
    
    ' Chercher le parametre
    For i = 1 To wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
        If Trim(UCase(wsConfig.Cells(i, 1).Value)) = UCase(nomParam) Then
            Dim valeur As Double
            valeur = wsConfig.Cells(i, 2).Value
            
            If valeur > 0 Then
                LireParametreConfig = valeur
            Else
                LireParametreConfig = valeurDefaut
            End If
            
            Exit Function
        End If
    Next i
    
    ' Si non trouve, retourner valeur par defaut
    LireParametreConfig = valeurDefaut
    On Error GoTo 0
End Function

'===============================================================================
' FONCTION: ObtenirDureeVisite
' DESCRIPTION: Calcule la duree d'une visite en heures
'===============================================================================
Private Function ObtenirDureeVisite(idVisite As String) As Double
    Dim wsVisites As Worksheet
    Dim i As Long
    Dim heureDebut As Date
    Dim heureFin As Date

    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    ObtenirDureeVisite = 1 ' Duree par defaut (1 heure)

    ' Chercher la visite
    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
        If wsVisites.Cells(i, 1).Value = idVisite Then
            On Error Resume Next
            heureDebut = CDate(wsVisites.Cells(i, 3).Value)
            heureFin = CDate(wsVisites.Cells(i, 4).Value)

            If Err.Number = 0 And heureFin > heureDebut Then
                ' Calculer la difference en heures
                ObtenirDureeVisite = (heureFin - heureDebut) * 24
            End If

            On Error GoTo 0
            Exit Function
        End If
    Next i
End Function


'===============================================================================
' FONCTION: ObtenirNomCompletGuide
' DESCRIPTION: Retourne le nom complet d'un guide
'===============================================================================
Private Function ObtenirNomCompletGuide(guideID As String) As String
    Dim wsGuides As Worksheet
    Dim i As Long

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    ObtenirNomCompletGuide = guideID ' Par defaut retourner l'ID

    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If Trim(wsGuides.Cells(i, 1).Value) = Trim(guideID) Then
            ObtenirNomCompletGuide = Trim(wsGuides.Cells(i, 2).Value) & " " & Trim(wsGuides.Cells(i, 3).Value)
            Exit Function
        End If
    Next i
End Function


'===============================================================================
' FONCTION: GenererFichePaieGuide
' DESCRIPTION: Genere une fiche de paie detaillee pour un guide
'===============================================================================
Public Sub GenererFichePaieGuide()
    Dim guideID As String
    Dim guideNom As String
    Dim wsPlanning As Worksheet
    Dim wsVisites As Worksheet
    Dim wbFiche As Workbook
    Dim wsFiche As Worksheet
    Dim i As Long
    Dim ligne As Long
    Dim moisFiltre As String
    Dim dateVisite As Date
    Dim totalVisites As Integer
    Dim totalMontant As Double
    Dim fichier As String
    Dim dictJours As Object
    Dim cleJour As String
    Dim typeVisite As String
    Dim dureeHeures As Double

    On Error GoTo Erreur

    ' Demander l'ID du guide
    guideID = InputBox("Entrez l'ID du guide:", "Fiche de paie")
    If guideID = "" Then Exit Sub

    ' Verifier que le guide existe
    guideNom = ObtenirNomCompletGuide(guideID)
    If guideNom = guideID Then
        MsgBox "Guide non trouve.", vbExclamation
        Exit Sub
    End If

    ' Demander le mois
    moisFiltre = InputBox("Mois (MM/AAAA):", "Periode", Format(Date, "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    Dim moisCible As Integer, anneeCible As Integer
    moisCible = CInt(Left(moisFiltre, 2))
    anneeCible = CInt(Right(moisFiltre, 4))

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set dictJours = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Grouper les visites par jour
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        If Trim(wsPlanning.Cells(i, 5).Value) = Trim(guideID) Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 Then
                If Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                    cleJour = Format(dateVisite, "yyyy-mm-dd")
                    
                    Dim idVisite As String
                    idVisite = wsPlanning.Cells(i, 1).Value
                    typeVisite = IdentifierTypeVisite(idVisite)
                    dureeHeures = ObtenirDureeVisite(idVisite)

                    If Not dictJours.exists(cleJour) Then
                        Dim infoJour As Variant
                        infoJour = Array(dateVisite, typeVisite, 1, dureeHeures)
                        dictJours.Add cleJour, infoJour
                    Else
                        Dim temp As Variant
                        temp = dictJours(cleJour)
                        temp(2) = temp(2) + 1
                        temp(3) = temp(3) + dureeHeures
                        dictJours(cleJour) = temp
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    If dictJours.Count = 0 Then
        MsgBox "Aucune visite trouvee pour ce guide ce mois-ci.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Créer un nouveau classeur pour la fiche
    Set wbFiche = Workbooks.Add
    Set wsFiche = wbFiche.Worksheets(1)
    wsFiche.Name = "Fiche_Paie"

    ' En-tete de la fiche
    With wsFiche
        .Range("A1").Value = "FICHE DE PAIE"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True

        .Range("A3").Value = "Guide :"
        .Range("B3").Value = guideNom & " (" & guideID & ")"
        .Range("B3").Font.Bold = True

        .Range("A4").Value = "Periode :"
        .Range("B4").Value = Format(DateSerial(anneeCible, moisCible, 1), "mmmm yyyy")

        .Range("A5").Value = "Date d'edition :"
        .Range("B5").Value = Format(Date, "dd/mm/yyyy")

        .Range("A7:F7").Value = Array("Date", "Type", "Nb visites", "Duree (h)", "Tarif journee", "Montant")
        .Range("A7:F7").Font.Bold = True
        .Range("A7:F7").Interior.Color = RGB(68, 114, 196)
        .Range("A7:F7").Font.Color = RGB(255, 255, 255)
    End With

    ' Lister les journees
    ligne = 8
    totalVisites = 0
    totalMontant = 0
    
    Dim keyJour As Variant
    For Each keyJour In dictJours.Keys
        Dim infoJour As Variant
        infoJour = dictJours(keyJour)
        
        Dim nbVisitesJour As Integer
        Dim typeJour As String
        Dim dureeJour As Double
        Dim montantJour As Double
        
        dateVisite = infoJour(0)
        typeJour = infoJour(1)
        nbVisitesJour = infoJour(2)
        dureeJour = infoJour(3)
        montantJour = CalculerTarifJournee(typeJour, nbVisitesJour, dureeJour)

        wsFiche.Cells(ligne, 1).Value = Format(dateVisite, "dd/mm/yyyy")
        wsFiche.Cells(ligne, 2).Value = typeJour
        wsFiche.Cells(ligne, 3).Value = nbVisitesJour
        wsFiche.Cells(ligne, 4).Value = dureeJour
        wsFiche.Cells(ligne, 5).Value = montantJour & " €"
        wsFiche.Cells(ligne, 6).Value = montantJour
        wsFiche.Cells(ligne, 6).NumberFormat = "#,##0.00 €"

        totalVisites = totalVisites + nbVisitesJour
        totalMontant = totalMontant + montantJour
        ligne = ligne + 1
    Next keyJour

    ' Totaux
    wsFiche.Cells(ligne, 2).Value = "TOTAL"
    wsFiche.Cells(ligne, 2).Font.Bold = True
    wsFiche.Cells(ligne, 3).Value = totalVisites
    wsFiche.Cells(ligne, 3).Font.Bold = True
    wsFiche.Cells(ligne, 6).Value = totalMontant
    wsFiche.Cells(ligne, 6).NumberFormat = "#,##0.00 €"
    wsFiche.Cells(ligne, 6).Font.Bold = True
    wsFiche.Range("A" & ligne & ":F" & ligne).Interior.Color = RGB(255, 242, 204)

    ' Informations supplementaires
    ligne = ligne + 2
    wsFiche.Cells(ligne, 1).Value = "Nombre de jours travailles :"
    wsFiche.Cells(ligne, 2).Value = dictJours.Count

    wsFiche.Columns.AutoFit

    ' Proposer de sauvegarder
    fichier = Application.GetSaveAsFilename("Fiche_Paie_" & guideID & "_" & Format(DateSerial(anneeCible, moisCible, 1), "yyyymm") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")
    If fichier <> "False" Then
        wbFiche.SaveAs fichier
        MsgBox "Fiche de paie generee avec succès !" & vbCrLf & fichier, vbInformation
    End If

    wbFiche.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbFiche Is Nothing Then wbFiche.Close SaveChanges:=False
    MsgBox "Erreur lors de la generation de la fiche : " & Err.Description, vbCritical
End Sub


'===============================================================================
' FONCTION: ExporterRecapitulatifPaie
' DESCRIPTION: Exporte un recapitulatif de paie pour tous les guides
'===============================================================================
Public Sub ExporterRecapitulatifPaie()
    Dim wsCalculs As Worksheet
    Dim fichier As String
    Dim wbExport As Workbook

    On Error GoTo Erreur

    ' D'abord calculer
    Call CalculerVisitesEtSalaires

    Set wsCalculs = ThisWorkbook.Worksheets(FEUILLE_CALCULS)

    Application.ScreenUpdating = False

    ' Créer un nouveau classeur
    Set wbExport = Workbooks.Add

    ' Copier les calculs
    wsCalculs.UsedRange.Copy
    wbExport.Worksheets(1).Range("A1").PasteSpecial xlPasteAll
    wbExport.Worksheets(1).Name = "Recapitulatif_Paie"
    wbExport.Worksheets(1).Columns.AutoFit

    Application.CutCopyMode = False

    ' Proposer de sauvegarder
    fichier = Application.GetSaveAsFilename("Recapitulatif_Paie_" & Format(Date, "yyyymmdd") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")
    If fichier <> "False" Then
        wbExport.SaveAs fichier
        MsgBox "Recapitulatif exporte avec succès !" & vbCrLf & fichier, vbInformation
    End If

    wbExport.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    MsgBox "Erreur lors de l'export : " & Err.Description, vbCritical
End Sub


