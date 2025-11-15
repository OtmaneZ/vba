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
' - Standards (45min): 1 visite=80, 2 visites=110, 3 visites=140
' - Branly (evenements): 2h=120, 3h=150, 4h=180
' - Hors-les-murs (deplacements): 1 visite=100, 2 visites=130, 3 visites=160
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
    Dim infoJour As Variant
    Dim temp As Variant

    On Error GoTo Erreur

    ' Demander le mois a calculer (optionnel)
    moisFiltre = InputBox("Filtrer par mois (MM/AAAA) ou laisser vide pour tout:", "Periode de calcul", Format(Date, "mm/yyyy"))

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsCalculs = ThisWorkbook.Worksheets(FEUILLE_CALCULS)
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set dictGuides = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Effacer les anciens calculs (conserver les en-tetes) - INCLUT colonnes F et G
    Dim derLigneCalcul As Long
    derLigneCalcul = wsCalculs.Cells(wsCalculs.Rows.Count, 1).End(xlUp).Row
    If derLigneCalcul > 1 Then
        wsCalculs.Range("A2:G" & derLigneCalcul).ClearContents
    End If

    ' Parcourir le planning et grouper par guide + jour
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = Trim(wsPlanning.Cells(i, 7).Value) ' Guide_Attribue (Col 7)

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
                    infoJour = Array(dateVisite, typeVisite, 1, dureeHeures)
                    dictJours.Add cleJour, infoJour
                Else
                    ' Incrementer les compteurs de cette journee
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

    ' Calculer les salaires pour chaque guide AVEC SYSTEME DE CACHETS
    ligneCalcul = 2
    Dim keyGuide As Variant
    Dim keyJour As Variant

    For Each keyGuide In dictGuides.Keys
        guideID = CStr(keyGuide)
        guideNom = ObtenirNomCompletGuide(guideID)

        Set dictJours = dictGuides(guideID)
        nbVisitesTotal = 0
        montantSalaire = 0
        Dim nbJoursTravailles As Integer
        nbJoursTravailles = dictJours.Count

        ' Calculer le montant TOTAL du mois
        For Each keyJour In dictJours.Keys
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

        ' SYSTEME DE CACHETS : Montant par cachet = Total  Nb jours (arrondi sup)
        Dim montantParCachet As Double
        If nbJoursTravailles > 0 Then
            montantParCachet = Application.WorksheetFunction.RoundUp(montantSalaire / nbJoursTravailles, 2)
        Else
            montantParCachet = 0
        End If

        ' Remplir la ligne dans Calculs_Paie
        wsCalculs.Cells(ligneCalcul, 1).Value = guideID
        wsCalculs.Cells(ligneCalcul, 2).Value = guideNom
        wsCalculs.Cells(ligneCalcul, 3).Value = nbVisitesTotal
        wsCalculs.Cells(ligneCalcul, 4).Value = nbJoursTravailles ' Nombre de jours = Nombre de cachets
        wsCalculs.Cells(ligneCalcul, 5).Value = montantSalaire
        wsCalculs.Cells(ligneCalcul, 5).NumberFormat = "#,##0.00 "

        ' NOUVELLE COLONNE F : Montant par cachet
        wsCalculs.Cells(ligneCalcul, 6).Value = montantParCachet
        wsCalculs.Cells(ligneCalcul, 6).NumberFormat = "#,##0.00 "

        ' NOUVELLE COLONNE G : Total recalcule (cachets  montant)
        Dim totalRecalcule As Double
        totalRecalcule = montantParCachet * nbJoursTravailles
        wsCalculs.Cells(ligneCalcul, 7).Value = totalRecalcule
        wsCalculs.Cells(ligneCalcul, 7).NumberFormat = "#,##0.00 "

        ' COLONNE N : Defraiements (initialise a 0, a remplir manuellement)
        If IsEmpty(wsCalculs.Cells(ligneCalcul, 14).Value) Then
            wsCalculs.Cells(ligneCalcul, 14).Value = 0
            wsCalculs.Cells(ligneCalcul, 14).NumberFormat = "#,##0.00 "
        End If

        ' COLONNE O : Total avec frais (Total_Brut + Defraiements)
        ' Formule Excel automatique : =I[ligne]+N[ligne]
        wsCalculs.Cells(ligneCalcul, 15).Formula = "=I" & ligneCalcul & "+N" & ligneCalcul
        wsCalculs.Cells(ligneCalcul, 15).NumberFormat = "#,##0.00 "

        ' Formater
        If nbVisitesTotal > 0 Then
            wsCalculs.Rows(ligneCalcul).Interior.Color = COULEUR_DISPONIBLE
        End If

        ligneCalcul = ligneCalcul + 1
    Next keyGuide

    ' Mettre a jour les en-tetes si besoin
    If wsCalculs.Cells(1, 6).Value = "" Then
        wsCalculs.Cells(1, 6).Value = "Montant/Cachet"
        wsCalculs.Cells(1, 6).Font.Bold = True
    End If
    If wsCalculs.Cells(1, 7).Value = "" Then
        wsCalculs.Cells(1, 7).Value = "Total Recalcule"
        wsCalculs.Cells(1, 7).Font.Bold = True
    End If

    ' Ajouter une ligne de total
    If ligneCalcul > 2 Then
        wsCalculs.Cells(ligneCalcul, 2).Value = "TOTAL"
        wsCalculs.Cells(ligneCalcul, 2).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 3).Formula = "=SUM(C2:C" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 3).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 4).Formula = "=SUM(D2:D" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 4).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 5).Formula = "=SUM(E2:E" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 5).NumberFormat = "#,##0.00 "
        wsCalculs.Cells(ligneCalcul, 5).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 7).Formula = "=SUM(G2:G" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 7).NumberFormat = "#,##0.00 "
        wsCalculs.Cells(ligneCalcul, 7).Font.Bold = True

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

    MsgBox "Calculs effectues avec succes" & msgPeriode & " !" & vbCrLf & vbCrLf & _
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
    '===============================================================================
    ' FONCTION: IdentifierTypeVisite
    ' DESCRIPTION: Identifie le type de visite depuis colonne Type_Prestation
    ' PARAMETRE: idVisite - ID de la visite (ex: V001)
    ' RETOUR: Type de prestation (VISITE CONTEE BRANLY, MARINE, HORS LES MURS, etc.)
    '===============================================================================
    Dim wsVisites As Worksheet
    Dim i As Long
    Dim typePrestation As String

    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    IdentifierTypeVisite = "AUTRE" ' Par defaut

    ' Chercher la visite dans la feuille Visites
    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
        If wsVisites.Cells(i, 1).Value = idVisite Then
            ' Lire la colonne Type_Prestation (colonne 6)
            typePrestation = UCase(Trim(wsVisites.Cells(i, 6).Value)) ' Type_Prestation

            ' Normaliser les types pour correspondre aux tarifs
            Select Case typePrestation
                Case "VISITE CONTEE BRANLY", "VISITE BRANLY"
                    IdentifierTypeVisite = "BRANLY"
                Case "VISITE CONTEE MARINE", "VISITE MARINE"
                    IdentifierTypeVisite = "MARINE"
                Case "HORS LES MURS", "HORS-LES-MURS", "HORSLEMURS"
                    IdentifierTypeVisite = "HORSLEMURS"
                Case "VISIO"
                    IdentifierTypeVisite = "VISIO"
                Case "EVENEMENT BRANLY", "EVENEMENT"
                    IdentifierTypeVisite = "EVENEMENT"
                Case Else
                    ' Retourner le type tel quel si non reconnu
                    If typePrestation <> "" Then
                        IdentifierTypeVisite = typePrestation
                    End If
            End Select

            Exit Function
        End If
    Next i
End Function

'===============================================================================
' FONCTION: CalculerTarifJournee
' DESCRIPTION: Calcule le tarif pour une journee selon le type et nb de visites
'===============================================================================
Private Function CalculerTarifJournee(typeVisite As String, nbVisites As Integer, dureeHeures As Double) As Double
    '===============================================================================
    ' LOGIQUE TARIFAIRE CLIENT (exemple mail):
    ' - 1 visite/jour = 80€
    ' - 2 visites/jour = 110€
    ' - 3+ visites/jour = 140€
    ' - EXCEPTION: Hors-les-murs = 100€ (peu importe le nombre)
    '===============================================================================

    ' CAS SPECIAL : Hors-les-murs = 100€ fixe
    If UCase(Trim(typeVisite)) = "HORS-LES-MURS" Or _
       UCase(Trim(typeVisite)) = "HORS LES MURS" Or _
       UCase(Trim(typeVisite)) = "HORSLEMURS" Then
        CalculerTarifJournee = LireParametreConfig("TARIF_HORSLEMURS", 100)
        Exit Function
    End If

    ' TARIFS STANDARDS selon nombre de visites PAR JOUR
    Select Case nbVisites
        Case 1
            CalculerTarifJournee = LireParametreConfig("TARIF_1_VISITE", 80)
        Case 2
            CalculerTarifJournee = LireParametreConfig("TARIF_2_VISITES", 110)
        Case Is >= 3
            CalculerTarifJournee = LireParametreConfig("TARIF_3_VISITES", 140)
        Case Else
            CalculerTarifJournee = 0
    End Select
End Function'===============================================================================
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
' DESCRIPTION: Retourne le nom complet d'un guide (Prenom Nom)
' NOTE: guideID contient deja "Prenom Nom", on cherche juste a le valider
'===============================================================================
Private Function ObtenirNomCompletGuide(guideID As String) As String
    Dim wsGuides As Worksheet
    Dim i As Long
    Dim nomComplet As String

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    ObtenirNomCompletGuide = guideID ' Par defaut retourner l'ID

    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        ' Construire le nom complet : Prenom (col A) + Nom (col B)
        nomComplet = Trim(wsGuides.Cells(i, 1).Value) & " " & Trim(wsGuides.Cells(i, 2).Value)

        If UCase(nomComplet) = UCase(Trim(guideID)) Then
            ObtenirNomCompletGuide = nomComplet
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
    Dim infoJour As Variant
    Dim temp As Variant

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
        If Trim(wsPlanning.Cells(i, 7).Value) = Trim(guideID) Then ' Guide_Attribue (Col 7)
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
                        infoJour = Array(dateVisite, typeVisite, 1, dureeHeures)
                        dictJours.Add cleJour, infoJour
                    Else
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

    ' Creer un nouveau classeur pour la fiche
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
        wsFiche.Cells(ligne, 5).Value = montantJour & " "
        wsFiche.Cells(ligne, 6).Value = montantJour
        wsFiche.Cells(ligne, 6).NumberFormat = "#,##0.00 "

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
    wsFiche.Cells(ligne, 6).NumberFormat = "#,##0.00 "
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
        MsgBox "Fiche de paie generee avec succes !" & vbCrLf & fichier, vbInformation
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
' FONCTION: GenererDecompteMensuel
' DESCRIPTION: Genere un decompte detaille avec statistiques par categorie
'===============================================================================
Public Sub GenererDecompteMensuel()
    Dim wsPlanning As Worksheet
    Dim wsVisites As Worksheet
    Dim wbDecompte As Workbook
    Dim wsDecompte As Worksheet
    Dim moisFiltre As String
    Dim moisCible As Integer, anneeCible As Integer
    Dim dictGuides As Object
    Dim dictStats As Object
    Dim i As Long, ligne As Long
    Dim guideID As String, dateVisite As Date, heureVisite As String
    Dim categorieVisite As String, idVisite As String

    On Error GoTo Erreur

    ' Demander le mois
    moisFiltre = InputBox("Mois pour le decompte (MM/AAAA):", "Decompte mensuel", Format(Date, "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    moisCible = CInt(Left(moisFiltre, 2))
    anneeCible = CInt(Right(moisFiltre, 4))

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set dictGuides = CreateObject("Scripting.Dictionary")
    Set dictStats = CreateObject("Scripting.Dictionary")

    ' Initialiser compteurs statistiques
    dictStats.Add "Branly", 0
    dictStats.Add "Marine", 0
    dictStats.Add "Hors-les-murs", 0
    dictStats.Add "Evenements", 0
    dictStats.Add "Visio", 0
    dictStats.Add "Autres", 0
    dictStats.Add "Total", 0

    Application.ScreenUpdating = False

    ' Creer nouveau classeur
    Set wbDecompte = Workbooks.Add
    Set wsDecompte = wbDecompte.Worksheets(1)
    wsDecompte.Name = "Decompte_" & Format(DateSerial(anneeCible, moisCible, 1), "yyyymm")

    ' Titre
    wsDecompte.Cells(1, 1).Value = "DECOMPTE DETAILLE - " & Format(DateSerial(anneeCible, moisCible, 1), "MMMM YYYY")
    wsDecompte.Cells(1, 1).Font.Size = 14
    wsDecompte.Cells(1, 1).Font.Bold = True
    wsDecompte.Range("A1:G1").Merge

    ' En-tetes
    ligne = 3
    wsDecompte.Cells(ligne, 1).Value = "Guide"
    wsDecompte.Cells(ligne, 2).Value = "Date"
    wsDecompte.Cells(ligne, 3).Value = "Heure"
    wsDecompte.Cells(ligne, 4).Value = "Type Visite"
    wsDecompte.Cells(ligne, 5).Value = "Categorie"
    wsDecompte.Cells(ligne, 6).Value = "Nb Jours"
    wsDecompte.Cells(ligne, 7).Value = "Montant Cachet"
    wsDecompte.Range("A" & ligne & ":G" & ligne).Font.Bold = True
    wsDecompte.Range("A" & ligne & ":G" & ligne).Interior.Color = RGB(200, 200, 200)
    ligne = ligne + 1

    ' Parcourir le planning
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = Trim(wsPlanning.Cells(i, 7).Value) ' Guide_Attribue (Col 7)

        If guideID <> "NON ATTRIBUE" And guideID <> "" Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)
            heureVisite = wsPlanning.Cells(i, 3).Value
            idVisite = wsPlanning.Cells(i, 1).Value

            If Err.Number = 0 And Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                ' Identifier categorie
                categorieVisite = IdentifierCategorieVisite(idVisite)

                ' Incrementer stats
                dictStats(categorieVisite) = dictStats(categorieVisite) + 1
                dictStats("Total") = dictStats("Total") + 1

                ' Ajouter ligne
                wsDecompte.Cells(ligne, 1).Value = ObtenirNomCompletGuide(guideID)
                wsDecompte.Cells(ligne, 2).Value = Format(dateVisite, "dd/mm/yyyy")
                wsDecompte.Cells(ligne, 3).Value = heureVisite
                wsDecompte.Cells(ligne, 4).Value = ObtenirNomVisite(idVisite)
                wsDecompte.Cells(ligne, 5).Value = categorieVisite

                ' Compter jours par guide
                If Not dictGuides.exists(guideID) Then
                    dictGuides.Add guideID, CreateObject("Scripting.Dictionary")
                End If
                Dim cleJour As String
                cleJour = Format(dateVisite, "yyyy-mm-dd")
                If Not dictGuides(guideID).exists(cleJour) Then
                    dictGuides(guideID).Add cleJour, True
                End If

                ligne = ligne + 1
            End If

            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    ' Ajouter statistiques
    ligne = ligne + 2
    wsDecompte.Cells(ligne, 1).Value = "STATISTIQUES PAR CATEGORIE"
    wsDecompte.Cells(ligne, 1).Font.Bold = True
    wsDecompte.Cells(ligne, 1).Font.Size = 12
    ligne = ligne + 1

    wsDecompte.Cells(ligne, 1).Value = "Visites Branly :"
    wsDecompte.Cells(ligne, 2).Value = dictStats("Branly")
    ligne = ligne + 1

    wsDecompte.Cells(ligne, 1).Value = "Visites Marine :"
    wsDecompte.Cells(ligne, 2).Value = dictStats("Marine")
    ligne = ligne + 1

    wsDecompte.Cells(ligne, 1).Value = "Hors-les-murs :"
    wsDecompte.Cells(ligne, 2).Value = dictStats("Hors-les-murs")
    ligne = ligne + 1

    wsDecompte.Cells(ligne, 1).Value = "Evenements :"
    wsDecompte.Cells(ligne, 2).Value = dictStats("Evenements")
    ligne = ligne + 1

    wsDecompte.Cells(ligne, 1).Value = "Visio :"
    wsDecompte.Cells(ligne, 2).Value = dictStats("Visio")
    ligne = ligne + 1

    wsDecompte.Cells(ligne, 1).Value = "Autres :"
    wsDecompte.Cells(ligne, 2).Value = dictStats("Autres")
    ligne = ligne + 1

    wsDecompte.Cells(ligne, 1).Value = "TOTAL :"
    wsDecompte.Cells(ligne, 2).Value = dictStats("Total")
    wsDecompte.Cells(ligne, 1).Font.Bold = True
    wsDecompte.Cells(ligne, 2).Font.Bold = True

    ' Stats guides
    ligne = ligne + 2
    wsDecompte.Cells(ligne, 1).Value = "NOMBRE DE JOURS PAR GUIDE"
    wsDecompte.Cells(ligne, 1).Font.Bold = True
    ligne = ligne + 1

    Dim keyGuide As Variant
    For Each keyGuide In dictGuides.Keys
        wsDecompte.Cells(ligne, 1).Value = ObtenirNomCompletGuide(CStr(keyGuide))
        wsDecompte.Cells(ligne, 2).Value = dictGuides(keyGuide).Count & " jours"
        ligne = ligne + 1
    Next keyGuide

    wsDecompte.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "Decompte mensuel genere avec succes !" & vbCrLf & vbCrLf & _
           "Total visites : " & dictStats("Total") & vbCrLf & _
           "Branly : " & dictStats("Branly") & vbCrLf & _
           "Marine : " & dictStats("Marine"), _
           vbInformation, "Decompte Mensuel"

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la generation du decompte : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: IdentifierCategorieVisite
' DESCRIPTION: Identifie la categorie d'une visite
'===============================================================================
Private Function IdentifierCategorieVisite(idVisite As String) As String
    Dim nomVisite As String
    nomVisite = UCase(Trim(ObtenirNomVisite(idVisite)))

    If InStr(nomVisite, "BRANLY") > 0 Then
        IdentifierCategorieVisite = "Branly"
    ElseIf InStr(nomVisite, "MARINE") > 0 Then
        IdentifierCategorieVisite = "Marine"
    ElseIf InStr(nomVisite, "HORS-LES-MURS") > 0 Or InStr(nomVisite, "HORS LES MURS") > 0 Then
        IdentifierCategorieVisite = "Hors-les-murs"
    ElseIf InStr(nomVisite, "EVENEMENT") > 0 Or InStr(nomVisite, "EVENT") > 0 Then
        IdentifierCategorieVisite = "Evenements"
    ElseIf InStr(nomVisite, "VISIO") > 0 Then
        IdentifierCategorieVisite = "Visio"
    Else
        IdentifierCategorieVisite = "Autres"
    End If
End Function

'===============================================================================
' FONCTION: ObtenirNomVisite
' DESCRIPTION: Retourne le nom d'une visite
'===============================================================================
Private Function ObtenirNomVisite(idVisite As String) As String
    Dim wsVisites As Worksheet
    Dim i As Long

    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    ObtenirNomVisite = idVisite

    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
        If wsVisites.Cells(i, 1).Value = idVisite Then
            ObtenirNomVisite = wsVisites.Cells(i, 2).Value
            Exit Function
        End If
    Next i
End Function


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

    ' Creer un nouveau classeur
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
        MsgBox "Recapitulatif exporte avec succes !" & vbCrLf & fichier, vbInformation
    End If

    wbExport.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    MsgBox "Erreur lors de l'export : " & Err.Description, vbCritical
End Sub


