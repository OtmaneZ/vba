Attribute VB_Name = "Module_Calculs"
'===============================================================================
' MODULE: Calculs de Paie
' DESCRIPTION: Calcul automatique des visites et des salaires
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
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
    Dim dictGuides As Object ' Dictionary
    Dim i As Long
    Dim guideID As String
    Dim guideNom As String
    Dim nbVisites As Integer
    Dim montantSalaire As Double
    Dim tarifHeure As Double
    Dim dureeVisite As Double
    Dim ligneCalcul As Long
    Dim moisFiltre As String
    Dim dateVisite As Date

    On Error GoTo Erreur

    ' Demander le mois a calculer (optionnel)
    moisFiltre = InputBox("Filtrer par mois (MM/AAAA) ou laisser vide pour tout:", "Periode de calcul", Format(Date, "mm/yyyy"))

    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set wsCalculs = ThisWorkbook.Worksheets(FEUILLE_CALCULS)
    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set wsVisites = ThisWorkbook.Worksheets(FEUILLE_VISITES)
    Set dictGuides = CreateObject("Scripting.Dictionary")

    ' Obtenir le tarif horaire
    tarifHeure = ObtenirTarifHeure()

    Application.ScréénUpdating = False

    ' Effacer les anciens calculs (conserver les en-tetes)
    Dim derLigneCalcul As Long
    derLigneCalcul = wsCalculs.Cells(wsCalculs.Rows.Count, 1).End(xlUp).Row
    If derLigneCalcul > 1 Then
        wsCalculs.Range("A2:D" & derLigneCalcul).ClearContents
    End If

    ' Parcourir le planning et compter les visites par guide
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = wsPlanning.Cells(i, 5).Value

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
                ' Calculer la duree de la visite en heures
                Dim idVisite As String
                idVisite = wsPlanning.Cells(i, 1).Value
                dureeVisite = ObtenirDureeVisite(idVisite)

                ' Ajouter au dictionnaire
                If Not dictGuides.exists(guideID) Then
                    Dim infoGuide As Variant
                    infoGuide = Array(0, 0) ' (nb_visites, total_heures)
                    dictGuides.Add guideID, infoGuide
                End If

                ' Incrementer les compteurs
                Dim temp As Variant
                temp = dictGuides(guideID)
                temp(0) = temp(0) + 1 ' Nombre de visites
                temp(1) = temp(1) + dureeVisite ' Total heures
                dictGuides(guideID) = temp
            End If

            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    ' Remplir la feuille Calculs_Paie
    ligneCalcul = 2
    Dim key As Variant

    For Each key In dictGuides.Keys
        guideID = CStr(key)
        guideNom = ObtenirNomCompletGuide(guideID)

        Dim stats As Variant
        stats = dictGuides(guideID)
        nbVisites = stats(0)

        ' Calcul du salaire avec GRILLE DEGRESSIVE
        montantSalaire = CalculerSalaireDegressif(stats(1), tarifHeure)

        ' Remplir la ligne
        wsCalculs.Cells(ligneCalcul, 1).Value = guideID
        wsCalculs.Cells(ligneCalcul, 2).Value = guideNom
        wsCalculs.Cells(ligneCalcul, 3).Value = nbVisites
        wsCalculs.Cells(ligneCalcul, 4).Value = montantSalaire
        wsCalculs.Cells(ligneCalcul, 4).NumberFormat = "#,##0.00 €"

        ' Formater
        If nbVisites > 0 Then
            wsCalculs.Rows(ligneCalcul).Interior.Color = COULEUR_DISPONIBLE
        End If

        ligneCalcul = ligneCalcul + 1
    Next key

    ' Ajouter une ligne de total
    If ligneCalcul > 2 Then
        wsCalculs.Cells(ligneCalcul, 2).Value = "TOTAL"
        wsCalculs.Cells(ligneCalcul, 2).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 3).Formula = "=SUM(C2:C" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 3).Font.Bold = True

        wsCalculs.Cells(ligneCalcul, 4).Formula = "=SUM(D2:D" & ligneCalcul - 1 & ")"
        wsCalculs.Cells(ligneCalcul, 4).NumberFormat = "#,##0.00 €"
        wsCalculs.Cells(ligneCalcul, 4).Font.Bold = True

        wsCalculs.Rows(ligneCalcul).Interior.Color = RGB(255, 242, 204)
    End If

    wsCalculs.Columns.AutoFit
    Application.ScréénUpdating = True

    Dim msgPeriode As String
    If moisFiltre <> "" Then
        msgPeriode = " pour " & moisFiltre
    Else
        msgPeriode = " (toutes periodes)"
    End If

    MsgBox "Calculs effectues avec succès" & msgPeriode & " !" & vbCrLf & vbCrLf & _
           "Nombre de guides : " & dictGuides.Count & vbCrLf & _
           "Tarif horaire : " & Format(tarifHeure, "#,##0.00 €"), _
           vbInformation, "Calculs Paie"

    Exit Sub

Erreur:
    Application.ScréénUpdating = True
    MsgBox "Erreur lors des calculs : " & Err.Description, vbCritical
End Sub

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
    ObtenirDureeVisite = 2 ' Duree par defaut si non trouvee (2 heures)

    ' Chercher la visite
    For i = 2 To wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
        If wsVisites.Cells(i, 1).Value = idVisite Then
            On Error Resume Next
            heureDebut = CDate(wsVisites.Cells(i, 3).Value)
            heureFin = CDate(wsVisites.Cells(i, 4).Value)

            If Err.Number = 0 Then
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
    ObtenirNomCompletGuide = ""

    For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
        If wsGuides.Cells(i, 1).Value = guideID Then
            ObtenirNomCompletGuide = wsGuides.Cells(i, 2).Value & " " & wsGuides.Cells(i, 3).Value
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
    Dim totalHeures As Double
    Dim totalMontant As Double
    Dim tarifHeure As Double
    Dim fichier As String

    On Error GoTo Erreur

    ' Demander l'ID du guide
    guideID = InputBox("Entrez l'ID du guide:", "Fiche de paie")
    If guideID = "" Then Exit Sub

    ' Verifier que le guide existe
    guideNom = ObtenirNomCompletGuide(guideID)
    If guideNom = "" Then
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

    tarifHeure = ObtenirTarifHeure()

    Application.ScréénUpdating = False

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

        .Range("A7:E7").Value = Array("Date", "Horaires", "Musee", "Duree (h)", "Montant")
        .Range("A7:E7").Font.Bold = True
        .Range("A7:E7").Interior.Color = RGB(68, 114, 196)
        .Range("A7:E7").Font.Color = RGB(255, 255, 255)
    End With

    ' Lister les visites du guide
    ligne = 8
    totalHeures = 0
    totalMontant = 0

    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        If wsPlanning.Cells(i, 5).Value = guideID Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 Then
                If Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                    Dim idVisite As String
                    Dim duree As Double
                    Dim montant As Double

                    idVisite = wsPlanning.Cells(i, 1).Value
                    duree = ObtenirDureeVisite(idVisite)
                    montant = duree * tarifHeure

                    wsFiche.Cells(ligne, 1).Value = Format(dateVisite, "dd/mm/yyyy")
                    wsFiche.Cells(ligne, 2).Value = wsPlanning.Cells(i, 3).Value
                    wsFiche.Cells(ligne, 3).Value = wsPlanning.Cells(i, 4).Value
                    wsFiche.Cells(ligne, 4).Value = duree
                    wsFiche.Cells(ligne, 5).Value = montant
                    wsFiche.Cells(ligne, 5).NumberFormat = "#,##0.00 €"

                    totalHeures = totalHeures + duree
                    totalMontant = totalMontant + montant
                    ligne = ligne + 1
                End If
            End If
            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    ' Totaux
    If ligne > 8 Then
        wsFiche.Cells(ligne, 3).Value = "TOTAL"
        wsFiche.Cells(ligne, 3).Font.Bold = True
        wsFiche.Cells(ligne, 4).Value = totalHeures
        wsFiche.Cells(ligne, 4).Font.Bold = True
        wsFiche.Cells(ligne, 5).Value = totalMontant
        wsFiche.Cells(ligne, 5).NumberFormat = "#,##0.00 €"
        wsFiche.Cells(ligne, 5).Font.Bold = True
        wsFiche.Range("A" & ligne & ":E" & ligne).Interior.Color = RGB(255, 242, 204)

        ' Informations supplementaires
        ligne = ligne + 2
        wsFiche.Cells(ligne, 1).Value = "Tarif horaire :"
        wsFiche.Cells(ligne, 2).Value = tarifHeure & " €/h"

        ligne = ligne + 1
        wsFiche.Cells(ligne, 1).Value = "Nombre de visites :"
        wsFiche.Cells(ligne, 2).Value = ligne - 9
    Else
        wsFiche.Cells(8, 1).Value = "Aucune visite ce mois-ci"
    End If

    wsFiche.Columns.AutoFit

    ' Proposer de sauvegarder
    fichier = Application.GetSaveAsFilename("Fiche_Paie_" & guideID & "_" & Format(DateSerial(anneeCible, moisCible, 1), "yyyymm") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")
    If fichier <> "False" Then
        wbFiche.SaveAs fichier
        MsgBox "Fiche de paie generee avec succès !" & vbCrLf & fichier, vbInformation
    End If

    wbFiche.Close SaveChanges:=False
    Application.ScréénUpdating = True

    Exit Sub

Erreur:
    Application.ScréénUpdating = True
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

    Application.ScréénUpdating = False

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
    Application.ScréénUpdating = True

    Exit Sub

Erreur:
    Application.ScréénUpdating = True
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    MsgBox "Erreur lors de l'export : " & Err.Description, vbCritical
End Sub

'===============================================================================
' FONCTION: CalculerSalaireDegressif
' DESCRIPTION: Calcule le salaire avec une grille tarifaire degressive
' NOTE: GRILLE DE TEST - A ADAPTER SELON LES BESOINS DU CLIENT
'===============================================================================
Private Function CalculerSalaireDegressif(totalHeures As Double, tarifBase As Double) As Double
    '===========================================================================
    ' GRILLE TARIFAIRE DEGRESSIVE DE TEST
    '===========================================================================
    ' Cette grille est un EXEMPLE pour les tests.
    ' Elle sera modifiee avec les vrais tarifs du client apres le call de mardi.
    '
    ' PRINCIPE : Plus un guide travaille d'heures dans le mois,
    '            plus son tarif horaire diminue (degressivite)
    '
    ' GRILLE ACTUELLE (EXEMPLE) :
    ' ┌──────────────────┬─────────────┬──────────────────────┐
    ' │ Tranche d'heures │ Tarif/heure │ Exemple              │
    ' ├──────────────────┼─────────────┼──────────────────────┤
    ' │ 0 - 10h          │ 100% (30€)  │ 8h × 30€ = 240€      │
    ' │ 10 - 20h         │ 90% (27€)   │ 15h -> 10×30 + 5×27   │
    ' │ 20 - 40h         │ 80% (24€)   │ 25h -> 10×30 + 10×27  │
    ' │ 40h+             │ 70% (21€)   │ + reste × 24€        │
    ' └──────────────────┴─────────────┴──────────────────────┘
    '
    ' AVANTAGE : Encourage la disponibilite tout en controlant les couts
    '===========================================================================

    Dim montantTotal As Double
    Dim heuresRestantes As Double
    Dim tarifActuel As Double

    montantTotal = 0
    heuresRestantes = totalHeures

    ' TRANCHE 1 : Premieres 10 heures a 100% du tarif
    If heuresRestantes > 0 Then
        If heuresRestantes <= 10 Then
            montantTotal = heuresRestantes * tarifBase
            heuresRestantes = 0
        Else
            montantTotal = 10 * tarifBase
            heuresRestantes = heuresRestantes - 10
        End If
    End If

    ' TRANCHE 2 : Heures 11 a 20 a 90% du tarif
    If heuresRestantes > 0 Then
        tarifActuel = tarifBase * 0.9
        If heuresRestantes <= 10 Then
            montantTotal = montantTotal + (heuresRestantes * tarifActuel)
            heuresRestantes = 0
        Else
            montantTotal = montantTotal + (10 * tarifActuel)
            heuresRestantes = heuresRestantes - 10
        End If
    End If

    ' TRANCHE 3 : Heures 21 a 40 a 80% du tarif
    If heuresRestantes > 0 Then
        tarifActuel = tarifBase * 0.8
        If heuresRestantes <= 20 Then
            montantTotal = montantTotal + (heuresRestantes * tarifActuel)
            heuresRestantes = 0
        Else
            montantTotal = montantTotal + (20 * tarifActuel)
            heuresRestantes = heuresRestantes - 20
        End If
    End If

    ' TRANCHE 4 : Heures 41+ a 70% du tarif
    If heuresRestantes > 0 Then
        tarifActuel = tarifBase * 0.7
        montantTotal = montantTotal + (heuresRestantes * tarifActuel)
    End If

    CalculerSalaireDegressif = montantTotal

    '===========================================================================
    ' NOTES POUR LE DEVELOPPEUR :
    '===========================================================================
    ' Pour modifier cette grille apres le call client :
    '
    ' 1. GRILLE PAR TRANCHES (comme ci-dessus) :
    '    - Modifier les seuils (10, 20, 40)
    '    - Modifier les pourcentages (1.0, 0.9, 0.8, 0.7)
    '
    ' 2. GRILLE PAR FORFAIT :
    '    If totalHeures <= 2 Then
    '        CalculerSalaireDegressif = 50  ' Forfait visite courte
    '    ElseIf totalHeures <= 4 Then
    '        CalculerSalaireDegressif = 90  ' Forfait demi-journee
    '    ElseIf totalHeures <= 8 Then
    '        CalculerSalaireDegressif = 150 ' Forfait journee
    '    Else
    '        CalculerSalaireDegressif = 150 + ((totalHeures - 8) * 20)
    '    End If
    '
    ' 3. GRILLE AVEC PALIERS ET BONUS :
    '    Dim montantBase As Double
    '    montantBase = totalHeures * tarifBase
    '    If totalHeures > 40 Then
    '        ' Bonus de 10% si plus de 40h dans le mois
    '        CalculerSalaireDegressif = montantBase * 1.1
    '    ElseIf totalHeures > 20 Then
    '        ' Bonus de 5% si plus de 20h
    '        CalculerSalaireDegressif = montantBase * 1.05
    '    Else
    '        CalculerSalaireDegressif = montantBase
    '    End If
    '===========================================================================
End Function

'===============================================================================
' FONCTION: AfficherExempleGrilleTarifaire
' DESCRIPTION: Affiche un exemple de calcul pour comprendre la grille
'===============================================================================
Public Sub AfficherExempleGrilleTarifaire()
    Dim tarifBase As Double
    Dim exemples() As Variant
    Dim i As Integer
    Dim message As String

    tarifBase = 30 ' Tarif de base pour l'exemple (30€/h)

    ' Exemples de calculs
    exemples = Array(5, 8, 15, 25, 35, 50)

    message = "GRILLE TARIFAIRE DEGRESSIVE - EXEMPLES" & vbCrLf & _
              "Tarif de base : " & Format(tarifBase, "#,##0.00 €") & "/heure" & vbCrLf & vbCrLf & _
              "┌────────────┬──────────────┬──────────────────────────────┐" & vbCrLf & _
              "│ Heures     │ Montant      │ Detail du calcul             │" & vbCrLf & _
              "├────────────┼──────────────┼──────────────────────────────┤" & vbCrLf

    For i = LBound(exemples) To UBound(exemples)
        Dim heures As Double
        Dim montant As Double
        Dim detail As String

        heures = exemples(i)
        montant = CalculerSalaireDegressif(heures, tarifBase)

        ' Generer le detail
        If heures <= 10 Then
            detail = Format(heures, "0") & "h × " & Format(tarifBase, "0") & "€"
        ElseIf heures <= 20 Then
            detail = "10h×" & Format(tarifBase, "0") & "€ + " & Format(heures - 10, "0") & "h×" & Format(tarifBase * 0.9, "0") & "€"
        ElseIf heures <= 40 Then
            detail = "10h×" & Format(tarifBase, "0") & "€ + 10h×" & Format(tarifBase * 0.9, "0") & "€ + " & Format(heures - 20, "0") & "h×" & Format(tarifBase * 0.8, "0") & "€"
        Else
            detail = "Tranches 1-3 + " & Format(heures - 40, "0") & "h×" & Format(tarifBase * 0.7, "0") & "€"
        End If

        message = message & _
                  "│ " & Format(heures, "0") & "h" & String(9 - Len(Format(heures, "0")), " ") & _
                  "│ " & Format(montant, "#,##0.00 €") & String(11 - Len(Format(montant, "#,##0.00 €")), " ") & _
                  "│ " & detail & String(28 - Len(detail), " ") & "│" & vbCrLf
    Next i

    message = message & _
              "└────────────┴──────────────┴──────────────────────────────┘" & vbCrLf & vbCrLf & _
              "[!] GRILLE DE TEST - A adapter selon les besoins du client"

    MsgBox message, vbInformation, "Grille Tarifaire Degressive"
End Sub

