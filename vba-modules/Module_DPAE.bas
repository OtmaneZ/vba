Attribute VB_Name = "Module_DPAE"
'===============================================================================
' MODULE: Export DPAE
' DESCRIPTION: Generation fichier Excel pre-rempli pour DPAE (Declaration Prealable A l'Embauche)
' AUTEUR: Systeme de Gestion Planning Guides
' DATE: Novembre 2025
'===============================================================================

Option Explicit

'===============================================================================
' FONCTION: ExporterDonneesDPAE
' DESCRIPTION: Genere un fichier Excel avec toutes les donnees pour DPAE
'===============================================================================
Public Sub ExporterDonneesDPAE()
    Dim wsGuides As Worksheet
    Dim wsPlanning As Worksheet
    Dim wbDPAE As Workbook
    Dim wsDPAE As Worksheet
    Dim moisFiltre As String
    Dim moisCible As Integer, anneeCible As Integer
    Dim i As Long, ligne As Long
    Dim guideID As String, guideNom As String, guidePrenom As String
    Dim dateVisite As Date, dateDebut As Date, dateFin As Date
    Dim dictGuides As Object
    Dim fichier As String

    On Error GoTo Erreur

    ' Demander le mois
    moisFiltre = InputBox("Mois pour export DPAE (MM/AAAA):", "Export DPAE", Format(Date, "mm/yyyy"))
    If moisFiltre = "" Then Exit Sub

    moisCible = CInt(Left(moisFiltre, 2))
    anneeCible = CInt(Right(moisFiltre, 4))

    Set wsGuides = ThisWorkbook.Worksheets(FEUILLE_GUIDES)
    Set wsPlanning = ThisWorkbook.Worksheets(FEUILLE_PLANNING)
    Set dictGuides = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    ' Collecter les guides ayant travaille ce mois
    For i = 2 To wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
        guideID = Trim(wsPlanning.Cells(i, 5).Value)

        If guideID <> "NON ATTRIBUE" And guideID <> "" Then
            On Error Resume Next
            dateVisite = CDate(wsPlanning.Cells(i, 2).Value)

            If Err.Number = 0 And Month(dateVisite) = moisCible And Year(dateVisite) = anneeCible Then
                If Not dictGuides.exists(guideID) Then
                    ' Stocker premiere et derniere date du mois
                    dictGuides.Add guideID, Array(dateVisite, dateVisite)
                Else
                    ' Mettre a jour les dates
                    Dim dates As Variant
                    dates = dictGuides(guideID)

                    If dateVisite < dates(0) Then dates(0) = dateVisite
                    If dateVisite > dates(1) Then dates(1) = dateVisite

                    dictGuides(guideID) = dates
                End If
            End If

            Err.Clear
            On Error GoTo Erreur
        End If
    Next i

    If dictGuides.Count = 0 Then
        MsgBox "Aucune visite trouvee pour ce mois.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Creer nouveau classeur
    Set wbDPAE = Workbooks.Add
    Set wsDPAE = wbDPAE.Worksheets(1)
    wsDPAE.Name = "DPAE_" & Format(DateSerial(anneeCible, moisCible, 1), "yyyymm")

    ' Titre
    wsDPAE.Cells(1, 1).Value = "DONNEES POUR DPAE - " & Format(DateSerial(anneeCible, moisCible, 1), "MMMM YYYY")
    wsDPAE.Cells(1, 1).Font.Size = 14
    wsDPAE.Cells(1, 1).Font.Bold = True
    wsDPAE.Range("A1:L1").Merge
    wsDPAE.Cells(1, 1).Interior.Color = RGB(0, 112, 192)
    wsDPAE.Cells(1, 1).Font.Color = RGB(255, 255, 255)

    wsDPAE.Cells(2, 1).Value = "A copier-coller dans le formulaire DPAE sur net-entreprises.fr"
    wsDPAE.Cells(2, 1).Font.Italic = True
    wsDPAE.Range("A2:L2").Merge

    ' En-tetes
    ligne = 4
    wsDPAE.Cells(ligne, 1).Value = "ID Guide"
    wsDPAE.Cells(ligne, 2).Value = "Nom"
    wsDPAE.Cells(ligne, 3).Value = "Prenom"
    wsDPAE.Cells(ligne, 4).Value = "Email"
    wsDPAE.Cells(ligne, 5).Value = "Telephone"
    wsDPAE.Cells(ligne, 6).Value = "Date debut contrat"
    wsDPAE.Cells(ligne, 7).Value = "Date fin contrat"
    wsDPAE.Cells(ligne, 8).Value = "Nature contrat"
    wsDPAE.Cells(ligne, 9).Value = "Qualification"
    wsDPAE.Cells(ligne, 10).Value = "Type emploi"
    wsDPAE.Cells(ligne, 11).Value = "Nb heures"
    wsDPAE.Cells(ligne, 12).Value = "Remuneration"

    wsDPAE.Range("A" & ligne & ":L" & ligne).Font.Bold = True
    wsDPAE.Range("A" & ligne & ":L" & ligne).Interior.Color = RGB(220, 220, 220)
    ligne = ligne + 1

    ' Remplir les donnees pour chaque guide
    Dim keyGuide As Variant
    For Each keyGuide In dictGuides.Keys
        guideID = CStr(keyGuide)

        ' Recuperer infos guide
        For i = 2 To wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
            If wsGuides.Cells(i, 1).Value = guideID Then
                guideNom = wsGuides.Cells(i, 3).Value  ' Nom
                guidePrenom = wsGuides.Cells(i, 2).Value  ' Prenom

                Dim dates As Variant
                dates = dictGuides(guideID)
                dateDebut = dates(0)
                dateFin = dates(1)

                ' Remplir la ligne
                wsDPAE.Cells(ligne, 1).Value = guideID
                wsDPAE.Cells(ligne, 2).Value = guideNom
                wsDPAE.Cells(ligne, 3).Value = guidePrenom
                wsDPAE.Cells(ligne, 4).Value = wsGuides.Cells(i, 4).Value  ' Email
                wsDPAE.Cells(ligne, 5).Value = wsGuides.Cells(i, 5).Value  ' Tel
                wsDPAE.Cells(ligne, 6).Value = Format(dateDebut, "dd/mm/yyyy")
                wsDPAE.Cells(ligne, 7).Value = Format(dateFin, "dd/mm/yyyy")
                wsDPAE.Cells(ligne, 8).Value = "CDD d'usage"
                wsDPAE.Cells(ligne, 9).Value = "Guide conferencier"
                wsDPAE.Cells(ligne, 10).Value = "Temps partiel"
                wsDPAE.Cells(ligne, 11).Value = "Variable"
                wsDPAE.Cells(ligne, 12).Value = "Selon cachet"

                ' Colorier en alternance
                If ligne Mod 2 = 0 Then
                    wsDPAE.Range("A" & ligne & ":L" & ligne).Interior.Color = RGB(242, 242, 242)
                End If

                ligne = ligne + 1
                Exit For
            End If
        Next i
    Next keyGuide

    ' Instructions
    ligne = ligne + 2
    wsDPAE.Cells(ligne, 1).Value = "INSTRUCTIONS :"
    wsDPAE.Cells(ligne, 1).Font.Bold = True
    wsDPAE.Cells(ligne, 1).Font.Size = 12
    ligne = ligne + 1

    wsDPAE.Cells(ligne, 1).Value = "1. Allez sur net-entreprises.fr"
    ligne = ligne + 1
    wsDPAE.Cells(ligne, 1).Value = "2. Connectez-vous avec vos identifiants"
    ligne = ligne + 1
    wsDPAE.Cells(ligne, 1).Value = "3. Allez dans 'DPAE - Declaration Prealable A l'Embauche'"
    ligne = ligne + 1
    wsDPAE.Cells(ligne, 1).Value = "4. Pour chaque guide, copiez-collez les informations du tableau ci-dessus"
    ligne = ligne + 1
    wsDPAE.Cells(ligne, 1).Value = "5. Validez chaque declaration"
    ligne = ligne + 2

    wsDPAE.Cells(ligne, 1).Value = "Note : Les champs 'Nature contrat', 'Qualification' et 'Type emploi' sont pre-remplis avec les valeurs standard."
    wsDPAE.Cells(ligne, 1).Font.Italic = True
    wsDPAE.Cells(ligne, 1).Font.Color = RGB(255, 0, 0)
    wsDPAE.Range("A" & ligne & ":L" & ligne).Merge
    ligne = ligne + 1

    wsDPAE.Cells(ligne, 1).Value = "Adaptez-les si necessaire selon votre convention collective."
    wsDPAE.Cells(ligne, 1).Font.Italic = True
    wsDPAE.Cells(ligne, 1).Font.Color = RGB(255, 0, 0)
    wsDPAE.Range("A" & ligne & ":L" & ligne).Merge

    ' Ajuster colonnes
    wsDPAE.Columns.AutoFit

    ' Figer les en-tetes
    wsDPAE.Range("A5").Select
    ActiveWindow.FreezePanes = True

    Application.ScreenUpdating = True

    ' Sauvegarder
    fichier = Application.GetSaveAsFilename("DPAE_" & Format(DateSerial(anneeCible, moisCible, 1), "yyyymm") & ".xlsx", _
                                            "Fichiers Excel (*.xlsx), *.xlsx")

    If fichier <> "False" Then
        wbDPAE.SaveAs fichier
        MsgBox "Fichier DPAE genere avec succes !" & vbCrLf & vbCrLf & _
               "Nombre de guides : " & dictGuides.Count & vbCrLf & _
               "Fichier : " & fichier & vbCrLf & vbCrLf & _
               "Vous pouvez maintenant copier-coller ces donnees dans le formulaire DPAE sur net-entreprises.fr", _
               vbInformation, "Export DPAE"
    End If

    wbDPAE.Close SaveChanges:=False

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    If Not wbDPAE Is Nothing Then wbDPAE.Close SaveChanges:=False
    MsgBox "Erreur lors de l'export DPAE : " & Err.Description, vbCritical
End Sub
