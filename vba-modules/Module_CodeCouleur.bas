Attribute VB_Name = "Module_CodeCouleur"
' ============================================
' MODULE CODE COULEUR PLANNING
' ============================================
' Applique automatiquement les couleurs au planning
' selon la categorie de visite
' ============================================

Option Explicit

' Fonction : Applique le code couleur a une cellule selon la categorie
' Parametres :
'   - cell : La cellule a formater
'   - categorie : La categorie (Individuel, Groupe, Evenement, Hors-les-murs, Marine)
Public Sub AppliquerCodeCouleur(cell As Range, categorie As String)
    On Error Resume Next
    
    Select Case Trim(UCase(categorie))
        Case "INDIVIDUEL"
            ' Bleu standard
            With cell.Interior
                .Color = RGB(0, 112, 192)
                .Pattern = xlSolid
            End With
            cell.Font.Color = RGB(255, 255, 255) ' Blanc
            cell.Font.Bold = False
            
        Case "GROUPE"
            ' Bleu clair
            With cell.Interior
                .Color = RGB(155, 194, 230)
                .Pattern = xlSolid
            End With
            cell.Font.Color = RGB(0, 0, 0) ' Noir
            cell.Font.Bold = False
            
        Case "EVENEMENT", "EVENEMENT"
            ' Rose
            With cell.Interior
                .Color = RGB(255, 192, 203)
                .Pattern = xlSolid
            End With
            cell.Font.Color = RGB(0, 0, 0) ' Noir
            cell.Font.Bold = False
            
        Case "HORS-LES-MURS", "HORS LES MURS"
            ' Rouge
            With cell.Interior
                .Color = RGB(255, 0, 0)
                .Pattern = xlSolid
            End With
            cell.Font.Color = RGB(255, 255, 255) ' Blanc
            cell.Font.Bold = False
            
        Case "MARINE"
            ' Bleu fonce + GRAS + MAJUSCULES
            With cell.Interior
                .Color = RGB(0, 32, 96)
                .Pattern = xlSolid
            End With
            cell.Font.Color = RGB(255, 255, 255) ' Blanc
            cell.Font.Bold = True
            cell.Font.Size = cell.Font.Size + 1
            
            ' Mettre en majuscules
            If Not IsEmpty(cell.Value) Then
                cell.Value = UCase(cell.Value)
            End If
            
        Case Else
            ' Pas de categorie ou inconnue : pas de formatage
            cell.Interior.ColorIndex = xlNone
            cell.Font.Color = RGB(0, 0, 0)
            cell.Font.Bold = False
    End Select
End Sub

' Fonction : Applique le code couleur a toute la feuille Planning
Public Sub AppliquerCodeCouleurPlanning()
    On Error Resume Next
    
    Dim wsPlanning As Worksheet
    Dim wsVisites As Worksheet
    Dim derLignePlanning As Long
    Dim derColPlanning As Long
    Dim i As Long, j As Long
    Dim cellPlanning As Range
    Dim idVisite As String
    Dim categorie As String
    
    Application.ScreenUpdating = False
    
    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)
    Set wsVisites = ThisWorkbook.Sheets(FEUILLE_VISITES)
    
    ' Trouver les dimensions du planning
    derLignePlanning = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
    derColPlanning = wsPlanning.Cells(4, wsPlanning.Columns.Count).End(xlToLeft).Column
    
    ' Parcourir toutes les cellules du planning (a partir de ligne 5, colonne 2)
    For i = 5 To derLignePlanning
        For j = 2 To derColPlanning
            Set cellPlanning = wsPlanning.Cells(i, j)
            
            ' Si la cellule contient une donnee
            If Not IsEmpty(cellPlanning.Value) And cellPlanning.Value <> "" Then
                ' Chercher la categorie correspondante dans Visites
                ' On suppose que l'ID ou type de visite est dans la cellule
                Dim typeVisite As String
                typeVisite = Trim(cellPlanning.Value)
                
                ' Chercher dans feuille Visites
                categorie = ChercherCategorieVisite(typeVisite)
                
                ' Appliquer le code couleur
                If categorie <> "" Then
                    AppliquerCodeCouleur cellPlanning, categorie
                End If
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Code couleur applique avec succes au planning !", vbInformation, "Formatage termine"
End Sub

' Fonction : Cherche la categorie d'une visite dans la feuille Visites
' Parametres :
'   - typeVisite : Type ou ID de la visite
' Retour : La categorie trouvee, ou "" si non trouvee
Private Function ChercherCategorieVisite(typeVisite As String) As String
    On Error Resume Next
    
    Dim wsVisites As Worksheet
    Dim derLigne As Long
    Dim i As Long
    Dim typeCol As Long
    Dim catCol As Long
    
    Set wsVisites = ThisWorkbook.Sheets(FEUILLE_VISITES)
    derLigne = wsVisites.Cells(wsVisites.Rows.Count, 1).End(xlUp).Row
    
    ' Trouver les colonnes Type et Categorie
    typeCol = 0
    catCol = 0
    
    For i = 1 To wsVisites.Cells(4, wsVisites.Columns.Count).End(xlToLeft).Column
        If InStr(1, wsVisites.Cells(4, i).Value, "Type", vbTextCompare) > 0 Then
            typeCol = i
        End If
        If InStr(1, wsVisites.Cells(4, i).Value, "Categorie", vbTextCompare) > 0 Or _
           InStr(1, wsVisites.Cells(4, i).Value, "Categorie", vbTextCompare) > 0 Then
            catCol = i
        End If
    Next i
    
    ' Si colonnes trouvees
    If typeCol > 0 And catCol > 0 Then
        ' Parcourir les visites
        For i = 5 To derLigne
            If InStr(1, wsVisites.Cells(i, typeCol).Value, typeVisite, vbTextCompare) > 0 Or _
               InStr(1, typeVisite, wsVisites.Cells(i, typeCol).Value, vbTextCompare) > 0 Then
                ' Visite trouvee, retourner categorie
                ChercherCategorieVisite = Trim(wsVisites.Cells(i, catCol).Value)
                Exit Function
            End If
        Next i
    End If
    
    ' Non trouve
    ChercherCategorieVisite = ""
End Function

' Fonction : Applique le code couleur a une ligne specifique du planning
' Utilise lors de la generation automatique ligne par ligne
Public Sub AppliquerCodeCouleurLigne(wsPlanning As Worksheet, ligneNum As Long, categorie As String)
    On Error Resume Next
    
    Dim derCol As Long
    Dim j As Long
    
    derCol = wsPlanning.Cells(4, wsPlanning.Columns.Count).End(xlToLeft).Column
    
    ' Appliquer a toutes les cellules de la ligne
    For j = 2 To derCol
        If Not IsEmpty(wsPlanning.Cells(ligneNum, j).Value) Then
            AppliquerCodeCouleur wsPlanning.Cells(ligneNum, j), categorie
        End If
    Next j
End Sub

' Fonction : Reinitialise le formatage du planning
Public Sub ReinitialiserFormatagePlanning()
    On Error Resume Next
    
    Dim wsPlanning As Worksheet
    Dim derLigne As Long
    Dim derCol As Long
    
    Set wsPlanning = ThisWorkbook.Sheets(FEUILLE_PLANNING)
    
    derLigne = wsPlanning.Cells(wsPlanning.Rows.Count, 1).End(xlUp).Row
    derCol = wsPlanning.Cells(4, wsPlanning.Columns.Count).End(xlToLeft).Column
    
    Application.ScreenUpdating = False
    
    ' Reinitialiser le formatage de toute la zone de donnees
    Dim rng As Range
    Set rng = wsPlanning.Range(wsPlanning.Cells(5, 2), wsPlanning.Cells(derLigne, derCol))
    
    With rng
        .Interior.ColorIndex = xlNone
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "Formatage du planning reinitialise.", vbInformation
End Sub
