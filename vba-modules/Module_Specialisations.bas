Attribute VB_Name = "Module_Specialisations"
' ============================================
' MODULE GESTION SPECIALISATIONS GUIDES
' ============================================
' Verifie qu'un guide peut effectuer un type de visite donne
' ============================================

Option Explicit

' Fonction : Verifie si un guide est autorise a faire une visite
' Parametres :
'   - nomGuide : Nom complet du guide (ex: "Peggy GENESTIE")
'   - typeVisite : Type de visite (ex: "Maman Serpent")
' Retour : True si autorise, False sinon
Public Function GuideAutoriseVisite(nomGuide As String, typeVisite As String) As Boolean
    On Error GoTo Erreur
    
    Dim ws As Worksheet
    Dim derLigne As Long
    Dim i As Long
    Dim guideSpec As String
    Dim visiteSpec As String
    Dim trouve As Boolean
    
    ' Par defaut, on autorise (pour les guides sans contraintes)
    GuideAutoriseVisite = True
    trouve = False
    
    ' Recuperer la feuille Specialisations
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Specialisations")
    On Error GoTo Erreur
    
    If ws Is Nothing Then
        ' Feuille n'existe pas, on autorise tout
        Exit Function
    End If
    
    ' Chercher le guide dans la feuille Specialisations
    derLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 4 To derLigne ' Commence a ligne 4 (apres en-tetes)
        guideSpec = Trim(ws.Cells(i, 1).Value)
        visiteSpec = Trim(ws.Cells(i, 2).Value)
        
        ' Si on trouve le guide
        If InStr(1, guideSpec, nomGuide, vbTextCompare) > 0 Or _
           InStr(1, nomGuide, guideSpec, vbTextCompare) > 0 Then
            
            trouve = True
            
            ' Verifier si cette visite est dans les autorisees
            If InStr(1, visiteSpec, typeVisite, vbTextCompare) > 0 Or _
               InStr(1, typeVisite, visiteSpec, vbTextCompare) > 0 Then
                ' Visite trouvee dans les autorisees
                GuideAutoriseVisite = True
                Exit Function
            End If
            
            ' Cas special : "Tous sauf"
            If InStr(1, visiteSpec, "Tous sauf", vbTextCompare) > 0 Or _
               InStr(1, visiteSpec, "tous les autres", vbTextCompare) > 0 Then
                ' Guide fait tout sauf certaines visites
                ' Verifier dans la colonne Notes (C) les exclusions
                Dim notes As String
                notes = Trim(ws.Cells(i, 3).Value)
                
                ' Si le type de visite est dans les exclusions
                If InStr(1, notes, typeVisite, vbTextCompare) > 0 Then
                    GuideAutoriseVisite = False
                    Exit Function
                Else
                    GuideAutoriseVisite = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' Si guide trouve dans Specialisations mais visite pas listee = NON autorise
    ' (car c'est une liste RESTRICTIVE pour ces guides)
    If trouve Then
        ' Verifier si le guide a des entrees "UNIQUEMENT"
        For i = 4 To derLigne
            guideSpec = Trim(ws.Cells(i, 1).Value)
            visiteSpec = Trim(ws.Cells(i, 2).Value)
            
            If InStr(1, guideSpec, nomGuide, vbTextCompare) > 0 Then
                Dim notes2 As String
                notes2 = Trim(ws.Cells(i, 3).Value)
                
                ' Si mention "UNIQUEMENT" ou "SEULEMENT"
                If InStr(1, notes2, "UNIQUEMENT", vbTextCompare) > 0 Or _
                   InStr(1, notes2, "SEULEMENT", vbTextCompare) > 0 Or _
                   InStr(1, visiteSpec, "UNIQUEMENT", vbTextCompare) > 0 Then
                    ' C'est une liste restrictive
                    GuideAutoriseVisite = False
                    Exit Function
                End If
            End If
        Next i
    End If
    
    Exit Function
    
Erreur:
    ' En cas d'erreur, on autorise par securite
    GuideAutoriseVisite = True
End Function

' Fonction : Retourne la liste des guides autorises pour une visite
' Parametres :
'   - typeVisite : Type de visite
' Retour : Collection de noms de guides autorises
Public Function ObtenirGuidesAutorises(typeVisite As String) As Collection
    On Error GoTo Erreur
    
    Dim wsGuides As Worksheet
    Dim derLigne As Long
    Dim i As Long
    Dim nomGuide As String
    Dim col As New Collection
    
    Set wsGuides = ThisWorkbook.Sheets(FEUILLE_GUIDES)
    derLigne = wsGuides.Cells(wsGuides.Rows.Count, 1).End(xlUp).Row
    
    ' Parcourir tous les guides
    For i = 5 To derLigne ' Ligne 5 = premier guide
        nomGuide = Trim(wsGuides.Cells(i, 1).Value)
        
        If nomGuide <> "" Then
            ' Verifier si le guide peut faire cette visite
            If GuideAutoriseVisite(nomGuide, typeVisite) Then
                col.Add nomGuide
            End If
        End If
    Next i
    
    Set ObtenirGuidesAutorises = col
    Exit Function
    
Erreur:
    Set ObtenirGuidesAutorises = New Collection
End Function

' Fonction : Affiche un rapport des contraintes d'un guide
Public Sub AfficherContraintesGuide(nomGuide As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim derLigne As Long
    Dim i As Long
    Dim msg As String
    Dim guideSpec As String
    
    Set ws = ThisWorkbook.Sheets("Specialisations")
    
    If ws Is Nothing Then
        MsgBox "Aucune contrainte definie pour ce guide.", vbInformation
        Exit Sub
    End If
    
    msg = "CONTRAINTES POUR : " & nomGuide & vbCrLf & vbCrLf
    derLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim trouve As Boolean
    trouve = False
    
    For i = 4 To derLigne
        guideSpec = Trim(ws.Cells(i, 1).Value)
        
        If InStr(1, guideSpec, nomGuide, vbTextCompare) > 0 Then
            trouve = True
            msg = msg & " " & ws.Cells(i, 2).Value
            If Trim(ws.Cells(i, 3).Value) <> "" Then
                msg = msg & " (" & ws.Cells(i, 3).Value & ")"
            End If
            msg = msg & vbCrLf
        End If
    Next i
    
    If Not trouve Then
        msg = msg & "Aucune contrainte specifique." & vbCrLf
        msg = msg & "Ce guide peut effectuer toutes les visites."
    End If
    
    MsgBox msg, vbInformation, "Contraintes guide"
End Sub
