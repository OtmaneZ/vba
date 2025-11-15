Attribute VB_Name = "Module_Specialisations"
Option Explicit

' ===== VERIFIER AUTORISATION GUIDE POUR TYPE VISITE =====
Public Function GuideAutoriseVisite(nomGuide As String, typeVisite As String) As Boolean
    On Error Resume Next

    Dim ws As Worksheet
    Dim derLigne As Long
    Dim i As Long
    Dim nomGuideSpec As String
    Dim typeVisiteSpec As String
    Dim autorise As String
    Dim trouve As Boolean

    ' Par defaut : autorise (si pas de regle specifique)
    GuideAutoriseVisite = True
    trouve = False

    ' Recuperer feuille Specialisations
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("Specialisations")

    If ws Is Nothing Then
        Exit Function
    End If

    derLigne = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' STRUCTURE REELLE (selon analyse):
    ' Col 1: ID_Specialisation
    ' Col 2: Nom_Guide (NOM uniquement)
    ' Col 3: Email_Guide
    ' Col 4: Type_Prestation
    ' Col 5: Autorise (OUI/NON)

    For i = 2 To derLigne
        nomGuideSpec = UCase(Trim(ws.Cells(i, 2).Value))
        typeVisiteSpec = UCase(Trim(ws.Cells(i, 4).Value))
        autorise = UCase(Trim(ws.Cells(i, 5).Value))

        ' Verifier correspondance (nom peut etre partiel)
        If (InStr(1, UCase(nomGuide), nomGuideSpec, vbTextCompare) > 0 Or _
            InStr(1, nomGuideSpec, UCase(nomGuide), vbTextCompare) > 0) And _
           (InStr(1, UCase(typeVisite), typeVisiteSpec, vbTextCompare) > 0 Or _
            InStr(1, typeVisiteSpec, UCase(typeVisite), vbTextCompare) > 0) Then

            trouve = True

            If autorise = "OUI" Then
                GuideAutoriseVisite = True
            Else
                GuideAutoriseVisite = False
            End If

            Exit Function
        End If
    Next i

    ' Si aucune regle trouvee, autoriser par defaut
    If Not trouve Then
        GuideAutoriseVisite = True
    End If

    On Error GoTo 0
End Function

' ===== OBTENIR SPECIALISATIONS D'UN GUIDE =====
Public Function ObtenirSpecialisationsGuide(nomGuide As String) As Collection
    On Error Resume Next

    Dim ws As Worksheet
    Dim col As New Collection
    Dim i As Long
    Dim nomGuideSpec As String
    Dim typeVisite As String
    Dim autorise As String

    Set ws = ThisWorkbook.Worksheets("Specialisations")

    If ws Is Nothing Then
        Set ObtenirSpecialisationsGuide = col
        Exit Function
    End If

    For i = 2 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        nomGuideSpec = UCase(Trim(ws.Cells(i, 2).Value))
        typeVisite = Trim(ws.Cells(i, 4).Value)
        autorise = UCase(Trim(ws.Cells(i, 5).Value))

        If InStr(1, UCase(nomGuide), nomGuideSpec, vbTextCompare) > 0 And autorise = "OUI" Then
            col.Add typeVisite
        End If
    Next i

    Set ObtenirSpecialisationsGuide = col

    On Error GoTo 0
End Function
