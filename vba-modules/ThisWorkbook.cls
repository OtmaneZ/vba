' ============================================
' EVENEMENTS DU CLASSEUR
' A copier dans l'objet "ThisWorkbook" (pas un module)
' ATTENTION : NE PAS copier la ligne "Attribute VB_Name"
' ============================================

Option Explicit

' ============================================
' Evenement : Ouverture du classeur
' ============================================
Private Sub Workbook_Open()
    On Error Resume Next

    ' Message de bienvenue confirmant que les macros fonctionnent
    MsgBox "[OK] Bienvenue dans le systeme de planning !" & vbCrLf & vbCrLf & _
           "Les macros sont activees correctement." & vbCrLf & _
           "Vous pouvez utiliser toutes les fonctionnalites.", _
           vbInformation, "Systeme operationnel"

    ' Creer/Afficher la feuille d'accueil
    Call CreerFeuilleAccueil

    ' Activer la feuille d'accueil
    ThisWorkbook.Sheets("Accueil").Activate

    On Error GoTo 0
End Sub

' ============================================
' Evenement : Avant la fermeture du classeur
' ============================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Deconnecter l'utilisateur (nettoyage)
    If niveauAcces <> "" Then
        ' Nettoyer les variables de session
        utilisateurConnecte = ""
        niveauAcces = ""
        emailUtilisateur = ""
    End If
End Sub

' ============================================
' Evenement : Avant la sauvegarde
' ============================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Optionnel : proteger certaines feuilles avant la sauvegarde
    On Error Resume Next

    ' Exemple : proteger la feuille Configuration
    ' ThisWorkbook.Sheets("Configuration").Protect Password:="admin123"

    On Error GoTo 0
End Sub

' ============================================
' Evenement : Apres la sauvegarde
' ============================================
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    ' Optionnel : deproteger apres sauvegarde
    On Error Resume Next

    ' ThisWorkbook.Sheets("Configuration").Unprotect Password:="admin123"

    On Error GoTo 0
End Sub

' ============================================
' Evenement : Activation d'une feuille
' ============================================
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Optionnel : actions lors du changement de feuille
    ' Par exemple, masquer certains onglets selon le niveau d'acces

    On Error Resume Next

    ' Si un guide est connecte, masquer les feuilles admin
    If niveauAcces = "GUIDE" Then
        ThisWorkbook.Sheets("Calculs_Paie").Visible = xlSheetVeryHidden
        ThisWorkbook.Sheets("Configuration").Visible = xlSheetVeryHidden
    ElseIf niveauAcces = "ADMIN" Then
        ' L'admin voit tout
        ThisWorkbook.Sheets("Calculs_Paie").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Configuration").Visible = xlSheetVisible
    End If

    On Error GoTo 0
End Sub
