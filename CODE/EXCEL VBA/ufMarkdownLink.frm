VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMarkdownLink 
   Caption         =   "Lien Markdown"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915.001
   OleObjectBlob   =   "ufMarkdownLink.frx":0000
End
Attribute VB_Name = "ufMarkdownLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================================================================================
'
'                                                               CATALOGUE
'
'Auteur :
'           Paul RENARD
'
'Fonction :
'           Catalogue Excel / HTML de fonctions, morceaux de codes, astuces, méthodes, fichiers, ... fréquemment utilisés
'           Capitalise le savoir tout en fonctionnant sur des systèmes d'exploitation avec des politiques de sécurité strictes
'           Nécessite Microsoft Excel (2010 ou supérieur), Visual Basic for Applications, ainsi qu'un navigateur internet récent
'
'Dépendances VBA :
'           - Microsoft Scripting Runtime
'           - Visual Basic for Applications
'           - Microsoft Forms XX Object Library
'           - Microsoft Excel XX Object Library
'           - Microsoft Office XX Object Library
'           - Microsoft VBScript Regular Expressions 5.5
'
'===============================================================================================================================================

Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Valider" : poursuite de la fonction cbLink_Click() si une adresse ainsi qu'un texte de lien sont renseignés
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbOKLink_Click()

    'Si pas d'adresse et pas de texte renseignés
    If delAllSpace(Me.tbLink.Value) = "" And delAllSpace(Me.tbText.Value) = "" Then
        MsgBox "Une adresse et un texte doivent être renseignés.", vbInformation, "Saisie incomplète"
    'Si pas d'adresse renseignée
    ElseIf delAllSpace(Me.tbLink.Value) = "" Then
        MsgBox "Une adresse doit être renseignée.", vbInformation, "Adresse invalide"
    'Si pas de texte renseigné
    ElseIf delAllSpace(Me.tbText.Value) = "" Then
        MsgBox "Un texte doit être renseigné.", vbInformation, "Texte invalide"
    'Si une adresse et un texte sont renseignés
    Else
        Me.Hide
        bCancelLink = False
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Annuler" : arrêt de la fonction cbLink_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCancelLink_Click()

    Me.Hide
    bCancelLink = True
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur ferme le UserForm (croix rouge) : arrêt de la fonction cbLink_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Me.Hide
    bCancelLink = True
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Sélection du contenu de la TextBox tbText lorsqu'elle est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbText_Enter()

    Me.tbText.SelStart = 0
    Me.tbText.SelLength = Len(Me.tbText.Text)
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Sélection du contenu de la TextBox tbLink lorsqu'elle est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbLink_Enter()

    Me.tbLink.SelStart = 0
    Me.tbLink.SelLength = Len(Me.tbLink.Text)
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Vérifie la saisie utilisateur dans la TextBox tbLink : si une URL de fiche est détectée, renseigne son titre dans la TextBox tbText
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbLink_AfterUpdate()

    'Définition de l'expression régulière de détection d'URL de fiche
    Dim regSheetURL As New VBScript_RegExp_55.RegExp
    regSheetURL.Pattern = ".*[\/\\]SHEETS*[\/\\]([0-9]{6}_.*_[0-9]+\.html)"
    'Définition de l'expression régulière d'ID de fiche
    Dim regSheetID As New VBScript_RegExp_55.RegExp
    regSheetID.Pattern = ".*[\/\\]SHEETS*[\/\\]([0-9]{6})_.*_[0-9]+\.html"
    'Définition de l'expression régulière de version de fiche
    Dim regSheetVersion As New VBScript_RegExp_55.RegExp
    regSheetVersion.Pattern = ".*[\/\\]SHEETS*[\/\\][0-9]{6}_.*_([0-9]+)\.html"
    'Récupération de l'éventuel texte sélectionné dans la TextBox active
    Dim sSelectedText As String
    If ufCATALOGUE.tbProblem.Height > 100 Then
        sSelectedText = ufCATALOGUE.tbProblem.SelText
    ElseIf ufCATALOGUE.tbSolution.Height > 100 Then
        sSelectedText = ufCATALOGUE.tbSolution.SelText
    ElseIf ufCATALOGUE.tbCode.Height > 100 Then
        sSelectedText = ufCATALOGUE.tbCode.SelText
    Else
        sSelectedText = ""
    End If
    'Si c'est une URL de fiche
    If regSheetURL.test(ufMarkdownLink.tbLink.Value) Then
        Dim lVersion As Long: lVersion = regSheetVersion.Execute(ufMarkdownLink.tbLink.Value)(0).SubMatches(0)
        Dim lID As Long: lID = regSheetID.Execute(ufMarkdownLink.tbLink.Value)(0).SubMatches(0)
        Dim freeLine As Integer: freeLine = 3
        Do While catalogueSheet.Cells(freeLine, colId).Value <> ""
            'Si la fiche est trouvée dans le catalogue Excel
            If catalogueSheet.Cells(freeLine, colId).Value = lID And catalogueSheet.Cells(freeLine, colVersion).Value = lVersion Then
                'Si aucun texte, on le renseigne avec le titre de la fiche
                If delAllSpace(Me.tbText) = "" Then
                    Me.tbText.Value = catalogueSheet.Cells(freeLine, colTitle).Value
                'Sinon, on demande s'il faut le remplacer par le titre de la fiche
                Else
                    If sSelectedText = "" Then
                        If delAllSpace(Me.tbText.Value) <> delAllSpace(catalogueSheet.Cells(freeLine, colTitle).Value) Then
                            If MsgBox("Remplacer le texte par le titre de la fiche ?", vbQuestion + vbYesNo, "Titre du lien") = vbYes Then
                                Me.tbText.Value = catalogueSheet.Cells(freeLine, colTitle).Value
                            End If
                        End If
                    End If
                End If
                Exit Do
            End If
            freeLine = freeLine + 1
        Loop
    'Si ce n'est pas une URL de fiche
    Else
        'Si aucun texte, on le renseigne avec l'adresse de la fiche
        If delAllSpace(Me.tbText) = "" Then
            Me.tbText.Value = Me.tbLink.Value
        End If
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Récupération du contenu du presse papier pour remplir le champ "Adresse" du lien (TextBox tbLink) lors du clic sur le bouton cbPasteLink
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbPasteLink_Click()

    'Si le champ d'adresse du lien n'est pas vide, on demande s'il faut remplacer son contenu
    If Replace(Me.tbLink.Value, " ", "") <> "" Then
        If MsgBox("Le champ d'adresse n'est pas vide, souhaitez vous le remplacer par le contenu du presse papier ? Les données existantes seront perdues.", vbYesNo + vbQuestion, "Continuer ?") = vbNo Then
            Exit Sub
        End If
    End If
    'Copie du presse papier dans le champ d'adresse du lien
    Me.tbLink.Value = getClipboardContent()
    'Remplissage automatique du champ de texte du lien
    Call tbLink_AfterUpdate
    'Focus sur le champ de texte du lien
    Me.tbText.SetFocus

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Positionnement du UserForm au milieu du UserForm de création / édition de fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_activate()

    With Me
        'Positionnement du UserForm au milieu du UserForm de création / édition de fiche
        .StartUpPosition = 0
        .Left = ufCATALOGUE.Left + (0.5 * ufCATALOGUE.Width) - (0.5 * .Width)
        .Top = ufCATALOGUE.Top + (0.5 * ufCATALOGUE.Height) - (0.5 * .Height)
        'Focus sur la TextBox d'adresse de lien
        .tbLink.SetFocus
    End With
    
End Sub

