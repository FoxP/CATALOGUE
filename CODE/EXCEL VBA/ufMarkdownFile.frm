VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMarkdownFile 
   Caption         =   "Pi�ce jointe Markdown"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915.001
   OleObjectBlob   =   "ufMarkdownFile.frx":0000
End
Attribute VB_Name = "ufMarkdownFile"
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
'           Catalogue Excel / HTML de fonctions, morceaux de codes, astuces, m�thodes, fichiers, ... fr�quemment utilis�s
'           Capitalise le savoir tout en fonctionnant sur des syst�mes d'exploitation avec des politiques de s�curit� strictes
'           N�cessite Microsoft Excel (2010 ou sup�rieur), Visual Basic for Applications, ainsi qu'un navigateur internet r�cent
'
'D�pendances VBA :
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
'L'utilisateur appuie sur "Valider" : poursuite de la fonction cbFile_Click() si une pi�ce jointe ainsi qu'une description sont renseign�es
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbOKFile_Click()

    'Si pas de pi�ce jointe et pas de description renseign�es
    If delAllSpace(Me.cbFile.Value) = "" And delAllSpace(Me.tbDescription.Value) = "" Then
        MsgBox "Une pi�ce jointe et une description doivent �tre renseign�es.", vbInformation, "Saisie incompl�te"
    'Si pas de pi�ce jointe s�lectionn�e
    ElseIf delAllSpace(Me.cbFile.Value) = "" Then
        MsgBox "Une pi�ce jointe doit �tre s�lectionn�e.", vbInformation, "Pi�ce jointe invalide"
    'Si pas de description renseign�e
    ElseIf delAllSpace(Me.tbDescription.Value) = "" Then
        MsgBox "Une description doit �tre renseign�e.", vbInformation, "Description invalide"
    'Si une pi�ce jointe et une description sont renseign�es
    Else
        Me.Hide
        bCancelFile = False
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Annuler" : arr�t de la fonction cbFile_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCancelFile_Click()

    Me.Hide
    bCancelFile = True
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur ferme le UserForm (croix rouge) : arr�t de la fonction cbFile_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Me.Hide
    bCancelFile = True
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouverture de la pi�ce jointe s�lectionn�e lors du clic sur le bouton cbOpenFile
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbOpenFile_Click()

    Call OpenIt(strFile(Me.cbFile.Text))
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'S�lection du contenu de la TextBox tbDescription lorsqu'elle est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbDescription_Enter()

    Me.tbDescription.SelStart = 0
    Me.tbDescription.SelLength = Len(Me.tbDescription.Text)
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'S�lection du contenu de la ComboBox cbFile lorsqu'elle est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbFile_Enter()

    Me.cbFile.SelStart = 0
    Me.cbFile.SelLength = Len(Me.cbFile.Text)
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Lors de la s�lection d'une pi�ce jointe dans la ComboBox cbFile, renseigne son nom dans la TextBox tbDescription
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbFile_Change()

    'R�cup�ration de l'�ventuel texte s�lectionn� dans la zone de saisie active
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
    'Si aucune description, on la renseigne avec l'adresse de la fiche
    If delAllSpace(Me.tbDescription) = "" Then
        Me.tbDescription.Value = Me.cbFile.Value
    Else
        'Si pas de texte s�lectionn�, on remplace la description actuelle
        If sSelectedText = "" Then
            If delAllSpace(Me.tbDescription.Value) <> delAllSpace(Me.cbFile.Value) Then
'                If MsgBox("Remplacer le texte par le titre de la fiche ?", vbQuestion + vbYesNo, "Titre du lien") = vbYes Then
                    Me.tbDescription.Value = Me.cbFile.Value
'                End If
            End If
        End If
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Positionnement du UserForm au milieu du UserForm de cr�ation / �dition de fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_activate()

    With Me
        'Positionnement du UserForm au milieu du UserForm de cr�ation / �dition de fiche
        .StartUpPosition = 0
        .Left = ufCATALOGUE.Left + (0.5 * ufCATALOGUE.Width) - (0.5 * .Width)
        .Top = ufCATALOGUE.Top + (0.5 * ufCATALOGUE.Height) - (0.5 * .Height)
        'Focus sur la ComboBox de s�lection de pi�ce jointe
        .cbFile.SetFocus
    End With
     
End Sub

