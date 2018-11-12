VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMarkdownTable 
   Caption         =   "Tableau Markdown"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5070
   OleObjectBlob   =   "ufMarkdownTable.frx":0000
End
Attribute VB_Name = "ufMarkdownTable"
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

'Nombre de colonnes du tableau
Public plColsNbr As Long
'Nombre de lignes du tableau
Public plRowsNbr As Long

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Valider" : poursuite de la fonction cbTable_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbOKTable_Click()

    Me.Hide
    bCancelTable = False
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Annuler" : arrêt de la fonction cbTable_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCancelTable_Click()

    Me.Hide
    bCancelTable = True
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur ferme le UserForm (croix rouge) : arrêt de la fonction cbTable_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Me.Hide
    bCancelTable = True
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Vérifie la saisie utilisateur dans la TextBox tbRowsNbr par le biais de la fonction OnlyNumbers()
'N'accepte que les nombres entiers >= 1, sinon remplace par la valeur par défaut de la TextBox : 1
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbRowsNbr_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If OnlyNumbers(Replace(Me.tbRowsNbr.Value, ".", ",")) Then
        If CInt(Replace(Me.tbRowsNbr.Value, ".", ",")) >= 1 Then
            plRowsNbr = CInt(Replace(Me.tbRowsNbr.Value, ".", ","))
            Me.tbRowsNbr.Value = CInt(Replace(Me.tbRowsNbr.Value, ".", ","))
        Else
            Me.tbRowsNbr.Value = 1
        End If
    Else
        Me.tbRowsNbr.Value = plRowsNbr
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Sélection du contenu de la TextBox tbRowsNbr lorsqu'elle est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbRowsNbr_Enter()

    plRowsNbr = CInt(Replace(Me.tbRowsNbr.Value, ".", ","))
    Me.tbRowsNbr.SelStart = 0
    Me.tbRowsNbr.SelLength = Len(Me.tbRowsNbr.Text)
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Vérifie la saisie utilisateur dans la TextBox tbColsNbr par le biais de la fonction OnlyNumbers()
'N'accepte que les nombres entiers >= 2, sinon remplace par la valeur par défaut de la TextBox : 2
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbColsNbr_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If OnlyNumbers(Replace(Me.tbColsNbr.Value, ".", ",")) Then
        If CInt(Replace(Me.tbColsNbr.Value, ".", ",")) >= 2 Then
            plColsNbr = CInt(Replace(Me.tbColsNbr.Value, ".", ","))
            Me.tbColsNbr.Value = CInt(Replace(Me.tbColsNbr.Value, ".", ","))
        Else
            Me.tbColsNbr.Value = 2
        End If
    Else
        Me.tbColsNbr.Value = plColsNbr
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Sélection du contenu de la TextBox tbColsNbr lorsqu'elle est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub tbColsNbr_Enter()

    plColsNbr = CInt(Replace(Me.tbColsNbr.Value, ".", ","))
    Me.tbColsNbr.SelStart = 0
    Me.tbColsNbr.SelLength = Len(Me.tbColsNbr.Text)
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Positionnement du UserForm au milieu du UserForm de création / édition de fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_activate()

    With Me
        .StartUpPosition = 0
        .Left = ufCATALOGUE.Left + (0.5 * ufCATALOGUE.Width) - (0.5 * .Width)
        .Top = ufCATALOGUE.Top + (0.5 * ufCATALOGUE.Height) - (0.5 * .Height)
        .tbRowsNbr.SetFocus
    End With
    
End Sub

