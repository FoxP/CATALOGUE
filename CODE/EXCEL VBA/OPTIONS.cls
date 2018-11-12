VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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

'-----------------------------------------------------------------------------------------------------------------------------------------------
'S'active lorsque l'utilisateur modifie une des options de la feuille "OPTIONS"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)

    'Mise à jour des ComboBox logiciels, langages, et types
    Call initializeComboBox
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'S'active lorsque l'utilisateur clique sur un lien hypertexte de la feuille "OPTIONS"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    If ActiveCell.Row = 2 Then
        'Reconstruit une copie du catalogue HTML à partir du catalogue Excel
        Call rebuildHtmlCatalogue
    End If
    
End Sub
