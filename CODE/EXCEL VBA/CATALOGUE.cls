VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
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

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Lors d'un clic droit dans la feuille "CATALOGUE", ajoute ou supprime des entr�es dans le menu contextuel du clic droit
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    'Si clic en dehors du header du tableau
    If ActiveCell.Row > 2 Then
        'Si clic sur une ligne vide
        If catalogueSheet.Cells(ActiveCell.Row, colId).Value = "" Or catalogueSheet.Cells(ActiveCell.Row, colTitle).Value = "" Or catalogueSheet.Cells(ActiveCell.Row, colVersion).Value = "" Then
            Call addCommandBar(False)
        Else
            Call addCommandBar
        End If
    Else
        Call addCommandBar(False)
    End If
End Sub
