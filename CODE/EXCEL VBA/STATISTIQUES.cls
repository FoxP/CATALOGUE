VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil3"
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
'Lorsque la feuille "STATISTIQUES" est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Activate()

    Application.StatusBar = "Mise � jour de l'onglet STATISTIQUES en cours..."

    'Met � jour les tableaux crois�s dynamiques ainsi que les graphiques de statistiques du catalogue Excel
    Call updateWorkbookData
    
    Application.StatusBar = ""
    
End Sub
