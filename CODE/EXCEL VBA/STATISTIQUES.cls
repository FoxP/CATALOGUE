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
'Lorsque la feuille "STATISTIQUES" est active
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Activate()

    Application.StatusBar = "Mise à jour de l'onglet STATISTIQUES en cours..."

    'Met à jour les tableaux croisés dynamiques ainsi que les graphiques de statistiques du catalogue Excel
    Call updateWorkbookData
    
    Application.StatusBar = ""
    
End Sub
