VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
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
'A la fermeture du classeur
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    'Destruction du menu contextuel du clic droit
    Call deleteCommandBar

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'A l'ouverture du classeur
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_Open()

    'D�finition des variables publiques
    Call definePublicVar
    
    'Si le classeur est partag�, on masque l'onglet "STATISTIQUES"
    If ActiveWorkbook.MultiUserEditing Then
        statsSheet.Visible = xlSheetVeryHidden
    Else
        statsSheet.Visible = xlSheetVisible
    End If
    
    'Suppression des filtres du classeur
    On Error Resume Next
    If catalogueSheet.AutoFilterMode Then
        catalogueSheet.ShowAllData
    End If
    On Error GoTo 0
    
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Si le classeur a �t� ex�cut� depuis un fichier batch (.bat), on lance la cr�ation d'une fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
    
    If InStr(1, GetEnvironmentVariable("InBatch"), "TRUE", vbTextCompare) Then
        'N'affiche pas Excel
        'Application.Visible = False
        
        'Cr�ation d'une fiche
        Call addToCatalogue
        
        'Quitte Excel et ferme le classeur
        ThisWorkbook.Saved = True
        Application.Quit
        Application.ActiveWindow.Close SaveChanges:=False
        ActiveWorkbook.Close SaveChanges:=False
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Lors du changement de feuille dans le classeur
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_SheetActivate(ByVal Sh As Object)

    'Si ce n'est pas la feuille "CATALOGUE"
    If Not Sh Is catalogueSheet Then
        'Destruction du menu contextuel du clic droit
        Call deleteCommandBar
    End If

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Lors du changement de classeur � l'�cran
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)

    'Destruction du menu contextuel du clic droit
    Call deleteCommandBar
    
End Sub
