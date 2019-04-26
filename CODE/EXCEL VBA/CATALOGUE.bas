Attribute VB_Name = "CATALOGUE"
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

'===============================================================================================================================================
'===============================================================================================================================================
'
'Déclaration des variables publiques partagées entre le module CATALOGUE et le UserForm ufCATALOGUE
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Variables de création / édition de fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public strTitle As String                                                       'Titre de la fiche
Public strType As String                                                        'Type de la fiche
Public strCode As String                                                        'Code de la fiche
Public dblId As Double                                                          'ID de la fiche
Public intVersion As Integer                                                    'Version de la fiche
Public strStatus As String                                                      'Statut de la fiche
Public strKeywords As String                                                    'Mots clés de la fiche
Public strSoftware As String                                                    'Logiciel de la fiche
Public strLanguage As String                                                    'Langage de la fiche
Public strProblem As String                                                     'Problème de la fiche
Public strSolution As String                                                    'Solution de la fiche
Public isSuperseded As Boolean                                                  'True si modification d'une fiche released
Public isObsolete As Boolean                                                    'True si obsolescence d'une fiche
Public strIDLanguage As String                                                  'Identifiant correspondant au langage pour highlight.js
Public strFile As New Collection                                                'Collection de pièces jointes à copier / conserver dans la fiche
Public strToDelete As New Collection                                            'Collection de pièces jointes à supprimer dans la fiche à éditer
Public cataloguePath As String                                                  'Emplacement du catalogue Excel
Public statsSheet As Worksheet                                                  'Feuille "STATISTIQUES"
Public optionsSheet As Worksheet                                                'Feuille "OPTIONS"
Public catalogueSheet As Worksheet                                              'Feuille "CATALOGUE"

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Constantes
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public Const filesPath As String = "FILES"                                      'Dossier des pièces jointes des fiches
Public Const sheetsPath As String = "\SHEETS\"                                  'Dossier des fiches
Public Const templateName As String = "TEMPLATE.html"                           'Nom du template des fiches
Public Const htmlCataloguePath As String = "\CATALOGUE\"                        'Dossier du catalogue HTML
Public Const htmlCatalogueName As String = "CATALOGUE.html"                     'Nom du catalogue HTML
Public Const imgArray As String = "jpg,jpeg,jpe,png,bmp,tif,tiff,gif,webp,svg"  'Extensions des images supportées par les fiches
Public Const htmlPreviewName As String = "TEMPLATE_PREVIEW.html"                'Nom du template de prévisualisation des fiches
Public Const htmlPreviewTmpName As String = "PREVIEW_TMP.html"                  'Nom du fichier temporaire de prévisualisation des fiches
Public Const htmlCatalogueTemplateName As String = "TEMPLATE_CATALOGUE.html"    'Nom du template du catalogue

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Identifiants de colonnes de la feuille Excel "CATALOGUE"
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public colId As Integer
Public colTitle As Integer
Public colVersion As Integer
Public colStatus As Integer
Public colSoftware As Integer
Public colLanguage As Integer
Public colType As Integer
Public colDatec As Integer
Public colDatem As Integer
Public colUser As Integer
Public colKeywords As Integer

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Identifiants de colonnes de la feuille Excel "OPTIONS"
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public colOptLanguages As Integer
Public colOptIdLanguages As Integer
Public colOptSoftwares As Integer
Public colOptTypes As Integer
Public colOptStatus As Integer

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Quitte ou non la fonction addToCatalogue()
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public bExit As Boolean

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Quitte ou non la fonction cbFile_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public bCancelFile As Boolean

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Quitte ou non la fonction cbLink_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public bCancelLink As Boolean

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Quitte ou non la fonction cbTable_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public bCancelTable As Boolean

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Détermine si l'application Microsoft Outlook est installée ou non
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public bIsOutlookInstalled As Boolean

'===============================================================================================================================================
'===============================================================================================================================================
'
'Définition / réitinitalisation des variables publiques partagées entre le module CATALOGUE et le UserForm ufCATALOGUE
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Définition de l'emplacement du catalogue et des identifiants de colonnes des feuilles Excel "CATALOGUE" et "OPTIONS"
'Vérifie également si l'application Microsoft Outlook est installée. Si ce n'est pas le cas, les fonctionnalités liées sont masquées
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub definePublicVar()

    'Emplacement du catalogue
    cataloguePath = Application.ThisWorkbook.Path
    
    'Feuilles du catalogue
    Set statsSheet = ThisWorkbook.Sheets("STATISTIQUES")
    Set optionsSheet = ThisWorkbook.Sheets("OPTIONS")
    Set catalogueSheet = ThisWorkbook.Sheets("CATALOGUE")
    
    'Définition des numéros de colonnes de la feuille "CATALOGUE" en fonction des noms de colonnes
    colId = catalogueSheet.Range("id").Column
    colTitle = catalogueSheet.Range("title").Column
    colVersion = catalogueSheet.Range("version").Column
    colStatus = catalogueSheet.Range("status").Column
    colSoftware = catalogueSheet.Range("software").Column
    colLanguage = catalogueSheet.Range("language").Column
    colType = catalogueSheet.Range("type").Column
    colDatec = catalogueSheet.Range("datec").Column
    colDatem = catalogueSheet.Range("datem").Column
    colUser = catalogueSheet.Range("user").Column
    colKeywords = catalogueSheet.Range("keywords").Column
    
    'Définition des numéros de colonnes de la feuille "OPTIONS" en fonction des noms de colonnes
    colOptLanguages = optionsSheet.Range("optlanguages").Column
    colOptIdLanguages = optionsSheet.Range("optidlanguages").Column
    colOptSoftwares = optionsSheet.Range("optsoftwares").Column
    colOptTypes = optionsSheet.Range("opttypes").Column
    colOptStatus = optionsSheet.Range("optstatus").Column
    
    'Vérification si l'application Microsoft Outlook est installée
    Err.Clear
    On Error Resume Next
    Dim MonAppliOutlook As Object
    Set MonAppliOutlook = CreateObject("Outlook.Application")
    'Si l'opération n'a pas levé d'erreur
    If Err.Number = 0 Then
        bIsOutlookInstalled = True
    Else
        bIsOutlookInstalled = False
    End If
    Err.Clear
    On Error GoTo 0
    Set MonAppliOutlook = Nothing

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Réinitialisation des variables publiques de création / édition de fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function resetPublicVar()

    strTitle = ""
    dblId = 0
    intVersion = 0
    strType = ""
    strStatus = ""
    strKeywords = ""
    strSoftware = ""
    strLanguage = ""
    strCode = ""
    strProblem = ""
    strSolution = ""
    isSuperseded = False
    isObsolete = False
    strIDLanguage = ""
    Set strFile = Nothing
    Set strToDelete = Nothing
    
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions liées à l'édition des catalogue Excel et HTML, ainsi que des fiches HTML
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Création / édition d'une fiche dans le catalogue (Excel + HTML) : visible dans la fenêtre de macros (Alt + F8)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub createSheet()

    Call addToCatalogue
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Création / édition d'une fiche dans le catalogue (Excel + HTML) : non visible dans la fenêtre de macros (Alt + F8)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub addToCatalogue(Optional ByVal openUserform = True)

    'Création / modification d'une fiche
    'Si fiche en train d'être passée Superseded ou Obsolete, on n'ouvre pas le UserForm
    If openUserform = True Then
        Application.StatusBar = "Edition de fiche en cours..."
        ufCATALOGUE.Show
        Application.StatusBar = ""
    End If
    
    'L'utilisateur a cliqué sur "Annuler" ou la croix rouge dans le UserForm
    If bExit = True Then
        bExit = False
        Exit Sub
    End If
    
    'Un des champs requis n'a pas été remplis, cette condition ne devrait jamais apparaître
    If strTitle = "" Or intVersion = 0 Or strStatus = "" Or strLanguage = "" Or strSoftware = "" Or strType = "" Or strKeywords = "" Then
        Exit Sub
    End If
    
modifyCatalogueAndSheet:

    Application.StatusBar = "Sauvegarde de fiche en cours..."

'###############################################################################################################################################
'1. Création / édition d'une fiche HTML
'###############################################################################################################################################
    
    'Récupération de l'identifiant correspondant au langage pour highlight.js
    'Voir : http://highlightjs.readthedocs.io/en/latest/css-classes-reference.html
    strIDLanguage = getIDLanguage(strLanguage)
    
    'Les mots clés sont séparés par des espaces dans le UserForm
    Dim tmpKeywords As Variant
    Dim htmlKeywords As String
    tmpKeywords = Split(strKeywords, " ")
    htmlKeywords = ""
    Dim k As Integer
    For k = 0 To UBound(tmpKeywords)
        'Balise <span>MOT_CLE_1</span><span>MOT_CLE_2</span><span>MOT_CLE_3</span>
        htmlKeywords = htmlKeywords & vbTab & vbTab & vbTab & "<span>" & tmpKeywords(k) & "</span>" & vbNewLine
    Next
    
    'Création du dossier des pièces jointes (s'il n'existe pas)
    Call createFolder(cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion))
    
    'Tri des pièces jointes par ordre alphabétique pour la fiche
    Call sortCollection(strFile)
    
    'Gestion des pièces jointes : images et documents
    Dim htmlPictures As String: htmlPictures = ""
    Dim htmlDocuments As String: htmlDocuments = ""
    Dim cptPictures As Integer: cptPictures = 0
    Dim cptDocuments As Integer: cptDocuments = 0
    Dim Filename As String: Filename = ""
    'Si l'utilisateur a supprimé des pièces jointes dans une fiche existante
    Dim i As Integer
    For i = 1 To strToDelete.Count
        'Ne pas supprimer les pièces jointes en dehors du dossier de pièces jointes de la fiche
        If InStr(strToDelete.item(i), replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion)) <> 0 Then
            'Suppression des pièces jointes
            deleteFile (strToDelete.item(i))
        End If
    Next
    'Parcourt de la collection de pièces jointes à copier / conserver
    For i = 1 To strFile.Count
        'Passage de l'extension de la pièce jointe en minuscules pour qu'elle soit bien repérée par la feuille de style "icons.css" de la fiche
        Filename = Replace(cleanNameForMarkdown(getFilenameFromPath(strFile.item(i))), getExtensionFromPath(strFile.item(i)), LCase(getExtensionFromPath(strFile.item(i))))
        'Déplacement de la pièce jointe vers le dossier de pièces jointes de la fiche
        Call copyFileFromTo(strFile.item(i), cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "\" & Filename, True)
        'Si la pièce jointe est une image
        If isInArray(LCase(getExtensionFromPath(strFile.item(i))), Split(imgArray, ",")) Then
            cptPictures = cptPictures + 1
            'Balise <a class='galery' rel='alternate' title='NOM' href='CHEMIN'><img src='CHEMIN' alt='NOM'/></a>
            htmlPictures = htmlPictures & vbTab & vbTab & vbTab & "<a class=" & Chr(34) & "galery" & Chr(34) & " rel=" & Chr(34) & "alternate" & Chr(34) & " title=" & Chr(34) & Filename & Chr(34) & " href=" & Chr(34) & filesPath & "/" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "/" & Filename & Chr(34) & ">" & _
                           "<img src=" & Chr(34) & filesPath & "/" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "/" & Filename & Chr(34) & " alt=" & Chr(34) & Filename & Chr(34) & "/></a>" & vbNewLine
        'Sinon, la pièce jointe est un document
        Else
            cptDocuments = cptDocuments + 1
            'Balise <a target='_blank' href='CHEMIN'>NOM</a>
            htmlDocuments = htmlDocuments & vbTab & vbTab & vbTab & "<a target=" & Chr(34) & "_blank" & Chr(34) & " href=" & Chr(34) & filesPath & "/" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "/" & Filename & Chr(34) & ">" & Filename & "</a>" & vbNewLine
        End If
    Next

    'Copie / remplissage du template HTML pour création / édition de fiche
    Dim strIn 'As String
    strIn = getFileContent(cataloguePath & sheetsPath & templateName)
    strIn = Replace(strIn, "{TEMPLATE_ID}", dblId)
    strIn = Replace(strIn, "{TEMPLATE_TITLE}", strTitle)
    strIn = Replace(strIn, "{TEMPLATE_VERSION}", intVersion)
    strIn = Replace(strIn, "{TEMPLATE_SNIPPET_CODE}", htmlSpecialChars(strCode))
    strIn = Replace(strIn, "{TEMPLATE_TYPE}", strType)
    strIn = Replace(strIn, "{TEMPLATE_PROBLEM}", strProblem)
    strIn = Replace(strIn, "{TEMPLATE_SOLUTION}", strSolution)
    strIn = Replace(strIn, "{TEMPLATE_STATUS}", strStatus)
    strIn = Replace(strIn, "{TEMPLATE_SOFTWARE}", strSoftware)
    strIn = Replace(strIn, "{TEMPLATE_LANGUAGE}", strLanguage)
    strIn = Replace(strIn, "{TEMPLATE_IDLANG}", strIDLanguage)
    strIn = Replace(strIn, "{TEMPLATE_KEYWORDS}", htmlKeywords)
    strIn = Replace(strIn, "{TEMPLATE_PICTURES}", htmlPictures)
    strIn = Replace(strIn, "{TEMPLATE_DOCUMENTS}", htmlDocuments)
    strIn = Replace(strIn, "{TEMPLATE_FOLDER}", filesPath & "/" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion))
    'Si pas d'image, on supprime la section de la fiche
    If cptPictures = 0 Then
        strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "pictures" & Chr(34) & ">", "</section>"), "")
    End If
    'Si pas de documents, on supprime la section de la fiche
    If cptDocuments = 0 Then
        strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "documents" & Chr(34) & ">", "</section>"), "")
    End If
    'Si pas de code, on supprime la section de la fiche
    If strCode = "" Then
        strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "code" & Chr(34) & ">", "</section>"), "")
    End If
    'Si pas de problème, on supprime la section de la fiche
    If strProblem = "" Then
        strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "problem" & Chr(34) & ">", "</section>"), "")
    End If
    'Si pas de solution, on supprime la section de la fiche
    If strSolution = "" Then
        strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "solution" & Chr(34) & ">", "</section>"), "")
    End If
    'Ecriture de la fiche HTML
    Call writeToFile(cataloguePath & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html", strIn)
    'Si la fiche n'est pas en train d'être passée Superseded ou Obsolete, on l'ouvre dans le navigateur
    If strStatus <> "Superseded" And strStatus <> "Obsolete" Then
        Call browseURL(cataloguePath & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html")
    End If
    
'###############################################################################################################################################
'2. Ajout / édition d'une fiche dans le catalogue Excel
'###############################################################################################################################################
    
    'Evite un conflit si entre temps un utilisateur a rajouté une ligne dans le catalogue
    'Quand un classeur Excel est partagé, effectuer une sauvegarde le met à jour automatiquement
    ThisWorkbook.Save
    
    'Ajout / mise à jour de la fiche dans le catalogue Excel
    Dim freeLine As Integer: freeLine = 3
    Do While catalogueSheet.Cells(freeLine, colId).Value <> ""
        'Recherche de la fiche dans le catalogue : si existe, on met à jour
        If catalogueSheet.Cells(freeLine, colId).Value = dblId And catalogueSheet.Cells(freeLine, colVersion).Value = intVersion Then
            Exit Do
        'Sinon, on continue à parcourir le catalogue pour trouver une ligne vide
        Else
            freeLine = freeLine + 1
        End If
    Loop
    'Colonne 3 : ID de la fiche
    catalogueSheet.Cells(freeLine, colId).Value = dblId
    'Colonne 4 : Titre de la fiche
    catalogueSheet.Cells(freeLine, colTitle).Value = strTitle
    'Colonne 5 : Version de la fiche
    catalogueSheet.Cells(freeLine, colVersion).Value = intVersion
    'Colonne 6 : Statut de la fiche
    catalogueSheet.Cells(freeLine, colStatus).Value = strStatus
    'Colonne 7 : Logiciel de la fiche
    catalogueSheet.Cells(freeLine, colSoftware).Value = strSoftware
    'Colonne 8 : Langage de la fiche
    catalogueSheet.Cells(freeLine, colLanguage).Value = strLanguage
    'Colonne 9 : Type de la fiche
    catalogueSheet.Cells(freeLine, colType).Value = strType
    'Colonnes 10 et 11 : Respectivement date de création et édition de la fiche
    If catalogueSheet.Cells(freeLine, colDatec).Value <> "" Then                                                'Si la fiche existe, on ajoute la date de mise à jour
        catalogueSheet.Cells(freeLine, colDatem).Value = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
    Else                                                                                                        'Sinon on ajoute la date de création
        catalogueSheet.Cells(freeLine, colDatec).Value = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
    End If
    'Colonne 12 : Créateur de la fiche
    If catalogueSheet.Cells(freeLine, colUser).Value = "" Then
        catalogueSheet.Cells(freeLine, colUser).Value = Application.UserName
    End If
    'Colonnes 13+ : Mots clés de la fiche
    For k = 0 To UBound(tmpKeywords)
        catalogueSheet.Cells(freeLine, colKeywords + k).Value = tmpKeywords(k)
    Next
    'En cas de mise à jour de la fiche, il peut arriver qu'il y ait moins de mots clés qu'avant
    If catalogueSheet.Cells(freeLine, colKeywords + UBound(tmpKeywords) + 1).Value <> "" Then
        k = colKeywords + UBound(tmpKeywords) + 1
        While catalogueSheet.Cells(freeLine, k).Value <> ""
            'On supprime donc les mots clés en trop
            catalogueSheet.Cells(freeLine, k).Value = ""
            k = k + 1
        Wend
    End If
    
'###############################################################################################################################################
'3. Ajout / édition d'une fiche dans le catalogue HTML
'###############################################################################################################################################
    
    'Mots clés
    htmlKeywords = ""
    For k = 0 To UBound(tmpKeywords)
        'Balise <span>MOT_CLE_1</span><span>MOT_CLE_2</span><span>MOT_CLE_3</span>
        htmlKeywords = htmlKeywords & "<span>" & tmpKeywords(k) & "</span>"
    Next
    
    'Créateur de la fiche
    Dim strUser As String
    'Si création de fiche
    If catalogueSheet.Cells(freeLine, colUser).Value = "" Then
        'Récupération du nom d'utilisateur défini dans l'application Excel
        strUser = Application.UserName
    'Si édition de fiche
    Else
        'Récupération du nom du créateur de la fiche
        strUser = catalogueSheet.Cells(freeLine, colUser).Value
    End If
    
    'Récupération du contenu du catalogue HTML
    strIn = getFileContent(cataloguePath & htmlCataloguePath & htmlCatalogueName)
    'Si la fiche n'existe pas dans le catalogue HTML, on l'ajoute
    If InStr(strIn, "<tr class='" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "'>") = 0 Then
        strIn = Replace(strIn, "</tbody>", _
            vbTab & "<tr class='" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "'>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & dblId & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td><a href='.." & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html'>" & strTitle & "</a></td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & intVersion & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strStatus & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strSoftware & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strLanguage & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strType & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & VBA.Left(Year(Date) & "/" & Month(Date) & "/" & Day(Date), 10) & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td></td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strUser & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & htmlKeywords & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "</tr>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & "</tbody>")
    'Sinon la fiche existe donc on la met à jour (remplace) dans le catalogue
    Else
        Dim strReplaced, strReplacing As String
        strReplaced = GetBetween(strIn, "<tr class='" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "'>", "</tr>")
        strReplacing = vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & dblId & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td><a href='.." & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html'>" & strTitle & "</a></td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & intVersion & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strStatus & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strSoftware & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strLanguage & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strType & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & VBA.Left(catalogueSheet.Cells(freeLine, colDatec).Value, 10) & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & VBA.Left(Year(Date) & "/" & Month(Date) & "/" & Day(Date), 10) & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strUser & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & htmlKeywords & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab
        strIn = Replace(strIn, strReplaced, strReplacing)
    End If
    'Ecriture du catalogue HTML
    Call writeToFile(cataloguePath & htmlCataloguePath & htmlCatalogueName, strIn)
    
    'Si nouvelle fiche Released, passage de la release précédente en Superseded
    If isSuperseded = True Then
        'Modification de l'ancienne fiche
        isSuperseded = False
        intVersion = intVersion - 1
        strStatus = "Superseded"
        GoTo modifyCatalogueAndSheet
    Else
        'Vidage des variables publiques
        Call resetPublicVar
    End If
    
    'Sauvegarde du classeur Excel
    ThisWorkbook.Save
    
    Application.StatusBar = ""

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renvoie le numéro de la ligne Excel de la dernière version d'un ID de fiche donné
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function getLastVersionRow(ByVal dblId As Double) As Integer
    Dim tmpVersion As Integer: tmpVersion = 0
    Dim tmpRow As Integer: tmpRow = 0
    Dim freeLine As Integer: freeLine = 3
    Do While catalogueSheet.Cells(freeLine, colId).Value <> ""
        If catalogueSheet.Cells(freeLine, colId).Value = dblId Then
            If catalogueSheet.Cells(freeLine, colVersion).Value > tmpVersion Then
                tmpVersion = catalogueSheet.Cells(freeLine, colVersion).Value
                tmpRow = freeLine
            End If
        End If
            
        freeLine = freeLine + 1
    Loop
    
    getLastVersionRow = tmpRow
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Récupération de l'identifiant correspondant au langage pour highlight.js
'Voir : http://highlightjs.readthedocs.io/en/latest/css-classes-reference.html
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function getIDLanguage(ByVal sLanguageName As String) As String
    Dim i As Integer: i = 2
    While optionsSheet.Cells(i, colOptLanguages).Value <> sLanguageName And optionsSheet.Cells(i, colOptLanguages).Value <> ""
        i = i + 1
    Wend
    getIDLanguage = optionsSheet.Cells(i, colOptIdLanguages).Value
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Vérifie si un ID de fiche existe déjà ou non dans le catalogue Excel
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function isUniqueID(ByVal dblId As Double) As Boolean
    isUniqueID = True
    Dim i As Integer: i = 3
    Do While catalogueSheet.Cells(i, colId).Value <> ""
        If catalogueSheet.Cells(i, colId).Value = dblId Then
            isUniqueID = False
            Exit Do
        End If
        i = i + 1
    Loop
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Met à jour la totalité des données du classeur Excel actif : tableaux croisés dynamiques, graphiques
'Ne fonctionne pas si le partage multi-utilisateur est activé
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub updateWorkbookData()
    'Si le classeur est partagé
    If ActiveWorkbook.MultiUserEditing Then
        MsgBox "Le classeur est partagé, les statistiques ne peuvent être actualisés. Retirez le du partage, puis réessayez. Enfin, n'oubliez pas de le partager à nouveau puis de sauvegarder.", vbInformation, "Mise à jour impossible"
    Else
        Application.ScreenUpdating = False
        'Si la feuille "STATISTIQUES" est masquée, on l'affiche
        If statsSheet.Visible = xlSheetVeryHidden Then
            statsSheet.Visible = xlSheetVisible
        End If
        ActiveWorkbook.RefreshAll
        Application.ScreenUpdating = True
        'MsgBox "Mise à jour effectuée avec succès !", vbInformation, "Classeur à jour"
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Reconstruit une copie du catalogue HTML à partir du catalogue Excel
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub rebuildHtmlCatalogue()

    Application.StatusBar = "Reconstruction du catalogue HTML en cours..."

    'Verrouillage d'Excel
    Application.ScreenUpdating = False
    'Récupération de la feuille active
    Dim oldSheet As Worksheet
    Set oldSheet = Application.ActiveSheet
    'Activation de la feuille "CATALOGUE"
    catalogueSheet.Activate
    
    Dim strCatalogueContent As String
    strCatalogueContent = ""
    
    'Parcourt des fiches de la feuille "CATALOGUE"
    'Définition des variables publiques de chaque fiche
    Dim i As Integer: i = 3
    Do While catalogueSheet.Cells(i, colId).Value <> ""
        'Titre de la fiche
        strTitle = catalogueSheet.Cells(i, colTitle).Value
        'ID de la fiche
        dblId = catalogueSheet.Cells(i, colId).Value
        'Logiciel de la fiche
        strSoftware = catalogueSheet.Cells(i, colSoftware).Value
        'Langage de la fiche
        strLanguage = catalogueSheet.Cells(i, colLanguage).Value
        'Version de la fiche
        intVersion = catalogueSheet.Cells(i, colVersion).Value
        'Statut de la fiche
        strStatus = catalogueSheet.Cells(i, colStatus).Value
        'Type de la fiche
        strType = catalogueSheet.Cells(i, colType).Value
        'Créé le
        Dim strAddedDate As String
        strAddedDate = VBA.Left(catalogueSheet.Cells(i, colDatec).Value, 10)
        'Modifié le
        Dim strModifiedDate As String
        strModifiedDate = VBA.Left(catalogueSheet.Cells(i, colDatem).Value, 10)
        'Utilisateur
        Dim strUser As String
        strUser = catalogueSheet.Cells(i, colUser).Value
        
        'Mots clés de la fiche
        Dim m As Integer: m = 0
        While catalogueSheet.Cells(i, colKeywords + m).Value <> ""
            If m > 0 Then
                strKeywords = strKeywords & " " & catalogueSheet.Cells(i, colKeywords + m).Value
            Else
                strKeywords = catalogueSheet.Cells(i, colKeywords + m).Value
            End If
            m = m + 1
        Wend
        
        Dim tmpKeywords As Variant
        Dim htmlKeywords As String
        tmpKeywords = Split(strKeywords, " ")
        htmlKeywords = ""
        Dim k As Integer
        For k = 0 To UBound(tmpKeywords)
            htmlKeywords = htmlKeywords & "<span>" & tmpKeywords(k) & "</span>"
        Next
        
        'Informations de la fiche au format HTML à partir de ses variables publiques
        strCatalogueContent = strCatalogueContent & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "<tr class='" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "'>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & dblId & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td><a href='.." & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html'>" & strTitle & "</a></td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & intVersion & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strStatus & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strSoftware & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strLanguage & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strType & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strAddedDate & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strModifiedDate & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & strUser & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<td>" & htmlKeywords & "</td>" & _
            vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "</tr>"
        
        i = i + 1
    Loop
    
    Dim strIn 'As String
    'Ouverture du template du catalogue HTML
    strIn = getFileContent(cataloguePath & htmlCataloguePath & htmlCatalogueTemplateName)
    'Remplissage du template du catalogue HTML
    strIn = Replace(strIn, "{TEMPLATE_CATALOGUE_CONTENT}", strCatalogueContent)
    'Copie de l'ancien catalogue HTML avec la date du jour
    Call copyFileFromTo(cataloguePath & htmlCataloguePath & htmlCatalogueName, cataloguePath & htmlCataloguePath & "CATALOGUE" & "_" & Replace(Date, "/", "-") & ".html", True)
    'Enregistrement du nouveau catalogue HTML
    Call writeToFile(cataloguePath & htmlCataloguePath & htmlCatalogueName, strIn)
    
    'Ouverture du nouveau catalogue HTML dans le navigateur
    browseURL (cataloguePath & htmlCataloguePath & htmlCatalogueName)
    
    'Vidage des variables publiques
    Call resetPublicVar

    'Réactivation de l'ancienne feuille active
    oldSheet.Activate
    'Déverrouillage d'Excel
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Exporte une copie du catalogue pour la version 2 de l'application : REXit
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub exportCatalogue()

    Application.StatusBar = "Export du catalogue en cours..."

    'Verrouillage d'Excel
    Application.ScreenUpdating = False
    'Récupération de la feuille active
    Dim oldSheet As Worksheet
    Set oldSheet = Application.ActiveSheet
    'Activation de la feuille "CATALOGUE"
    catalogueSheet.Activate
    
    Dim sExportDirectoryPath As String
    sExportDirectoryPath = Environ("USERPROFILE") & "\Desktop" & "\CATALOGUE"
    
    If folderExists(sExportDirectoryPath) Then
        deleteFolder (sExportDirectoryPath)
    End If
    
    Call createFolder(sExportDirectoryPath)
    
    'Parcourt des fiches de la feuille "CATALOGUE"
    'Définition des variables publiques de chaque fiche
    Dim i As Integer: i = 3
    Do While catalogueSheet.Cells(i, colId).Value <> ""
        'ID de la fiche
        dblId = catalogueSheet.Cells(i, colId).Value
        'Version de la fiche
        intVersion = catalogueSheet.Cells(i, colVersion).Value
        
        Dim sSheetPath As String
        sSheetPath = sExportDirectoryPath & "\" & CStr(dblId) & "_" & CStr(intVersion)
        Call createFolder(sSheetPath)
        
        'Titre de la fiche
        strTitle = catalogueSheet.Cells(i, colTitle).Value
        Call writeToFile(sSheetPath & "\" & "title.txt", strTitle)
        'Logiciel de la fiche
        strSoftware = catalogueSheet.Cells(i, colSoftware).Value
        Call writeToFile(sSheetPath & "\" & "software.txt", strSoftware)
        'Langage de la fiche
        strLanguage = catalogueSheet.Cells(i, colLanguage).Value
        Call writeToFile(sSheetPath & "\" & "language.txt", strLanguage)
        'Identifiant correspondant au langage pour highlight.js
        strIDLanguage = getIDLanguage(strLanguage)
        Call writeToFile(sSheetPath & "\" & "idLanguage.txt", strIDLanguage)
        'Identifiant correspondant au langage pour highlight.js
        strIDLanguage = getIDLanguage(strLanguage)
        'Statut de la fiche
        strStatus = catalogueSheet.Cells(i, colStatus).Value
        Call writeToFile(sSheetPath & "\" & "status.txt", strStatus)
        'Type de la fiche
        strType = catalogueSheet.Cells(i, colType).Value
        Call writeToFile(sSheetPath & "\" & "type.txt", strType)
        'Créé le
        Dim strAddedDate As String
        strAddedDate = VBA.Left(catalogueSheet.Cells(i, colDatec).Value, 10)
        Call writeToFile(sSheetPath & "\" & "creationDate.txt", strAddedDate)
        'Modifié le
        Dim strModifiedDate As String
        strModifiedDate = VBA.Left(catalogueSheet.Cells(i, colDatem).Value, 10)
        Call writeToFile(sSheetPath & "\" & "modificationDate.txt", strModifiedDate)
        'Utilisateur
        Dim strUser As String
        strUser = catalogueSheet.Cells(i, colUser).Value
        Call writeToFile(sSheetPath & "\" & "user.txt", strUser)
        
        'Mots clés de la fiche
        Dim m As Integer: m = 0
        While catalogueSheet.Cells(i, colKeywords + m).Value <> ""
            If m > 0 Then
                strKeywords = strKeywords & "," & catalogueSheet.Cells(i, colKeywords + m).Value
            Else
                strKeywords = catalogueSheet.Cells(i, colKeywords + m).Value
            End If
            m = m + 1
        Wend
        
        Call writeToFile(sSheetPath & "\" & "tags.txt", strKeywords)
        
        'Code de la fiche
        strCode = GetBetween(getFileContent(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html"), "<code class=""" & strIDLanguage & """>" & vbNewLine, vbNewLine & vbTab & vbTab & vbTab & "</code>")
        Call writeToFile(sSheetPath & "\" & "code.txt", specialCharsHtml(strCode))
        'Problème de la fiche
        strProblem = GetBetween(getFileContent(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html"), "<h1>Problème</h1>" & vbNewLine & vbTab & vbTab & "<span>" & vbNewLine, vbNewLine & vbTab & vbTab & "</span>")
        Call writeToFile(sSheetPath & "\" & "problem.txt", strProblem)
        'Solution de la fiche
        strSolution = GetBetween(getFileContent(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html"), "<h1>Solution</h1>" & vbNewLine & vbTab & vbTab & "<span>" & vbNewLine, vbNewLine & vbTab & vbTab & "</span>")
        Call writeToFile(sSheetPath & "\" & "solution.txt", strSolution)
        'Pièces jointes
        Call copyFolderFromTo(cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "\", sSheetPath & "\files\")
        
        i = i + 1
    Loop
    
    'Vidage des variables publiques
    Call resetPublicVar

    'Réactivation de l'ancienne feuille active
    oldSheet.Activate
    'Déverrouillage d'Excel
    Application.ScreenUpdating = True
    
    Application.StatusBar = ""
    
    MsgBox "CATALOGUE exporté avec succès :" & vbNewLine & vbNewLine & sExportDirectoryPath, vbInformation, "Export du CATALOGUE"
    
End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions liées aux différentes entrées ajoutées au menu contextuel du clic droit d'Excel dans la feuille "CATALOGUE"
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Envoi de l'url d'une fiche par email via Microsoft Outlook : l'utilisateur a cliqué sur "Envoyer la fiche par e-mail"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicSendMail()
    'Si la fiche existe
    If fileExists(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
        'Export du contenu de la fiche pour la mettre en pièce jointe si souhaité
        Dim sArchivePath As String
        If MsgBox("Inclure la totalité des fichiers de la fiche en pièce jointe ?" & vbNewLine & "Une archive compressée sera générée au préalable.", vbQuestion + vbYesNo, "Export de la fiche") = vbYes Then
             'Chemin de l'archive exportée de la fiche
             sArchivePath = rightClicExportSheet(True)
        End If
        Application.StatusBar = "Génération de l'e-mail en cours..."
        Call sendMail(Array(), _
                      "<body style='font-size:11pt;font-family:Calibri'>" & "<a href='" & path2UNC(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") & "'>" & path2UNC(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") & "</a>" & "</body>", _
                      catalogueSheet.Cells(ActiveCell.Row, colTitle).Value, , sArchivePath, 2)
        Application.StatusBar = ""
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Copie du code d'une fiche dans le presse papier : l'utilisateur a cliqué sur "Copier le contenu de la fiche"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicCopyToClipboard()
    'Si la fiche existe
    If fileExists(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
        Call CopyToClipboard(specialCharsHtml(GetBetween(getFileContent(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html"), "<code class=""" & getIDLanguage(catalogueSheet.Cells(ActiveCell.Row, colLanguage).Value) & """>" & vbNewLine, vbNewLine & vbTab & vbTab & vbTab & "</code>")))
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouverture de fiche dans le navigateur par défaut du système : l'utilisateur a cliqué sur "Ouvrir la fiche"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicOpenSheet()
    'Si la fiche existe
    If fileExists(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
        Application.StatusBar = "Ouverture de fiche en cours..."
        Call browseURL(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html")
        Application.StatusBar = ""
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouverture du dossier de pièces jointes dans l'Explorateur Windows : l'utilisateur a cliqué sur "Afficher les pièces jointes"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicOpenFiles()
    'Si la fiche existe
    If fileExists(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
        Dim sFolderPath As String
        sFolderPath = cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value)
        If folderExists(sFolderPath) Then
            Application.StatusBar = "Ouverture du dossier de pièces jointes en cours..."
            Shell Environ("WINDIR") & "\explorer.exe " & sFolderPath, vbNormalFocus
            Application.StatusBar = ""
        End If
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Export de fiche au format archive ZIP : l'utilisateur a cliqué sur "Exporter la fiche"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicExportSheet(Optional ByVal bIsMailAttached As Boolean = False) As String

    'Nom de la fiche
    Dim sSheetName As String
    sSheetName = replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value)
    'Nom de fichier de la fiche
    Dim sSheetFileName As String
    sSheetFileName = sSheetName & ".html"
    'Chemin de la fiche
    Dim sSheetPath As String
    sSheetPath = Application.ThisWorkbook.Path & sheetsPath & sSheetFileName
    'Si la fiche existe
    If fileExists(sSheetPath) Then
    
        'Vérification si le logiciel de compression de données 7-Zip est installé
        Dim bIs7zipInstalled As Boolean
        If get7zipExePath <> "" Then
            bIs7zipInstalled = True
        Else
            bIs7zipInstalled = False
        End If
    
        'Emplacement de l'export
        Dim sZipFilePath As Variant
        'Si le logiciel de compression de données 7-Zip n'est pas installé, format .zip
        If Not bIs7zipInstalled Then
            sZipFilePath = Application.GetSaveAsFilename(FileFilter:="Archive ZIP (*.zip), *.zip", Title:="Sélectionnez un fichier d'export", InitialFileName:=sSheetName)
        'Sinon, le logiciel de compression de données 7-Zip est installé, format .7z
        Else
            sZipFilePath = Application.GetSaveAsFilename(FileFilter:="Archive 7z (*.7z), *.7z", Title:="Sélectionnez un fichier d'export", InitialFileName:=sSheetName)
        End If
        If sZipFilePath <> False Then
            'Vérification avant écrasement
            If fileExists(sZipFilePath) Then
                If MsgBox("Le fichier suivant existe déjà :" & vbNewLine & vbNewLine & getFilenameFromPath(sZipFilePath) & vbNewLine & vbNewLine & "Souhaitez vous l'écraser ?", vbQuestion + vbYesNo, "Confirmation d'écrasement") = vbNo Then
                    Exit Function
                Else
                    Call deleteFile(sZipFilePath)
                End If
            End If
        Else
            MsgBox "Aucun fichier sélectionné, arrêt de l'export.", vbInformation, "Export annulé"
            Exit Function
        End If
    
        'Chemin du dossier de pièces jointes de la fiche
        Dim sFolderPath As String
        sFolderPath = cataloguePath & sheetsPath & filesPath & "\" & sSheetName
        'Chemin du dossier temporaire où seront copiés les fichiers de la fiche
        Dim sTempFolder As String
        sTempFolder = Environ("temp") & "\" & sSheetName & "\"
        
        Application.StatusBar = "Extraction de la fiche en cours..."

        'Création du dossier temporaire où seront copiés les fichiers de la fiche
        Call createFolder(sTempFolder)
        'Copie du fichier HTML de la fiche
        Call copyFileFromTo(sSheetPath, sTempFolder & sSheetFileName, True)
        'Copie du fichier favicon de la fiche
        Call copyFileFromTo(Application.ThisWorkbook.Path & sheetsPath & "favicon.png", sTempFolder & "favicon.png", True)
        'Copie du dossier de pièces jointes de la fiche :
        '...s'il existe
        If folderExists(sFolderPath) Then
            '...et qu'il n'est pas vide
            If Not isEmptyDirectory(sFolderPath) Then
                Call createFolder(sTempFolder & filesPath)
                Call copyFolderFromTo(sFolderPath, sTempFolder & filesPath & "\" & sSheetName)
            End If
        End If
        'Copie des dépendances JavaScript de la fiche
        Call copyFolderFromTo(Application.ThisWorkbook.Path & sheetsPath & "JS", sTempFolder & "JS")
        'Copie des dépendances CSS de la fiche
        Call copyFolderFromTo(Application.ThisWorkbook.Path & sheetsPath & "CSS", sTempFolder & "CSS")
        'Ne pas inclure les dépendances CSS et JavaScript de la fiche d'aide au langage Markdown inutilement
        If Not catalogueSheet.Cells(ActiveCell.Row, colId).Value = "187028" Then
            Call deleteFolder(sTempFolder & "CSS" & "\" & "editor")
            Call deleteFile(sTempFolder & "JS" & "\" & "simplemde.min.js")
        End If

        Application.StatusBar = "Compression de la fiche en cours..."

        'Compression du dossier temporaire où ont été copiés les fichiers de la fiche
        'Si le logiciel de compression de données 7-Zip n'est pas installé, compression zip
        If Not bIs7zipInstalled Then
            Call addFileOrFolderToZipFile(sZipFilePath, sTempFolder)
        'Sinon, le logiciel de compression de données 7-Zip est installé, compression 7z
        Else
            Call addFolderTo7ZipFile(sZipFilePath, sTempFolder)
        End If
        
        'Suppression du dossier temporaire où ont été copiés les fichiers de la fiche
        Call deleteFolder(sTempFolder)

        Application.StatusBar = ""
        
        'Fin de l'opération, ouverture du dossier d'export ?
        'Si on n'a pas généré l'archive en vue de la mettre dans un mail, on propose l'ouverture du dossier d'export
        If Not bIsMailAttached Then
            If MsgBox("Fiche exportée avec succès !" & vbNewLine & "Ouvrir le dossier d'export maintenant ?", vbQuestion + vbYesNo, "Export terminé") = vbYes Then
                Call OpenIt(getDirectoryFromPath(sZipFilePath))
            End If
        'Sinon, on renvoie l'emplacement de l'archive générée pour la mettre en pièce jointe d'un mail
        Else
            rightClicExportSheet = sZipFilePath
        End If
        
    End If
    
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Edition de fiche : l'utilisateur a cliqué sur "Editer la fiche"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicEditSheet()

    'Si la fiche existe
    If fileExists(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
        
        Application.StatusBar = "Edition de fiche en cours..."
        
        'Récupération de la dernière version de la fiche
        Dim rowToEdit As Integer
        rowToEdit = getLastVersionRow(catalogueSheet.Cells(ActiveCell.Row, colId).Value)
    
        'Remplissage des variables publiques pour le UserForm
        'Note : en cas de création d'une fiche, les variables publiques sont vides
        
        'Titre de la fiche
        strTitle = catalogueSheet.Cells(rowToEdit, colTitle).Value
        
        'ID de la fiche
        dblId = catalogueSheet.Cells(rowToEdit, colId).Value
        
        'Mots clés de la fiche
        Dim m As Integer: m = 0
        While catalogueSheet.Cells(rowToEdit, colKeywords + m).Value <> ""
            If m > 0 Then
                strKeywords = strKeywords & " " & catalogueSheet.Cells(rowToEdit, colKeywords + m).Value
            Else
                strKeywords = catalogueSheet.Cells(rowToEdit, colKeywords + m).Value
            End If
            m = m + 1
        Wend
        'Logiciel de la fiche
        strSoftware = catalogueSheet.Cells(rowToEdit, colSoftware).Value
        'Langage de la fiche
        strLanguage = catalogueSheet.Cells(rowToEdit, colLanguage).Value
        'Identifiant correspondant au langage pour highlight.js
        strIDLanguage = getIDLanguage(strLanguage)
        'Version de la fiche
        intVersion = catalogueSheet.Cells(rowToEdit, colVersion).Value
        'Statut de la fiche
        If isObsolete = True Then
            'Suppression d'une fiche Released, passage en Obsolete
            strStatus = "Obsolete"
        Else
            strStatus = catalogueSheet.Cells(rowToEdit, colStatus).Value
        End If
        'Type de la fiche
        strType = catalogueSheet.Cells(rowToEdit, colType).Value
        'Code de la fiche
        strCode = GetBetween(getFileContent(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html"), "<code class=""" & strIDLanguage & """>" & vbNewLine, vbNewLine & vbTab & vbTab & vbTab & "</code>")
        'Problème de la fiche
        strProblem = GetBetween(getFileContent(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html"), "<h1>Problème</h1>" & vbNewLine & vbTab & vbTab & "<span>" & vbNewLine, vbNewLine & vbTab & vbTab & "</span>")
        'Solution de la fiche
        strSolution = GetBetween(getFileContent(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html"), "<h1>Solution</h1>" & vbNewLine & vbTab & vbTab & "<span>" & vbNewLine, vbNewLine & vbTab & vbTab & "</span>")
        'Pièces jointes
        Call listFilesFromFolder(cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & "\", strFile)
        'Mise à jour de la fiche
        If isObsolete = True Then
            'Suppression d'une fiche Released, passage en Obsolete : on ne raffiche pas le UserForm d'édition de fiche
            Call addToCatalogue(False)
        Else
            Call addToCatalogue(True)
        End If
        
        Application.StatusBar = ""
        
    End If
    
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Suppression de fiche : l'utilisateur a cliqué sur "Supprimer la fiche"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicDeleteSheet()
    
    'Si la fiche existe
    If fileExists(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
        If MsgBox("Souhaitez vous vraiment supprimer la fiche suivante :" & vbNewLine & vbNewLine & """" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & """", vbYesNo + vbQuestion, "Confirmation de suppression") = vbYes Then
            
            Application.StatusBar = "Suppression de fiche en cours..."
            
            'Si la fiche est l'aide au langage Markdown, on propose à l'utilisateur de ne la retirer que des catalogues Excel et HTML
            'Les différents fichiers de la fiche ne seront pas supprimés, elle restera consultable depuis l'interface de création / édition de fiche
            If catalogueSheet.Cells(ActiveCell.Row, colId).Value = "187028" Then
                If MsgBox("L'aide au langage Markdown ne peut être supprimée. Souhaitez vous malgré tout la retirer des catalogues Excel et HTML ? La fiche restera accessible depuis l'interface de création / édition de fiche.", vbYesNo + vbQuestion, "Suppression du fichier d'aide") = vbYes Then
                    GoTo deleteSheetFromCatalogs
                Else
                    Exit Function
                End If
            End If
            
            'Si la fiche n'est pas en draft, on la supprime. Sinon on la passe obsolète et on barre la ligne
            If catalogueSheet.Cells(ActiveCell.Row, colStatus).Value <> "Draft" Then
                isObsolete = True
                Call rightClicEditSheet
            Else
                'Suppression des pièces jointes de la fiche
                Call deleteFolder(cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value))
                
                'Si l'opération de suppression s'est mal déroulée
                If folderExists(cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value)) Then
                    MsgBox "Le dossier de pièces jointes n'a pas pu être supprimé intégralement, un ou plusieurs de ses fichiers sont en cours d'utilisation par un autre programme, ou en cours d'exécution." & vbNewLine & vbNewLine & "Veuillez fermer le dossier s'il est ouvert dans l'Explorateur Windows, ainsi que chacune de ses pièces jointes. Puis réessayez.", vbExclamation, "Fichier(s) en cours d'utilisation"
                    Exit Function
                End If
                
                'Suppression du fichier HTML de la fiche
                Call deleteFile(cataloguePath & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html")
                
                'Si l'opération de suppression s'est mal déroulée
                If fileExists(cataloguePath & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
                    MsgBox "La fiche n'a pu être supprimée." & vbNewLine & "Veuillez la fermer avant de réessayer.", vbExclamation, "Fiche en cours d'utilisation"
                    Exit Function
                End If
                
deleteSheetFromCatalogs:
                
                'Suppression de la fiche dans le catalogue HTML
                Dim strIn 'As String
                strIn = getFileContent(cataloguePath & htmlCataloguePath & htmlCatalogueName)
                Dim strDeleted As String
                strDeleted = GetBetween(strIn, "<tr class='" & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & "'>", "</tr>")
                strDeleted = vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "<tr class='" & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & "'>" & _
                    strDeleted & _
                    "</tr>"
                strIn = Replace(strIn, strDeleted, "")
                Call writeToFile(cataloguePath & htmlCataloguePath & htmlCatalogueName, strIn)
                
                'Suppression de la fiche dans le catalogue Excel
                Rows(ActiveCell.Row).EntireRow.Delete
            End If
                'Sauvegarde du classeur
            ThisWorkbook.Save
            
            Application.StatusBar = ""
            
        End If
    End If
    
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Copie l'URL de la fiche dans le presse papier : l'utilisateur a cliqué sur "Copier l'URL de la fiche"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function rightClicCopyUrlToClipboard()
    'Si la fiche existe
    If fileExists(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html") Then
        Call CopyToClipboard(path2UNC(Application.ThisWorkbook.Path & sheetsPath & replaceIllegalChar(catalogueSheet.Cells(ActiveCell.Row, colId).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colTitle).Value & "_" & catalogueSheet.Cells(ActiveCell.Row, colVersion).Value) & ".html"))
    End If
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions liées au paramétrage du catalogue Excel depuis la feuille "OPTIONS"
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Initialisation des ComboBox logiciels, langages, et types
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function initializeComboBox()
    'Remplissage de la ComboBox de logiciels
    ufCATALOGUE.cbSoftware.Clear
    Dim i As Integer: i = 2
    While optionsSheet.Cells(i, colOptSoftwares).Value <> ""
        ufCATALOGUE.cbSoftware.AddItem optionsSheet.Cells(i, colOptSoftwares).Value
        i = i + 1
    Wend
    
    'Remplissage de la ComboBox de langages
    ufCATALOGUE.cbLanguage.Clear
    i = 2
    While optionsSheet.Cells(i, colOptLanguages).Value <> ""
        ufCATALOGUE.cbLanguage.AddItem optionsSheet.Cells(i, colOptLanguages).Value
        i = i + 1
    Wend
    
    'Remplissage de la ComboBox de types
    ufCATALOGUE.cbType.Clear
    i = 2
    While optionsSheet.Cells(i, colOptTypes).Value <> ""
        ufCATALOGUE.cbType.AddItem optionsSheet.Cells(i, colOptTypes).Value
        i = i + 1
    Wend
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions liées à l'ajout d'entrées dans le menu contextuel du clic droit d'Excel dans la feuille "CATALOGUE"
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ajout de 8 entrées dans le menu contextuel du clic droit :
'   - Ouvrir la fiche
'   - Editer la fiche
'   - Exporter la fiche"
'   - Supprimer la fiche
'   - Copier l'URL de la fiche
'   - Afficher les pièces jointes
'   - Envoyer la fiche par e-mail
'   - Copier le contenu de la fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function addCommandBar(Optional ByVal bSheetOptions As Boolean = True)

    On Error Resume Next

    Dim cBar As CommandBar
    Set cBar = Application.CommandBars("Cell")
    
    'Reset du menu pour éviter les doublons
    cBar.Reset
    
    'Si clic si une fiche
    If bSheetOptions Then
    
        'Entrée "Voir la fiche"
        Dim cbbOpen As CommandBarControl
        'Définition du menu : ajout d'une entrée en 1ière position du menu
        Set cbbOpen = cBar.Controls.Add(temporary:=True, before:=1)
        With cbbOpen
            'Au début du menu contextuel du clic droit
            .BeginGroup = True
            'Nom de l'entrée
            .Caption = "Voir la fiche"
            .Style = msoButtonIconAndCaption
            'Icone de copie
            .FaceId = 2937
            'Exécute rightClicCopyToClipboard()
            .OnAction = "rightClicOpenSheet"
        End With
        
        'Entrée "Editer la fiche"
        Dim cbbEdit As CommandBarControl
        'Définition du menu : ajout d'une entrée en 2ième position du menu
        Set cbbEdit = cBar.Controls.Add(temporary:=True, before:=2)
        With cbbEdit
            'Dans le même groupe que la 1ière entrée ajoutée
            .BeginGroup = False
            'Nom de l'entrée
            .Caption = "Editer la fiche"
            .Style = msoButtonIconAndCaption
            'Icone d'édition
            .FaceId = 162
            'Exécute rightClicEditSheet()
            .OnAction = "rightClicEditSheet"
        End With
        
        'Entrée "Exporter la fiche"
        Dim cbbExport As CommandBarControl
        Set cbbExport = cBar.Controls.Add(temporary:=True, before:=3)
        With cbbExport
            .BeginGroup = False
            .Caption = "Exporter la fiche"
            .Style = msoButtonIconAndCaption
            .FaceId = 2109
            .OnAction = "rightClicExportSheet"
        End With
        
        'Entrée "Supprimer la fiche"
        Dim cbbDelete As CommandBarControl
        Set cbbDelete = cBar.Controls.Add(temporary:=True, before:=4)
        With cbbDelete
            .BeginGroup = False
            .Caption = "Supprimer la fiche"
            .Style = msoButtonIconAndCaption
            .FaceId = 67
            .OnAction = "rightClicDeleteSheet"
        End With
        
        'Entrée "Copier l'URL de la fiche"
        Dim cbbCopy As CommandBarControl
        Set cbbCopy = cBar.Controls.Add(temporary:=True, before:=5)
        With cbbCopy
            .BeginGroup = False
            .Caption = "Copier l'URL de la fiche"
            .Style = msoButtonIconAndCaption
            .FaceId = 2159
            .OnAction = "rightClicCopyUrlToClipboard"
        End With
        
        'Entrée "Afficher les pièces jointes"
        Dim cbbPJ As CommandBarControl
        Set cbbPJ = cBar.Controls.Add(temporary:=True, before:=6)
        With cbbPJ
            .BeginGroup = False
            .Caption = "Afficher les pièces jointes"
            .Style = msoButtonIconAndCaption
            .FaceId = 931
            .OnAction = "rightClicOpenFiles"
        End With
        
        'Entrée "Copier le code de la fiche"
        Dim cbbClipboard As CommandBarControl
        Set cbbClipboard = cBar.Controls.Add(temporary:=True, before:=7)
        With cbbClipboard
            .BeginGroup = False
            .Caption = "Copier le code de la fiche"
            .Style = msoButtonIconAndCaption
            .FaceId = 19
            .OnAction = "rightClicCopyToClipboard"
        End With
        
        'Prochaine position dans le menu contextuel
        Dim index As Integer
        
        'Si l'application Microsoft Outlook est installée
        If bIsOutlookInstalled Then
        
            'Entrée "Envoyer la fiche par e-mail"
            Dim cbbMail As CommandBarControl
            Set cbbMail = cBar.Controls.Add(temporary:=True, before:=8)
            With cbbMail
                .BeginGroup = False
                .Caption = "Envoyer la fiche par e-mail"
                .Style = msoButtonIconAndCaption
                .FaceId = 1983
                .OnAction = "rightClicSendMail"
            End With
            
            index = 9
        Else
            index = 8
        End If
        
        'Ajout d'un séparateur sous la dernière entrée ajoutée dans le menu
        cBar.Controls(index).BeginGroup = True
    
    End If
    
    'Entrée "Créer une nouvelle fiche"
    Dim cbbNewSheet As CommandBarControl
    Set cbbNewSheet = cBar.Controls.Add(temporary:=True, before:=1)
    With cbbNewSheet
        .BeginGroup = True
        .Caption = "Créer une nouvelle fiche"
        .Style = msoButtonIconAndCaption
        .FaceId = 2646
        .OnAction = "addToCatalogue"
    End With
    
    'Ajout d'un séparateur sous la 1ière entrée ajoutée dans le menu
    cBar.Controls(2).BeginGroup = True
    
    On Error GoTo 0
    
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Destruction du menu contextuel du clic droit
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function deleteCommandBar()
    On Error Resume Next
    
    Dim cBar As CommandBar
    Set cBar = Application.CommandBars("Cell")

    'Destruction du menu contextuel du clic droit
    cBar.Reset
    cBar.Delete
    
    On Error GoTo 0
End Function
