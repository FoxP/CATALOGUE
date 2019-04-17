Attribute VB_Name = "SHAREDCODE"
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

'===============================================================================================================================================
'===============================================================================================================================================
'
'D�claration des fonctions issues des API Win32
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Pr�requis de la fonction browseURL()
'-----------------------------------------------------------------------------------------------------------------------------------------------
#If Win64 Then
    Private Declare PtrSafe Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal Operation As String, _
        ByVal Filename As String, _
        Optional ByVal Parameters As String, _
        Optional ByVal Directory As String, _
        Optional ByVal WindowStyle As Long = vbMinimizedFocus _
        ) As Long
#Else
    Private Declare Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal Operation As String, _
        ByVal Filename As String, _
        Optional ByVal Parameters As String, _
        Optional ByVal Directory As String, _
        Optional ByVal WindowStyle As Long = vbMinimizedFocus _
        ) As Long
#End If

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Pr�requis de la fonction GetEnvironmentVariable()
'-----------------------------------------------------------------------------------------------------------------------------------------------
#If Win64 Then
    Private Declare PtrSafe Function GetEnvVar Lib "kernel32" Alias "GetEnvironmentVariableA" _
        (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
    Private Declare Function GetEnvVar Lib "kernel32" Alias "GetEnvironmentVariableA" _
        (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Pr�requis de la fonction ShellAndWait()
'-----------------------------------------------------------------------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#End If

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouvre une URL donn�e dans le navigateur par d�faut
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub browseURL(ByVal url As String)
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", url)
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'R�cup�ration de la valeur d'une variable d'environnement Windows
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function GetEnvironmentVariable(var As String) As String
    Dim numChars As Long
    GetEnvironmentVariable = String(255, " ")
    numChars = GetEnvVar(var, GetEnvironmentVariable, 255)
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ex�cution d'une commande Shell, et attente de la fin de son ex�cution
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'D�finition des arguments manquants et ex�cution du programme
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg : ID de processus sous Win32
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'D�finition de la variable Exitcode
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es aux fichiers : lecture, �criture, copie, suppression, v�rification d'existence, r�cup�ration du nom, de l'extension, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouvre un fichier avec le programme par d�faut du syst�me
'Exemple : ouvrir un .txt avec Notepad
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function OpenIt(ByVal sPathToFile As String)
   Dim Shex As Object
   Set Shex = CreateObject("Shell.Application")
   Shex.Open (sPathToFile)
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renvoie le contenu d'un fichier dont le chemin et l'encodage (par d�faut, utf-8) sont donn�s
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function getFileContent(ByVal filePath As String, Optional charset As String = "utf-8") As String
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.charset = charset
    objStream.Open
    objStream.LoadFromFile (filePath)
    getFileContent = objStream.ReadText()
    objStream.Close
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ecrit dans un fichier dont le chemin est donn�, dans un encodage donn� (par d�faut, utf-8)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function writeToFile(ByVal filePath As String, ByVal strContent As String, Optional charset As String = "utf-8")
    Dim FileOut As Object
    Set FileOut = CreateObject("ADODB.Stream")
    FileOut.Type = 2
    FileOut.charset = charset
    FileOut.Open
    On Error Resume Next
    FileOut.WriteText strContent
    FileOut.SaveToFile filePath, 2
    FileOut.Close
End Function
 
'-----------------------------------------------------------------------------------------------------------------------------------------------
'V�rifie l'existence d'un fichier � partir de son chemin
'Retourne True si le fichier existe, sinon retourne False
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function fileExists(ByVal filePath As String) As Boolean
    If Not Dir(filePath, vbDirectory) = vbNullString Then
        fileExists = True
    Else
        fileExists = False
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Suppression d'un fichier donn� � partir de son chemin
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub deleteFile(ByVal fileToDeletePath As String)
    If fileExists(fileToDeletePath) Then
        SetAttr fileToDeletePath, vbNormal
        Kill fileToDeletePath
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Copie un fichier d'un emplacement � un autre
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function copyFileFromTo(ByVal sourceFile As String, ByVal destFile As String, ByVal overwrite As Boolean)
    On Error GoTo errorMsg
    Dim FSO As Object
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    Call FSO.CopyFile(sourceFile, destFile, overwrite)
    Exit Function
errorMsg:
    MsgBox "Le fichier " & getFilenameFromPath(sourceFile) & " n'a pas pu �tre sauvegard�.", vbCritical, "Erreur lors de la copie"
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renvoie le nom d'un fichier donn� � partir de son chemin
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function getFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        getFilenameFromPath = getFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renvoie le dossier d'un fichier donn� � partir de son chemin
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function getDirectoryFromPath(ByVal strPath As String) As String
    getDirectoryFromPath = Left(strPath, InStrRev(strPath, "\"))
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renvoie l'extension d'un fichier donn� � partir de son chemin
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function getExtensionFromPath(ByVal strPath As String) As String
    getExtensionFromPath = Split(strPath, ".")(UBound(Split(strPath, ".")))
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Retourne l'adresse compl�te d'un chemin (r�seau ou non) donn�, au format UNC (Universal Naming Convention)
'Exemple :
'    Debug.Print path2UNC("C:\Users\Paul\Downloads\test.xml")
'C:\Users\Paul\Downloads\test.xml
'    Debug.Print path2UNC("Z:\02_REX\02_Fiches_REX\CATALOGUE.xlsm")
'\\silvershare.fr.astrium.corp\PART\A\alten_pmt\Com\02_REX\02_Fiches_REX\CATALOGUE.xlsm
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function path2UNC(ByVal sFullName As String) As String
    Dim sDrive      As String
    Dim i           As Long
    Dim ModDrive1 As String

    Application.Volatile

    sDrive = UCase(Left(sFullName, 2))

    With CreateObject("WScript.Network").EnumNetworkDrives
        For i = 0 To .Count - 1 Step 2
            If .item(i) = sDrive Then
                path2UNC = .item(i + 1) & Mid(sFullName, 3)
                Exit For
            End If
        Next
    End With

    ModDrive1 = Replace(path2UNC, " ", "%20")
    If ModDrive1 = "" Then
        If Not VBA.Left(sFullName, 2) = "\\" Then
            path2UNC = "file:///" & sFullName
        Else
            path2UNC = sFullName
        End If
    Else
        path2UNC = ModDrive1
    End If
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es aux dossiers : copie, cr�ation, suppression, v�rification d'existence, r�cup�ration de son contenu, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Suppression d'un dossier donn� � partir de son chemin
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub deleteFolder(ByVal folderPath As String)
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    'Si le dossier n'existe pas, on quitte
    If FSO.folderExists(folderPath) = False Then
        Exit Sub
    End If
    On Error Resume Next
    'Supprimer tous les fichiers du dossier
    FSO.deleteFile folderPath & "\*.*", True
    'Supprimer les sous-dossiers du dossier
    FSO.deleteFolder folderPath & "\*.*", True
    'Supprimer le dossier
    FSO.deleteFolder folderPath
    On Error GoTo 0
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Copie un dossier d'un emplacement � un autre
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub copyFolderFromTo(ByVal sSourceFolderPath As String, ByVal sDestFolderPath As String)
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")

    If Right(sSourceFolderPath, 1) = "\" Then
        sSourceFolderPath = Left(sSourceFolderPath, Len(sSourceFolderPath) - 1)
    End If

    If Right(sDestFolderPath, 1) = "\" Then
        sDestFolderPath = Left(sDestFolderPath, Len(sDestFolderPath) - 1)
    End If

    'Si le dossier n'existe pas, on quitte
    If FSO.folderExists(sSourceFolderPath) = False Then
        Exit Sub
    End If

    On Error Resume Next
    FSO.CopyFolder Source:=sSourceFolderPath, Destination:=sDestFolderPath
    On Error GoTo 0
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Liste tous les fichiers d'un dossier donn� dans une Collection donn�e
'ATTENTION : N�cessite "Microsoft Scripting Runtime" dans les r�f�rences
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub listFilesFromFolder(ByVal strPath As String, ByRef col As Collection)
    Dim objFSO As FileSystemObject
    Dim objFolder 'As Folder
    Dim objFile As File
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.folderExists(strPath) = False Then
        Exit Sub
    End If
    Set objFolder = objFSO.GetFolder(strPath)
    If objFolder.Files.Count = 0 Then
        Exit Sub
    End If
    For Each objFile In objFolder.Files
        If objFile.Name <> "Thumbs.db" Then
            col.Add strPath & objFile.Name, objFile.Name
        End If
    Next objFile
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'V�rifie l'existence d'un dossier � partir de son chemin
'Retourne True si le dossier existe, sinon retourne False
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function folderExists(ByVal folderPath As String) As Boolean
 Dim lngAttr As Long
     On Error GoTo noFolder
     lngAttr = GetAttr(folderPath)
     If (lngAttr And vbDirectory) = vbDirectory Then
         folderExists = True
     End If
noFolder:
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Cr�e un dossier � un emplacement donn�
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function createFolder(ByVal folderPath As String)
    Dim FSO As Object
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    If Len(folderPath) > 0 Then
        If FSO.folderExists(folderPath) = False Then
            FSO.createFolder (folderPath)
        End If
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'V�rifie si un dossier donn� est vide ou non
'Un dossier ne contenant que des fichiers cach�s est consid�r� vide
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function isEmptyDirectory(ByVal folderPath As String) As Boolean
    isEmptyDirectory = (Dir(folderPath & "\*.*") = "")
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es � la r�cup�ration du contenu ainsi qu'� la modification du presse papier
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Copie une chaine de caract�res donn�e dans le presse papier
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub CopyToClipboard(ByVal strToCopy As String)
    
    'Fichier texte temporaire contenant la chaine de texte � copier dans le presse papier
    'Impossible de copier une chaine de texte de plusieurs lignes sans utiliser un fichier texte
    Dim pathToTmpTxtFile As String
    pathToTmpTxtFile = Environ("Temp") & "\tmpTxtFile.txt"

    'Seul l'encodage UTF-16LE est support� par la commande "clip" de Windows
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2 'Type du stream : Text/String
    fsT.charset = "utf-16" 'Encodage : UTF-16LE
    fsT.Open
    fsT.WriteText strToCopy
    fsT.SaveToFile pathToTmpTxtFile, 2

    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    'Attend la fin de l'ex�cution pour passer � la suite
    Dim waitOnReturn As Boolean: waitOnReturn = True
    'Invite de commande masqu�e � l'ex�cution
    Dim WindowStyle As Integer: WindowStyle = 0
    
    'Copie du contenu du fichier texte temporaire dans le presse papier
    wsh.Run "cmd.exe /S /C clip < " & """" & pathToTmpTxtFile & """", WindowStyle, waitOnReturn
    
    'Suppression du fichier texte temporaire
    deleteFile (pathToTmpTxtFile)
    
'    'Si la commande "clip" n'est pas accessible, utiliser l'utilitaire CLIP.EXE de Dave Navarro, Jr :
    
'    Dim pathToClipExe As String
'    pathToClipExe = Application.ThisWorkbook.Path & "\LIBRARIES\CLIP.EXE"
'
'    Call writeToFile(pathToTmpTxtFile, strToCopy, "iso-8859-1")
'
'    Dim wsh As Object
'    Set wsh = VBA.CreateObject("WScript.Shell")
'    Dim waitOnReturn As Boolean: waitOnReturn = True
'    Dim windowStyle As Integer: windowStyle = 0
'
'    wsh.Run "cmd.exe /S /C " & pathToClipExe & " " & """" & pathToTmpTxtFile & """", windowStyle, waitOnReturn

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'R�cup�re le contenu du presse papier
'D�pendance : Microsoft Forms 2.0 Object Library
'Attention, renseigner la variable pathToClipExe si l'utilitaire CLIP.EXE est utilis�
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function getClipboardContent() As String

    On Error Resume Next
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.GetFromClipboard
    getClipboardContent = clipboard.GetText
    On Error GoTo 0
    
'    'Si la r�f�rence "Microsoft Forms 2.0 Object Library" n'est pas accessible, utiliser l'utilitaire CLIP.EXE de Dave Navarro, Jr :
'
'    'Emplacement de l'utilitaire CLIP.EXE de Dave Navarro, Jr (dave@basicguru.com)
'    'La commande "clip" de Windows ne permet pas de lire le contenu du presse papier
'    Dim pathToClipExe As String
'    pathToClipExe = Application.ThisWorkbook.Path & "\LIBRARIES\CLIP.EXE"
'
'    'Emplacement du fichier texte temporaire contenant la chaine de texte issue du presse papier
'    'Impossible de r�cup�rer une chaine de texte de plusieurs lignes � partir du presse papier sans utiliser un fichier texte
'    Dim pathToTmpTxtFile As String
'    pathToTmpTxtFile = Environ("Temp") & "\tmpTxtFile.txt"
'
'    Dim wsh As Object
'    Set wsh = VBA.CreateObject("WScript.Shell")
'    'Attend la fin de l'ex�cution pour passer � la suite
'    Dim waitOnReturn As Boolean: waitOnReturn = True
'    'Invite de commande masqu�e � l'ex�cution
'    Dim WindowStyle As Integer: WindowStyle = 0
'
'    'Copie du contenu du presse papier dans le fichier texte temporaire
'    wsh.Run "cmd.exe /S /C " & pathToClipExe & " " & """" & pathToTmpTxtFile & """" & " /r", WindowStyle, waitOnReturn
'
'    'R�cup�ration du contenu du fichier texte temporaire
'    'L'utilitaire CLIP.EXE ne lit/�crit qu'au format iso-8859-1
'    If fileExists(pathToTmpTxtFile) Then
'        getClipboardContent = getFileContent(pathToTmpTxtFile, "iso-8859-1")
'    End If
'
'    'Suppression du fichier texte temporaire
'    deleteFile (pathToTmpTxtFile)

End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es aux tableaux et aux collections : recherche de valeur, tri, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'V�rifie si un tableau donn� contient une valeur donn�e, et renvoie True si oui
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function isInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    On Error GoTo IsInArrayError: 'Tableau vide
        For Each element In arr
            If element = valToBeFound Then
                isInArray = True
                Exit Function
            End If
        Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    isInArray = False
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'D�termine si une cl� donn�e existe dans un object Collection donn�
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function keyExistsInColl(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    keyExistsInColl = (Err.Number = 0)
    Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Trie le contenu d'une collection donn�e par ordre alphab�tique
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub sortCollection(ByRef collectionObject As Collection)
    If IsNull(collectionObject) Then
        Exit Sub
    End If

    Dim item As Variant
    Dim innerItem As Variant

    Dim i As Long
    Dim j As Long
    Dim index As Long

    Dim hasSwapped As Boolean
    Dim collectionCount As Long

    collectionCount = collectionObject.Count

    Do

        hasSwapped = False

        For i = 1 To collectionCount
            index = i

            If IsObject(collectionObject(i)) Then
                Set item = collectionObject(i)
            Else
                item = collectionObject(i)
            End If

            For j = i + 1 To collectionCount
                If IsObject(collectionObject(j)) Then
                    Set innerItem = collectionObject(j)
                Else
                    innerItem = collectionObject(j)
                End If

                If item > innerItem Then
                    collectionObject.Add item, After:=j
                    collectionObject.Remove index
                    index = j
                    hasSwapped = True
                End If
            Next j
        Next i

    Loop Until Not hasSwapped
End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es au texte : suppression de caract�res sp�ciaux, des espaces inutiles, des accents, v�rification de l'orthographe, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'G�n�ration d'un identifiant unique
'Exemple
'    Dim UID As Double
'    UID = 100
'    Do While Len(CStr(UID)) <> 6
'        UID = GetUniqueID
'    Loop
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function GetUniqueID() As Double
    On Error Resume Next
    'Initialise le compteur
    Randomize
    'Petite boucle pour temporiser et d�cor�ler le tick machine et le random
    Dim i As Long
    For i = 1 To 20000
        i = i + 1
    Next
    'Retourne un double compris entre 1 et 999999
    Dim oDate As String
    oDate = Now
    GetUniqueID = Round(Now * Rnd * 10, 0)
    
    On Error GoTo 0
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renvoie la chaine de texte contenue entre deux autres chaines de texte donn�es
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function GetBetween(ByVal sSearch As String, ByVal sStart As String, ByVal sStop As String, Optional ByVal lSearch As Long = 1) As String
    lSearch = InStr(lSearch, sSearch, sStart)
    If lSearch > 0 Then
        lSearch = lSearch + Len(sStart)
        Dim lTemp As Long
        lTemp = InStr(lSearch, sSearch, sStop)
        If lTemp > lSearch Then
            GetBetween = Mid$(sSearch, lSearch, lTemp - lSearch)
        End If
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ne conserve que les caract�res suivants d'une chaine de texte donn�e : [A-Z], [0-9], "_"
'Tr�s restrictif, mais permet d'�viter les noms de fichiers sources d'erreurs tr�s facilement
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function replaceIllegalChar(strIn As String) As String
    Dim j As Integer
    Dim varStr As String, xStr As String
    varStr = deleteAccents(strIn)
    For j = 1 To Len(varStr)
        Select Case Asc(Mid(varStr, j, 1))
            Case 48 To 57, 65 To 90, 97 To 122
            xStr = xStr & Mid(varStr, j, 1)
        Case Else
            xStr = xStr & "_"
        End Select
    Next
    replaceIllegalChar = xStr
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Supprime les espaces en trop d'une chaine de caract�res donn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function delAllSpace(strParamString As String) As String
    Dim strTempString As String
    Dim i As Integer
 
    strTempString = LTrim(strParamString)
    strTempString = RTrim(strTempString)
 
    i = InStr(1, strTempString, "  ")
 
    While i <> 0
        strTempString = Replace(strTempString, "  ", " ")
        i = InStr(1, strTempString, "  ")
        DoEvents
    Wend
 
    delAllSpace = strTempString
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Convertit les caract�res sp�ciaux en entit�s HTML
' & (ET commercial)         ->     &amp;
' " (double guillement)     ->     &quot;
' ' (simple guillemet)      ->     &#039;
' < (inf�rieur �)           ->     &lt;
' > (sup�rieur �)           ->     &gt;
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function htmlSpecialChars(ByVal strCode As String) As String
    strCode = Replace(strCode, "<", "&lt;")
    strCode = Replace(strCode, ">", "&gt;")
    strCode = Replace(strCode, "'", "&#039;")
    strCode = Replace(strCode, """", "&quot;")
'    strCode = Replace(strCode, "&", "&amp;")
    htmlSpecialChars = strCode
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Convertit entit�s HTML en caract�res sp�ciaux
' & (ET commercial)         <-     &amp;
' " (double guillement)     <-     &quot;
' ' (simple guillemet)      <-     &#039;
' < (inf�rieur �)           <-     &lt;
' > (sup�rieur �)           <-     &gt;
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function specialCharsHtml(ByVal strCode As String) As String
    strCode = Replace(strCode, "&lt;", "<")
    strCode = Replace(strCode, "&gt;", ">")
    strCode = Replace(strCode, "&#039;", "'")
    strCode = Replace(strCode, "&quot;", """")
'    strCode = Replace(strCode, "&amp;", "&")
    specialCharsHtml = strCode
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Substitue les caract�res accentu�s d'une chaine de texte donn�e en conservant ou non la casse des caract�res
'Exemple :
'           Debug.Print deleteAccents("������������", True)
'   AAAAAAcDeeee
'           Debug.Print deleteAccents("������������", False)
'   AAAAAACDEEEE
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function deleteAccents(ByVal stringToModify As String, Optional ByVal keepCase As Boolean = True) As String

  Const sWithAccents As String = "���������������������֊����ݟ�"
  Const sWithoutAccents As String = "AAAAAACDEEEEIIIINOOOOOSUUUUYYZ"
  Dim f As String, i As Long
 
   If Not keepCase Then
      stringToModify = UCase$(stringToModify)
      For i = 1 To Len(sWithAccents)
         f = Mid$(sWithAccents, i, 1)
         If InStr(1, stringToModify, f, vbBinaryCompare) > 0 Then
            stringToModify = Replace(stringToModify, f, Mid$(sWithoutAccents, i, 1), , , vbBinaryCompare)
         End If
      Next i
   Else
      For i = 1 To Len(sWithAccents)
         f = Mid$(sWithAccents, i, 1)
         If InStr(1, stringToModify, f, vbBinaryCompare) > 0 Then
            stringToModify = Replace(stringToModify, f, Mid$(sWithoutAccents, i, 1), , , vbBinaryCompare)
         End If
         If InStr(1, stringToModify, LCase$(f), vbBinaryCompare) > 0 Then
            stringToModify = Replace(stringToModify, LCase$(f), LCase$(Mid$(sWithoutAccents, i, 1)), , , vbBinaryCompare)
         End If
      Next i
   End If
 
   deleteAccents = stringToModify
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Supprime les caract�res "(", ")", "[" et "]" d'un nom de fichier donn� pour qu'il soit bien interpr�t� en Markdown si dans un lien ou une image
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function cleanNameForMarkdown(ByVal sFileName As String) As String
    cleanNameForMarkdown = Replace(Replace(Replace(Replace(sFileName, "(", ""), ")", ""), "[", ""), "]", "")
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'V�rifie l'orthographe d'une TextBox donn�e. N�cessite une feuille Excel pour y �crire temporairement son contenu
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub checkSpellTextbox(tbToCheck As Control, xlSheet As Worksheet)
    If delAllSpace(tbToCheck.Value) <> "" Then
        With xlSheet
            Application.EnableEvents = False
            With .Range("IV1")
                .Value = tbToCheck.Text
                .CheckSpelling IgnoreUppercase:=False, AlwaysSuggest:=True, SpellLang:=1036
                tbToCheck = .Text
                .Value = ""
            End With
            Application.EnableEvents = True
        End With
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'V�rifie si une variable donn�e est un nombre ou pas. Si oui, v�rifie ou non si c'est un nombre n�gatif
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function OnlyNumbers(ByVal newValue, Optional ByVal canBeNegative As Boolean) As Boolean
    If IsNumeric(newValue) And newValue <> vbNullString Then
        If Not canBeNegative And newValue < 0 Then
            OnlyNumbers = False
        Else
            OnlyNumbers = True
        End If
    Else
        OnlyNumbers = False
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'D�code une URL �chapp�e / %-encod�e / url-encod�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function URLDecode(sEncodedURL As String) As String

    Dim sTemp As String
    Dim iCurChr As Integer
    
    iCurChr = 1
    
    Do Until iCurChr - 1 = Len(sEncodedURL)
      Select Case Mid(sEncodedURL, iCurChr, 1)
        Case "+"
          sTemp = sTemp & " "
        Case "%"
          sTemp = sTemp & Chr(Val("&h" & _
             Mid(sEncodedURL, iCurChr + 1, 2)))
           iCurChr = iCurChr + 2
        Case Else
          sTemp = sTemp & Mid(sEncodedURL, iCurChr, 1)
      End Select
    
    iCurChr = iCurChr + 1
    Loop
    
    URLDecode = sTemp

End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Echappe / %-encode / url-encode une URL donn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function URLEncode(sURLToEncode As String, Optional UsePlusRatherThanHexForSpace As Boolean = False) As String

    Dim sTemp As String
    Dim iCurChr As Integer
    iCurChr = 1
    Do Until iCurChr - 1 = Len(sURLToEncode)
      Select Case Asc(Mid(sURLToEncode, iCurChr, 1))
        Case 48 To 57, 65 To 90, 97 To 122
          sTemp = sTemp & Mid(sURLToEncode, iCurChr, 1)
        Case 32
          If UsePlusRatherThanHexForSpace = True Then
            sTemp = sTemp & "+"
          Else
            sTemp = sTemp & "%" & Hex(32)
          End If
       Case Else
             sTemp = sTemp & "%" & _
                  Format(Hex(Asc(Mid(sURLToEncode, _
                  iCurChr, 1))), "00")
    End Select
    
      iCurChr = iCurChr + 1
    Loop
    
    URLEncode = sTemp
    
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es aux couleurs : conversion entre diff�rents formats de couleurs, fen�tre de s�lection de couleur, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Retourne une couleur d�cimale depuis une fen�tre de s�lection de couleurs (color picker d'Excel)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function pickNewColor(Optional i_OldColor As Double = xlNone) As Double
    Const BGColor As Long = 13160660  'Couleur de fond de la fen�tre de s�lection de couleur
    Const ColorIndexLast As Long = 32 'Index de la derni�re couleur personnalis�e de la palette
    
    Dim myOrgColor As Double          'Couleur originale de l'index
    Dim myNewColor As Double          'Couleur r�cup�r�e depuis la fen�tre de s�lection de couleur
    'Couleur RGB affich�e dans la fen�tre de s�lection de couleur comme couleur courante
    Dim myRGB_R As Integer
    Dim myRGB_G As Integer
    Dim myRGB_B As Integer
  
    'Sauvegarde de la palette de couleurs originale
    myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
    
    If i_OldColor = xlNone Then
        'Couleur RGB du fond de la fen�tre de s�lection de couleur
        'Pour que la couleur courante semble "vide"
        decimalColor2RGB BGColor, myRGB_R, myRGB_G, myRGB_B
    Else
        'R�cup�re la couleur RGB de i_OldColor
        decimalColor2RGB i_OldColor, myRGB_R, myRGB_G, myRGB_B
    End If
    
    'Affichage de la fen�tre de s�lection de couleur
    If Application.Dialogs(xlDialogEditColor).Show(ColorIndexLast, _
        myRGB_R, myRGB_G, myRGB_B) = True Then
        '"OK" s�lectionn�, Excel a modifi� la palette de couleurs
        'Lecture de la nouvelle couleur dans la palette
        pickNewColor = ActiveWorkbook.Colors(ColorIndexLast)
        'R�initialisation de la palette � sa valeur initiale
        ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
    Else
        '"Annul�" s�lectionn�, la palette de couleurs n'a pas chang�
        'On retourne l'ancienne couleur (ou xlNone si aucune couleur pass�e en param�tre de la fonction)
        pickNewColor = i_OldColor
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Convertit une couleur d�cimale au format RGB
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub decimalColor2RGB(ByVal i_Color As Long, ByRef o_R As Integer, ByRef o_G As Integer, ByRef o_B As Integer)
    o_R = i_Color Mod 256
    i_Color = i_Color \ 256
    o_G = i_Color Mod 256
    i_Color = i_Color \ 256
    o_B = i_Color Mod 256
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Convertit une couleur d�cimale au format Hexadecimal (HEX)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function decimalColor2Hex(ByVal iColor As Long) As String
    Dim sColor As String
    sColor = Right("000000" & Hex(iColor), 6)
    decimalColor2Hex = Right(sColor, 2) & Mid(sColor, 3, 2) & Left(sColor, 2)
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es aux UserForms : tri d'une d'une ListBox, encadrer un texte s�lectionn� dans une TextBox par un pr�fixe / suffixe, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Encadre du texte s�lectionn� dans une TextBox donn�e d'un UserForm par un pr�fixe ainsi qu'un suffixe donn�s
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub addTagToSelectedText(ByRef oTextBox As Control, ByVal sBefore As String, ByVal sAfter As String)
    If Len(oTextBox.SelText) > 0 Then
        Dim lPos As Long
        lPos = oTextBox.SelStart
        Dim lLength As Long
        lLength = oTextBox.SelLength
        oTextBox.SelText = sBefore & oTextBox.SelText & sAfter
        oTextBox.SelStart = lPos + Len(sBefore)
        oTextBox.SelLength = lLength
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Trie les �l�ments d'une ListBox donn�e par ordre alphab�tique
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub sortListBox(listBoxToSort As Control)
    Dim i As Long
    Dim j As Long
    Dim temp As Variant
        
    With listBoxToSort
        For j = 0 To listBoxToSort.ListCount - 2
            For i = 0 To listBoxToSort.ListCount - 2
                If .List(i) > .List(i + 1) Then
                    temp = .List(i)
                    .List(i) = .List(i + 1)
                    .List(i + 1) = temp
                End If
            Next i
        Next j
    End With
End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Autres fonctions sp�cifiques aux logiciels Microsoft Office : Excel, Outlook, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'D�sactive le partage multi-utilisateur dans le classeur Excel actif
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub makeWorkbookExclusive()
    If ActiveWorkbook.MultiUserEditing Then
        Application.DisplayAlerts = False
        ActiveWorkbook.ExclusiveAccess
        Application.DisplayAlerts = True
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Active le partage multi-utilisateur dans le classeur Excel actif
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub makeWorkbookShared()
    If Not ActiveWorkbook.MultiUserEditing Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs ActiveWorkbook.Name, accessmode:=xlShared
        Application.DisplayAlerts = True
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Envoi d'un mail via Microsoft Outlook
'Param�tre bodyFormat : 2 = olFormatHTML (HTML ), 1 = olFormatPlain (Texte brut), ou 3 = olFormatRichText (Texte enrichi)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub sendMail(ByVal mailsArray As Variant, ByVal sMessage As String, ByVal sTopic As String, _
             Optional ByVal bSendToUser As Boolean = False, Optional ByVal sAttachmentPath As String = "", _
             Optional ByVal bodyFormat As Integer = 3, Optional ByVal bSendNow As Boolean = False)

    On Error GoTo noOutlookApp

    'Cr�ation de l'application Outlook et du message

    Dim MonAppliOutlook As Object
    Set MonAppliOutlook = CreateObject("Outlook.Application")
    Dim MonMail As Object
    Set MonMail = MonAppliOutlook.CreateItem(0)
    
    On Error GoTo 0

    'Liste de destinataires

    Dim mailList As String
    mailList = ""
    Dim i As Integer
    For i = 0 To UBound(mailsArray)
        mailList = mailList + mailsArray(i) & ";"
    Next

    'Connexion � Outlook

    MonAppliOutlook.Session.Logon

    'D�finiton du message

    With MonMail
        'Destinataires du mail
        .To = mailList
        'Mail de l'utilisateur Outlook en cours dans le champ CC
        If bSendToUser = True Then
            .CC = MonAppliOutlook.Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        End If
        'Objet du mail
        .Subject = sTopic
        'Format du corps du mail : 2 = olFormatHTML (HTML ), 1 = olFormatPlain (Texte brut), ou 3 = olFormatRichText (Texte enrichi)
        .bodyFormat = bodyFormat
        'Corps du mail
        If bodyFormat = 2 Then
            .HTMLBody = sMessage
        Else
            .Body = sMessage
        End If
        'Pi�ce jointe
        If sAttachmentPath <> "" Then
            Err.Clear
            On Error Resume Next
            
            .Attachments.Add sAttachmentPath
            
            If Err.Number <> 0 Then
                MsgBox "La pi�ce jointe n'a pas pu �tre ajout�e � l'e-mail." & vbNewLine & "Sa taille est la suivante : " & Round(FileLen(sAttachmentPath) / 1024 ^ 2, 1) & " Mo.", vbInformation, "Pi�ce jointe"
            End If
            
            Err.Clear
            On Error GoTo 0
        End If

        'Envoi depuis le compte principal
        .SendUsingAccount = MonAppliOutlook.Session.Accounts.item(1)
        If bSendNow Then
            .Send
        Else
            .Display
        End If
    End With
    
    Set MonMail = Nothing
    Set MonAppliOutlook = Nothing
    
noOutlookApp:

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Affiche la table des caract�res de Windows
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function showCharMap()
    Dim WsShell As Object
    Set WsShell = CreateObject("WScript.Shell")
    WsShell.Run ("charmap")
    Set WsShell = Nothing
End Function

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es � la compression de fichiers au format .zip (archiveur int�gr� � Windows) et .7z (archiveur 7-Zip, compression LZMA2 ultra)
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Cr�ation d'une archive vide au format .zip � un emplacement donn�
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub createZipFile(ByVal sArchivePath As String)
    If Len(Dir(sArchivePath)) > 0 Then Kill sArchivePath
    Open sArchivePath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ajout d'un fichier ou d'un dossier donn� dans une archive au format .zip donn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub addFileOrFolderToZipFile(ByVal sArchivePath As String, ByVal fileOrFolderToAdd As Variant)

    Call createZipFile(sArchivePath)

    Dim sZipFile As String
    sZipFile = sArchivePath
    
    Dim objShell As Object
    Set objShell = CreateObject("Shell.Application")
    
    Dim varZipFile As Variant
    varZipFile = sZipFile
        
    If folderExists(fileOrFolderToAdd) Then
        If Right$(fileOrFolderToAdd, 1) <> "\" Then
            fileOrFolderToAdd = fileOrFolderToAdd & "\"
        End If
        objShell.Namespace(varZipFile).CopyHere objShell.Namespace(fileOrFolderToAdd).Items
        On Error Resume Next
        Do Until objShell.Namespace(varZipFile).Items.Count = objShell.Namespace(fileOrFolderToAdd).Items.Count
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        On Error GoTo 0
    ElseIf fileExists(fileOrFolderToAdd) Then
        objShell.Namespace(varZipFile).CopyHere fileOrFolderToAdd
        On Error Resume Next
        Do Until objShell.Namespace(varZipFile).Items.Count = 1
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        On Error GoTo 0
    End If

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renvoie le dossier d'installation du logiciel de compression de donn�es 7-Zip, si install�. Sinon, renvoie ""
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function get7zipExePath() As String
    If folderExists("C:\Program Files\7-Zip") Then
        get7zipExePath = "C:\Program Files\7-Zip"
    ElseIf folderExists("C:\Program Files (x86)\7-Zip") Then
        get7zipExePath = "C:\Program Files (x86)\7-Zip"
    Else
        get7zipExePath = ""
    End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ajout d'un dossier donn� dans une archive au format .7z donn�e : m�thode de compression LZMA2 ultra
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub addFolderTo7ZipFile(ByVal sArchivePath As String, ByVal sFolderPathToAdd As String)

    'Dossier d'installation du logiciel de compression de donn�es 7-Zip
    Dim s7zipProgramPath As String
    s7zipProgramPath = get7zipExePath()
    If s7zipProgramPath = "" Then
        MsgBox "Dossier d'installation de 7-Zip introuvable.", vbExclamation, "Abandon de l'op�ration"
        Exit Sub
    End If
    If Right(s7zipProgramPath, 1) <> "\" Then
        s7zipProgramPath = s7zipProgramPath & "\"
    End If

    'V�rification de la pr�sence de l'ex�cutable "7z.exe"
    If Dir(s7zipProgramPath & "7z.exe") = "" Then
        MsgBox "7z.exe n'est pas trouvable dans le dossier suivant :" & vbNewLine & vbNewLine & s7zipProgramPath, vbExclamation, "Abandon de l'op�ration"
        Exit Sub
    End If

    'Chemin du dossier � compresser
    Dim sFolderPath As String
    sFolderPath = sFolderPathToAdd
    If Right(sFolderPath, 1) <> "\" Then
        sFolderPath = sFolderPath & "\"
    End If

    'M�thode de compression : ajout, r�cursif, �crasement (a -r -aoa)
    'Compression au format d'archive : 7z (-t7z)
    'M�thode de compression : LZMA2 (-m0=lzma2)
    'Niveau de compression : ultra (-mx=9)
    'Type d'archive : solide (-ms=on)
    Dim sShellCmd As String
    sShellCmd = s7zipProgramPath & "7z.exe -t7z a -r -m0=lzma2 -mx=9 -aoa -ms=on" _
             & " " & Chr(34) & sArchivePath & Chr(34) _
             & " " & Chr(34) & sFolderPath & "*.*" & Chr(34)

    ShellAndWait sShellCmd, vbHide
    
End Sub

