'======================================================================================================================================
'                                                  CATALOGUE Launcher
'
'Fonction :
'           Launcher VB.NET pour le catalogue Excel / HTML de fonctions et bouts de codes couramment utilisés
'
'Versions :
'           - v1.0 : 03-02-2017 : Paul RENARD : Code initial :
'                                                   - Lancement des catalogues HTML et XLS, et création de fiches
'                                                   - Minimisation dans la zone de notifications, menu du clic droit
'                                                   - Notification si nouvelle(s) fiche(s) et modification du catalogue
'                                                   - Gestion des erreurs et des temps de latence des accès aux fichiers
'                                                   - Aperçu du nombre de fiches du catalogue (MAJ toutes les 10 minutes)
'                                                   - Ajout automatique du launcher au démarrage de la session utilisateur
'                                                   - Une seule instance du launcher possible. Si seconde, focus sur l'actuelle
'                                                   - Gestion de paramètres stockés dans un fichier de configuration au format .ini
'           - v1.1 : 17-02-2017 : Paul RENARD : Activation automatique des macros du catalogue au format Excel (fonction expérimentale)
'           - v1.2 : 02-04-2017 : Paul RENARD : Au démarrage de la session utilisateur, le launcher démarre désormais minimisé
'           - v1.3 : 20-05-2017 : Paul RENARD : Ajout de la recherche d'un / plusieurs mots-clés dans le contenu des fiches
'           - v1.4 : 02-07-2017 : Paul RENARD : Lors du clic sur la popup de la zone de notifications, selon le cas :
'                                                   - Ouverture du catalogue HTML si modifié
'                                                   - Ouverture de la / des dernière(s) fiche(s) créée(s)
'
'======================================================================================================================================

'Code partagé
Module SHAREDCODE

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Ouvre une URL donnée dans un navigateur donné ou dans celui par défaut
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub browseURL(ByVal url As String, Optional ByVal browser As String = "default")
        If Not (browser = "default") Then
            Try
                Process.Start(browser, url)
            Catch ex As Exception
                Process.Start(url)
            End Try
        Else
            Process.Start(url)
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Supprime les espaces en trop d'une chaine de caractères donnée
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function delAllSpace(ByVal strParamString As String) As String

        Dim strTempString As String
        Dim i As Integer

        strTempString = LTrim(strParamString)
        strTempString = RTrim(strTempString)

        i = InStr(1, strTempString, "  ")

        While i <> 0
            strTempString = Replace(strTempString, "  ", " ")
            i = InStr(1, strTempString, "  ")
            Application.DoEvents()
        End While

        delAllSpace = strTempString

    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Ouvre un fichier avec le programme par défaut du système
    'Exemple : ouvrir un .txt avec Notepad
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub openFile(ByVal sFilePath As String)
        Dim myProcess As New Process
        myProcess.StartInfo.FileName = sFilePath
        myProcess.StartInfo.UseShellExecute = True
        myProcess.StartInfo.RedirectStandardOutput = False
        myProcess.Start()
        myProcess.Dispose()
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Vérifie l'existance d'un fichier à partir de son chemin
    'Retourne True si le fichier existe, sinon retourne False
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function doesFileExist(ByVal sFilePath As String) As Boolean
        If Not System.IO.File.Exists(sFilePath) Then
            doesFileExist = False
        Else
            doesFileExist = True
        End If
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Renvoie le contenu d'un fichier dont le chemin et l'encodage (par défaut, utf-8) sont donnés
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function getFileContent(ByVal filePath As String, Optional ByVal charset As String = "utf-8") As String
        Dim objStream
        objStream = CreateObject("ADODB.Stream")
        objStream.charset = charset
        objStream.Open()
        objStream.LoadFromFile(filePath)
        getFileContent = objStream.ReadText()
        objStream.Close()
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Crée un raccourci donné à un endroit donné
    'Exemple : Call createShortCut(Application.ExecutablePath, Environment.GetFolderPath(Environment.SpecialFolder.Startup), "RACCOURCI")
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub createShortCut(ByVal TargetPath As String, ByVal ShortCutPath As String, ByVal ShortCutName As String, Optional ByVal Arguments As String = "")
        Dim oShell As Object
        Dim oLink As Object
        Try
            oShell = CreateObject("WScript.Shell")
            oLink = oShell.CreateShortcut(ShortCutPath & "\" & ShortCutName & ".lnk")
            oLink.Arguments = Arguments
            oLink.TargetPath = TargetPath
            oLink.WindowStyle = 1
            oLink.Save()
        Catch ex As Exception
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Renvoie la cible d'un raccourci donné
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function getShortCutTarget(ByVal ShortCutPath As String) As String
        Dim shell = CreateObject("WScript.Shell")
        getShortCutTarget = shell.CreateShortcut(ShortCutPath).TargetPath
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Renvoie les arguments d'un raccourci donné
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function getShortCutArguments(ByVal ShortCutPath As String) As String
        Dim shell = CreateObject("WScript.Shell")
        getShortCutArguments = shell.CreateShortcut(ShortCutPath).Arguments
    End Function

    'Dépendance fonction isWindowOpenedOrNot()
    Public Declare Function FindWindowA Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Détermine si une fenêtre donnée est ouverte
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function isWindowOpenedOrNot(ByVal sWindowsName As String) As Boolean
        Dim windowhandle As IntPtr = FindWindowA(Nothing, sWindowsName)
        If windowhandle = Nothing Then
            isWindowOpenedOrNot = False
        Else
            isWindowOpenedOrNot = True
        End If
    End Function

    'Dépendance fonction isWindowVisibleOrNot()
    Public Declare Function IsWindowVisible Lib "User32" Alias "IsWindowVisible" (ByVal hWnd As IntPtr) As Boolean

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Détermine si une fenêtre donnée est visible
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function isWindowVisibleOrNot(ByVal sWindowsName As String) As Boolean
        Dim windowhandle As IntPtr = FindWindowA(Nothing, sWindowsName)
        If windowhandle = Nothing Then
            isWindowVisibleOrNot = False
        Else
            If IsWindowVisible(windowhandle) = True Then
                isWindowVisibleOrNot = True
            Else
                isWindowVisibleOrNot = False
            End If
        End If
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Retourne le chemin du dossier %APPDATA%
    'Exemple : C:\Users\USERNAME\AppData\Roaming
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function getAppDataPath() As String
        Return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Vérifie l'existance d'un dossier donné à partir de son chemin
    'Retourne True si le dossier existe, sinon retourne False
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function doesDirectoryExists(ByVal sDirectoryPath As String) As Boolean
        If Not System.IO.Directory.Exists(sDirectoryPath) Then
            Return False
        Else
            Return True
        End If
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Crée un dossier d'un nom donné à un emplacement donné
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub createDirectory(ByVal sDirectoryPath As String)
        System.IO.Directory.CreateDirectory(sDirectoryPath)
    End Sub

    'Dépendance fonction leftClick()
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Effectue un clic gauche de la souris
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub leftClick()
        mouse_event(&H2, 0, 0, 0, 0)
        mouse_event(&H4, 0, 0, 0, 0)
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Déplace le curseur de la souris à une position donnée, clique si souhaité, et le replace à sa position d'origine si souhaité
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub moveMouseAndClick(ByVal dXPosition As Double, ByVal dYPosition As Double, ByVal bClick As Boolean, ByVal bGetBackToPreviousPosition As Boolean)
        'Position d'origine du curseur
        Dim dOldXPosition As Double
        dOldXPosition = Windows.Forms.Cursor.Position.X
        Dim dOldYPosition As Double
        dOldYPosition = Windows.Forms.Cursor.Position.Y

        'Déplacement du curseur
        Windows.Forms.Cursor.Position = New Point(dXPosition, dYPosition)

        'Clic gauche
        If bClick Then
            Call leftClick()
        End If

        'Retour du curseur à sa position d'origine
        If bGetBackToPreviousPosition Then
            Windows.Forms.Cursor.Position = New Point(dOldXPosition, dOldYPosition)
        End If
    End Sub

    'Dépendance function getWindowPosition()
    Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    'Dépendance function getWindowPosition()
    Private Declare Function GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Retourne l'ID d'une fenêtre dont une partie du nom est donné
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function findPartialTitle(ByVal partialTitle As String) As IntPtr
        For Each p As Process In Process.GetProcesses()
            If p.MainWindowTitle.IndexOf(partialTitle, 0, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                Return p.MainWindowHandle
            End If
        Next
        Return IntPtr.Zero
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Retourne la position d'une fenêtre dont une partie du nom est donné
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function getWindowPosition(ByVal sWindowPartialName As String) As Double()
        Dim appRect As RECT
        Dim appHandle As Integer
        appHandle = findPartialTitle(sWindowPartialName)
        GetWindowRect(appHandle, appRect)
        Dim positionArray(3) As Double
        positionArray(0) = appRect.Top
        positionArray(1) = appRect.Bottom
        positionArray(2) = appRect.Left
        positionArray(3) = appRect.Right
        Return positionArray
    End Function

    'Dépendance function GetForegroundText()
    Private Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As IntPtr
    Public Declare Auto Function GetWindowText Lib "user32" (ByVal hWnd As System.IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Retourne le nom de la fenêtre au premier plan (c.a.d, le nom de la fenêtre qui a le focus)
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function GetForegroundText() As String
        Dim Caption As New System.Text.StringBuilder(256)
        Dim hWnd As IntPtr = GetForegroundWindow()
        GetWindowText(hWnd, Caption, Caption.Capacity)
        Return Caption.ToString()
    End Function

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Vérifie si un classeur Excel donné (sWorkbookName) est ouvert, si oui, clique sur le bouton "Activer les macros"
    'Nécessite la function activateExcelWorkbookMacros()
    'Exemple : sWorkbookName = "CATALOGUE.xlsm"
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub timerToActivateExcelWorkbookMacros()
        '<IMPORTANT>
        'Déclarer timerWindow dans un module publique (ex : SHAREDVAR) dédié :
        '   - Public timerWindow As New System.Windows.Forms.Timer
        'Lancer l'ouverture du classeur Excel juste avant l'appel de cette fonction :
        '   - Call openFile(sWorkbookPath & "\" & sWorkbookName)
        '<IMPORTANT>

        'Toutes les 1/2 secondes
        timerWindow.Interval = 500
        timerWindow.Start()
        'Si le classeur Excel est ouvert, clique sur le bouton "Activer les macros"
        AddHandler timerWindow.Tick, AddressOf activateExcelWorkbookMacros

        Try
            'Récupération du catalogue Excel
            Dim excelAppHandle As Integer
            excelAppHandle = findPartialTitle("CATALOGUE.xlsm")
            'S'il a été trouvé, c'est qu'il est ouvert
            If excelAppHandle <> 0 Then
                'Récupération du titre de la fenêtre Excel
                Dim Caption As New System.Text.StringBuilder(256)
                GetWindowText(excelAppHandle, Caption, Caption.Capacity)

                'Fenêtre du catalogue Excel au premier plan
                AppActivate(Caption.ToString())
            End If
        Catch ex As Exception
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Vérifie si un classeur Excel donné (sWorkbookName) est ouvert, si oui, clique sur le bouton "Activer les macros"
    'Appelée par la function timerToActivateExcelWorkbookMacros()
    'Exemple : sWorkbookName = "CATALOGUE.xlsm"
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub activateExcelWorkbookMacros()
        'Si le nom de la fenêtre au 1er plan contient sWorkbookName
        If GetForegroundText.Contains(sWorkbookName) Then
            Dim positionArray(3) As Double
            'Récupération de la position de la fenêtre du classeur Excel
            positionArray = getWindowPosition(sWorkbookName)
            'Si la position de la fenêtre est détectée, c'est qu'elle existe
            If Math.Abs(positionArray(0)) + Math.Abs(positionArray(1)) + Math.Abs(positionArray(2)) + Math.Abs(positionArray(3)) <> 0 Then
                'Déplacement du curseur sur le bouton "Activer les macros", clic, puis retour à sa position de départ
                Call moveMouseAndClick(positionArray(2) + 470, positionArray(0) + 165, True, True)
                'Arrêt du timer qui exécute cette fonction
                timerWindow.Stop()
            End If
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Effectue une pause d'un temps donné en millisecondes
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Sub sleep(ByVal iTimeoutInMilliseconds As Integer)
        System.Threading.Thread.Sleep(iTimeoutInMilliseconds)
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    'Retourne la date / heure de compilation d'un EXE ou d'une DLL VB.NET donnée
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Function retrieveLinkerTimestamp(ByVal sFilePath As String) As DateTime

        Const PortableExecutableHeaderOffset As Integer = 60
        Const LinkerTimestampOffset As Integer = 8

        Dim b(2047) As Byte
        Dim s As IO.Stream = Nothing

        Try
            s = New IO.FileStream(sFilePath, IO.FileMode.Open, IO.FileAccess.Read)
            s.Read(b, 0, 2048)
        Finally
            If Not s Is Nothing Then s.Close()
        End Try

        Dim i As Integer = BitConverter.ToInt32(b, PortableExecutableHeaderOffset)
        Dim secondsSince1970 As Integer = BitConverter.ToInt32(b, i + LinkerTimestampOffset)
        Dim dt As New DateTime(1970, 1, 1, 0, 0, 0)

        dt = dt.AddSeconds(secondsSince1970)
        dt = dt.AddHours(TimeZone.CurrentTimeZone.GetUtcOffset(dt).Hours)

        Return dt

    End Function

    'Ajoute l'application au démarrage de Windows
    Sub addExeToStartupFolder()
        'Si le raccourci n'existe pas, on le crée dans le dossier de démarrage de la session
        If Not doesFileExist(Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "\" & "CATALOGUE.lnk") Then
            Call createShortCut(Application.ExecutablePath, Environment.GetFolderPath(Environment.SpecialFolder.Startup), "CATALOGUE", "-startminimized")
        Else
            'Si le raccourci existe, on vérifie qu'il pointe bien vers le bon endroit, et on le réécrit si besoin
            If (UCase(getShortCutTarget(Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "\" & "CATALOGUE.lnk")) <> UCase(Application.ExecutablePath)) _
            Or (UCase(getShortCutArguments(Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "\" & "CATALOGUE.lnk")) <> "-STARTMINIMIZED") Then
                Call createShortCut(Application.ExecutablePath, Environment.GetFolderPath(Environment.SpecialFolder.Startup), "CATALOGUE", "-startminimized")
            End If
        End If
    End Sub

    'Supprime l'application du démarrage de Windows
    Sub deleteExeFromStartupFolder()
        If System.IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "\" & "CATALOGUE.lnk") Then
            System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Startup) & "\" & "CATALOGUE.lnk")
        End If
    End Sub

End Module
