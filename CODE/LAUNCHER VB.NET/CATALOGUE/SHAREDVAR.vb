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

'Variables partagées entre tous les modules / classes
Public Module SHAREDVAR

    'Chemins
    Public sLauncherPath As String = My.Application.Info.DirectoryPath
    Public Const sHtmlCataloguePath As String = "CATALOGUE"
    Public Const sSheetsPath As String = "SHEETS"
    Public Const sHtmlCatalogueName As String = "CATALOGUE.html"
    Public Const sWorkbookName As String = "CATALOGUE.xlsm"

    'Nombre de fiches
    Public iSheetsNbr As Integer
    'Date de modification du catalogue
    Public sLastModifDate As DateTime

    'Fichier de configuration
    Public ini As New IniFile
    'Démarrage automatique
    Public bAutoStart As Boolean
    'Notifications de modification du catalogue
    Public bNotifications As Boolean
    'Minimisation dans la zone de notifications
    Public bMinimize As Boolean
    'Activation automatique de la macro Excel
    Public bAutoActivateExcelMacro As Boolean

    'Entrées du menu de l'icone de la zone de notifications
    Public menuHTMLCatalogue As New ToolStripMenuItem
    Public menuXLSCatalogue As New ToolStripMenuItem
    Public menuAddSheet As New ToolStripMenuItem
    Public menuSearch As New ToolStripMenuItem
    Public menuAbout As New ToolStripMenuItem
    Public menuExit As New ToolStripMenuItem
    Public menuOptions As New ToolStripMenuItem

    'Entrée du menu du clic droit de la ListView de la Form SEARCH
    Public menuOpenSheet As New ToolStripMenuItem
    Public menuCopySheetUrl As New ToolStripMenuItem

    'Timer permettant de vérifier si le catalogue Excel est ouvert
    Public timerWindow As New System.Windows.Forms.Timer

    'Dictionnaire de résultats de recherche
    'Clé : Nom de la fiche
    'Valeur : URL de la fiche
    Public searchResultsDic As New Dictionary(Of String, String)

End Module
