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

'Dépendances
Imports System.IO                       'File
Imports System.Text.RegularExpressions  'Regex

'Code principal
'Form "CATALOGUE"
Public Class CATALOGUE

    'A l'initialisation de la fenêtre
    Private Sub CATALOGUE_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Met à jour la zone d'infos de la fenêtre
        Call updateLabelInfo()

        'Répète l'opération toutes les 10 minutes
        Dim myTimer As New System.Windows.Forms.Timer
        myTimer.Interval = 300000
        myTimer.Start()
        AddHandler myTimer.Tick, AddressOf updateLabelInfo

        'Gestion des paramètres depuis le fichier CATALOGUE.ini
        'Si le fichier CATALOGUE.ini n'existe pas dans "C:\Users\USERNAME\AppData\Roaming\CATALOGUE", on le crée
        If Not doesFileExist(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini") Then
createIniFile:
            'Création du fichier .ini, valeurs par défaut à True
            '
            '[OPTIONS]
            'NOTIFICATIONS = True
            'AUTOSTART = True
            'ACTIVATEMACRO = True
            'MINIMIZE = True
            '
            ini.AddSection("OPTIONS").AddKey("AUTOSTART").Value = True
            ini.AddSection("OPTIONS").AddKey("NOTIFICATIONS").Value = True
            ini.AddSection("OPTIONS").AddKey("MINIMIZE").Value = True
            ini.AddSection("OPTIONS").AddKey("ACTIVATEMACRO").Value = True
            'Si le dossier "C:\Users\USERNAME\AppData\Roaming\CATALOGUE" n'existe pas, on le crée
            If Not doesDirectoryExists(getAppDataPath() & "\" & "CATALOGUE") Then
                Call createDirectory(getAppDataPath() & "\" & "CATALOGUE")
            End If
            'Enregistrement du fichier .ini créé
            ini.Save(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini")
        Else
            'Le fichier CATALOGUE.ini existe dans "C:\Users\USERNAME\AppData\Roaming\CATALOGUE", on le charge
            ini.Load(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini")
        End If

        Try
            'Démarrage automatique
            bAutoStart = ini.GetSection("OPTIONS").GetKey("AUTOSTART").GetValue
            'Notifications de modification du catalogue
            bNotifications = ini.GetSection("OPTIONS").GetKey("NOTIFICATIONS").GetValue
            'Minimisation dans la zone de notifications
            bMinimize = ini.GetSection("OPTIONS").GetKey("MINIMIZE").GetValue
            'Activation automatique de la macro Excel
            bAutoActivateExcelMacro = ini.GetSection("OPTIONS").GetKey("ACTIVATEMACRO").GetValue
            'Si l'intégrité du fichier .ini n'est pas correcte, on le regénère
        Catch ex As Exception
            'Suppression de l'ancien fichier .ini
            System.IO.File.Delete(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini")
            'Création d'un nouveau fichier .ini
            GoTo createIniFile
        End Try

        If bAutoStart = True Then
            'Ajoute l'application au démarrage de Windows
            Call addExeToStartupFolder()
        Else
            'Supprime l'application du démarrage de Windows
            Call deleteExeFromStartupFolder()
        End If

        'Menu de l'icone dans la zone de notifications :

        'Définition des différentes entrées (ToolStripMenuItem) du menu (ContextMenuStrip)
        menuHTMLCatalogue.Text = "Parcourir le catalogue"
        menuHTMLCatalogue.Image = My.Resources.book_open_page_variant_small
        menuXLSCatalogue.Text = "Modifier le catalogue"
        menuXLSCatalogue.Image = My.Resources.database_small
        menuAddSheet.Text = "Ajouter une fiche"
        menuAddSheet.Image = My.Resources.plus_circle_small
        menuSearch.Text = "Rechercher"
        menuSearch.Image = My.Resources.search_small
        menuAbout.Text = "A propos"
        menuAbout.Image = My.Resources.information_small_black
        menuOptions.Text = "Options"
        menuOptions.Image = My.Resources.settings_small
        menuExit.Text = "Quitter"
        menuExit.Image = My.Resources.exit_small

        'Ajout des entrées (ToolStripMenuItem) au menu (ContextMenuStrip)
        ContextMenuStrip1.Items.Add(menuHTMLCatalogue)
        ContextMenuStrip1.Items.Add(menuXLSCatalogue)
        ContextMenuStrip1.Items.Add(menuAddSheet)
        ContextMenuStrip1.Items.Add(menuSearch)
        ContextMenuStrip1.Items.Add(New ToolStripSeparator())
        ContextMenuStrip1.Items.Add(menuAbout)
        ContextMenuStrip1.Items.Add(menuOptions)
        ContextMenuStrip1.Items.Add(menuExit)

        'Actions lors du clic sur les entrées (ToolStripMenuItem) du menu (ContextMenuStrip)
        AddHandler menuHTMLCatalogue.Click, AddressOf menuHTMLCatalogue_Click
        AddHandler menuXLSCatalogue.Click, AddressOf menuXLSCatalogue_Click
        AddHandler menuAddSheet.Click, AddressOf menuAddSheet_Click
        AddHandler menuSearch.Click, AddressOf menuSearch_Click
        AddHandler menuAbout.Click, AddressOf menuAbout_Click
        AddHandler menuExit.Click, AddressOf menuExit_Click
        AddHandler menuOptions.Click, AddressOf menuOptions_Click

        'Ajout du menu (ContextMenuStrip) à l'icone de la zone de notifications (NotifyIcon)
        NotifyIcon1.ContextMenuStrip = ContextMenuStrip1

        'Si au moins un argument est passé à l'exe
        If My.Application.CommandLineArgs.Count > 0 Then
            'Si le 1er argument est "-startminimized"
            If UCase(My.Application.CommandLineArgs.Item(0)) = "-STARTMINIMIZED" Then
                'Ouverture de la Form minimisée dans la barre de tâches
                Me.WindowState = FormWindowState.Minimized
                ' Si minimisation dans la zone de notifications activée
                If bMinimize Then
                    Me.ShowInTaskbar = False
                End If
            End If
        End If

    End Sub

#Region "Boutons interface"

    'Bouton "Parcourir le catalogue HTML"
    Private Sub cbHTMLCatalogue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbHTMLCatalogue.Click

        'Activation gif d'attente
        cbHTMLCatalogue.BackColor = Color.WhiteSmoke
        labelWaitHtml.Visible = True

        Call HTMLCatalogue()

        'Désactivation gif d'attente
        labelWaitHtml.Visible = False
        cbHTMLCatalogue.BackColor = Color.White

    End Sub

    'Bouton "Modifier le catalogue EXCEL"
    Private Sub cbXLSCatalogue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbXLSCatalogue.Click

        'Activation gif d'attente
        cbXLSCatalogue.BackColor = Color.WhiteSmoke
        labelWaitXsl.Visible = True

        Call XLSCatalogue()

        If bAutoActivateExcelMacro Then
            'Activation automatique de la macro Excel
            Call timerToActivateExcelWorkbookMacros()
        End If

        'Désactivation gif d'attente
        labelWaitXsl.Visible = False
        cbXLSCatalogue.BackColor = Color.White

    End Sub

    'Bouton "Ajouter une fiche au catalogue"
    Private Sub cbAddSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAddSheet.Click

        'Activation gif d'attente
        cbAddSheet.BackColor = Color.WhiteSmoke
        labelWaitAddSheet.Visible = True

        Call AddSheet()

        If bAutoActivateExcelMacro Then
            'Activation automatique de la macro Excel
            Call timerToActivateExcelWorkbookMacros()
        End If

        'Désactivation gif d'attente
        labelWaitAddSheet.Visible = False
        cbAddSheet.BackColor = Color.White

    End Sub

    'Bouton "Rechercher dans les fiches"
    Private Sub cbSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSearch.Click

        'Affichage de la zone de saisie des mots à rechercher
        tbSearch.Visible = True
        tbSearch.Select()

    End Sub

    'Si la zone de saisie des mots à rechercher n'a plus le focus, on la masque
    Private Sub tbSearch_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSearch.LostFocus
        tbSearch.Visible = False
    End Sub

    'Lors de l'appui sur Echap ou Entrée dans la zone de saisie des mots à rechercher
    Private Sub tbSearch_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbSearch.KeyPress

        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then            'Entrée

            'Activation gif d'attente
            cbSearch.BackColor = Color.WhiteSmoke
            labelWaitSearch.Visible = True

            tbSearch.Visible = False

            SEARCH.Enabled = False

            'Pour ne pas bloquer la Form, on lance la recherche dans un nouveau thread
            Dim t As New Threading.Thread(AddressOf SearchSheets)
            t.Start()

        ElseIf e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Escape) Then       'Echap

            tbSearch.Visible = False

        End If

    End Sub

#End Region

#Region "Fonctions spécifiques"

    'Parcourir le catalogue HTML
    Sub HTMLCatalogue()
        If doesFileExist(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName) Then
            Call browseURL(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName)
        Else
            MsgBox("Le fichier " & sHtmlCatalogueName & " est introuvable." & vbNewLine & "Assurez vous que le launcher n'a pas été déplacé.", vbCritical, "Erreur, fichier introuvable")
        End If
    End Sub

    'Modifier le catalogue EXCEL
    Sub XLSCatalogue()
        'Variable d'environnement à Nothing : Ouverture du catalogue
        If Not Environment.GetEnvironmentVariable("InBatch") Is Nothing Then
            Environment.SetEnvironmentVariable("InBatch", Nothing)
        End If

        If doesFileExist(sLauncherPath & "\" & sWorkbookName) Then
            Call openFile(sLauncherPath & "\" & sWorkbookName)
        Else
            MsgBox("Le fichier " & sWorkbookName & " est introuvable." & vbNewLine & "Assurez vous que le launcher n'a pas été déplacé.", vbCritical, "Erreur, fichier introuvable")
        End If
    End Sub

    'Ajouter une fiche au catalogue
    Sub AddSheet()
        'Variable d'environnement à True : Création d'une fiche
        If Environment.GetEnvironmentVariable("InBatch") Is Nothing Then
            Environment.SetEnvironmentVariable("InBatch", True)
        End If

        If doesFileExist(sLauncherPath & "\" & sWorkbookName) Then
            Call openFile(sLauncherPath & "\" & sWorkbookName)
        Else
            MsgBox("Le fichier " & sWorkbookName & " est introuvable." & vbNewLine & "Assurez vous que le launcher n'a pas été déplacé.", vbCritical, "Erreur, fichier introuvable")
        End If
    End Sub

    'Rechercher un ou plusieurs mots dans le contenu des fiches
    Sub SearchSheets()

        'Si l'utilisateur a saisi au moins un mot
        If delAllSpace(tbSearch.Text) <> String.Empty Then

            'Si le dossier contenant les fiches est accessible
            If doesDirectoryExists(sLauncherPath & "\" & sSheetsPath) Then

                'Tableau de mots saisis par l'utilisateur
                Dim sWordsToSearch()
                sWordsToSearch = Split(delAllSpace(tbSearch.Text), " ")
                Dim sWord As String

                'Dossier contenant les fiches
                Dim sheetsDirectory As New IO.DirectoryInfo(sLauncherPath & "\" & sSheetsPath)

                'Récupération des fichiers au format .html uniquement
                Dim allSheets As IO.FileInfo() = sheetsDirectory.GetFiles("*.html")
                Dim singleSheet As IO.FileInfo
                Dim sContent As String
                Dim sTitle As String

                'Regex de récupération du titre d'une fiche
                Dim regexObj As New Regex("(?s)(?<=<h1>)(.+?)(?=</h1>)")

                'Vidage du dictionnaire
                searchResultsDic.Clear()

                'Pour chaque fichier .html du dossier de fiches
                For Each singleSheet In allSheets
                    'On ignore les différents templates de fiches
                    If singleSheet.Name <> "PREVIEW_TMP.html" And singleSheet.Name <> "TEMPLATE.html" And singleSheet.Name <> "TEMPLATE_PREVIEW.html" Then
                        'Pour chaque mot saisi par l'utilisateur
                        For Each sWord In sWordsToSearch
                            'Si le mot existe dans le fichier .html
                            sContent = getFileContent(singleSheet.FullName)
                            If sContent.ToLower.Contains(sWord.ToLower) Then
                                'Récupération du titre de la fiche : balise <h1>
                                sTitle = Regex.Replace((regexObj.Matches(sContent).Item(0).Groups(1).Value), "<.*?>", "")
                                'Ajout du titre et de l'URL de la fiche au dictionnaire
                                If Not searchResultsDic.ContainsKey(sTitle) Then
                                    searchResultsDic.Add(sTitle, singleSheet.FullName)
                                End If
                            End If
                        Next
                    End If
                Next

            Else
                MsgBox("Le dossier " & sSheetsPath & " est introuvable." & vbNewLine & "Assurez vous que le launcher n'a pas été déplacé.", vbCritical, "Erreur, dossier introuvable")
            End If

        End If

        'Masque le gif d'attente du bouton "Rechercher dans les fiches" une fois la recherche terminée
        Call hideWaitAnimation()

    End Sub

    'Le gif d'attente du bouton "Rechercher dans les fiches" est verrouillé par le thread de la Form
    Sub hideWaitAnimation()

        'Si le gif d'attente est verrouillé par un autre thread
        If labelWaitSearch.InvokeRequired Then
            labelWaitSearch.Invoke(New MethodInvoker(AddressOf hideWaitAnimation))
        Else
            'Désactivation gif d'attente
            labelWaitSearch.Visible = False
            cbSearch.BackColor = Color.White

            'Vidage de la ListView
            SEARCH.lvSearchResults.Items.Clear()

            If searchResultsDic.Count <> 0 Then

                'Header de la ListView
                With SEARCH.lvSearchResults
                    'S'il n'a pas déjà été configuré
                    If .Columns.Count = 0 Then
                        .Columns.Add("Titre de la fiche")
                    End If
                End With

                'Titre de la fenêtre
                If searchResultsDic.Count > 1 Then
                    SEARCH.Text = "Résultats de la recherche : " & searchResultsDic.Count & " éléments"
                Else
                    SEARCH.Text = "Résultat de la recherche : 1 élément"
                End If

                'Remplissage de la ListView
                For Each sKey In searchResultsDic.Keys
                    SEARCH.lvSearchResults.Items.Add(sKey)
                Next

                'Autosize des colonnes de la ListView en fonction du contenu
                SEARCH.lvSearchResults.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
                SEARCH.lvSearchResults.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
                SEARCH.lvSearchResults.Columns.Item(0).Width = SEARCH.lvSearchResults.Columns.Item(0).Width - 6

                SEARCH.Enabled = True

                SEARCH.Show()
            Else
                SEARCH.Hide()
                MsgBox("Aucun résultat pour cette recherche !", vbInformation, "Pas de résultat")
            End If
        End If

    End Sub

    'Récupère et retourne le nombre de fiches du catalogue
    Function getSheetsNbr() As String
        If doesFileExist(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName) Then
            Dim sHtmlCatalogueContent As String
            sHtmlCatalogueContent = getFileContent(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName)

            Dim colMatchResults As MatchCollection
            Dim regexObj As New Regex("</tr>")
            colMatchResults = regexObj.Matches(sHtmlCatalogueContent)

            getSheetsNbr = colMatchResults.Count - 1
        Else
            getSheetsNbr = 0
        End If
    End Function

    'Récupère et retourne la date de modification du catalogue
    Function getCatalogueLastModifDate() As DateTime
        Dim lastModifDT As DateTime
        If doesFileExist(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName) Then
            lastModifDT = File.GetLastWriteTime(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName)
        End If
        getCatalogueLastModifDate = lastModifDT
    End Function

    'Met à jour la zone d'infos de la fenêtre :
    '   - Nombre de fiches du catalogue
    '   - Date de modification du catalogue
    'Affiche une popup d'information si le catalogue a été modifié
    '   - Nombre de nouvelles fiches
    '   - Date de modification du catalogue
    Sub updateLabelInfo()

        'Debug : Si un fichier nommé "END" existe à la racine du programme, quitter le programme
        'Permet de le mettre à jour sans que les utilisateurs aient à quitter manuellement le programme
        If doesFileExist(sLauncherPath & "\" & "END") Then
            End
        End If

        'Le programme vient d'être exécuté
        If iSheetsNbr = 0 And sLastModifDate.ToString = "01/01/0001 00:00:00" Then
            iSheetsNbr = CInt(getSheetsNbr())
            sLastModifDate = getCatalogueLastModifDate()
        Else
            Dim iSheetsNbrTmp As Integer
            iSheetsNbrTmp = CInt(getSheetsNbr())
            Dim sLastModifDateTmp As DateTime
            sLastModifDateTmp = getCatalogueLastModifDate()
            'Le catalogue a été modifié, deux cas possibles
            If iSheetsNbrTmp <> iSheetsNbr Or sLastModifDateTmp <> sLastModifDate Then
                If iSheetsNbrTmp > iSheetsNbr Then
                    'Nouvelle(s) fiche(s)
                    If iSheetsNbrTmp - iSheetsNbr > 1 Then
                        NotifyIcon1.BalloonTipText = iSheetsNbrTmp - iSheetsNbr & " nouvelles fiches dans le catalogue"
                    Else
                        NotifyIcon1.BalloonTipText = iSheetsNbrTmp - iSheetsNbr & " nouvelle fiche dans le catalogue"
                    End If
                Else
                    'Edition / suppression de fiche(s) : pas de nouvelle(s) fiche(s), mais catalogue modifié
                    NotifyIcon1.BalloonTipText = "Catalogue modifié le " & sLastModifDateTmp.ToString("dd/MM/yy") & " à " & sLastModifDateTmp.ToString("HH:mm")
                End If
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipTitle = "CATALOGUE"
                'Si les notifications sont activées dans le fichier de configuration CATALOGUE.ini
                If bNotifications Then
                    'Affichage d'une popup dans la zone de notifications
                    NotifyIcon1.ShowBalloonTip(0)
                End If
                iSheetsNbr = iSheetsNbrTmp
                sLastModifDate = sLastModifDateTmp
            End If
        End If
        'Mise à jour de la zone d'infos de la fenêtre
        If iSheetsNbr > 1 Then
            labelInfo.Text = "Catalogue édité le " & sLastModifDate.ToString("dd/MM/yyyy") & " : " & iSheetsNbr & " fiches"
        Else
            labelInfo.Text = "Catalogue édité le " & sLastModifDate.ToString("dd/MM/yyyy") & " : " & iSheetsNbr & " fiche"
        End If
    End Sub

#End Region

#Region "Icone zone de notification"

    'Lors de la minimisation de la fenêtre, la masquer et afficher un message dans la zone de notifications
    Private Sub CATALOGUE_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        'Si la minimisation dans la zone de notifications est activée dans le fichier de configuration CATALOGUE.ini
        If bMinimize Then
            If Me.WindowState = FormWindowState.Minimized Then
                Me.Hide()
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipTitle = "CATALOGUE"
                NotifyIcon1.BalloonTipText = "Minimisé dans la zone de notifications"
                NotifyIcon1.ShowBalloonTip(250)
                'Si la fenêtre de recherche est ouverte, on la réduit
                SEARCH.Hide()
            End If
        End If
    End Sub

    'Lors du clic sur la popup de la zone de notifications, ouvre les dernières fiches créées ou le catalogue HTML selon le cas
    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As Object, ByVal e As EventArgs) Handles NotifyIcon1.BalloonTipClicked
        If doesFileExist(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName) Then
            If (NotifyIcon1.BalloonTipText).Contains("nouvelle") Then
                'Si nouvelle(s) fiche(s), on les ouvre lors du clic sur la popup de la zone de notifications

                'Nombre de nouvelles fiches
                Dim colMatchNbr As MatchCollection
                Dim regexNbr As New Regex("([0-9]+)")
                colMatchNbr = regexNbr.Matches(NotifyIcon1.BalloonTipText)

                'Fiches du catalogue HTML
                Dim sHtmlCatalogueContent As String
                sHtmlCatalogueContent = getFileContent(sLauncherPath & "\" & sHtmlCataloguePath & "\" & sHtmlCatalogueName)
                Dim colMatchResults As MatchCollection
                Dim regexObj As New Regex("<tr class='(.*)'>")
                colMatchResults = regexObj.Matches(sHtmlCatalogueContent)

                'Ouverture des dernières fiches créées
                For i As Integer = 1 To colMatchNbr.Item(0).Groups(1).Value
                    Call browseURL(sLauncherPath & "\" & sSheetsPath & "\" & colMatchResults.Item(colMatchResults.Count - i).Groups(1).Value & ".html")
                Next
            ElseIf (NotifyIcon1.BalloonTipText).Contains("modifié") Then
                'Si modification du catalogue, on l'ouvre lors du clic sur la popup de la zone de notifications
                Call HTMLCatalogue()
            End If
        Else
            Exit Sub
        End If
    End Sub


    'Lors du clic sur l'icone de la zone de notifications, réafficher la fenêtre
    Private Sub NotifyIcon1_MouseUpk(ByVal sender As Object, ByVal e As MouseEventArgs) Handles NotifyIcon1.MouseUp
        'Seulement si clic gauche
        'Clic droit ne doit qu'afficher le menu
        If (e.Button = MouseButtons.Left) Then
            Me.Show()
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
        End If
    End Sub

    'Lors de la fermeture de la fenêtre, détruire l'icone de la zone de notification
    Private Sub CATALOGUE_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        NotifyIcon1.Visible = False
        NotifyIcon1.Icon = Nothing
        NotifyIcon1.Dispose()
    End Sub

    'Action lors du clic sur le menu "Parcourir le catalogue" de l'icone de la zone de notifications 
    Private Sub menuHTMLCatalogue_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Activation fenêtre d'attente
        WAIT.Show()
        Call HTMLCatalogue()
        'Fermeture fenêtre d'attente
        WAIT.Hide()
    End Sub

    'Action lors du clic sur le menu "Modifier le catalogue" de l'icone de la zone de notifications
    Private Sub menuXLSCatalogue_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Activation fenêtre d'attente
        WAIT.Show()
        Call XLSCatalogue()
        If bAutoActivateExcelMacro Then
            'Activation automatique de la macro Excel
            Call timerToActivateExcelWorkbookMacros()
        End If
        'Fermeture fenêtre d'attente
        WAIT.Hide()
    End Sub

    'Action lors du clic sur le menu "Ajouter une fiche" de l'icone de la zone de notifications
    Private Sub menuAddSheet_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Activation fenêtre d'attente
        WAIT.Show()
        Call AddSheet()
        If bAutoActivateExcelMacro Then
            'Activation automatique de la macro Excel
            Call timerToActivateExcelWorkbookMacros()
        End If
        'Fermeture fenêtre d'attente
        'WAIT.Hide()
        'Bugfix : La fenêtre d'attente se ferme toute seul lorsqu'elle détecte la présence du catalogue Excel
    End Sub

    'Action lors du clic sur le menu "Rechercher" de l'icone de la zone de notifications
    Private Sub menuSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Affichage du launcher si masqué
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.ShowInTaskbar = True
        'Focus sur le launcher
        Me.Activate()
        'Focus sur la zone de saisie
        tbSearch.Visible = True
        tbSearch.Select()
    End Sub

    'Action lors du clic sur le menu "A propos" de l'icone de la zone de notifications
    Private Sub menuAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ABOUT.Show()
    End Sub

    'Action lors du clic sur le menu "Options" de l'icone de la zone de notifications
    Private Sub menuOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        OPTIONS.Show()
    End Sub

    'Action lors du clic sur le menu "Quitter" de l'icone de la zone de notifications
    Private Sub menuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Easter egg : Avant de quitter, sort la réplique "Bye bye !" issue du jeu Worms
        My.Computer.Audio.Play(My.Resources.worm_sound_bye_bye, AudioPlayMode.WaitToComplete)
        Close()
    End Sub

#End Region

#Region "Fenêtre « A propos »"

    'Curseur d'aide au survol de la zone d'infos de la fenêtre
    Private Sub panelInfo_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles panelInfo.MouseMove
        Cursor.Current = Cursors.Help
    End Sub

    'Curseur d'aide au survol de la zone d'infos de la fenêtre
    Private Sub labelInfo_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles labelInfo.MouseMove
        Cursor.Current = Cursors.Help
    End Sub

    'Ouverture de la fenêtre "A propos" lors du clic sur la zone d'infos de la fenêtre
    Private Sub panelInfo_MouseUp(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles panelInfo.MouseUp
        'Seulement si clic gauche
        If (e.Button = MouseButtons.Left) Then
            ABOUT.Show()
        End If
    End Sub

    'Ouverture de la fenêtre "A propos" lors du clic sur la zone d'infos de la fenêtre
    Private Sub labelInfo_MouseUp(ByVal sender As System.Object, ByVal e As MouseEventArgs) Handles labelInfo.MouseUp
        'Seulement si clic gauche
        If (e.Button = MouseButtons.Left) Then
            ABOUT.Show()
        End If
    End Sub

#End Region

End Class

'Si l'application est exécutée alors qu'une instance est déjà en cours :
'   - on affiche la fenêtre 
'   - on met la fenêtre au premier plan
Namespace My
    Partial Friend Class MyApplication
        Private Sub MyApplication_StartupNextInstance(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupNextInstanceEventArgs) Handles Me.StartupNextInstance
            CATALOGUE.Show()
            e.BringToForeground = True
        End Sub
    End Class
End Namespace