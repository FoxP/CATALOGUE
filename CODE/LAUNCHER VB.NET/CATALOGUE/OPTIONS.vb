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

'Form d'options "OPTIONS"
Public Class OPTIONS

    'A l'ouverture de la Form "OPTIONS"
    Private Sub OPTIONS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Si démarrage automatique activé, on coche la case "Démarrage à l'ouverture de session"
        If bAutoStart Then
            cbAutostart.Checked = True
            'Sinon, on la décoche
        Else
            cbAutostart.Checked = False
        End If

        'Si notifications activées, on coche la case "Alertes de mise à jour du catalogue"
        If bNotifications Then
            cbNotifications.Checked = True
            'Sinon, on la décoche
        Else
            cbNotifications.Checked = False
        End If

        'Si minimisation activée, on coche la case "Réduire dans la zone de notifications"
        If bMinimize Then
            cbMinimize.Checked = True
            'Sinon, on la décoche
        Else
            cbMinimize.Checked = False
        End If

        'Si activation automatique de la macro Excel activée, on coche la case "Activation automatique de la macro Excel"
        If bAutoActivateExcelMacro Then
            cbAutoActivateExcelMacro.Checked = True
            'Sinon, on la décoche
        Else
            cbAutoActivateExcelMacro.Checked = False
        End If

    End Sub

    'Clic sur la case "Alertes de mise à jour du catalogue"
    Private Sub cbNotifications_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbNotifications.CheckedChanged

        'Si cochée
        If cbNotifications.Checked Then
            bNotifications = True
            ini.AddSection("OPTIONS").AddKey("NOTIFICATIONS").Value = True
        Else
            bNotifications = False
            ini.AddSection("OPTIONS").AddKey("NOTIFICATIONS").Value = False
        End If
        'Modification du fichier de configuration CATALOGUE.ini
        ini.Save(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini")

    End Sub

    'Clic sur la case "Démarrage à l'ouverture de session"
    Private Sub cbAutostart_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAutostart.CheckedChanged

        'Si cochée
        If cbAutostart.Checked Then
            bAutoStart = True
            ini.AddSection("OPTIONS").AddKey("AUTOSTART").Value = True
            'Ajout de l'application au démarrage de session
            Call addExeToStartupFolder()
        Else
            bAutoStart = False
            ini.AddSection("OPTIONS").AddKey("AUTOSTART").Value = False
            'Suppression de l'application au démarrage de session
            Call deleteExeFromStartupFolder()
        End If
        'Modification du fichier de configuration CATALOGUE.ini
        ini.Save(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini")

    End Sub

    'Clic sur la case "Réduire dans la zone de notifications"
    Private Sub cbMinimize_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMinimize.CheckedChanged

        'Si cochée
        If cbMinimize.Checked Then
            bMinimize = True
            ini.AddSection("OPTIONS").AddKey("MINIMIZE").Value = True
        Else
            bMinimize = False
            ini.AddSection("OPTIONS").AddKey("MINIMIZE").Value = False
        End If
        'Modification du fichier de configuration CATALOGUE.ini
        ini.Save(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini")

    End Sub

    'Clic sur la case "Activation automatique de la macro Excel"
    Private Sub cbAutoActivateExcelMacro_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAutoActivateExcelMacro.CheckedChanged

        'Si cochée
        If cbAutoActivateExcelMacro.Checked Then
            bAutoActivateExcelMacro = True
            ini.AddSection("OPTIONS").AddKey("ACTIVATEMACRO").Value = True
        Else
            bAutoActivateExcelMacro = False
            ini.AddSection("OPTIONS").AddKey("ACTIVATEMACRO").Value = False
        End If
        'Modification du fichier de configuration CATALOGUE.ini
        ini.Save(getAppDataPath() & "\" & "CATALOGUE" & "\" & "CATALOGUE.ini")

    End Sub

End Class