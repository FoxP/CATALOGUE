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

'Form d'info sur le programme
Public Class ABOUT

    Private Sub ABOUT_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        labelExeInfo.Text = "Version de l'exécutable : " & "v" & My.Application.Info.Version.ToString & vbNewLine & _
                            "Date de compilation : " & retrieveLinkerTimestamp(Application.ExecutablePath).ToString("dd/MM/yyyy")
    End Sub

End Class