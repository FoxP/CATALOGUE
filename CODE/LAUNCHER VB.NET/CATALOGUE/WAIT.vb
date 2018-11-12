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

'Form d'attente "WAIT"
Public Class WAIT

    'A l'ouverture de la Form "WAIT"
    Private Sub WAIT_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim myTimerWait As New Timer
        'Toutes les 1/2 secondes
        myTimerWait.Interval = 500
        myTimerWait.Start()
        'Si le catalogue Excel est ouvert, ferme la Form "WAIT"
        AddHandler myTimerWait.Tick, AddressOf isExcelFormOpened
    End Sub

    'Vérifie si le catalogue Excel est ouvert, si oui, ferme la Form "WAIT"
    Sub isExcelFormOpened()
        Dim positionArray(3) As Double
        positionArray = getWindowPosition("CATALOGUE.xlsm")
        If Math.Abs(positionArray(0)) + Math.Abs(positionArray(1)) + Math.Abs(positionArray(2)) + Math.Abs(positionArray(3)) <> 0 Then
            Me.Hide()
        End If
    End Sub

End Class