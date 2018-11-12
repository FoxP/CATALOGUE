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
Imports System.ComponentModel   'CancelEventArgs

'Form de résultats de recherche "SEARCH"
Public Class SEARCH

    'Au clic sur un élément de la ListView, ouverture de la fiche associée
    Private Sub lvSearchResults_ItemActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvSearchResults.ItemActivate
        If searchResultsDic.ContainsKey(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text) Then
            If doesFileExist(searchResultsDic(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text)) Then
                Call browseURL(searchResultsDic(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text))
            End If
        End If
    End Sub

    'Au chargement de la Form, définition du menu contextuel du clic droit
    Private Sub SEARCH_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Menu contextuel du clic droit
        Dim ContextMenuStripResults As New ContextMenuStrip

        'Menu "Ouvrir la fiche"
        menuOpenSheet.Text = "Ouvrir la fiche"
        menuOpenSheet.Image = My.Resources.open_small
        'Menu "Copier l'URL de la fiche"
        menuCopySheetUrl.Text = "Copier l'URL de la fiche"
        menuCopySheetUrl.Image = My.Resources.copy_small

        'Ajout des menus au ContextMenuStrip
        ContextMenuStripResults.Items.Add(menuOpenSheet)
        ContextMenuStripResults.Items.Add(menuCopySheetUrl)

        'Evénements lors du clic sur les menus
        AddHandler menuOpenSheet.Click, AddressOf menuOpenSheet_Click
        AddHandler menuCopySheetUrl.Click, AddressOf menuCopySheetUrl_Click

        'Ajout du ContextMenuStrip à la ListView
        Me.lvSearchResults.ContextMenuStrip = ContextMenuStripResults
    End Sub

    'A la fermeture de la Form, destruction des événements du menu contextuel du clic droit
    Private Sub SEARCH_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        RemoveHandler menuOpenSheet.Click, AddressOf menuOpenSheet_Click
        RemoveHandler menuCopySheetUrl.Click, AddressOf menuCopySheetUrl_Click
    End Sub

    'Clic sur le menu "Ouvrir la fiche" du menu contextuel du clic droit
    Private Sub menuOpenSheet_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If searchResultsDic.ContainsKey(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text) Then
            If doesFileExist(searchResultsDic(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text)) Then
                Call browseURL(searchResultsDic(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text))
            End If
        End If
    End Sub

    'Clic sur le menu "Copier l'URL de la fiche" du menu contextuel du clic droit
    Private Sub menuCopySheetUrl_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If searchResultsDic.ContainsKey(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text) Then
            If doesFileExist(searchResultsDic(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text)) Then
                'Vidage du presse papier
                Clipboard.Clear()
                'Copie de l'URL de la fiche dans le presse papier
                Clipboard.SetText(searchResultsDic(Me.lvSearchResults.Items(Me.lvSearchResults.FocusedItem.Index).SubItems(0).Text))
            End If
        End If
    End Sub

    'Empêche le menu contextuel du clic droit de la ListView de s'activer si aucun élément n'est sélectionné
    Private Sub ContextMenuStripResults_Opening(ByVal sender As System.Object, ByVal e As CancelEventArgs)
        If Me.lvSearchResults.SelectedItems.Count = 0 Then
            e.Cancel = True
        End If
    End Sub

End Class