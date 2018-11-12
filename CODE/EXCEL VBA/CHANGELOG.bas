Attribute VB_Name = "CHANGELOG"
'==============================================================================================================================================================
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
'Versions :
'           - v1.0 : 02/11/2016 : Code initial
'           - v1.5 : 21/11/2016 : Support du Markdown pour du contenu riche dans les champs "Problème" et "Solution"
'                                 Gestion de la maturité des fiches : Draft, Submitted, Release, Superseded, Obsolete
'           - v2.0 : 14/02/2017 : Launcher VB.NET :
'                                     - lancement des catalogues HTML et XLS, et création de fiches
'                                     - minimisation dans la zone de notifications, menu du clic droit
'                                     - notification si nouvelle(s) fiche(s) et modification du catalogue
'                                     - aperçu du nombre de fiches du catalogue (MAJ toutes les 10 minutes)
'                                     - ajout automatique du launcher au démarrage de la session utilisateur
'                                     - gestion de paramètres stockés dans un fichier de configuration au format .ini
'                                 Easter egg de Noël : durant le mois de décembre, fait tomber de la neige sur les fiches
'                                 Catalogue Excel désormais compatible multi-utilisateurs, gestion des fiches au clic droit
'                                 Robustification / refactoring du code, correction de multiples bugs, quelques optimisations
'                                 Support des pièces jointes (images / documents de tous types / tailles), galerie photo automatique
'                                 Possibilité de re-générer manuellement une copie complète du catalogue HTML depuis le catalogue Excel
'                                 Création d'un certificat permettant l'utilisation du catalogue sans avertissements de sécurité d'Excel
'                                 Correcteur orthographique et fonction de prévisualisation lors de la création ou l'édition d'une fiche
'                                 Ajout de statistiques dans une feuille dédiée du catalogue Excel : pourcentage de chaque catégorie, etc
'                                     - nécessite cependant de dé-partager temporairement le classeur Excel pour effectuer une mise à jour
'                                 Enrichissement du template HTML des fiches :
'                                     - responsive design, s'adapte à tous types d'écrans
'                                     - comptage du nombre total d'images et de documents
'                                     - création d'une feuille de style dédiée pour l'impression
'                                     - nuage de mots clés permettant de rechercher des fiches liées
'                                     - affichage / masquage rapide des différentes sections de la fiche
'                                     - refonte graphique : tooltips, notifications, effets de survol, ombres, ...
'                                     - meilleur affichage des tableaux, images et portions de code insérés en Markdown
'                                     - éditeur Markdown WYSIWYG (What You See Is What You Get) dans la fiche d'aide au language
'           - v2.1 : 02/04/2017 : A la suppression d'une fiche, vérification si l'opération a réussi ou non
'                                 Suppression des caractères "," et ";" du champ des mots clés (au cas où...)
'                                 Boutons d'ouverture du dossier des pièces jointes, et tri par ordre alphabétique
'           - v3.0 : 26/08/2017 : Ouverture des fiches dans le navigateur défini par défaut au lieu d'Internet Explorer
'                                 Mise à jour automatique du contenu de la feuille Excel de statistiques à son ouverture
'                                 Si le classeur Excel est partagé, la feuille Excel de statistiques est masquée automatiquement
'                                     - un classeur partagé désactive la mise à jour des tableaux croisés dynamiques / graphiques
'                                 Re-configuration automatique du catalogue lors de la modification de la feuille Excel d'options
'                                     - plus besoin de fermer puis de réouvrir le classeur Excel après modification de ses options
'                                 Possibilité de fermer la UserForm de création / édition de fiche via la touche "Echap" du clavier
'                                 Correction du mauvais positionnement de la UserForm lors de l'ouverture du panel de pièces jointes
'                                 Depuis le panel de pièces jointes, affichage de l'une d'entre elle via le clic molette de la souris
'                                 Depuis le panel de pièces jointes, copie de son code Markdown dans le presse papier via double-clic
'                                 Le menu déroulant du clic droit pour l'édition / ... de fiche n'apparait désormais que là où il le doit
'                                 La fonction de prévisualisation d'une fiche en cours de création / édition affiche désormais les images
'                                 Ajout de différents boutons permettant l'ajout rapide de balises Markdown : souligné, gras, couleur, ...
'                                 Enrichissement du template HTML des fiches :
'                                     - ajout d'un bouton pour revenir directement au catalogue HTML depuis la fiche elle-même
'                                     - les caractères "<" ou ">" ne sont plus affichés en "&lt;" ou "&gt;" dans les balises code
'                                     - les anciennes versions d'une fiche sont accessibles directement depuis la fiche elle-même
'                                     - les fiches liées ainsi que les hyperliens sont accessibles depuis la fiche via un menu dédié
'                                     - lors du clic sur le statut, le type, le logiciel, le langage ou bien un mot clé dans une fiche,
'                                       ouverture du catalogue HTML et présentation des différences fiches ayant la même caractéristique
'                                 Enrichissement du template HTML du catalogue :
'                                     - amélioration des performances : minification et regroupement des fichiers JavaScript et CSS
'                                     - graphiques de statistiques dynamiques en bas de page : pourcentage de chaque catégorie, ...
'                                     - refonte graphique du style CSS : plus clair, plus lisible, s'adapte à tous types d'écrans, ...
'                                     - compression sans perte visuelle des images du template via "quantification" : images 8 bits indexées
'                                 Tests de compatibilité du catalogue HTML et des fiches sur les principaux navigateurs mobiles Android :
'                                     - Opera : fonctionne parfaitement, affichage strictement correct, très bonnes performances
'                                     - Firefox : fonctionne parfaitement, affichage strictement correct, lenteurs dans le catalogue
'                                     - Google Chrome : fonctionne parfaitement, tailles de polices inégales, très bonnes performances
'                                     - Gello (navigateur CyanogenMod / LineageOS) : fonctionne parfaitement, très bonnes performances
'           - v3.1 : 25/09/2017 : Refactoring du code VBA du catalogue Excel : mise au propre, relocalisation de chaque procédure
'                                 Création de fiche possible depuis le menu contextuel du clic droit : "Créer une nouvelle fiche"
'                                 Possibilité d'exporter sous forme d'archive une fiche donnée depuis le menu contextuel du clic droit :
'                                     - embarque la totalité des dépendances de la fiche : HTML, CSS, JavaScript, images et pièces jointes
'                                     - lors de l'envoi d'une fiche par e-mail, ajout automatique de l'archive en pièce jointe si souhaité
'                                     - si le logiciel de compression de données 7-Zip est installé, compression .7z LZMA2 ultra, sinon .zip
'                                 Suppression de fiche plus robuste : gestion du cas où une pièce jointe de fiche est verrouillée / utilisée
'                                 Compatibilité du code VBA du catalogue Excel avec les systèmes d'exploitation 32 bits (oui, ça existe...)
'                                 Hyperlien automatique lors du partage de l'URL d'une fiche par e-mail : formatage HTML du corps de l'e-email
'                                 Correction du titre de l'e-mail qui était en majuscules si partage d'une URL de fiche depuis sa version HTML
'                                 Suppression de la dépendance VBA Microsoft Outlook XX.X Object Library : plus robuste au changement de version
'                                     - si le logiciel de messagerie Microsoft Outlook n'est pas installé, les fonctions d'e-mails sont masquées
'                                 Nettoyage des fonctions accessibles depuis la fenêtre Excel de liste des macros (Alt + F8) pour plus de clarté
'                                 Utilisation de la barre d'état d'Excel pour afficher des informations sur la progression des opérations longues
'                                 La fiche d'aide au langage Markdown peut désormais être supprimée (masquée, plutôt) des catalogues Excel et HTML
'                                 Enrichissement de la fiche d'aide au langage Markdown :
'                                     - mise à jour des captures d'écran, meilleure intégration CSS de l'éditeur WYSIWYG
'                                     - liens vers la spécification du langage, ajout de l'article Wikipédia en pièce jointe
'                                     - documentation de l'ajout d'un code sur une ligne, de références de liens, de code HTML
'                                 Enrichissement du template HTML des fiches :
'                                     - copie rapide du contenu du champ "Code" dans le presse papier si clic sans sélection
'                                     - support des formats d'image WebP et SVG (Scalable Vector Graphics) dans la galerie photo
'                                     - génération d'un QR Code permettant de partager l'URL de la fiche sur des périphériques mobiles
'                                     - mise en avant des différentes sections cliquables par des animations CSS au survol de la souris
'                                     - modifications mineures de la galerie photo : traduction française, CSS, animations plus rapides
'                                     - ajout d'une tooltip au survol des caractéristiques de la fiche : statut, logiciel, type, langage
'                                     - correction des ombres portées des différentes sections de la fiche, et de l'espacement entre eux
'                                     - ajout d'icones de formats de fichiers pour les pièces jointes : 280 extensions de fichiers gérées
'                                     - uniformisation du clic souris : clic gauche exécute dans l'onglet actif, clic droit dans un nouvel onglet
'                                 Enrichissement du template HTML du catalogue :
'                                     - feuille de style dédiée pour l'impression : plus de clarté, évite de vider une cartouche d'encre
'                                     - compatibilité avec toutes les résolutions d'écrans : masque les colonnes moins importantes dynamiquement
'                                     - affichage de la page uniquement lorsque son contenu est entièrement chargé et à jour : plus agréable à l'oeil
'                                     - bouton de retour en haut de page plus intelligent : accès rapide au haut ou au bas de page suivant la situation
'           - v3.5 : 05/11/2017 : Amélioration et correctifs de l'interface graphique VBA de création / édition de fiche du catalogue Excel :
'                                     - regroupement des boutons de formatage Markdown des champs "Problème" et "Solution" sur une seule ligne
'                                     - l'appui sur un bouton de formatage Markdown ne fait plus perdre le focus sur la zone de saisie en cours
'                                     - correction du double clic dans la liste de pièces jointes qui était actif même s'il n'y en avait aucune
'                                     - possibilité d'ouvrir la table des caractères de Windows depuis la fenêtre de création / édition de fiche
'                                     - ajout d'un bouton permettant d'ouvrir le catalogue HTML depuis la fenêtre de création / édition de fiche
'                                     - agrandissement de la fenêtre de création / édition de fiche : compatible avec une résolution de 720p au minimum
'                                     - remise à zéro de la position du curseur dans les champs "Problème", "Solution" et "Code" à l'édition d'une fiche
'                                     - possibilité d'agrandir le champ "Problème", "Solution" ou "Code" pour qu'il remplisse la majorité de l'interface
'                                     - les boutons de réinitialisation / insertion du presse papier fonctionnent dans les champs "Problème et "Solution"
'                                     - agrandissement automatique de la hauteur des champs "Problème", "Solution" et "Code" lorsqu'ils possèdent le focus
'                                     - affichage du nombre de pièces jointes de la fiche sur le bouton (le trombone) permettant de les afficher / masquer
'                                     - le 1er affichage de la fenêtre de création / édition de fiche du catalogue Excel est désormais deux fois plus rapide
'                                 Ajout de différents boutons de formatage Markdown et HTML pour les champs "Problème" et "Solution" :
'                                     - bouton "Citation" : met en relief / retrait une ligne ou un paragraphe de texte donné
'                                     - bouton "Ligne de séparation" : comme son nom l'indique, ligne de séparation horizontale
'                                     - bouton "Tableau" : insertion d'un tableau dont le nombre de lignes et de colonnes est configurable
'                                     - bouton "Entités HTML" : convertit les caractères spéciaux "<", ">", "'" et """ en entités HTML et inversement
'                                     - bouton "Lien" : insertion d'un lien vers une fiche du catalogue (conversion en relatif), un site internet, ...
'                                     - bouton "Pièce jointe" : insertion d'une image ou d'un document issu de la liste des pièces jointes de la fiche
'                                     - bouton "Statistiques" : affichage du nombre de mots, lignes et caractères du champ actif (inutile mais indispensable)
'
'Dépendances VBA :
'              - Microsoft Scripting Runtime
'              - Visual Basic for Applications
'              - Microsoft Forms XX Object Library
'              - Microsoft Excel XX Object Library
'              - Microsoft Office XX Object Library
'              - Microsoft VBScript Regular Expressions 5.5
'
'==============================================================================================================================================================
