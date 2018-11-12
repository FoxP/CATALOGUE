Attribute VB_Name = "CHANGELOG"
'==============================================================================================================================================================
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
'Versions :
'           - v1.0 : 02/11/2016 : Code initial
'           - v1.5 : 21/11/2016 : Support du Markdown pour du contenu riche dans les champs "Probl�me" et "Solution"
'                                 Gestion de la maturit� des fiches : Draft, Submitted, Release, Superseded, Obsolete
'           - v2.0 : 14/02/2017 : Launcher VB.NET :
'                                     - lancement des catalogues HTML et XLS, et cr�ation de fiches
'                                     - minimisation dans la zone de notifications, menu du clic droit
'                                     - notification si nouvelle(s) fiche(s) et modification du catalogue
'                                     - aper�u du nombre de fiches du catalogue (MAJ toutes les 10 minutes)
'                                     - ajout automatique du launcher au d�marrage de la session utilisateur
'                                     - gestion de param�tres stock�s dans un fichier de configuration au format .ini
'                                 Easter egg de No�l : durant le mois de d�cembre, fait tomber de la neige sur les fiches
'                                 Catalogue Excel d�sormais compatible multi-utilisateurs, gestion des fiches au clic droit
'                                 Robustification / refactoring du code, correction de multiples bugs, quelques optimisations
'                                 Support des pi�ces jointes (images / documents de tous types / tailles), galerie photo automatique
'                                 Possibilit� de re-g�n�rer manuellement une copie compl�te du catalogue HTML depuis le catalogue Excel
'                                 Cr�ation d'un certificat permettant l'utilisation du catalogue sans avertissements de s�curit� d'Excel
'                                 Correcteur orthographique et fonction de pr�visualisation lors de la cr�ation ou l'�dition d'une fiche
'                                 Ajout de statistiques dans une feuille d�di�e du catalogue Excel : pourcentage de chaque cat�gorie, etc
'                                     - n�cessite cependant de d�-partager temporairement le classeur Excel pour effectuer une mise � jour
'                                 Enrichissement du template HTML des fiches :
'                                     - responsive design, s'adapte � tous types d'�crans
'                                     - comptage du nombre total d'images et de documents
'                                     - cr�ation d'une feuille de style d�di�e pour l'impression
'                                     - nuage de mots cl�s permettant de rechercher des fiches li�es
'                                     - affichage / masquage rapide des diff�rentes sections de la fiche
'                                     - refonte graphique : tooltips, notifications, effets de survol, ombres, ...
'                                     - meilleur affichage des tableaux, images et portions de code ins�r�s en Markdown
'                                     - �diteur Markdown WYSIWYG (What You See Is What You Get) dans la fiche d'aide au language
'           - v2.1 : 02/04/2017 : A la suppression d'une fiche, v�rification si l'op�ration a r�ussi ou non
'                                 Suppression des caract�res "," et ";" du champ des mots cl�s (au cas o�...)
'                                 Boutons d'ouverture du dossier des pi�ces jointes, et tri par ordre alphab�tique
'           - v3.0 : 26/08/2017 : Ouverture des fiches dans le navigateur d�fini par d�faut au lieu d'Internet Explorer
'                                 Mise � jour automatique du contenu de la feuille Excel de statistiques � son ouverture
'                                 Si le classeur Excel est partag�, la feuille Excel de statistiques est masqu�e automatiquement
'                                     - un classeur partag� d�sactive la mise � jour des tableaux crois�s dynamiques / graphiques
'                                 Re-configuration automatique du catalogue lors de la modification de la feuille Excel d'options
'                                     - plus besoin de fermer puis de r�ouvrir le classeur Excel apr�s modification de ses options
'                                 Possibilit� de fermer la UserForm de cr�ation / �dition de fiche via la touche "Echap" du clavier
'                                 Correction du mauvais positionnement de la UserForm lors de l'ouverture du panel de pi�ces jointes
'                                 Depuis le panel de pi�ces jointes, affichage de l'une d'entre elle via le clic molette de la souris
'                                 Depuis le panel de pi�ces jointes, copie de son code Markdown dans le presse papier via double-clic
'                                 Le menu d�roulant du clic droit pour l'�dition / ... de fiche n'apparait d�sormais que l� o� il le doit
'                                 La fonction de pr�visualisation d'une fiche en cours de cr�ation / �dition affiche d�sormais les images
'                                 Ajout de diff�rents boutons permettant l'ajout rapide de balises Markdown : soulign�, gras, couleur, ...
'                                 Enrichissement du template HTML des fiches :
'                                     - ajout d'un bouton pour revenir directement au catalogue HTML depuis la fiche elle-m�me
'                                     - les caract�res "<" ou ">" ne sont plus affich�s en "&lt;" ou "&gt;" dans les balises code
'                                     - les anciennes versions d'une fiche sont accessibles directement depuis la fiche elle-m�me
'                                     - les fiches li�es ainsi que les hyperliens sont accessibles depuis la fiche via un menu d�di�
'                                     - lors du clic sur le statut, le type, le logiciel, le langage ou bien un mot cl� dans une fiche,
'                                       ouverture du catalogue HTML et pr�sentation des diff�rences fiches ayant la m�me caract�ristique
'                                 Enrichissement du template HTML du catalogue :
'                                     - am�lioration des performances : minification et regroupement des fichiers JavaScript et CSS
'                                     - graphiques de statistiques dynamiques en bas de page : pourcentage de chaque cat�gorie, ...
'                                     - refonte graphique du style CSS : plus clair, plus lisible, s'adapte � tous types d'�crans, ...
'                                     - compression sans perte visuelle des images du template via "quantification" : images 8 bits index�es
'                                 Tests de compatibilit� du catalogue HTML et des fiches sur les principaux navigateurs mobiles Android :
'                                     - Opera : fonctionne parfaitement, affichage strictement correct, tr�s bonnes performances
'                                     - Firefox : fonctionne parfaitement, affichage strictement correct, lenteurs dans le catalogue
'                                     - Google Chrome : fonctionne parfaitement, tailles de polices in�gales, tr�s bonnes performances
'                                     - Gello (navigateur CyanogenMod / LineageOS) : fonctionne parfaitement, tr�s bonnes performances
'           - v3.1 : 25/09/2017 : Refactoring du code VBA du catalogue Excel : mise au propre, relocalisation de chaque proc�dure
'                                 Cr�ation de fiche possible depuis le menu contextuel du clic droit : "Cr�er une nouvelle fiche"
'                                 Possibilit� d'exporter sous forme d'archive une fiche donn�e depuis le menu contextuel du clic droit :
'                                     - embarque la totalit� des d�pendances de la fiche : HTML, CSS, JavaScript, images et pi�ces jointes
'                                     - lors de l'envoi d'une fiche par e-mail, ajout automatique de l'archive en pi�ce jointe si souhait�
'                                     - si le logiciel de compression de donn�es 7-Zip est install�, compression .7z LZMA2 ultra, sinon .zip
'                                 Suppression de fiche plus robuste : gestion du cas o� une pi�ce jointe de fiche est verrouill�e / utilis�e
'                                 Compatibilit� du code VBA du catalogue Excel avec les syst�mes d'exploitation 32 bits (oui, �a existe...)
'                                 Hyperlien automatique lors du partage de l'URL d'une fiche par e-mail : formatage HTML du corps de l'e-email
'                                 Correction du titre de l'e-mail qui �tait en majuscules si partage d'une URL de fiche depuis sa version HTML
'                                 Suppression de la d�pendance VBA Microsoft Outlook XX.X Object Library : plus robuste au changement de version
'                                     - si le logiciel de messagerie Microsoft Outlook n'est pas install�, les fonctions d'e-mails sont masqu�es
'                                 Nettoyage des fonctions accessibles depuis la fen�tre Excel de liste des macros (Alt + F8) pour plus de clart�
'                                 Utilisation de la barre d'�tat d'Excel pour afficher des informations sur la progression des op�rations longues
'                                 La fiche d'aide au langage Markdown peut d�sormais �tre supprim�e (masqu�e, plut�t) des catalogues Excel et HTML
'                                 Enrichissement de la fiche d'aide au langage Markdown :
'                                     - mise � jour des captures d'�cran, meilleure int�gration CSS de l'�diteur WYSIWYG
'                                     - liens vers la sp�cification du langage, ajout de l'article Wikip�dia en pi�ce jointe
'                                     - documentation de l'ajout d'un code sur une ligne, de r�f�rences de liens, de code HTML
'                                 Enrichissement du template HTML des fiches :
'                                     - copie rapide du contenu du champ "Code" dans le presse papier si clic sans s�lection
'                                     - support des formats d'image WebP et SVG (Scalable Vector Graphics) dans la galerie photo
'                                     - g�n�ration d'un QR Code permettant de partager l'URL de la fiche sur des p�riph�riques mobiles
'                                     - mise en avant des diff�rentes sections cliquables par des animations CSS au survol de la souris
'                                     - modifications mineures de la galerie photo : traduction fran�aise, CSS, animations plus rapides
'                                     - ajout d'une tooltip au survol des caract�ristiques de la fiche : statut, logiciel, type, langage
'                                     - correction des ombres port�es des diff�rentes sections de la fiche, et de l'espacement entre eux
'                                     - ajout d'icones de formats de fichiers pour les pi�ces jointes : 280 extensions de fichiers g�r�es
'                                     - uniformisation du clic souris : clic gauche ex�cute dans l'onglet actif, clic droit dans un nouvel onglet
'                                 Enrichissement du template HTML du catalogue :
'                                     - feuille de style d�di�e pour l'impression : plus de clart�, �vite de vider une cartouche d'encre
'                                     - compatibilit� avec toutes les r�solutions d'�crans : masque les colonnes moins importantes dynamiquement
'                                     - affichage de la page uniquement lorsque son contenu est enti�rement charg� et � jour : plus agr�able � l'oeil
'                                     - bouton de retour en haut de page plus intelligent : acc�s rapide au haut ou au bas de page suivant la situation
'           - v3.5 : 05/11/2017 : Am�lioration et correctifs de l'interface graphique VBA de cr�ation / �dition de fiche du catalogue Excel :
'                                     - regroupement des boutons de formatage Markdown des champs "Probl�me" et "Solution" sur une seule ligne
'                                     - l'appui sur un bouton de formatage Markdown ne fait plus perdre le focus sur la zone de saisie en cours
'                                     - correction du double clic dans la liste de pi�ces jointes qui �tait actif m�me s'il n'y en avait aucune
'                                     - possibilit� d'ouvrir la table des caract�res de Windows depuis la fen�tre de cr�ation / �dition de fiche
'                                     - ajout d'un bouton permettant d'ouvrir le catalogue HTML depuis la fen�tre de cr�ation / �dition de fiche
'                                     - agrandissement de la fen�tre de cr�ation / �dition de fiche : compatible avec une r�solution de 720p au minimum
'                                     - remise � z�ro de la position du curseur dans les champs "Probl�me", "Solution" et "Code" � l'�dition d'une fiche
'                                     - possibilit� d'agrandir le champ "Probl�me", "Solution" ou "Code" pour qu'il remplisse la majorit� de l'interface
'                                     - les boutons de r�initialisation / insertion du presse papier fonctionnent dans les champs "Probl�me et "Solution"
'                                     - agrandissement automatique de la hauteur des champs "Probl�me", "Solution" et "Code" lorsqu'ils poss�dent le focus
'                                     - affichage du nombre de pi�ces jointes de la fiche sur le bouton (le trombone) permettant de les afficher / masquer
'                                     - le 1er affichage de la fen�tre de cr�ation / �dition de fiche du catalogue Excel est d�sormais deux fois plus rapide
'                                 Ajout de diff�rents boutons de formatage Markdown et HTML pour les champs "Probl�me" et "Solution" :
'                                     - bouton "Citation" : met en relief / retrait une ligne ou un paragraphe de texte donn�
'                                     - bouton "Ligne de s�paration" : comme son nom l'indique, ligne de s�paration horizontale
'                                     - bouton "Tableau" : insertion d'un tableau dont le nombre de lignes et de colonnes est configurable
'                                     - bouton "Entit�s HTML" : convertit les caract�res sp�ciaux "<", ">", "'" et """ en entit�s HTML et inversement
'                                     - bouton "Lien" : insertion d'un lien vers une fiche du catalogue (conversion en relatif), un site internet, ...
'                                     - bouton "Pi�ce jointe" : insertion d'une image ou d'un document issu de la liste des pi�ces jointes de la fiche
'                                     - bouton "Statistiques" : affichage du nombre de mots, lignes et caract�res du champ actif (inutile mais indispensable)
'
'D�pendances VBA :
'              - Microsoft Scripting Runtime
'              - Visual Basic for Applications
'              - Microsoft Forms XX Object Library
'              - Microsoft Excel XX Object Library
'              - Microsoft Office XX Object Library
'              - Microsoft VBScript Regular Expressions 5.5
'
'==============================================================================================================================================================
