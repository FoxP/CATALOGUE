/* ----------------------------------------------------------------------------------------------------------------------------------------------- */
/*  Fonctionnalités JavaScript du catalogue	HTML					   			   				   			   				   			           */
/* ----------------------------------------------------------------------------------------------------------------------------------------------- */

/* ------------------------------------------------------------------------------------ */
/*  Fonctions génériques                                    							*/
/* ------------------------------------------------------------------------------------ */

	/* Retourne une couleur aléatoire */
	function getRandomColor() {
		var letters = '0123456789ABCDEF';
		var color = '#';
		for (var i = 0; i < 6; i++) {
			color += letters[Math.floor(Math.random() * 16)];
		}
		return color;
	}
	
	/* Retourne la taille d'un dictionaire donné */
	function size_dict(d){c=0; for (i in d) ++c; return c}
	
/* ------------------------------------------------------------------------------------ */
/* Lorsque la page est entièrement chargée (HTML, CSS, JS							    */
/* ------------------------------------------------------------------------------------ */
	
	$(window).load(function() {
		/* Par défaut, le contenu de la page est caché pendant son chargement */
		$("section").show();
	});

/* ------------------------------------------------------------------------------------ */
/* Lorsque le contenu HTML de la page est entièrement chargé							*/
/* ------------------------------------------------------------------------------------ */
	
	$(document).ready(function(){
		
		/* Dépendances : Yet Another DataTables Column Filter
			jQuery DataTables 
			jQuery UI 
			jQuery */
		oTable = $('#catalogueTable').dataTable({
			"bPaginate": false,
			"bJQueryUI": true,
			"autoWidth": false,
			"bStateSave": true,
			"language": {
				"sProcessing":     "Traitement en cours...",
				"sSearch":         "Rechercher&nbsp;: ",
				"sLengthMenu":     "Afficher _MENU_ &eacute;l&eacute;ments",
				"sInfo":           "&nbsp;&nbsp;&nbsp;Affichage de l'&eacute;l&eacute;ment _START_ &agrave; _END_ sur _TOTAL_ &eacute;l&eacute;ments",
				"sInfoEmpty":      "&nbsp;&nbsp;&nbsp;Affichage de l'&eacute;l&eacute;ment 0 &agrave; 0 sur 0 &eacute;l&eacute;ment",
				"sInfoFiltered":   "(filtr&eacute; de _MAX_ &eacute;l&eacute;ments au total)",
				"sInfoPostFix":    "",
				"sLoadingRecords": "Chargement en cours...",
				"sZeroRecords":    "Aucun &eacute;l&eacute;ment &agrave; afficher",
				"sEmptyTable":     "Aucune donn&eacute;e disponible dans le tableau",
				"oPaginate": {
					"sFirst":      "Premier",
					"sPrevious":   "Pr&eacute;c&eacute;dent",
					"sNext":       "Suivant",
					"sLast":       "Dernier"
				},
				"oAria": {
					"sSortAscending":  ": activer pour trier la colonne par ordre croissant",
					"sSortDescending": ": activer pour trier la colonne par ordre d&eacute;croissant"
				}
			}
		}).yadcf([
			{column_number : 0, filter_type: "number", filter_default_label: ""},
			{column_number : 1, filter_type: "text", filter_default_label: ""},
			{column_number : 2, filter_type: "number", filter_default_label: ""},
			{column_number : 3, filter_default_label: ""},
			{column_number : 4, filter_default_label: ""},
			{column_number : 5, filter_default_label: ""},
			{column_number : 6, filter_default_label: ""},
			{column_number : 7, filter_type: "range_date", date_format: "yyyy/mm/dd"},
			{column_number : 8, filter_type: "range_date", date_format: "yyyy/mm/dd"},
			{column_number : 9, filter_default_label: ""},
			{column_number : 10, column_data_type: "html", html_data_type: "text", filter_default_label: "", filter_type: "text"}
		]);
	
	/* Bouton retour en haut de l'écran */
	
	var el = $(window),
	/* Dernière position avant scroll */
    lastY = el.scrollTop();
	
	$(window).scroll(function () {
		var currY = el.scrollTop(),
        /* Détermination de la direction de scroll */
        y = (currY > lastY)&&(lastY > 0) ? 'down' : ((currY === lastY) ? 'none' : 'up');
		if (y == 'up') {
			$('.scrollup').text('˄');
		} else if (y == 'down') {
			$('.scrollup').text('˅');
		}
		/* Dernière position avant scroll = position actuelle */
		lastY = currY;
		
		if ($(this).scrollTop() > 100) {
			$('.scrollup').fadeIn();
		} else {
			$('.scrollup').fadeOut();
		}
	});
	
	$('.scrollup').click(function () {
		if ($(this).text() == '˄') {
			$("html, body").animate({
				scrollTop: 0
			}, 600);
		} else {
			$("html, body").animate({
			scrollTop: $(document).height()
		}, 600);
		}
		return false;
	});
	
	/* Lecture de l'ancre de l'url du catalogue pour recherche depuis une fiche */
	/* Ancre encodée en base64 */
	
	var hash = atob(window.location.hash.substr(1));

	/* Statut de la fiche */
	if (hash.indexOf("status=") != -1) {
		document.getElementById("yadcf-filter--catalogueTable-3").value = hash.replace("status=","");
		setTimeout(function(){document.getElementById("yadcf-filter--catalogueTable-3").onchange();},100);
	/* Type de la fiche */
	} else if (hash.indexOf("type=") != -1) {
		document.getElementById("yadcf-filter--catalogueTable-6").value = hash.replace("type=","");
		setTimeout(function(){document.getElementById("yadcf-filter--catalogueTable-6").onchange();},100);
	/* Langage de la fiche */
	} else if (hash.indexOf("lang=") != -1) {
		document.getElementById("yadcf-filter--catalogueTable-5").value = hash.replace("lang=","");
		setTimeout(function(){document.getElementById("yadcf-filter--catalogueTable-5").onchange();},100);
	/* Logiciel de la fiche */
	} else if (hash.indexOf("soft=") != -1) {
		document.getElementById("yadcf-filter--catalogueTable-4").value = hash.replace("soft=","");
		setTimeout(function(){document.getElementById("yadcf-filter--catalogueTable-4").onchange();},100);
	/* Mot clé de la fiche */
	} else {
		if (hash != "") {
			document.getElementById("yadcf-filter--catalogueTable-10").value = hash;
			setTimeout(function(){yadcf.textKeyUP("keyup", '-catalogueTable', 10);},100);
		}
	}
	
	/* Création de graphiques statistiques */
	
	/* Dépendances : Chart.js
	chartjs-plugin-deferred 
	Chart.PieceLabel
	jQuery */
	function getStatsFromTable() {
	
		var SoftwaresDict = {};		/* Liste de logiciels et le nombre de fiches associé */
		var UsersDict = {};			/* Liste d'utilisateurs et le nombre de fiches associé */
		var LanguagesDict = {};		/* Liste de langages et le nombre de fiches associé */
		var TypesDict = {};			/* Liste de types et le nombre de fiches associé */
		var StatusDict = {};		/* Liste de status et le nombre de fiches associé */
		
		var tmpSoftware, tmpUser, tmpLanguage, tmpType, tmpStatus
		
		var tableBody = $('#catalogueTable > tbody > tr').each(function() {
			$this = $(this);
			//Logiciel
			tmpSoftware = $this.find("td:eq(4)").html()
			if (SoftwaresDict[tmpSoftware] === undefined) {
				SoftwaresDict[tmpSoftware] = 1;
			} else {
				SoftwaresDict[tmpSoftware] = SoftwaresDict[tmpSoftware] + 1;
			}
			//Langage
			tmpLanguage = $this.find("td:eq(5)").html()
			if (LanguagesDict[tmpLanguage] === undefined) {
				LanguagesDict[tmpLanguage] = 1;
			} else {
				LanguagesDict[tmpLanguage] = LanguagesDict[tmpLanguage] + 1;
			}
			//Utilisateur
			tmpUser = $this.find("td:eq(9)").html()
			if (UsersDict[tmpUser] === undefined) {
				UsersDict[tmpUser] = 1;
			} else {
				UsersDict[tmpUser] = UsersDict[tmpUser] + 1;
			}
			//Type
			tmpType = $this.find("td:eq(6)").html()
			if (TypesDict[tmpType] === undefined) {
				TypesDict[tmpType] = 1;
			} else {
				TypesDict[tmpType] = TypesDict[tmpType] + 1;
			}
			//Statut
			tmpStatus = $this.find("td:eq(3)").html()
			if (StatusDict[tmpStatus] === undefined) {
				StatusDict[tmpStatus] = 1;
			} else {
				StatusDict[tmpStatus] = StatusDict[tmpStatus] + 1;
			}
		});
		
		//Langages
		
		var array_keys = new Array();
		var array_values = new Array();
		var array_colors = new Array();

		for (var key in LanguagesDict) {
			array_keys.push(key);
			array_values.push(LanguagesDict[key]);
			array_colors.push(getRandomColor());
		}
		
		var config = {
			type: 'pie',
			data: {
				datasets: [{
					label: 'Langages',
					data: array_values,
					backgroundColor: array_colors
				}],
				labels: array_keys
			},
			options: {
				title: {
				display: true,
				fontSize: 13,
				fontColor: "#444",
				text: 'Langages (' + size_dict(LanguagesDict) + ') :' 
				},
				responsive: true,
				tooltips: {
					callbacks: {
						label: function(tooltipItem, data) {
							var allData = data.datasets[tooltipItem.datasetIndex].data;
							var tooltipLabel = data.labels[tooltipItem.index];
							var tooltipData = allData[tooltipItem.index];
							var total = 0;
							for (var i in allData) {
								total += allData[i];
							}
							var tooltipPercentage = Math.round((tooltipData / total) * 100);
							return tooltipLabel + ' : ' + tooltipData + ' (' + tooltipPercentage + '%)';
						}
					}
				},
				legend: {
					labels: {
						fontSize: 13,
						fontColor: "#444"
					},
				},
				pieceLabel: {
					mode: 'label',
					fontSize: 13,
					fontFamily: '"Helvetica Neue", Helvetica, Arial, sans-serif'
				}
			}
		};
		
		/* Suppression du canvas avant mise à jour du graph */
		var chartContainer = document.getElementById("languagesChart").parentNode;
		chartContainer.innerHTML = '&nbsp;';
		chartContainer.innerHTML = '<canvas id="languagesChart"></canvas>';
		
		/* Création du graph */
		var ctx = document.getElementById("languagesChart").getContext("2d");
		window.myPie = new Chart(ctx, config);
		
		//Logiciels
		
		array_keys = new Array();
		array_values = new Array();
		array_colors = new Array();

		for (var key in SoftwaresDict) {
			array_keys.push(key);
			array_values.push(SoftwaresDict[key]);
			array_colors.push(getRandomColor());
		}
		
		config = {
			type: 'pie',
			data: {
				datasets: [{
					label: 'Logiciels',
					data: array_values,
					backgroundColor: array_colors
				}],
				labels: array_keys
			},
			options: {
				title: {
				display: true,
				fontSize: 13,
				fontColor: "#444",
				text: 'Logiciels (' + size_dict(SoftwaresDict) + ') :' 
				},
				responsive: true,
				tooltips: {
					callbacks: {
						label: function(tooltipItem, data) {
							var allData = data.datasets[tooltipItem.datasetIndex].data;
							var tooltipLabel = data.labels[tooltipItem.index];
							var tooltipData = allData[tooltipItem.index];
							var total = 0;
							for (var i in allData) {
								total += allData[i];
							}
							var tooltipPercentage = Math.round((tooltipData / total) * 100);
							return tooltipLabel + ' : ' + tooltipData + ' (' + tooltipPercentage + '%)';
						}
					}
				},
				legend: {
					labels: {
						fontSize: 13,
						fontColor: "#444"
					},
				},
				pieceLabel: {
					mode: 'label',
					fontSize: 13,
					fontFamily: '"Helvetica Neue", Helvetica, Arial, sans-serif'
				}
			}
		};
		
		/* Suppression du canvas avant mise à jour du graph */
		var chartContainer = document.getElementById("softwaresChart").parentNode;
		chartContainer.innerHTML = '&nbsp;';
		chartContainer.innerHTML = '<canvas id="softwaresChart"></canvas>';
		
		/* Création du graph */
		ctx = document.getElementById("softwaresChart").getContext("2d");
		window.myPie = new Chart(ctx, config);
		
		//Types
		
		array_keys = new Array();
		array_values = new Array();
		array_colors = new Array();

		for (var key in TypesDict) {
			array_keys.push(key);
			array_values.push(TypesDict[key]);
			array_colors.push(getRandomColor());
		}
		
		config = {
			type: 'pie',
			data: {
				datasets: [{
					label: 'Types',
					data: array_values,
					backgroundColor: array_colors
				}],
				labels: array_keys
			},
			options: {
				title: {
				display: true,
				fontSize: 13,
				fontColor: "#444",
				text: 'Types (' + size_dict(TypesDict) + ') :' 
				},
				responsive: true,
				tooltips: {
					callbacks: {
						label: function(tooltipItem, data) {
							var allData = data.datasets[tooltipItem.datasetIndex].data;
							var tooltipLabel = data.labels[tooltipItem.index];
							var tooltipData = allData[tooltipItem.index];
							var total = 0;
							for (var i in allData) {
								total += allData[i];
							}
							var tooltipPercentage = Math.round((tooltipData / total) * 100);
							return tooltipLabel + ' : ' + tooltipData + ' (' + tooltipPercentage + '%)';
						}
					}
				},
				legend: {
					labels: {
						fontSize: 13,
						fontColor: "#444"
					},
				},
				pieceLabel: {
					mode: 'label',
					fontSize: 13,
					fontFamily: '"Helvetica Neue", Helvetica, Arial, sans-serif'
				}
			}
		};
		
		/* Suppression du canvas avant mise à jour du graph */
		var chartContainer = document.getElementById("typesChart").parentNode;
		chartContainer.innerHTML = '&nbsp;';
		chartContainer.innerHTML = '<canvas id="typesChart"></canvas>';
		
		/* Création du graph */
		ctx = document.getElementById("typesChart").getContext("2d");
		window.myPie = new Chart(ctx, config);
		
		//Utilisateurs
		
		array_keys = new Array();
		array_values = new Array();
		array_colors = new Array();

		for (var key in UsersDict) {
			array_keys.push(key);
			array_values.push(UsersDict[key]);
			array_colors.push(getRandomColor());
		}
		
		config = {
			type: 'pie',
			data: {
				datasets: [{
					label: 'Utilisateurs',
					data: array_values,
					backgroundColor: array_colors
				}],
				labels: array_keys
			},
			options: {
				title: {
				display: true,
				fontSize: 13,
				fontColor: "#444",
				text: 'Utilisateurs (' + size_dict(UsersDict) + ') :' 
				},
				responsive: true,
				tooltips: {
					callbacks: {
						label: function(tooltipItem, data) {
							var allData = data.datasets[tooltipItem.datasetIndex].data;
							var tooltipLabel = data.labels[tooltipItem.index];
							var tooltipData = allData[tooltipItem.index];
							var total = 0;
							for (var i in allData) {
								total += allData[i];
							}
							var tooltipPercentage = Math.round((tooltipData / total) * 100);
							return tooltipLabel + ' : ' + tooltipData + ' (' + tooltipPercentage + '%)';
						}
					}
				},
				legend: {
					labels: {
						fontSize: 13,
						fontColor: "#444"
					},
				},
				pieceLabel: {
					mode: 'label',
					fontSize: 13,
					fontFamily: '"Helvetica Neue", Helvetica, Arial, sans-serif'
				}
			}
		};
		
		/* Suppression du canvas avant mise à jour du graph */
		var chartContainer = document.getElementById("usersChart").parentNode;
		chartContainer.innerHTML = '&nbsp;';
		chartContainer.innerHTML = '<canvas id="usersChart"></canvas>';
		
		/* Création du graph */
		ctx = document.getElementById("usersChart").getContext("2d");
		window.myPie = new Chart(ctx, config);
		
		//Statuts
		
		array_keys = new Array();
		array_values = new Array();
		array_colors = new Array();

		for (var key in StatusDict) {
			array_keys.push(key);
			array_values.push(StatusDict[key]);
			array_colors.push(getRandomColor());
		}
		
		config = {
			type: 'pie',
			data: {
				datasets: [{
					label: 'Statuts',
					data: array_values,
					backgroundColor: array_colors
				}],
				labels: array_keys
			},
			options: {
				title: {
				display: true,
				fontSize: 13,
				fontColor: "#444",
				text: 'Statuts (' + size_dict(StatusDict) + ') :'
				},
				responsive: true,
				tooltips: {
					callbacks: {
						label: function(tooltipItem, data) {
							var allData = data.datasets[tooltipItem.datasetIndex].data;
							var tooltipLabel = data.labels[tooltipItem.index];
							var tooltipData = allData[tooltipItem.index];
							var total = 0;
							for (var i in allData) {
								total += allData[i];
							}
							var tooltipPercentage = Math.round((tooltipData / total) * 100);
							return tooltipLabel + ' : ' + tooltipData + ' (' + tooltipPercentage + '%)';
						}
					}
				},
				legend: {
					labels: {
						fontSize: 13,
						fontColor: "#444"
					},
				},
				pieceLabel: {
					mode: 'label',
					fontSize: 13,
					fontFamily: '"Helvetica Neue", Helvetica, Arial, sans-serif'
				}
			}
		};
		
		/* Suppression du canvas avant mise à jour du graph */
		var chartContainer = document.getElementById("statusChart").parentNode;
		chartContainer.innerHTML = '&nbsp;';
		chartContainer.innerHTML = '<canvas id="statusChart"></canvas>';
		
		/* Création du graph */
		ctx = document.getElementById("statusChart").getContext("2d");
		window.myPie = new Chart(ctx, config);
		
	}
	
	getStatsFromTable()		/* A l'ouverture de la page */
	
	$('#catalogueTable').on('draw.dt',  function () {getStatsFromTable()})	/* Lorsque le tableau est trié */
	
});