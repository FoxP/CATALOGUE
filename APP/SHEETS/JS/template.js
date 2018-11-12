/* ----------------------------------------------------------------------------------------------------------------------------------------------- */
/* Fonctionnalités JavaScript des fiches HTML du catalogue                                                                                         */
/* ----------------------------------------------------------------------------------------------------------------------------------------------- */

/* ------------------------------------------------------------------------------------ */
/* Fonctions génériques 															    */
/* ------------------------------------------------------------------------------------ */

/* Fonction permettant de charger un .js dynamiquement */
/* Utile pour ne pas surcharger une page avec un .js peu utilisé */

function getScript(source, callback) {
    var script = document.createElement('script');
    var prior = document.getElementsByTagName('script')[0];
    script.async = 1;
    prior.parentNode.insertBefore(script, prior);

    script.onload = script.onreadystatechange = function( _, isAbort ) {
        if(isAbort || !script.readyState || /loaded|complete/.test(script.readyState) ) {
            script.onload = script.onreadystatechange = null;
            script = undefined;

            if(!isAbort) { if(callback) callback(); }
        }
    };

    script.src = source;
}

/* Fonction permettant de copier le contenu d'un élément dans le presse papier */

function copyToClipboard(contentHolder) {
	var range = document.createRange(),
		selection = window.getSelection();
	selection.removeAllRanges();
	range.selectNodeContents(contentHolder);
	selection.addRange(range);
	document.execCommand('copy');
	selection.removeAllRanges();
};

/* Fonction permettant de détecter si le navigateur est Internet Explorer, Edge ou autre chose */
/* var version = detectIE();
if (version === false) {
	//Autre navigateur
} else if (version >= 12) {
	//Edge
} else {
	//Internet Explorer
} */

function detectIE() {
	var ua = window.navigator.userAgent;

	var msie = ua.indexOf('MSIE ');
	if (msie > 0) {
		// IE 10 ou inférieur => retourne le numéro de version
		return parseInt(ua.substring(msie + 5, ua.indexOf('.', msie)), 10);
	}

	var trident = ua.indexOf('Trident/');
	if (trident > 0) {
		// IE 11 => retourne le numéro de version
		var rv = ua.indexOf('rv:');
		return parseInt(ua.substring(rv + 3, ua.indexOf('.', rv)), 10);
	}

	var edge = ua.indexOf('Edge/');
	if (edge > 0) {
		// Edge (IE 12+) => retourne le numéro de version
		return parseInt(ua.substring(edge + 5, ua.indexOf('.', edge)), 10);
	}

	// Autre navigateur
	return false;
}

/* ------------------------------------------------------------------------------------ */
/* Fonctions spécifiques aux fiches HTML du catalogue									*/
/* ------------------------------------------------------------------------------------ */

/* Lorsque la page est entièrement chargée */
	
$(document).ready(function() {
	/* Création d'une galerie photo */
	/* Dépendance : jQuery fancyBox
					jQuery */
	
	$("a.galery").fancybox({
		'transitionIn'	:	'elastic',
		'transitionOut'	:	'elastic',
		'speedIn'		:	200, 
		'speedOut'		:	200, 
		'overlayShow'	:	false,
	});
	
	/* Barre de miniatures de la galerie photo */
	/* Dépendance : Thumbnail helper for fancyBox */
	$("a.galery").fancybox({
		padding		: 2,
		margin		: 20,
		openSpeed	: 150,
		closeSpeed	: 50,
		openEffect	: 'elastic',
		closeEffect : 'fade',
		tpl : {
			closeBtn : '<a title="Fermer" class="fancybox-item fancybox-close" href="javascript:;"></a>',
			next     : '<a title="Suivant" class="fancybox-nav fancybox-next" href="javascript:;"><span></span></a>',
			prev     : '<a title="Précédent" class="fancybox-nav fancybox-prev" href="javascript:;"><span></span></a>'
		},
		prevEffect	: 'none',
		nextEffect	: 'none',
		helpers	: {
			thumbs	: {
				width	: 50,
				height	: 50
			}
		}
	});
	
	/* Coloration syntaxique du champ code */
	/* Dépendance : highlight.js */
	
	$('#code pre code').each(function(i, block) {
		hljs.highlightBlock(block);
	});
	
	/* Conversion Markdown --> HTML pour les champs problème et solution */
	
	/* Librairie "markdown-it" */
	/* Dépendance : markdown-it */
	
	/* var defaults = {
		html:         true,        	// Enable HTML tags in source
		xhtmlOut:     false,        // Use '/' to close single tags (<br />)
		breaks:       true,        	// Convert '\n' in paragraphs into <br>
		linkify:      true,         // autoconvert URL-like texts to links
		typographer:  false,        // Enable smartypants and other sweet transforms
	};
	
	if (document.getElementById('problem').getElementsByTagName("span")[0] != null) {
		var md1 = window.markdownit(defaults);
		var result1 = md1.render(document.getElementById('problem').getElementsByTagName("span")[0].innerHTML);
		document.getElementById('problem').getElementsByTagName("span")[0].innerHTML = result1;
	}
	if (document.getElementById('solution').getElementsByTagName("span")[0] != null) {
		var md2 = window.markdownit(defaults);
		var result2 = md2.render(document.getElementById('solution').getElementsByTagName("span")[0].innerHTML);
		document.getElementById('solution').getElementsByTagName("span")[0].innerHTML = result2;
	} */
	
	/* Librairie "marked" */
	/* Dépendance : marked */
	
	marked.setOptions({
		breaks: true	// Sauts de lignes
		// highlight: function (code) {
			// return hljs.highlightAuto(code).value;	 // Coloration syntaxique automatique du code
		// }
	});

	if (document.getElementById('problem').getElementsByTagName("span")[0] != null) {	// Champ "Problème"
		document.getElementById('problem').getElementsByTagName("span")[0].innerHTML = marked(document.getElementById('problem').getElementsByTagName("span")[0].innerHTML);
	}
	if (document.getElementById('solution').getElementsByTagName("span")[0] != null) {	// Champ "Solution"
		document.getElementById('solution').getElementsByTagName("span")[0].innerHTML = marked(document.getElementById('solution').getElementsByTagName("span")[0].innerHTML);
	}
	
	/* Easter egg : en décembre, fait tomber de la neige */
	/* Dépendance : WP Super Snow */
	
	getScript('./JS/snow.min.js', function(){
		var d = new Date();
		var n = d.getMonth();
		if (n == 11) {	// Si mois de décembre
			$('body').wpSuperSnow({
				flakes: ['./CSS/snow/IMG/snowflake.png','./CSS/snow/IMG/snowball.png'],
				totalFlakes: '100',
				zIndex: '999999',
				maxSize: '20',
				maxDuration: '50',
				useFlakeTrans: false
			});
		}
	});
	
	/* Nuage de mots clés */
	/* Dépendances : D3.js, Word Cloud Layout */
	
	var words = [""]
	$("#tags span").each(function (k, v) {
		if (($("#tags span:nth-child(4)").text() === $("#tags span:nth-child(2)").text()) || ($("#tags span:nth-child(4)").text() === $("#tags span:nth-child(3)").text()) || ($("#tags span:nth-child(3)").text() === $("#tags span:nth-child(2)").text())) {
			if ((v.innerHTML != $("#tags span:nth-child(2)").text()) && (v.innerHTML != $("#tags span:nth-child(3)").text()) && (v.innerHTML != $("#tags span:nth-child(2)").text())) {
				if (words[0] === "") {
					words[0] = v.innerHTML;
				} else {
					words[0] = words[0] + "|" + v.innerHTML;
				}
			}
		} else {
			if (words[0] === "") {
				words[0] = v.innerHTML;
			} else {
				words[0] = words[0] + "|" + v.innerHTML;
			}
		}
	});
	
	function getWords(i) {
		return words[i]
				.replace(/[!\,:;\?]/g, '')
				.split('|')
				.map(function(d) {
					return {text: d, size: 10 + Math.random() * 60};
				})
	}

	function showNewWords(vis, i) {
		i = i || 0;

		vis.update(getWords(i))
	}

	if (typeof(d3) != "undefined") {
		var myWordCloud = wordCloud('body');
		showNewWords(myWordCloud);
	}
		
	/* Redimensionnement du SVG à son contenu */
	/* Par défaut, la zone de dessin du SVG est plus grande que le dessin lui-même */
	
	if (typeof(d3) != "undefined") {
		setTimeout(function(){
			var svg = document.getElementsByTagName("svg")[0];
			var bbox = svg.getBBox();

			svg.setAttribute("viewBox", (bbox.x-10)+" "+(bbox.y-10)+" "+(bbox.width+20)+" "+(bbox.height+20));
			svg.setAttribute("width", (bbox.width+20)  + "px");
			svg.setAttribute("height",(bbox.height+20) + "px");
		},100);
	}
	
	/* Hotfix : si la fiche est ouverte dans un nouvel onglet en arrière plan, getBBox renvoie 0 */
	/* On réapplique donc la fonction ci-dessus dès que la page / l'onglet a le focus dans le navigateur */
	
	window.onfocus = function () { 
		if (typeof(d3) != "undefined") {
			setTimeout(function(){
				var svg = document.getElementsByTagName("svg")[0];
				var bbox = svg.getBBox();

				svg.setAttribute("viewBox", (bbox.x-10)+" "+(bbox.y-10)+" "+(bbox.width+20)+" "+(bbox.height+20));
				svg.setAttribute("width", (bbox.width+20)  + "px");
				svg.setAttribute("height",(bbox.height+20) + "px");
			},100);
		}
	}; 
	
	/* Recherche d'un mot clé issu du nuage de mots clés en bas de page dans le catalogue */
	/* On convertit le mot clé en base64, et on le passe dans l'ancre de l'url du catalogue */
	
	$("text").on('mousedown', function (e) {
		/* Si clic molette de la souris */
		if (e.which === 2) {
			/* Empêche l'activation du scroll */
			e.preventDefault();
			/* Ouverture du catalogue HTML dans un nouvel onglet */
			if ($("#tags span:nth-child(1)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("status=" + $(this).text()));
			} else if ($("#tags span:nth-child(4)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("type=" + $(this).text()));
			} else if ($("#tags span:nth-child(2)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("soft=" + $(this).text()));
			} else if ($("#tags span:nth-child(3)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("lang=" + $(this).text()));
			} else {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa($(this).text()));
			}
		/* Si clic gauche de la souris */
		} else if (e.which === 1) {
			/* Ouverture du catalogue HTML dans l'onglet en cours */
			if ($("#tags span:nth-child(1)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("status=" + $(this).text()), "_self");
			} else if ($("#tags span:nth-child(4)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("type=" + $(this).text()), "_self");
			} else if ($("#tags span:nth-child(2)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("soft=" + $(this).text()), "_self");
			} else if ($("#tags span:nth-child(3)").text() === $(this).text()) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("lang=" + $(this).text()), "_self");
			} else {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa($(this).text()), "_self");
			}
		}
	});
	
	/* Recherche d'un mot clé issu de la liste de mots clés en haut de page dans le catalogue */
	/* On convertit le mot clé en base64, et on le passe dans l'ancre de l'url du catalogue */
	
	$("#tags span:nth-child(n+1)").on('mousedown', function (e) {
		/* Si clic molette de la souris */
		if (e.which === 2) {
			/* Empêche l'activation du scroll */
			e.preventDefault();
			/* Ouverture du catalogue HTML dans un nouvel onglet */
			if ($(this).is('#tags span:nth-child(1)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("status=" + $(this).text()));
			} else if ($(this).is('#tags span:nth-child(4)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("type=" + $(this).text()));
			} else if ($(this).is('#tags span:nth-child(2)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("soft=" + $(this).text()));
			} else if ($(this).is('#tags span:nth-child(3)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("lang=" + $(this).text()));
			} else {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa($(this).text()));
			}
		/* Si clic gauche de la souris */
		} else if (e.which === 1) {
			/* Ouverture du catalogue HTML dans l'onglet en cours */
			if ($(this).is('#tags span:nth-child(1)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("status=" + $(this).text()), "_self");
			} else if ($(this).is('#tags span:nth-child(4)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("type=" + $(this).text()), "_self");
			} else if ($(this).is('#tags span:nth-child(2)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("soft=" + $(this).text()), "_self");
			} else if ($(this).is('#tags span:nth-child(3)')) {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa("lang=" + $(this).text()), "_self");
			} else {
				window.open("../CATALOGUE/CATALOGUE.html#" + btoa($(this).text()), "_self");
			}
		}
	});
	
	/* Tooltips : légendes des boutons du menu de la fiche */
	/* Dépendance : Tipped */

	var versionTag = document.getElementsByTagName("h1")[0].getElementsByTagName("span")[1];
	if (versionTag != null) {
		var versionTagText = versionTag.innerHTML;
		var urlWithoutVersion = window.location.href;
		urlWithoutVersion = urlWithoutVersion.replace(versionTagText + '.html', '');
		var tippedText = '<ul>';
		if (parseInt(versionTagText) > 1) {
			tippedText = tippedText + '<li>Versions antérieures :</li><ul>';
			var i = parseInt(versionTagText) - 1;
			while (i != 0) {
				tippedText = tippedText + '<li><a href=\'' + urlWithoutVersion + i + '.html\'>Version ' + i + '</a></li>';
				i = i - 1;
			}
			tippedText = tippedText + '</ul>';
		} else {
			tippedText = tippedText + '<li>Aucune version antérieure</li>';
		}
		var linksArray = [];
		if (document.getElementById('problem').getElementsByTagName("span")[0] != null) {	// Champ "Problème"
			var linksProblem = document.getElementById('problem').getElementsByTagName("span")[0].getElementsByTagName("a");
			for(var i=0; i<linksProblem.length; i++) {
				linksArray.push(linksProblem[i]);
			}
		}	
		if (document.getElementById('solution').getElementsByTagName("span")[0] != null) {	// Champ "Solution"
			linksSolution = document.getElementById('solution').getElementsByTagName("span")[0].getElementsByTagName("a");
			for(var i=0; i<linksSolution.length; i++) {
				linksArray.push(linksSolution[i]);
			}
		}
		if (linksArray.length > 0) {
			tippedText = tippedText + '<li>Fiches / liens hypertextes liés :</li><ul>';
			for(var i=0; i<linksArray.length; i++) {
				tippedText = tippedText + '<li><a href=\'' + linksArray[i].href + '\'>' + linksArray[i].innerHTML + '</a></li>';
			}
			tippedText = tippedText + '</ul>';
		} else {
			tippedText = tippedText + '<li>Aucune fiche / lien hypertexte lié</li>';
		}
		tippedText = tippedText	+ '</ul>';
		
		Tipped.create('#versions', tippedText, { position: 'bottomright', showDelay: 0, hideDelay: 0, size: 'large' });
	} else {
		$("#versions").hide();
	}
	
	Tipped.create('#print',{ position: 'bottomright', showDelay: 0, hideDelay: 0, size: 'large' });
	Tipped.create('#mail',{ position: 'bottomright', showDelay: 0, hideDelay: 0, behavior: 'hide', size: 'large' });
	Tipped.create('#folder',{ position: 'bottomright', showDelay: 0, hideDelay: 0, behavior: 'hide', size: 'large' });
	Tipped.create('#clipboard',{ position: 'bottomright', showDelay: 0, hideDelay: 0, behavior: 'hide', size: 'large' });
	Tipped.create('#catalogue',{ position: 'bottomright', showDelay: 0, hideDelay: 0, behavior: 'hide', size: 'large' });
	Tipped.create('#qrcode', "<a id='qrpng'><div id='qr'></div></a>", { position: 'bottomright', showDelay: 0, hideDelay: 0, size: 'large', afterUpdate:
		function(content, element){
			/* Génération du QRCode */
			new QRCode(document.getElementById("qr"), {
				text: document.location.href,
				width: 256,
				height: 256,
				colorDark : "#000000",
				colorLight : "#ffffff",
				correctLevel : QRCode.CorrectLevel.L
			});
			/* Lors du clic sur le QRCode, téléchargement au format PNG */
			/* Détection du navigateur */
			var version = detectIE();
			/* Si ce n'est pas Internet Explorer ou Edge (non compatibles avec l'attribut 'download') */
			if (version === false) {
				document.getElementById("qrpng").href = document.getElementById('qr').getElementsByTagName("canvas")[0].toDataURL();
				document.getElementById("qrpng").download = document.getElementsByTagName("h1")[0].textContent + ".png";
			}
			document.getElementById('qr').getElementsByTagName("img")[0].style.width = '256px';
			document.getElementById('qr').getElementsByTagName("img")[0].style.height = '256px';
		}
	});

	Tipped.create('#tags span:nth-child(1)', "Statut", { position: 'bottom', showDelay: 0, hideDelay: 0, fadeIn: 0, fadeOut: 0, behavior: 'hide', size: 'large' });
	Tipped.create('#tags span:nth-child(2)', "Logiciel", { position: 'bottom', showDelay: 0, hideDelay: 0, fadeIn: 0, fadeOut: 0, behavior: 'hide', size: 'large' });
	Tipped.create('#tags span:nth-child(3)', "Langage", { position: 'bottom', showDelay: 0, hideDelay: 0, fadeIn: 0, fadeOut: 0, behavior: 'hide', size: 'large' });
	Tipped.create('#tags span:nth-child(4)', "Type", { position: 'bottom', showDelay: 0, hideDelay: 0, fadeIn: 0, fadeOut: 0, behavior: 'hide', size: 'large' });
	
});

/* Plier / déplier les champs photos, documents, problème, solution et code lors du clic sur leur titre respectif */

$("#pictures h1, #documents h1, #code h1, #solution h1, #problem h1").click(function (e) {
	if(e.target == this){
		$(this).toggleClass("down");	// Permet d'inverser le sens de la flèche dans le titre
		$header = $(this);
		$content = $header.next();
		$content.slideToggle(200);
	}
});

/* Bouton imprimer la fiche */

$("#print").click(function (e) {
	window.print();
});

/* Bouton ouvrir le catalogue */

$("#catalogue").on('mousedown', function (e) {
	/* Si clic molette de la souris */
	if (e.which === 2) {
		/* Empêche l'activation du scroll */
		e.preventDefault();
		/* Ouverture du catalogue HTML dans un nouvel onglet */
		window.open("../CATALOGUE/CATALOGUE.html");
	/* Si clic gauche de la souris */
	} else if (e.which === 1) {
		/* Ouverture du catalogue HTML dans l'onglet en cours */
		window.open("../CATALOGUE/CATALOGUE.html", "_self");
	}
});

/* Copier le code dans le presse papier */

var contentHolder = document.getElementById('code').getElementsByTagName("code")[0];

/* S'il y a du code */
if (contentHolder != null) {
	/* Au double clic sur le code dans la fiche */
	/* contentHolder.addEventListener("dblclick", function() {
		copyToClipboard(contentHolder);
		if($('#ohsnap').length == 0) {
			$("body").append("<div id='ohsnap'></div>");
		}
		ohSnap('Code copié dans le presse papier !', {color: 'black', icon: 'icon-alert'});
	}, false); */
	
	/* Au clic sur le code dans la fiche */
	contentHolder.addEventListener("click", function() {
		/* Si pas de texte sélectionné */
		var selection = window.getSelection();
		if(selection.toString().length === 0) {
			copyToClipboard(contentHolder);
			if($('#ohsnap').length == 0) {
				$("body").append("<div id='ohsnap'></div>");
			}
			/* Popup de notification */
			/* Dépendance : Oh Snap! */
			ohSnap('Code copié dans le presse papier !', {color: 'black', icon: 'icon-alert'});
		}
	}, false);
	
	/* Bouton copier le code dans le presse papier */
	$("#clipboard").click(function (e) {
		copyToClipboard(contentHolder);
		if($('#ohsnap').length == 0) {
			$("body").append("<div id='ohsnap'></div>");
		}
		ohSnap('Code copié dans le presse papier !', {color: 'black', icon: 'icon-alert'});
	});
/* Sinon */
} else {
	$("#clipboard").hide();		// On masque le bouton
}

/* Bouton pièces jointes de la fiche */

/* S'il n'y a pas d'images et pas de documents */
if (!($("#pictures h1").length || $("#documents h1").length)) {
	$("#folder").hide();	// On masque le bouton
} else {
	$("#folder").on('mousedown', function (e) {
		/* Si clic molette de la souris */
		if (e.which === 2) {
			/* Empêche l'activation du scroll */
			e.preventDefault();
		}
		if ($("#pictures h1").length) {
			window.open($("#pictures h1 a").prop('href'));
		} else {
			window.open($("#documents h1 a").prop('href'));
		}
	});
}

/* Envoyer la fiche par mail */

function sendMail(title) {
	x=window.open("mailto:"+""+'&subject='+encodeURIComponent(title.textContent)+'&body='+encodeURIComponent(document.location.href));
	setTimeout(function(){x.close();},1000);
}

var title = document.getElementsByTagName("h1")[0];

/* Bouton envoyer la fiche par mail */
$("#mail").click(function (e) {
	sendMail(title);
});

/* Au double clic sur le titre de la fiche */
document.getElementsByTagName("h1")[0].addEventListener("dblclick", function() {
	sendMail(title)
}, false);

/* Compte et affiche le nombre de photos */

$("#pictures h1").html($("#pictures h1").html() + " (" + ($("#pictures a").length - 1) + ")");

/* Compte et affiche le nombre de documents */

$("#documents h1").html($("#documents h1").html() + " (" + ($("#documents a").length - 1) + ")");

/* Nuage de mots-clés */
/* Dépendances : D3.js, Word Cloud Layout */

function wordCloud(selector) {

    var fill = d3.scale.category20();

    var svg = d3.select(selector).append("svg")
        .attr("width", 500)
        .attr("height", 500)
		.attr("id", "svg")
        .append("g")
        .attr("transform", "translate(250,250)");

    function draw(words) {
        var cloud = svg.selectAll("g text")
                        .data(words, function(d) { return d.text; })

        cloud.enter()
            .append("text")
            .style("font-family", "Impact")
            .style("fill", function(d, i) { return fill(i); })
            .attr("text-anchor", "middle")
            .attr('font-size', 1)
            .text(function(d) { return d.text; });

        cloud.transition()
			.duration(0) /*Animation d'apparition désactivée */
			.style("font-size", function(d) { return d.size + "px"; })
			.attr("transform", function(d) {
				return "translate(" + [d.x, d.y] + ")rotate(" + d.rotate + ")";
			})
			.style("fill-opacity", 1);

        cloud.exit()
            .transition()
                .duration(200)
                .style('fill-opacity', 1e-6)
                .attr('font-size', 1)
                .remove();
    }

    return {
        update: function(words) {
            d3.layout.cloud().size([500, 500])
                .words(words)
                .padding(5)
                .rotate(function() { return ~~(Math.random() * 2) * 90; })
                .font("Impact")
                .fontSize(function(d) { return d.size; })
                .on("end", draw)
                .start();
        }
    }
}