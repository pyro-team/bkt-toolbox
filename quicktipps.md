# Quick-Tipps
{:.no_toc}

<style>
	#toc {
		list-style-type: upper-roman;
	}
	.quicktipps {
		counter-reset: section;
	}
	.quicktipp {
		display:inline-block;
	}
	h3::before {
		counter-increment: section;
		content: counter(section) ". ";
	}
	video {
		width: 100%;
	}

	@media screen and (min-width: 42em) {
		#toc li > ol {
			column-count: 2;
      margin-bottom:10px;
		}
		.quicktipps {
			column-count: 2;
		}
	}
</style>

Auf diese Seite befinden sich diverse kleine Animationen, die Tipps im Umgang mit PowerPoint und der BKT geben. Da die Animationen zu unterschiedlichen Zeitpunkten und mit verschiedenen BKT-Versionen aufgezeichnet wurden, kann es im Vergleich zur aktuellsten Version zu kleineren Abweichungen kommen.

<ol id="toc">
{% assign cats_sorted = site.data.tipps | sort %}
{% for cat_hash in cats_sorted %}
{% assign cat = cat_hash[1] %}
  <li><a href="#{{ cat.name | slugify }}">{{ cat.name }}</a><ol>
  {% for tipp in cat.tipps %}
    <li><a href="#{{ tipp.id }}">{{ tipp.name }}</a></li>
  {% endfor %}
  </ol></li>
{% endfor %}
</ol>

<!-- 
1. [Spezielle BKT-Shapes](#spezielle-bkt-shapes)
   1. Agenda einfügen und aktualisieren
   1. Shape-Tabelle anlegen
   1. Aktualisierbare Folien-Thumbnails anlegen
1. [Weitere BKT-Funktionen](#weitere-bkt-funktionen)
   1. Eigene Shape-Library anlegen
   1. Chart-Library mit Folienmastern
   1. Shapes gezielt auswählen
   1. Icons mit Icon-Fonts
   1. Shape-Statistiken anzeigen
   1. Folien-Notizen anlegen und löschen
   1. Toolbar-Themes und Einstellungen
1. *more to come...* -->

---


{% for cat_hash in cats_sorted %}
{% assign cat = cat_hash[1] %}
  <h2 id="{{ cat.name | slugify }}">{{ cat.name }}</h2>

<section class="quicktipps">
  {% for tipp in cat.tipps %}
  <div class="quicktipp">
    <h3 id="{{ tipp.id }}">{{ tipp.name }}</h3>
    <video loop muted autoplay playsinline controls>
      <source src="documentation/quicktipps/{{ tipp.id }}.webm" type="video/webm">
      <source src="documentation/quicktipps/{{ tipp.id }}.mp4" type="video/mp4">
    </video>
    {{ tipp.description | markdownify }}
    {% if tipp.note %}
      {{ tipp.note | markdownify }}
    {% endif %}
  </div>
  {% endfor %}
</section>
{% endfor %}
