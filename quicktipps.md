# Quick-Tipps
{:.no_toc}

<style>
	#toc {
		list-style-type: upper-roman;
	}
	#toc li > ol {
		column-count: 2;
	}
	.quicktipps {
		column-count: 2;
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
</style>

Auf diese Seite befinden sich diverse kleine Animationen, die Tipps im Umgang mit PowerPoint und der BKT geben. Da die Animationen zu unterschiedlichen Zeitpunkten und mit verschiedenen BKT-Versionen aufgezeichnet wurden, kann es im Vergleich zur aktuellsten Version zu kleineren Abweichungen kommen.

<ol id="toc">
{% assign cats_sorted = site.data.tipps | sort %}
{% for cat_hash in cats_sorted %}
{% assign cat = cat_hash[1] %}
  <li><a href="#{{ cat.name | slugify }}">{{ cat.name }}</a></li>
  <ol>
  {% for tipp in cat.tipps %}
    <li><a href="#{{ tipp.id }}">{{ tipp.name }}</a></li>
  {% endfor %}
  </ol>
{% endfor %}
</ol>

<!-- 
1. [Shape-Inhalte verändern (Text, Format)](#shape-inhalte-verändern-text-format)
   1. [Text mehrerer Shapes ersetzen und löschen](#text-mehrerer-shapes-ersetzen-und-löschen)
   1. Sprache für Rechtschreibprüfung festlegen
   1. Formate mehrerer Shapes angleichen
   1. Shapes skalieren
1. [Spezielle BKT-Shapes](#spezielle-bkt-shapes)
   1. [Ampel-Shape mit Popup zum schnellen Wechseln](#ampel-shape-mit-popup-zum-schnellen-wechseln)
   1. Harvey-Balls einfügen
   1. Agenda einfügen und aktualisieren
   1. Shape-Tabelle anlegen
   1. Aktualisierbare Folien-Thumbnails anlegen
1. [Weitere BKT-Funktionen](#weitere-bkt-funktionen)
   1. [Tastenkombinationen beim Anordnen und Spinner-Boxen](#tastenkombinationen-beim-anordnen-und-spinner-boxen)
   1. [Schnellanleitung der QuickEdit Toolbar](#schnellanleitung-der-quickedit-toolbar)
   1. [Benutzerdefinierte Formate/Style](#benutzerdefinierte-formate)
   1. Ungenutzte Folienlayouts löschen
   1. Eigene Shape-Library anlegen
   1. Chart-Library mit Folienmastern
   1. Shapes gezielt auswählen
   1. Icons mit Icon-Fonts
   1. Shapes teilen oder vervielfachen
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
    <p>
      {{ tipp.description | markdownify }}
      {% if tipp.note %}
        <br><em>{{ tipp.note | markdownify }}</em>
      {% endif %}
    </p>
  </div>
  {% endfor %}
</section>
{% endfor %}
